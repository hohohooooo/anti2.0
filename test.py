from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import glob
import uvicorn
import json
import time
from openai import OpenAI
from datetime import datetime


# 初始化時讀取所有 Excel 檔案
def load_excel_data(path_to_excel_folder):
    # 全局變數：在應用啟動時載入所有公司資料

    files = glob.glob(f"{path_to_excel_folder}/*.xlsx")
    data_content = []
    data_score = []
    for file in files:
        df_content = pd.read_excel(file, sheet_name='表1')
        df_score = pd.read_excel(file, sheet_name='表2')
        data_content.append(df_content)
        data_score.append(df_score)


    # 將所有資料合併到一個 DataFrame
    all_data_content = pd.concat(data_content, ignore_index=True)
    all_data_score = pd.concat(data_score, ignore_index=True)
    return all_data_content, all_data_score


def df2json(df_content, df_score):
    # content
    content_list = df_content['相關內容'].to_list()
    yesno_list = df_content['是否符合'].to_list()
    converted_list = list(map(lambda x: "提及 (點擊查看細節)" if x == "是" else "未提及 (點擊查看細節)", yesno_list))

    # score
    company_info = df_score.iloc[0].to_dict()

    # 構建 "評分" 字典，移除 "公司名稱" 並保留其他欄位
    score_info = {key: company_info[key] for key in company_info if key != '公司名稱' and key != '摘要'}

    output = {
        "公司名稱": company_info['公司名稱'],
        "評分": score_info,
        "是否符合驗測項目":converted_list,
        "相關內容": content_list,
        "摘要":company_info['摘要']
    }
    # 輸出結果
    return output

def assistant(messages):
    ## with plugin
    ASSISTANT_API = 'https://prod.dvcbot.net/api/assts/v1'
    API_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJEVkNBU1NJIiwic3ViIjoiVVNFUjAxQFNLRkhDT1JQLkNPTSIsImF1ZCI6WyJEVkNBU1NJIl0sImlhdCI6MTcyMTM3OTQ0NywianRpIjoiOTk4NTZiY2EtYjFkMC00OWJkLWFhMDctYjY0ZDBmNGE3NzJhIn0.ATijnDev3sOYrAPpNDZO4_r18kWiHOq39znqqeDrrO0'
    ASSISTANT_ID = 'asst_dvc_wNbWLbW9BjbGhJIGH8QGM48X'

    client = OpenAI(
        base_url=ASSISTANT_API,
        api_key=API_KEY,
    )
    # 建立 thread
    thread = client.beta.threads.create(messages=[])

    # 連續發送訊息
    for message in messages:
        client.beta.threads.messages.create(thread_id=thread.id, role='user', content=[message])

    # 執行 assistant
    run = client.beta.threads.runs.create_and_poll(thread_id=thread.id, assistant_id=ASSISTANT_ID, additional_instructions=f"\nThe current time is: {datetime.now()}")

    while run.status == 'requires_action' and run.required_action:
        outputs = []
        for call in run.required_action.submit_tool_outputs.tool_calls:
            resp = client._client.post(ASSISTANT_API + '/pluginapi', params={"tid": thread.id, "aid": ASSISTANT_ID, "pid": call.function.name}, headers={"Authorization": "Bearer " + API_KEY}, json=json.loads(call.function.arguments))
            outputs.append({"tool_call_id": call.id, "output": resp.text[:8000]})
        run = client.beta.threads.runs.submit_tool_outputs_and_poll(run_id=run.id, thread_id=thread.id, tool_outputs=outputs)

    if run.status == 'failed' and run.last_error:
        print(run.last_error.model_dump_json())

    msgs = client.beta.threads.messages.list(thread_id=thread.id, order='desc')
    client.beta.threads.delete(thread_id=thread.id)
    return msgs.data[0].content[0].text.value


app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.post("/check/{name}")

def find_company(search_query):
    # company dictionary
    company_dict = {
        "新光": ['新光','新光金', '新光金控','2888'],
        "中信": ['中信','中信金','中信金控','2891'],
        "元大": ['元大','元大金','元大金控','2885'],
        "台新": ['台新','台新金','台新金控','2887'],
        "永豐": ['永豐','永豐金','永豐金控','2890'],
        "玉山": ['玉山','玉山金','玉山金控','2884'],
        "兆豐": ['兆豐','兆豐金','兆豐金控','2886'],
        "合庫": ['合庫','合庫金','合庫金控','5880'],
        "國泰": ['國泰','國泰金','國泰金控','2882'],
        "國票": ['國票','國票金','國票金控','2889'],
        "第一": ['第一','第一金','第一金控','2892'],
        "富邦": ['富邦','富邦金','富邦金控','2881'],
        "華南": ['華南','華南金','華南金控','2880'],
        "開發": ['開發','開發金','開發金控','2883'],
        "久陽精密":["久陽精密","久陽","5011"],
        "台鋼":["台鋼","台灣鋼聯","鋼聯",'6581'],
        "台灣苯乙烯":["台灣苯乙烯","台苯","1310"],
        "名軒開發":["名軒開發","名軒","1442"],
        "至上":["至上","至上電子","8112"],
        "東和鋼鐵":["東和鋼鐵","2006"],
        "建新":["建新","建新國際","8367"],
        "泰鼎":["泰鼎","泰鼎國際","泰鼎-KY","4927"],
        "訊芯":['訊芯',"訊芯科技","訊芯-KY","6451"],
        "雲豹能源":['雲豹能源',"雲豹能源創","雲豹能源-創","6869"],
        "榮剛":['榮剛',"榮剛材料科技","5009"]

    }

    # for main_company, aliases in company_dict.items():
    #     if any(alias in search_query for alias in aliases):
    #         return "1"
    # return "0"
    for main_company, aliases in company_dict.items():
        if search_query in aliases:
            return "1"
    return "0"

@app.post("/company/{name}")

def get_company_data(name: str):
    
    # path = "/DEMO/DATA"
    path = "./DATA"
    company_data_content, company_data_score = load_excel_data(path)
    
    # check
    company_dict = {
        "新光": ['新光','新光金', '新光金控','2888'],
        "中信": ['中信','中信金','中信金控','2891'],
        "元大": ['元大','元大金','元大金控','2885'],
        "台新": ['台新','台新金','台新金控','2887'],
        "永豐": ['永豐','永豐金','永豐金控','2890'],
        "玉山": ['玉山','玉山金','玉山金控','2884'],
        "兆豐": ['兆豐','兆豐金','兆豐金控','2886'],
        "合庫": ['合庫','合庫金','合庫金控','5880'],
        "國泰": ['國泰','國泰金','國泰金控','2882'],
        "國票": ['國票','國票金','國票金控','2889'],
        "第一": ['第一','第一金','第一金控','2892'],
        "富邦": ['富邦','富邦金','富邦金控','2881'],
        "華南": ['華南','華南金','華南金控','2880'],
        "開發": ['開發','開發金','開發金控','2883'],
        "久陽精密":["久陽精密","久陽","5011"],
        "台鋼":["台鋼","台灣鋼聯","鋼聯",'6581'],
        "台灣苯乙烯":["台灣苯乙烯","台苯","1310"],
        "名軒開發":["名軒開發","名軒","1442"],
        "至上":["至上","至上電子","8112"],
        "東和鋼鐵":["東和鋼鐵","2006"],
        "建新":["建新","建新國際","8367"],
        "泰鼎":["泰鼎","泰鼎國際","泰鼎-KY","4927"],
        "訊芯":['訊芯',"訊芯科技","訊芯-KY","6451"],
        "雲豹能源":['雲豹能源',"雲豹能源創","雲豹能源-創","6869"],
        "榮剛":['榮剛',"榮剛材料科技","5009"]

    }
    for main_company, aliases in company_dict.items():
        if name in aliases:
            name2 = main_company 


    # 根據公司名稱查詢資料
    result_content = company_data_content[company_data_content['公司名稱'] == name2]
    result_score = company_data_score[company_data_score['公司名稱'] == name2]   
    if result_content.empty:
        raise HTTPException(status_code=404, detail="Company not found")
    # df2json
    output = df2json(result_content, result_score)
    return output

@app.post("/internet/{message}")

def internet_search(message: str):
    ##

    messages = [
        {"type": "text", "text": f"{message}"}
    ]
    output = assistant(messages)
    output = output.replace('```json', '').replace('```', '')
    data = json.loads(output)
    return [data['環境'], data['社會'], data['公司'], data['摘要']]


if __name__ == "__main__":
    uvicorn.run("test:app", host="0.0.0.0", port=9453, reload=True)
