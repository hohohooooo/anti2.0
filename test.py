from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import glob
import uvicorn

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 初始化時讀取所有 Excel 檔案
def load_excel_data(path_to_excel_folder):

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

# 全局變數：在應用啟動時載入所有公司資料
path = "./DATA"
company_data_content, company_data_score = load_excel_data(path)

def df2json(df_content, df_score):
    # content
    content_list = df_content['相關內容'].to_list()
    yesno_list = df_content['是否符合'].to_list()
    converted_list = list(map(lambda x: "1" if x == "是" else "0", yesno_list))

    # score
    company_info = df_score.iloc[0].to_dict()

    # 構建 "評分" 字典，移除 "公司名稱" 並保留其他欄位
    score_info = {key: company_info[key] for key in company_info if key != '公司名稱'}

    output = {
        "公司名稱": company_info['公司名稱'],
        "評分": score_info,
        "是否符合驗測項目":converted_list,
        "相關內容": content_list
    }
    # 輸出結果
    return output


@app.post("/company/{name}")

def get_company_data(name: str):
    # 根據公司名稱查詢資料
    result_content = company_data_content[company_data_content['公司名稱'] == name]
    result_score = company_data_score[company_data_score['公司名稱'] == name]   
    if result_content.empty:
        raise HTTPException(status_code=404, detail="Company not found")
    # df2json
    output = df2json(result_content, result_score)
    return output


if __name__ == "__main__":
    uvicorn.run("test:app", host="0.0.0.0", port=9453, reload=True)
