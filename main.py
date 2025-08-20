
import pandas as pd
import os

def generate_report(source_file):
    df = pd.read_excel(source_file, engine="openpyxl")
    sales_col = df.iloc[:, 23]  # X欄
    category_col = df.iloc[:, 34]  # AI欄

    data = pd.DataFrame({"業務姓名": sales_col, "工單類型": category_col})
    data = data.dropna()
    data = data[data["業務姓名"] != "開單者中文姓名"]

    total_counts = data["業務姓名"].value_counts().reset_index()
    total_counts.columns = ["業務姓名", "總筆數"]

    keywords = ["銷售相關問題", "(加盟)辦A給B/盜辦/特殊身份辦理爭議", "辦門號換現金"]
    filtered = data[data["工單類型"].astype(str).apply(lambda x: any(k in x for k in keywords))]
    sales_counts = filtered["業務姓名"].value_counts().reset_index()
    sales_counts.columns = ["業務姓名", "銷售類工單筆數"]

    merged = pd.merge(total_counts, sales_counts, on="業務姓名", how="left")
    merged["銷售類工單筆數"] = merged["銷售類工單筆數"].fillna(0).astype(int)

    merged.to_excel("業務工單筆數比較.xlsx", index=False)
    print("報表已產出：業務工單筆數比較.xlsx")

if __name__ == "__main__":
    for file in os.listdir():
        if file.endswith(".xlsx") and "Export" in file:
            generate_report(file)
            break
