import streamlit as st
import pandas as pd
import base64
import os
import io

from test import generate_codebook  # 確保 test.py 有放對位置並含有該函式
st.set_page_config(page_title="Codebook 產生器", layout="wide")
# ✅ 🚨 請確保這段放在所有 tab1/tab2 之前！
def read_uploaded_csv(uploaded_file):
    for enc in ["utf-8", "utf-8-sig", "cp950", "big5"]:
        try:
            return pd.read_csv(io.TextIOWrapper(uploaded_file, encoding=enc))
        except Exception:
            uploaded_file.seek(0)
            continue
    st.error("❌ 檔案無法讀取，請確認是否為有效的 CSV 並使用常見編碼（UTF-8、BIG5、CP950）")
    return None
tab1, tab2 = st.tabs(["📄 Codebook 產生器","📊 進階分析工具", ])


with tab1:
    st.title("📄 自動化 Codebook 產生工具")

    st.markdown("""### 📘 功能說明
    本工具可根據 `code.csv` 中的 Type 欄位，對主資料進行以下轉換：

    - `0`：略過
    - `1`：數值型
    - `2`：類別型
    """)

    import streamlit as st
import pandas as pd

# 📁 第一步：上傳主資料
st.header("📁 資料上傳")
data_file = st.file_uploader("請上傳主資料 CSV", type=["csv"], key="data")

df = None
code_df = None

if data_file:
    df = pd.read_csv(data_file)
    df = df.dropna(how="all")
    df.columns = df.columns.str.strip()  # 去除主資料欄位空白
    st.success("✅ 主資料上傳成功！")
    st.dataframe(df.head())

    st.markdown("---")
    st.info("📌 若需產出 Codebook，請繼續上傳 code.csv")

    # 📄 第二步：上傳 code.csv
    code_file = st.file_uploader("📄 請上傳 Codebook 設定檔（code.csv）", type=["csv"], key="code")

    if code_file:
        code_df = pd.read_csv(code_file)
        code_df = code_df.dropna(how="all")
        code_df.columns = code_df.columns.str.strip().str.lower()

        if "variable" not in code_df.columns or "type" not in code_df.columns:
            st.error("❌ code.csv 檔案中需包含 'Variable' 與 'Type' 欄位")
        else:
            # ➤ 抓取交集變數
            code_vars = code_df["variable"].astype(str).str.strip().tolist()
            df_vars = df.columns.tolist()
            common_vars = list(set(code_vars) & set(df_vars))

            st.info(f"✅ 同時存在於主資料與 Codebook 的變數數量：{len(common_vars)}")

            # ➤ 過濾 code_df 只保留交集變數
            code_df = code_df[code_df["variable"].astype(str).str.strip().isin(common_vars)].reset_index(drop=True)

            # 🧩 處理變數屬性
            column_types = {}
            variable_names = {}
            column_roles = {}
            x_counter = y_counter = 1

            for _, row in code_df.iterrows():
                col = str(row["variable"]).strip()
                t = str(row.get("type", "")).strip().lower()
                target = str(row.get("target", "") if "target" in row else "").strip().lower()

                if col not in df.columns:
                    continue  # 雙保險防呆

                if target:
                    column_roles[col] = f"Y{y_counter}"
                    variable_names[col] = f"Y{y_counter}"
                    y_counter += 1
                    continue

                if t == "numerical":
                    column_roles[col] = f"X{x_counter}"
                    column_types[col] = 1
                    x_counter += 1
                elif t == "categorical":
                    column_roles[col] = f"X{x_counter}"
                    column_types[col] = 2
                    x_counter += 1
                else:
                    st.warning(f"⚠️ Unknown Type '{t}' for column '{col}' — skipped.")
                    continue

                variable_names[col] = column_roles.get(col, col)

            # 📊 顯示變數類型統計
            st.subheader("📊 變數類型統計")
            type_count = pd.Series(column_types).value_counts().sort_index()
            type_label_map = {1: "數值型 (Numerical)", 2: "類別型 (Categorical)"}
            type_summary = pd.DataFrame({
                "變數類型": [type_label_map.get(t, f"其他 ({t})") for t in type_count.index],
                "欄位數": type_count.values
            })
            st.dataframe(type_summary)

            # 📤 產出報告按鈕
            st.markdown("---")
            st.subheader("📤 Codebook 報告產出")
            if st.button("🚀 產出 Codebook 報告"):
                with st.spinner("📄 報告產出中，請稍候..."):
                    try:
                        output_path = "codebook.docx"

                        # 🧠 假設你已經有這個函數
                        output_path = generate_codebook(
                            df, column_types, variable_names, {},
                            code_df=code_df, output_path=output_path
                        )

                        with open(output_path, "rb") as f:
                            b64 = base64.b64encode(f.read()).decode()
                            href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{output_path}">📥 點我下載 Codebook 報告</a>'
                            st.markdown(href, unsafe_allow_html=True)

                        st.success("✅ 報告產出完成！")
                    except Exception as e:
                        st.error(f"❌ 報告產出失敗：{e}")




# ---------- Tab 2 ----------
with tab2:
    st.title("📊 進階分析工具")

    st.markdown("""
    ### 📘 功能說明
    本工具可根據 `code.csv` 中的 Transform 欄位，對主資料進行以下轉換：

    - 若 Transform 欄為空或 'none'，則不進行任何轉換。
    - 若為 `cut:[0,100,200,300]`，將以手動分箱方式進行區間分類（含邊界），自動轉為 0, 1, 2...。
    - 如: `cut:[0,100,200,300]` 會將數值分為 0-99, 100-199, 200-300 三個區間。
    - 若 Transform 欄為 `onehot`，則會進行 one-hot encoding，並轉為 0/1。
    

    所有轉換後的欄位名稱將自動加上 `_binned` 或對應欄位前綴，原始欄位將被移除。
    轉換後將依據原始 code.csv 更新 Type 資訊並自動產出 Codebook。
    """)

    def read_uploaded_csv(uploaded_file):
        for enc in ["utf-8", "utf-8-sig", "cp950", "big5"]:
            try:
                return pd.read_csv(io.TextIOWrapper(uploaded_file, encoding=enc))
            except Exception:
                uploaded_file.seek(0)
                continue
        st.error("❌ 檔案無法讀取，請確認是否為有效的 CSV 並使用常見編碼（UTF-8、BIG5、CP950）")
        return None

    uploaded_main = st.file_uploader("📂 請上傳主資料（CSV）", type=["csv"], key="main2")
    uploaded_code = st.file_uploader("📋 請上傳 code.csv（需包含 Column、 Type、 Transform欄位）", type=["csv"], key="code2")

    df2, code2 = None, None

    if uploaded_main:
        df2 = read_uploaded_csv(uploaded_main)
    if uploaded_code:
        code2 = read_uploaded_csv(uploaded_code)

    if df2 is not None and code2 is not None:
        st.success("✅ 資料與 code.csv 載入成功")
        st.success(f"✅主資料共 {df2.shape[0]} 筆")
        code2 = code2[~code2["Type"].astype(str).str.lower().eq("0")]

        variable_names = {}
        column_types = {}
        category_definitions = {}

        for _, row in code2.iterrows():
            col = row["Column"]
            t = str(row["Type"]).lower()
            transform = str(row.get("Transform", "")).strip()

            if col not in df2.columns:
                continue

            if transform == '' or transform.lower() == 'none':
                column_types[col] = int(t[-1]) if t.startswith("y") else int(t)
                variable_names[col] = col
                continue

            if transform.lower().startswith("cut:["):
                try:
                    bins = eval(transform[4:])
                    new_col = col + "_binned"
                    df2[new_col] = pd.cut(df2[col], bins=bins, include_lowest=True, labels=False)
                    column_types[new_col] = 2
                    variable_names[new_col] = col
                    df2.drop(columns=[col], inplace=True)
                except Exception as e:
                    st.warning(f"🔸 {col} 分箱失敗：{e}")
            
            elif df2[col].dtype == 'object' or transform.lower() == 'onehot':
                try:
                    onehot = pd.get_dummies(df2[col], prefix=col, dtype=int)
                    for new_col in onehot.columns:
                        column_types[new_col] = 2
                        variable_names[new_col] = new_col
                    df2 = pd.concat([df2.drop(columns=[col]), onehot], axis=1)
                except Exception as e:
                    st.warning(f"🔸 {col} one-hot 編碼失敗：{e}")
            else:
                st.warning(f"🔸 未知 Transform 指令：{transform}（欄位 {col}）")

        st.markdown("---")
        st.subheader("🔍 預覽轉換後資料")
        st.dataframe(df2.head())
                # 🔎 遺失值統計
        st.subheader("📉 遺失值統計")
        na_counts = df2.isnull().sum()
        na_percent = df2.isnull().mean() * 100
        na_df = pd.DataFrame({
            "欄位名稱": na_counts.index,
            "遺失數": na_counts.values,
            "遺失比例 (%)": na_percent.round(2).values
        })
        na_df = na_df[na_df["遺失數"] > 0]
        if na_df.empty:
            st.info("✅ 無遺失值")
        else:
            st.warning("⚠️ 以下欄位有遺失值：")
            st.dataframe(na_df)
            rows_after_na = df2.dropna().shape[0]
            st.write(f"📦 刪除所有含遺失值的資料後，剩餘筆數為 {rows_after_na} 筆資料")
        # 🔎 變數類型統計
        st.markdown("---")
        st.subheader("📊 變數類型統計")
        type_count = pd.Series(column_types).value_counts().sort_index()
        type_label_map = {
            1: "數值型 (Type 1)",
            2: "類別型 (Type 2)"
        }
        type_summary = pd.DataFrame({
            "變數類型": [type_label_map.get(t, f"其他 ({t})") for t in type_count.index],
            "欄位數": type_count.values
        })
        st.dataframe(type_summary)
        csv = df2.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 下載轉換後的 CSV", data=csv, file_name="transformed_data.csv", mime="text/csv")
        # ✅ 製作轉換後的 code_df
        code_df_transformed = pd.DataFrame({
            "Column": list(variable_names.keys()),
            "Type": [column_types[col] for col in variable_names.keys()]
        })

        # ✅ 輸出轉換後的 code.csv
        code_df_transformed.to_csv("code_transformed.csv", index=False, encoding="utf-8-sig")

        # ✅ 加入下載按鈕
        csv_code = code_df_transformed.to_csv(index=False, encoding='utf-8-sig')
        st.download_button("📥 下載轉換後的 code.csv", data=csv_code, file_name="code_transformed.csv", mime="text/csv")
        st.markdown("---")
        st.subheader("📘 自動產出 Codebook")
        if st.button("📄 產生報告（轉換後資料）"):
            with st.spinner("產出中..."):
                try:
                    output_path = "transformed_codebook.docx"
                    code_df_transformed = pd.DataFrame({
                    "Column": list(column_types.keys()),
                    "Type": [column_types[k] for k in column_types],
                })
                    generate_codebook(df2, column_types, variable_names, category_definitions, code_df=code_df_transformed, output_path=output_path)
                    
                    with open(output_path, "rb") as f:
                        b64 = base64.b64encode(f.read()).decode()
                        href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{output_path}">📥 下載 Codebook 報告（轉換後）</a>'
                        st.markdown(href, unsafe_allow_html=True)
                    st.success("✅ Codebook 已產出！")
                except Exception as e:
                    st.error(f"❌ 報告產出失敗：{e}")




