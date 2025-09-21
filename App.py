import streamlit as st
import pandas as pd
import base64
import os
import io
from fast import generate_codebook_fast  # 確保 fast.py 有放對位置並含有該函式
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
tab1, tab2 = st.tabs(["📄 Codebook 產生器","📊 進階分析工具(尚在處理)", ])


with tab1:
    st.title("📄 自動化 Codebook 產生工具")

    with st.expander("📘 查看工具使用說明與檔案規格", expanded=False):
        st.markdown("""
### 功能說明
本工具將協助你依據主資料與對應的 `code.csv` 設定檔，自動產出變數摘要報告（Codebook）。

---

## 📂 上傳檔案規格

### 1️⃣ 主資料檔
- **格式**：CSV（UTF-8 編碼建議）
- **內容**：
    - 每一欄為一個變數（欄位名稱將與 Codebook 設定檔的 `Variable` 欄比對）。
    - 可包含數值型與類別型資料。
    - 欄位名稱可含中英文，但會自動去除多餘空白與 `Unnamed` 欄位。
    - 系統會自動移除整列全為空值的資料。

---

### 2️⃣ Codebook 設定檔
- **格式**：CSV（UTF-8 編碼建議）
- **必備欄位**：
    1. `Variable`（必填）：對應主資料的欄位名稱
    2. `Type`（必填）：
        - `0`、空白、`none`、`skip` → 略過該變數
        - `1`、`numerical`、`連續`、`數值` → 數值型
        - `2`、`categorical`、`類別` → 類別型
    3. `Description`（選填）：變數說明
    4. `Target`（選填）：`y`、`yes`、`true`、`1`、`target` → 目標變數（Y）

---

## 🔄 系統自動處理內容
1. 清理資料與欄位名稱
2. 僅分析主資料與 Codebook 同時存在的欄位
3. 自動解析 `Type` 與 `Target`
4. 顯示數值型、類別型、略過變數的數量統計
5. 列出目標變數清單

---

## 📄 報告內容
1. **缺失值統計**
2. **變數類型統計**（含目標變數清單）
3. **變數詳細資訊（依 Codebook 順序）**：
    - **數值型**：統計值、Boxplot、Histogram
    - **類別型**：類別分布、百分比、Count Plot

---
📥 完成設定後，點擊「🚀 產出 Codebook 報告」按鈕，即可下載 Word 格式報告。
    """)

    import streamlit as st
    import pandas as pd

    # 📁 第一步：上傳主資料
    st.header("📁 資料上傳")
    data_file = st.file_uploader("請上傳主資料 CSV", type=["csv"], key="data")

    df = None
    code_df = None

    if data_file:
        df = read_uploaded_csv(data_file)
        if df is not None:
            df = df.dropna(how="all")
            df.columns = df.columns.str.strip()
            df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
            st.success("✅ 主資料上傳成功！")
            st.dataframe(df.head())

        st.markdown("---")
        st.info("📌 若需產出 Codebook，請繼續上傳 code.csv")


    # 📄 第二步：上傳 code.csv
    code_file = st.file_uploader("📄 請上傳 Codebook 設定檔（code.csv）", type=["csv"], key="code")

    if code_file:
        code_df = read_uploaded_csv(code_file)
        if code_df is not None:
            code_df = code_df.dropna(how="all")
            # 去除 Variable 欄為空白或僅含空格的列
            
            code_df.columns = code_df.columns.str.strip().str.lower()
        if "target" not in code_df.columns:
            st.info("🔍 未偵測到 `Target` 欄位，預設所有變數皆為自變數（X）")
        
        if "variable" not in code_df.columns or "type" not in code_df.columns:
            st.error("❌ code.csv 檔案中需包含 'Variable' 與 'Type' 欄位")
        else:
            # ➤ 抓取交集變數
            code_vars = code_df["variable"].astype(str).str.strip().tolist()
            df_vars = df.columns.tolist()
            common_vars = list(set(code_vars) & set(df_vars))

            # ➤ 顯示落選的變數（只在 code.csv 裡但主資料中找不到）
            excluded_vars = sorted(set(code_vars) - set(df_vars))
            if excluded_vars:
                st.warning(f"⚠️ 有 {len(excluded_vars)} 個變數未在主資料中找到，已被略過：")
                st.code(", ".join(excluded_vars), language="text")

            excluded_code_vars = sorted(set(df_vars) - set(code_vars))
            if excluded_code_vars:
                st.warning(f"⚠️ 有 {len(excluded_code_vars)} 個變數未在 code 中找到，已被略過：")
                st.code(", ".join(excluded_code_vars), language="text")

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

                if target in ["y", "yes", "target", "1"]:
                    column_roles[col] = f"Y{y_counter}"
                    variable_names[col] = f"Y{y_counter}"
                    # 根據 type 欄位設定 column_types
                    if t in ["2", "categorical", "類別"]:
                        column_types[col] = 2
                    else:
                        column_types[col] = 1  # 預設為數值型
                    
                    y_counter += 1
                    continue
                    
                if t in ["", "0", "none"]:
                    continue  # 自動略過

                if t in ["1", "numerical","連續"]:
                    column_roles[col] = f"X{x_counter}"
                    column_types[col] = 1
                    x_counter += 1
                elif t in ["2", "categorical","類別"]:
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
            elif st.button("🚀 快速產出 Codebook 報告 (Fast Mode)"):
                with st.spinner("📄 快速報告產出中，請稍候..."):
                    try:
                        output_path = "codebook_fast.docx"

                        # 使用快速版函數
                        output_path = generate_codebook_fast(
                            df, column_types, variable_names, {},
                            code_df=code_df, output_path=output_path
                        )

                        with open(output_path, "rb") as f:
                            b64 = base64.b64encode(f.read()).decode()
                            href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{output_path}">📥 點我下載快速版 Codebook 報告</a>'
                            st.markdown(href, unsafe_allow_html=True)

                        st.success("✅ 快速版報告產出完成！")
                    except Exception as e:
                        st.error(f"❌ 快速版報告產出失敗：{e}")




# ---------- Tab 2 ----------
with tab2:
    st.title("📊 進階分析工具")

    st.markdown("""
    ### 📘 功能說明
    本工具可根據 `code.csv` 中的 **Transform 欄位**，對主資料進行以下轉換：

    - 若 Transform 欄為空或 'none'，則不進行任何轉換。
    - 若為 `cut:[0,100,200,300]` → 依指定切點分箱。
    - 若為 `cut:3` → 依分位數切成 3 組。
    - 若為 `24,30` → 依數字切成 `<24`, `24–30`, `>30` 三組。
    - 若為 `onehot` → 進行 one-hot 編碼。

    轉換後：
    - 原始欄位將被移除
    - 新欄位將自動加上 `_binned` 或 onehot 前綴
    - 可下載轉換後的 CSV 與新的 code.csv
    """)

    # === 檔案上傳 ===
    uploaded_main = st.file_uploader("📂 請上傳主資料（CSV）", type=["csv"], key="main2")
    uploaded_code = st.file_uploader("📋 請上傳 code.csv（需包含 Variable、Transform 欄位）", type=["csv"], key="code2")

    df2, code2 = None, None
    if uploaded_main:
        df2 = read_uploaded_csv(uploaded_main)
    if uploaded_code:
        code2 = read_uploaded_csv(uploaded_code)

    if df2 is not None and code2 is not None:
        st.success(f"✅ 主資料與 code.csv 載入成功，共 {df2.shape[0]} 筆資料")

        # 標準化欄位名稱
        code2.columns = code2.columns.str.strip().str.lower()

        variable_names = {}  # 映射新舊欄位名稱
        transformed_vars = []  # 儲存轉換後的變數資訊
        for _, row in code2.iterrows():
            col = str(row.get("variable", "")).strip()
            transform = str(row.get("transform", "")).strip()

            if not col or col not in df2.columns:
                continue

            # === case 1: 無 Transform → 保留原始欄位 ===
            if transform.lower() in ["", "nan", "none"]:
                orig_type = str(row.get("type", "1")).strip().lower()

                # 做 mapping，避免直接轉 int 出錯
                type_map = {
                    "1": 'Numerical', "numerical": 'Numerical', "連續": 1, "數值": 'Numerical',
                    "2": 'Categorical', "categorical": 'Categorical', "類別": 'Categorical'
                }

                # 如果沒辦法辨識，就預設 1
                t_val = type_map.get(orig_type, 1)

                transformed_vars.append({"Variable": col, "Type": t_val,"Description": row.get("description"), "Transform": row.get("transform")})
                continue


            # === case 2: cut:[…] → 手動分箱 ===
            if transform.lower().startswith("cut:["):
                try:
                    bins = eval(transform[4:])
                    new_col = col + "_binned"
                    df2[new_col] = pd.cut(df2[col], bins=bins, include_lowest=True, labels=False)
                    variable_names[new_col] = col
                    df2.drop(columns=[col], inplace=True)
                except Exception as e:
                    st.warning(f"🔸 {col} 分箱失敗：{e}")
                continue

            # === case 3: cut:k → 分位數切分 ===
            if transform.lower().startswith("cut:"):
                try:
                    k = int(transform.split(":")[1])
                    new_col = col + "_binned"
                    df2[new_col] = pd.qcut(df2[col], q=k, labels=False, duplicates="drop")
                    variable_names[new_col] = col
                    df2.drop(columns=[col], inplace=True)
                except Exception as e:
                    st.warning(f"🔸 {col} 分位數切分失敗：{e}")
                continue

            # === case 4: 逗號分隔數字 → 自訂切分點 ===
            if "," in transform:
                try:
                    cuts = [float(x.strip()) for x in transform.split(",") if x.strip()]
                    bins = [-float("inf")] + cuts + [float("inf")]
                    new_col = col + "_binned"
                    df2[new_col] = pd.cut(df2[col], bins=bins, labels=False)
                    variable_names[new_col] = col
                    df2.drop(columns=[col], inplace=True)
                except Exception as e:
                    st.warning(f"🔸 {col} 自訂切分失敗：{e}")
                continue

            # === case 5: onehot ===
            if transform.lower() == "onehot" or df2[col].dtype == "object":
                try:
                    onehot = pd.get_dummies(df2[col], prefix=col, dtype=int)
                    for new_col in onehot.columns:
                        variable_names[new_col] = col
                    df2 = pd.concat([df2.drop(columns=[col]), onehot], axis=1)
                except Exception as e:
                    st.warning(f"🔸 {col} one-hot 編碼失敗：{e}")
                continue
            # === case: 單一數字 → 切兩類 (<cut_point → 0, >=cut_point → 1) ===
            if transform.replace(".", "", 1).isdigit():  # 判斷是不是數字 (含小數)
                try:
                    cut_point = float(transform)
                    new_col = col + "_binned"

                    # 分箱：左閉右開，確保 <cut_point = 0, >=cut_point = 1
                    df2[new_col] = (df2[col] >= cut_point).astype(int)

                    variable_names[new_col] = col
                    df2.drop(columns=[col], inplace=True)
                except Exception as e:
                    st.warning(f"🔸 {col} 單一數字分界失敗：{e}")
                continue

            # === case 6: 未知指令 ===
            st.warning(f"🔸 未知 Transform 指令：{transform}（欄位 {col}）")

        # === 預覽結果 ===
        st.markdown("---")
        st.subheader("🔍 預覽轉換後資料")
        st.dataframe(df2.head())

        # === 提供下載轉換後 CSV ===
        csv = df2.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 下載轉換後的資料 (CSV)", data=csv, file_name="transformed_data.csv", mime="text/csv")
        #transformed_vars = []

        # 轉成 DataFrame
        code_df_transformed = pd.DataFrame(transformed_vars)

        # === 產生轉換後的 code.xlsx（只包含 Variable）===
        import io

        code_df_transformed = pd.DataFrame(transformed_vars)

        # 存成 Excel
        output = io.BytesIO()
        code_df_transformed.to_excel(output, index=False)
        st.download_button(
            "📥 下載轉換後的 code.xlsx",
            data=output.getvalue(),
            file_name="code_transformed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)






