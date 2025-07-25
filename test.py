from docx import Document
from docx.shared import Inches
import matplotlib.pyplot as plt
import pandas as pd
import tempfile
import os

def generate_codebook(df, column_types, variable_names, category_definitions, code_df=None, output_path="codebook.docx", preview_mode=False):
    if output_path is None:
        output_path = "codebook.docx"

    if code_df is not None:
        code_df.columns = code_df.columns.str.strip().str.lower()

    doc = Document()
    doc.add_heading("Codebook Summary Report", level=1)

    # ✅ 只統計實際存在欄位的缺失值
    valid_cols = [col for col in column_types.keys() if col in df.columns]
    na_counts = df[valid_cols].isnull().sum()
    na_percent = df[valid_cols].isnull().mean() * 100

    doc.add_heading("Missing Value Summary", level=2)
    na_df = pd.DataFrame({
        "column": na_counts.index,
        "missing_count": na_counts.values,
        "missing_rate (%)": na_percent.round(2).values
    }).query("`missing_count` > 0").reset_index(drop=True)

    if not na_df.empty:
        table = doc.add_table(rows=1 + len(na_df), cols=3)
        table.style = "Table Grid"
        table.cell(0, 0).text = "Column"
        table.cell(0, 1).text = "Missing Count"
        table.cell(0, 2).text = "Missing Rate (%)"
        for i, row in na_df.iterrows():
            table.cell(i + 1, 0).text = str(row["column"])
            table.cell(i + 1, 1).text = str(row["missing_count"])
            table.cell(i + 1, 2).text = str(row["missing_rate (%)"])
    else:
        doc.add_paragraph("No missing values in any columns.")

    # 🔹 變數類型統計區塊
    doc.add_heading("Variable Type Summary", level=2)
    type_count = pd.Series(column_types).value_counts().sort_index()
    type_label_map = {1: "數值型 (Numerical)", 2: "類別型 (Categorical)"}

    table = doc.add_table(rows=1 + len(type_count), cols=2)
    table.style = "Table Grid"
    table.cell(0, 0).text = "變數類型"
    table.cell(0, 1).text = "欄位數"
    for i, (type_code, count) in enumerate(type_count.items()):
        label = type_label_map.get(type_code, f"其他 ({type_code})")
        table.cell(i + 1, 0).text = label
        table.cell(i + 1, 1).text = str(count)

    # 🔹 欄位細節處理
    columns = code_df["variable"] if code_df is not None and "variable" in code_df.columns else df.columns

    for col in columns:
        col = str(col).strip()
        if col not in column_types or col not in df.columns:
            continue

        type_code = column_types[col]
        if type_code == 0:
            continue

        var_name = variable_names.get(col, col)
        doc.add_heading(f"Variable: {col} ({var_name})", level=2)

        # 🟦 類別型
        if type_code == 2:
            value_counts = df[col].value_counts(dropna=False).sort_index()
            total = len(df)
            defs = category_definitions.get(col, {})
            lines = [
                f"{int(k) if isinstance(k, float) and k.is_integer() else k}: {defs.get(k, '')} → {v} ({v/total:.2%})"
                for k, v in value_counts.items()
            ]
            summary_text = "\n".join(lines)

            table = doc.add_table(rows=2, cols=2)
            table.style = "Table Grid"
            table.cell(0, 0).text = "Variable Name"
            table.cell(0, 1).text = f"{col} ({var_name})"
            table.cell(1, 0).text = "Categories Summary"
            table.cell(1, 1).text = summary_text

            fig, ax = plt.subplots()
            value_counts.plot(kind="bar", color="cornflowerblue", ax=ax)
            ax.set_title(f"Count Plot of {col}")
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            plt.tight_layout()
            plt.savefig(tmp.name)
            plt.close()
            doc.add_picture(tmp.name, width=Inches(4.5))
            try: os.unlink(tmp.name)
            except PermissionError: pass

        # 🟩 數值型
        elif type_code == 1:
            try:
                df[col] = pd.to_numeric(df[col], errors="coerce")
            except Exception:
                continue
            if df[col].dropna().empty:
                continue

            desc = df[col].describe()
            table = doc.add_table(rows=4, cols=4)
            table.style = "Table Grid"
            table.cell(0, 0).text = "Index"
            table.cell(0, 1).text = var_name
            table.cell(0, 2).text = "Variable Name"
            table.cell(0, 3).text = col

            table.cell(1, 0).text = "Mean"
            table.cell(1, 1).text = f"{desc['mean']:.3f}"
            table.cell(1, 2).text = "Std Dev"
            table.cell(1, 3).text = f"{desc['std']:.3f}"

            table.cell(2, 0).text = "Max"
            table.cell(2, 1).text = f"{desc['max']:.3f}"
            table.cell(2, 2).text = "Min"
            table.cell(2, 3).text = f"{desc['min']:.3f}"

            table.cell(3, 0).text = "Q1 (25%)"
            table.cell(3, 1).text = f"{desc['25%']:.3f}"
            table.cell(3, 2).text = "Q2 (50%)"
            table.cell(3, 3).text = f"{desc['50%']:.3f}"

            table.cell(4, 0).text = "Q3 (75%)"
            table.cell(4, 1).text = f"{desc['75%']:.3f}"
            table.cell(4, 2).text = "Range)"
            table.cell(4, 3).text = f"{desc['max'] - desc['min']:.3f}"

            fig, ax = plt.subplots()
            df[col].plot(kind="hist", bins=10, color="skyblue", edgecolor="black", ax=ax)
            ax.set_title(f"Histogram of {col}")
            tmp1 = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            plt.tight_layout()
            plt.savefig(tmp1.name)
            plt.close()
            doc.add_picture(tmp1.name, width=Inches(4.5))
            try: os.unlink(tmp1.name)
            except PermissionError: pass

            fig2, ax2 = plt.subplots()
            box = df.boxplot(column=col, ax=ax2, grid=False, patch_artist=True, boxprops=dict(facecolor='lightblue'))

            # ➤ 計算五數摘要
            q1 = desc['25%']
            q2 = desc['50%']
            q3 = desc['75%']
            minimum = desc['min']
            maximum = desc['max']

            # ➤ 加上標註
            x = 1  # boxplot 當作只有一組資料，x軸位置為1
            ax2.text(x + 0.05, minimum, f"Min: {minimum:.2f}", va='center', fontsize=8)
            ax2.text(x + 0.05, q1, f"Q1: {q1:.2f}", va='center', fontsize=8)
            ax2.text(x + 0.05, q2, f"Median: {q2:.2f}", va='center', fontsize=8)
            ax2.text(x + 0.05, q3, f"Q3: {q3:.2f}", va='center', fontsize=8)
            ax2.text(x + 0.05, maximum, f"Max: {maximum:.2f}", va='center', fontsize=8)

            ax2.set_title(f"Boxplot of {col}")
            tmp2 = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            plt.tight_layout()
            plt.savefig(tmp2.name)
            plt.close()
            doc.add_picture(tmp2.name, width=Inches(4.5))
            try: os.unlink(tmp2.name)
            except PermissionError: pass

    doc.save(output_path)
    return output_path
