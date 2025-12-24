import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, RichText
import io

# ===============================
# 1️⃣ 格式規則設定（唯一維護點）
# ===============================
FORMAT_RULES = {
    "me_prefix": {"description": "ME 類：千分位 + 保留原始小數"},
    "decimal_2": {
        "keywords": ["_rate", "elec_price", "new_cop_std", "new_eff_std"],
        "description": "2 位小數"
                 },
    "decimal_1": {"keywords": ["_year"], "description": "1 位小數"},
    "integer": {"description": "整數（預設）"},
}

# ===============================
# 2️⃣ 單一變數處理（會格式化 + 紅字）
# ===============================
def process_value_to_richtext(val, key_name="", debug=False):
    if pd.isna(val):
        return ""

    val_str = str(val).strip()
    if val_str == "":
        return ""

    # 區間值不變紅
    if "~" in val_str or "～" in val_str:
        rt = RichText()
        rt.add(val_str, color="000000", bold=False)
        if debug:
            st.write(f"[DEBUG] {key_name} → 區間值")
        return rt

    # 嘗試轉為數字（排除日期）
    try:
        if "/" in val_str:
            return val_str
        if "-" in val_str and not val_str.startswith("-"):
            return val_str
        float_val = float(val_str)
    except ValueError:
        return val_str

    key_lower = str(key_name).lower()

    # ---------- 規則判斷 ----------
    if key_lower.startswith("me_"):
        rule_desc = FORMAT_RULES["me_prefix"]["description"]
        if "." in val_str:
            int_part, dec_part = val_str.split(".", 1)
            formatted = f"{int(int_part):,}.{dec_part}"
        else:
            formatted = f"{int(float_val):,}"

    elif any(k in key_lower for k in FORMAT_RULES["decimal_2"]["keywords"]):
        rule_desc = FORMAT_RULES["decimal_2"]["description"]
        formatted = f"{float_val:,.2f}"

    elif any(k in key_lower for k in FORMAT_RULES["decimal_1"]["keywords"]):
        rule_desc = FORMAT_RULES["decimal_1"]["description"]
        formatted = f"{float_val:,.1f}"

    else:
        rule_desc = FORMAT_RULES["integer"]["description"]
        formatted = f"{float_val:,.0f}"

    if debug:
        st.write(f"[DEBUG] {key_name} → {rule_desc} → {formatted}")

    rt = RichText()
    rt.add(formatted, color="FF0000", bold=False)
    return rt


# ===============================
# 3️⃣ 表格 sheet 的值：完全不更動
# ===============================
def keep_table_value_raw(val):
    # 表格欄位：不做任何格式化、不做紅字，僅處理空值
    if pd.isna(val):
        return ""
    s = str(val)
    # 防止 dtype=str 後出現 'nan'
    if s.strip().lower() == "nan":
        return ""
    return s


# ===============================
# Streamlit UI（原本顯示頁面）
# ===============================
st.set_page_config(page_title="節能績效計劃書生成器", page_icon="📊")

st.title("📊 HWsmart節能績效計劃書生成器")
st.markdown("""
此工具支援 **Excel 表格同步** 功能：

1. **單一變數**（例如：COP、效率、kWh 等）標示為 **紅字**。
   - 請放在 Excel Sheet 的 **第一個分頁**。  
   - 第 1 欄為「變數名稱」，第 2 欄為「數值」，其餘欄位會被忽略。  
   - 在 Word 中使用：`{{r 變數名稱}}`。

2. **表格資料（例如：改善前冰水機、改善前水泵…）**
   - 每個表格放在獨立的 Sheet  
   - **Sheet 名稱 = Word 中的分頁名稱**  
   - Word 表格內使用（docxtpl row 擴充）

3. **RichText（紅字）**
   - Python 端處理成 RichText
   - Word 模板請使用 `{{r 變數}}` 或 `{{row.欄位}}`
""")


debug_mode = st.checkbox("🧪 Debug 模式（顯示規則判斷）")

col1, col2 = st.columns(2)
with col1:
    uploaded_word = st.file_uploader("1️⃣ 上傳 Word 模板 (.docx)", type="docx")
with col2:
    uploaded_excel = st.file_uploader("2️⃣ 上傳 Excel 數據 (.xlsx)", type="xlsx")

# ===============================
# 主流程
# ===============================
if uploaded_word and uploaded_excel:
    st.divider()

    if st.button("🚀 開始生成報告", type="primary"):
        try:
            uploaded_word.seek(0)
            uploaded_excel.seek(0)

            word_bytes = uploaded_word.read()
            excel_file = pd.ExcelFile(uploaded_excel)

            context = {}
            st.toast("🔍 正在解析 Excel 資料...")

            for idx, sheet_name in enumerate(excel_file.sheet_names):

                # -------- 變數 Sheet（會格式化 + 紅字）--------
                if idx == 0:
                    df_var = excel_file.parse(sheet_name, header=None)
                    for _, row in df_var.iterrows():
                        if pd.isna(row[0]):
                            continue
                        key = str(row[0]).strip()
                        val = row[1]
                        context[key] = process_value_to_richtext(val, key, debug=debug_mode)

                # -------- 表格 Sheet（完全不更動值）--------
                else:
                    # 用 dtype=str 讀，盡量保留原始樣子（不套格式化規則）
                    df = excel_file.parse(sheet_name, dtype=str).fillna("")
                    df.columns = [str(c).strip() for c in df.columns]

                    # 刪除整列皆空（字串）列
                    df = df[df.apply(lambda r: any(str(x).strip() for x in r.values), axis=1)]

                    table = []
                    for _, row in df.iterrows():
                        row_dict = {col: keep_table_value_raw(row[col]) for col in df.columns}
                        table.append(row_dict)

                    context[sheet_name] = table

            # -------- Word Render --------
            doc = DocxTemplate(io.BytesIO(word_bytes))
            doc.render(context)

            output = io.BytesIO()
            doc.save(output)

            st.session_state["generated_doc"] = output.getvalue()
            st.session_state["download_name"] = "Generated_Report.docx"

            st.success("✅ 報告生成成功！請下載檔案。")

        except Exception as e:
            st.error(f"❌ 發生錯誤：{e}")

    if "generated_doc" in st.session_state:
        st.download_button(
            label="📥 下載生成的報告",
            data=st.session_state["generated_doc"],
            file_name=st.session_state["download_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

