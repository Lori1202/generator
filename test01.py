import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, RichText
import io

# ---------------- 輔助函式：處理紅字邏輯 ----------------
def process_value_to_richtext(val):
    """
    判斷數值是否需要變紅：
    - 空值 / NaN：回傳空字串
    - 純數字（不含 - /，避免把日期當數字）：回傳 RichText 紅字粗體
    - 其他：回傳字串
    """
    if pd.isna(val):
        return ""

    val_str = str(val).strip()
    if val_str == "":
        return ""

    is_number = False
    try:
        float(val_str)
        # 排除日期格式（含 - 或 /）
        if "-" not in val_str and "/" not in val_str:
            is_number = True
    except ValueError:
        is_number = False

    if is_number:
        rt = RichText()
        rt.add(val_str, color="FF0000", bold=True)
        return rt
    else:
        return val_str


# ---------------- 主程式 ----------------
st.set_page_config(page_title="節能績效計劃書生成器", page_icon="📊")

st.title("📊 HWsmart節能績效計劃書生成器")
st.markdown("""
此工具支援 **Excel 表格同步** 功能：

1. **單一變數（例如：COP、效率、kWh 等）**  
   - 請放在 Excel 的 `變數` 或 `Variables` 工作表中。  
   - 第 1 欄為「變數名稱」，第 2 欄為「數值」，其餘欄位會被忽略。  
   - 在 Word 中使用：`{{r 變數名稱}}`。

2. **表格資料（例如：改善前冰水機、改善前水泵…）**  
   - 每個表格放在獨立的 Sheet，**Sheet 名稱 = Word 中的變數名稱**  
     （例如 Excel Sheet 叫 `改善前冰水機`，Word 中就寫 `改善前冰水機`）。
   - 在 Word 表格內使用（搭配 docxtpl 的 row 擴充）：  

     開頭列某一格寫：`{%tr for row in 改善前冰水機 %}`  
     中間每個儲存格：`{{ row.欄位名 }}` 或 `{{r row.欄位名 }}`  
     結尾列某一格寫：`{%tr endfor %}`

3. **RichText（紅字）**  
   - 只要 Python 端把某變數處理成 RichText，Word 模板要寫成 `{{r 變數}}` 或 `{{r row.欄位}}`。
""")

col1, col2 = st.columns(2)
with col1:
    uploaded_word = st.file_uploader("1️⃣ 上傳 Word 模板 (.docx)", type="docx")
with col2:
    uploaded_excel = st.file_uploader("2️⃣ 上傳 Excel 數據 (.xlsx)", type="xlsx")

if uploaded_word and uploaded_excel:
    st.divider()

    if st.button("🚀 開始生成報告", type="primary"):
        try:
            # ===== 先把上傳檔案讀成 bytes，避免被多次 read() 造成錯位 =====
            word_bytes = uploaded_word.read()
            excel_bytes = uploaded_excel.read()

            # 給 pandas 使用 BytesIO
            excel_io = io.BytesIO(excel_bytes)
            excel_file = pd.ExcelFile(excel_io)
            sheet_names = excel_file.sheet_names

            context = {}
            st.write("🔍 正在解析 Excel 資料...")

            for sheet_name in sheet_names:
                df = excel_file.parse(sheet_name=sheet_name)

                # === 1) 「變數」Sheet：一律當成變數清單，只吃前兩欄 ===
                #    -> 讓你第一章的數值可以直接在 Word 裡用 {{r 變數名}} 插入
                if sheet_name in ["變數", "Variables"]:
                    df_var = excel_file.parse(sheet_name=sheet_name, header=None)
                    count_vars = 0
                    for _, row in df_var.iterrows():
                        if pd.isna(row[0]):
                            continue
                        key = str(row[0]).strip()   # 例如 forging_eff_pre
                        val = row[1]                # 對應數值 1.04
                        context[key] = process_value_to_richtext(val)
                        count_vars += 1

                    st.success(f"✅ 已載入變數表：{sheet_name}（共 {count_vars} 個變數）")

                # === 2) 其他 Sheet：當成一般「表格列表」 ===
                else:
                    table_list = []
                    for _, row in df.iterrows():
                        row_dict = {}
                        for col_name in df.columns:
                            val = row[col_name]
                            row_dict[col_name] = process_value_to_richtext(val)
                        table_list.append(row_dict)

                    context[sheet_name] = table_list
                    st.success(f"✅ 已載入表格資料：{sheet_name}（共 {len(table_list)} 筆）")

            # ===== 使用 docxtpl 渲染 Word 模板 =====
            doc_stream = io.BytesIO(word_bytes)
            doc = DocxTemplate(doc_stream)
            doc.render(context)

            # ===== 輸出到記憶體，再提供下載 =====
            output_buffer = io.BytesIO()
            doc.save(output_buffer)
            doc_bytes = output_buffer.getvalue()

            # 檔名邏輯：如果有「檔名」這個變數且是一般字串，就用它當檔名
            download_name = "報告測試.docx"
            file_name_var = context.get("檔名", None)
            if isinstance(file_name_var, str) and file_name_var.strip():
                download_name = f"{file_name_var.strip()}.docx"

            st.download_button(
                label="📥 下載生成的報告",
                data=doc_bytes,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        except Exception as e:
            st.error(f"❌ 發生錯誤：{e}")
            st.info(
                "提示：請檢查 Word 模板中的標籤。\n"
                "1. 請確保使用 {{ 變數名稱 }} 而非 {{r 變數名稱}}。\n"
                "2. 若發生 'Table' 相關錯誤，請改用標準 Word 表格重新排版。"
            )





