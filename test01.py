import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, RichText
import io

# ---------------- 處理數字邏輯 ----------------
def process_value_to_richtext(val, key_name=""):
    """
    判斷數值是否需要變紅並格式化：
    Args:
        val: 數值內容
        key_name: 變數名稱 (用來判斷格式化規則)
    
    規則：
    1. 空值 / NaN：回傳空字串
    2. 純數字：
       - 若 key_name 結尾是 "_rate" 或開頭是 "elec_price_"：強制保留 2 位小數
       - 其他變數：四捨五入取整數 (無小數)
       - 格式化後皆標示為 RichText 紅字粗體
    3. 其他：回傳字串
    """
    if pd.isna(val):
        return ""

    val_str = str(val).strip()
    if val_str == "":
        return ""

    is_number = False
    float_val = 0.0
    
    try:
        # 排除日期格式邏輯
        if "/" not in val_str:
            # 處理負號邏輯 (避免將 2023-01-01 誤判為負數)
            if "-" in val_str:
                if val_str.count("-") == 1 and val_str.startswith("-"):
                    float_val = float(val_str)
                    is_number = True
                else:
                    is_number = False
            else:
                float_val = float(val_str)
                is_number = True
    except ValueError:
        is_number = False

    if is_number:
        # === 根據變數名稱決定小數位數 ===
        key_str = str(key_name).strip()
        
        # 條件 1: 變數名稱結尾是 "_rate" 或 開頭是 "elec_price_"
        if key_str.endswith("_rate}}") or key_str.startswith("{{elec_price_"):
            formatted_str = "{:,.2f}".format(float_val) # 強制 2 位小數
        else:
            # 條件 2: 其餘數值，四捨五入取整數
            # 注意：Python 的 format 在 .0f 時會自動四捨五入 (Rounding)
            formatted_str = "{:,.0f}".format(float_val) 
            
        rt = RichText()
        rt.add(formatted_str, color="FF0000", bold=True)
        return rt
    else:
        # 非數值 (如日期字串、文字)，直接回傳原始內容
        return val_str

# ---------------- 主程式 ----------------
st.set_page_config(page_title="節能績效計劃書生成器", page_icon="📊")

st.title("📊 HWsmart節能績效計劃書生成器")
st.markdown("""
此工具支援 **Excel 表格同步** 功能：

1. **單一變數（例如：COP、效率、kWh 等）**  
   - 自動將數值轉為 **千分位 + 兩位小數** 並標示為 **紅字**。
   - 例如：`1506877.247` → `1,506,877.25` (紅字)。
   - 請放在 Excel Sheet 的 `變數` 或 `Variables` 工作表中。  
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

    # 按鈕邏輯修正：使用 session_state 來處理生成狀態
    if st.button("🚀 開始生成報告", type="primary"):
        try:
            # 重置指標至開頭，確保重複執行時讀取正確
            uploaded_word.seek(0)
            uploaded_excel.seek(0)

            # 讀取檔案
            word_bytes = uploaded_word.read()
            excel_bytes = uploaded_excel.read()

            excel_io = io.BytesIO(excel_bytes)
            excel_file = pd.ExcelFile(excel_io)
            sheet_names = excel_file.sheet_names

            context = {}
            st.toast("🔍 正在解析 Excel 資料...") # 使用 toast 比較不干擾

            for sheet_name in sheet_names:
                # 1) 變數 Sheet
                if sheet_name in ["變數", "Variables"]:
                    df_var = excel_file.parse(sheet_name=sheet_name, header=None)
                    count_vars = 0
                    for _, row in df_var.iterrows():
                        if pd.isna(row[0]):
                            continue
                        key = str(row[0]).strip()
                        val = row[1]
                        context[key] = process_value_to_richtext(val)
                        count_vars += 1
                    # 存入 log 供除錯用，不直接 print
                    print(f"變數表載入: {count_vars} 筆")

                # 2) 表格 Sheet
                else:
                    df = excel_file.parse(sheet_name=sheet_name)

                    # 刪除整列都是 NaN (空值) 的列
                    df = df.dropna(how='all')
                    
                    # 去除欄位名稱的空格，避免 Jinja2 報錯 (Option)
                    df.columns = [str(c).strip() for c in df.columns]
                    
                    table_list = []
                    for _, row in df.iterrows():
                        row_dict = {}
                        for col_name in df.columns:
                            val = row[col_name]
                            row_dict[col_name] = process_value_to_richtext(val, key_name=col_name)
                        table_list.append(row_dict)

                    context[sheet_name] = table_list
                    print(f"已載入表格資料：{sheet_name}（共 {len(table_list)} 筆）")

            # 渲染 Word
            doc_stream = io.BytesIO(word_bytes)
            doc = DocxTemplate(doc_stream)
            doc.render(context)

            # 輸出
            output_buffer = io.BytesIO()
            doc.save(output_buffer)
            doc_bytes = output_buffer.getvalue()

            # 檔名邏輯
            download_name = "報告測試.docx"
            file_name_var = context.get("檔名", None)
            
            # 注意：如果 "檔名" 變數也被轉成 RichText，要取回純文字才能當檔名
            if isinstance(file_name_var, RichText):
                # 這裡簡單處理，RichText 很難直接轉回 string，建議檔名變數在 Excel 裡不要是純數字
                download_name = "Generated_Report.docx" 
            elif isinstance(file_name_var, str) and file_name_var.strip():
                download_name = f"{file_name_var.strip()}.docx"

            # 將結果存入 Session State ==
            st.session_state['generated_doc'] = doc_bytes
            st.session_state['download_name'] = download_name
            st.success("✅ 報告生成成功！請點擊下方按鈕下載。")

        except Exception as e:
            st.error(f"❌ 發生錯誤：{e}")
            
    # 只要 session_state 裡有檔案，就顯示下載按鈕
    if 'generated_doc' in st.session_state:
        st.download_button(
            label="📥 下載生成的報告",
            data=st.session_state['generated_doc'],
            file_name=st.session_state['download_name'],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )








