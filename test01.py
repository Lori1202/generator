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
       - 若 key_name 結尾是 "_rate" (不分大小寫) 或 包含 "elec_price" (不分大小寫)：強制保留 2 位小數
       - 其他變數：四捨五入取整數 (無小數)
       - 格式化後皆標示為 RichText 紅字粗體
    3. 其他：回傳字串
    """
    if pd.isna(val):
        return ""

    val_str = str(val).strip()
    if val_str == "":
        return ""

    if "~" in val_str or "～" in val_str:
        rt = RichText()
        rt.add(val_str, color="000000", bold=False) # 強制黑色、不加粗
        return rt

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
        # 轉小寫並去空白，增加比對成功率
        key_lower = str(key_name).strip().lower()
        

        if key_lower.startswith("me_"):

            #使用字串切割方式判斷原始數據為39.0與2
            if "." in val_str:
                # 如果原始資料有小數點 (例如 "39.0" 或 "1234.56")
                parts = val_str.split(".")
                integer_part = parts[0]
                decimal_part = parts[1]
                
                # 整數部分加千分位
                formatted_int = "{:,}".format(int(integer_part))
                
                # 拼回去：千分位整數 + "." + 原始小數部分
                formatted_str = f"{formatted_int}.{decimal_part}"
            else:
                # 如果原始資料沒有小數點 (例如 "2" 或 "1000")
                formatted_str = "{:,}".format(int(float_val))

        # 2. 結尾是 _rate 或 包含 elec_price (強制 2 位小數)
        elif key_lower.endswith("_rate") or 
             "elec_price" in key_lower or
             "new_cop_std" in key_lower or
             "new_eff_std" in key_lower):
            formatted_str = "{:,.2f}".format(float_val)

        # 3. 結尾是 _year (強制 1 位小數)
        elif key_lower.endswith("_year"):
            formatted_str = "{:,.1f}".format(float_val)
            
        # 4. 其他預設情況 (四捨五入取整數)
        else:
            formatted_str = "{:,.0f}".format(float_val)
            
        rt = RichText()
        rt.add(formatted_str, color="FF0000", bold=False)
        return rt
    else:
        return val_str

# ---------------- 主程式 ----------------
st.set_page_config(page_title="節能績效計劃書生成器", page_icon="📊")

st.title("📊 HWsmart節能績效計劃書生成器")
st.markdown("""
此工具支援 **Excel 表格同步** 功能：

1. **單一變數**（例如：COP、效率、kWh 等）標示為 **紅字**。
   - 請放在 Excel Sheet 的 第一個分頁中。  
   - 第 1 欄為「變數名稱」，第 2 欄為「數值」，其餘欄位會被忽略。  
   - 在 Word 中使用：`{{r 變數名稱}}`。

2. **表格資料（例如：改善前冰水機、改善前水泵…）** - 每個表格放在獨立的 Sheet，**Sheet 名稱 = Word 中的變數名稱** （例如 Excel Sheet 叫 `改善前冰水機`，Word 中就寫 `改善前冰水機`）。
   - 在 Word 表格內使用（搭配 docxtpl 的 row 擴充）：  

     開頭列某一格寫：`{%tr for row in 改善前冰水機 %}`  
     中間每個儲存格：`{{ row.欄位名 }}` 或 `{{r row.欄位名 }}`  
     結尾列某一格寫：`{%tr endfor %}`

3. **RichText（紅字）** - 只要 Python 端把某變數處理成 RichText，Word 模板要寫成 `{{r 變數}}` 或 `{{r row.欄位}}`。
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
            debug_logs = [] # 用來存變數讀取紀錄
            st.toast("🔍 正在解析 Excel 資料...") # 使用 toast 比較不干擾

            # 用 enumerate 來取得索引
            for i, sheet_name in enumerate(sheet_names):
                
                # 1) 變數 Sheet：只要是第 1 個 Sheet (Index 0)，不論名稱為何，都視為變數表
                if i == 0:
                    df_var = excel_file.parse(sheet_name=sheet_name, header=None)
                    count_vars = 0
                    for _, row in df_var.iterrows():
                        if pd.isna(row[0]):
                            continue
                        key = str(row[0]).strip()
                        val = row[1]
                        
                        # 處理變數 (傳入 key 進行判斷)
                        processed_val = process_value_to_richtext(val, key_name=key)
                        context[key] = processed_val
                        count_vars += 1

                        # 記錄 debug 資訊
                        val_display = val
                        is_decimal = False
                        key_lower = key.lower()
                        # debug 顯示邏輯與處理邏輯同步
                        if key_lower.endswith("_rate") or "elec_price" in key_lower:
                            is_decimal = True
                            
                        debug_logs.append(f"變數: {key} | 原始值: {val} | 判斷小數: {is_decimal}")

                # 2) 表格 Sheet：其餘的 Sheet
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











