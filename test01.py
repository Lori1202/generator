
import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, RichText
import io

# ---------------- 輔助函式：處理紅字邏輯 ----------------
def process_value_to_richtext(val):
    """判斷數值是否需要變紅，回傳 RichText 或原始值"""
    val_str = str(val).strip()
    if pd.isna(val) or val_str == "":
        return ""
    
    # 判斷是否為數字
    is_number = False
    try:
        float(val_str)
        # 排除包含日期分隔符號的字串
        if '-' not in val_str and '/' not in val_str:
            is_number = True
    except ValueError:
        is_number = False

    if is_number:
        rt = RichText()
        rt.add(val_str, color='FF0000', bold=True)
        return rt
    else:
        return val

# ---------------- 主程式 ----------------
st.set_page_config(page_title="節能績效計劃書生成器", page_icon="📊")

st.title("📊 節能績效計劃書生成器(表格連動版)")
st.markdown("""
此工具支援 **Excel 表格同步** 功能：
1. **單一變數**：請放在 Excel 第一個 Sheet (或命名為 '變數')。
2. **表格資料**：請將每個表格放在獨立的 Sheet，Sheet 名稱即為 Word 中的變數名稱 (例如 `冰水機表`)。
3. **Word 設定**：在表格列使用 `{% tr for item in 冰水機表 %}` ... `{% tr endfor %}`。
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
            # 讀取 Excel 所有工作表
            excel_file = pd.ExcelFile(uploaded_excel)
            sheet_names = excel_file.sheet_names
            
            context = {}
            st.write("🔍 正在解析 Excel 資料...")

            for sheet_name in sheet_names:
                # 讀取該 Sheet
                df = pd.read_excel(uploaded_excel, sheet_name=sheet_name)
                
                # --- 判斷是「變數清單」還是「表格數據」 ---
                # 規則：如果欄位少於等於 2 且第一欄像是 Key，視為單一變數
                # 但為了彈性，我們約定：名為 "變數" 或 "Variables" 的 Sheet 視為單一變數
                # 其他 Sheet 視為表格列表
                
                if sheet_name in ["變數", "Variables", "Sheet1"] and len(df.columns) <= 2:
                    # === 處理單一變數 ===
                    # 假設第一欄是 Key，第二欄是 Value
                    # 重新讀取，不設 header 以便抓取第一列
                    df_var = pd.read_excel(uploaded_excel, sheet_name=sheet_name, header=None)
                    for index, row in df_var.iterrows():
                        if pd.isna(row[0]): continue
                        key = str(row[0]).strip()
                        val = row[1]
                        context[key] = process_value_to_richtext(val)
                    st.success(f"✅ 已載入變數表：{sheet_name}")

                else:
                    # === 處理表格列表 (Table List) ===
                    table_list = []
                    # 逐列處理
                    for index, row in df.iterrows():
                        row_dict = {}
                        for col_name in df.columns:
                            val = row[col_name]
                            # 對表格內的每個儲存格也套用紅字邏輯
                            row_dict[col_name] = process_value_to_richtext(val)
                        table_list.append(row_dict)
                    
                    # 將整張表存入 Context，Key 就是 Sheet 名稱
                    context[sheet_name] = table_list
                    st.success(f"✅ 已載入表格資料：{sheet_name} (共 {len(table_list)} 筆)")

            # --- 渲染 Word ---
            doc = DocxTemplate(uploaded_word)
            doc.render(context)

            # --- 輸出 ---
            output_buffer = io.BytesIO()
            doc.save(output_buffer)
            output_buffer.seek(0)

            # 檔名邏輯
            download_name = "報告_表格連動版.docx"
            if "檔名" in context and not isinstance(context["檔名"], RichText):
                download_name = f"{context['檔名']}.docx"
            elif "檔名" in context and isinstance(context["檔名"], RichText):
                 # 如果檔名不小心變紅字了，取出純文字
                 # RichText 目前沒有直接取文字的方法，這裡做簡單防呆
                 download_name = "報告_表格連動版.docx"

            st.download_button(
                label="📥 下載生成的報告",
                data=output_buffer,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"❌ 發生錯誤：{e}")
            st.info("提示：請確認 Word 裡的表格標籤 `{% tr for ... %}` 是否與 Excel Sheet 名稱一致。")