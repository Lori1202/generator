import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, RichText
import io
import sys


try:
    st.set_page_config(page_title="節能績效計劃書生成器", page_icon="📊")
except Exception:
    pass


# 1️ 格式規則設定

FORMAT_RULES = {
    "me_prefix": {"description": "ME 類：千分位 + 保留原始小數"},
    "decimal_2": {
        "keywords": ["_rate", "elec_price", "new_cop_std", "new_eff_std"],
        "description": "2 位小數"
    },
    "decimal_1": {"keywords": ["_year"], "description": "1 位小數"},
    "integer": {"description": "整數（預設）"},
}


def clean_text(val):
    """ 強力清洗：處理 nan, None, 以及多餘空白 """
    if pd.isna(val): return ""
    s = str(val).strip()
    if s.lower() in ["nan", "none", "nat", ""]: return ""
    return s

def process_value_to_richtext(val, key_name="", debug=False):
    """ 處理變數 Sheet：格式化數字並轉為紅字 """
    val_str = clean_text(val)
    if val_str == "": return ""

    # 區間值與特殊文字 -> 回傳原色字串 (不變紅)
    text_markers = ["~", "～", "/", "&", "、", "New", "new", "主機", "型號", "CHP", "CWP", "CH-"]
    if any(marker in val_str for marker in text_markers):
        if debug: st.write(f"🔤 [文字] {key_name}: {val_str}")
        return val_str

    try:
        if "-" in val_str and not val_str.startswith("-"): raise ValueError
        
        float_val = float(val_str)
        key_lower = str(key_name).lower()
        rule_desc = ""

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
            st.write(f"🔢 [數字] {key_name} ({rule_desc}): {formatted} -> 🔴")

        # 數字 -> 轉紅字
        rt = RichText()
        rt.add(formatted, color="FF0000", bold=False)
        return rt

    except ValueError:
        return val_str


# 3️ UI 介面

st.title("📊 HWsmart 節能績效計劃書生成器")
st.markdown("""
### 使用說明

1.  **Word 模板變數寫法：** `{{變數名稱}}` 
2.  **Excel 設定：**
    * **Sheet 1**: 變數設定 (A欄名稱, B欄數值)。
    * **Sheet 2+**: 表格資料 (Sheet 名稱需對應 Word 標籤)。
""")

col1, col2 = st.columns(2)
with col1:
    uploaded_word = st.file_uploader("1️⃣ 上傳 Word 模板 (.docx)", type="docx")
with col2:
    uploaded_excel = st.file_uploader("2️⃣ 上傳 Excel 數據 (.xlsx)", type="xlsx")

debug_mode = st.checkbox("顯示除錯與變數清單 (Debug Mode)")

if uploaded_word and uploaded_excel:
    st.divider()
    if st.button("🚀 開始生成報告", type="primary"):
        try:
            uploaded_word.seek(0)
            uploaded_excel.seek(0)

            word_bytes = uploaded_word.read()
            excel_file = pd.ExcelFile(uploaded_excel)
            
            context = {}
            st.toast("🔍 正在處理資料...")

            # Debug: 顯示讀到的所有 Sheet 名稱
            if debug_mode:
                st.info(f"📂 偵測到的分頁清單：{excel_file.sheet_names}")

            # ---  Excel 分頁 ---
            for idx, sheet_name in enumerate(excel_file.sheet_names):
                

                # Sheet 1: 變數

                if idx == 0:
                    df_var = excel_file.parse(sheet_name, header=None)
                    for i, row in df_var.iterrows():
                        if pd.isna(row[0]): continue
                        key = str(row[0]).strip()
                        if not key or key.lower() == "nan": continue

                        val_b = row.iloc[1] if len(row) > 1 else None
                        val_c = row.iloc[2] if len(row) > 2 else None
                        
                        final_val = val_b if clean_text(val_b) != "" else val_c
                        context[key] = process_value_to_richtext(final_val, key, debug=debug_mode)


                # Sheet 2+: 表格

                else:
                    # 1. 全部讀取為字串，避免格式跑掉
                    df = excel_file.parse(sheet_name, dtype=str)
                    
                    # 2. 清洗欄位名稱
                    df.columns = [str(c).strip() for c in df.columns]

                    # 3.批量清洗內容
                    # 使用 Pandas 原生方法一次處理所有 nan, None, <NA>
                    df = df.fillna("")
                    df = df.replace([r"^nan$", r"^NaN$", r"^None$", r"^<NA>$"], "", regex=True)

                    # 4. 過濾有效列
                    # 只要該列「任一欄位」有內容，就保留 (避免誤刪)
                    valid_rows = []
                    for _, row in df.iterrows():
                        # 建立該列的字典
                        row_dict = {col: str(row[col]).strip() for col in df.columns}
                        
                        # 檢查整列是否全是空字串
                        # join 所有的值，如果長度 > 0 代表有東西
                        if "".join(row_dict.values()) != "":
                            valid_rows.append(row_dict)

                    context[sheet_name] = valid_rows
                    
                    # Debug 訊息
                    msg = f"✅ 表格 [{sheet_name}]：保留 {len(valid_rows)} 筆資料"
                    if len(valid_rows) == 0:
                        st.warning(f"⚠️ 表格 [{sheet_name}] 似乎是空的？(0 筆資料)")
                    elif debug_mode:
                        st.success(msg)

            # --- 生成 Word ---
            doc = DocxTemplate(io.BytesIO(word_bytes))
            doc.render(context)
            
            output = io.BytesIO()
            doc.save(output)
            
            st.session_state["generated_doc"] = output.getvalue()
            st.success("🎉 報告生成成功！請下載。")

        except Exception as e:
            st.error(f"❌ 發生錯誤：{e}")
            st.warning("💡 提示：請檢查 Word 模板的標籤是否正確，或是否有多餘的 {{r ...}}")

    if "generated_doc" in st.session_state:
        st.download_button(
            label="📥 下載 Word 報告",
            data=st.session_state["generated_doc"],
            file_name="Generated_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
