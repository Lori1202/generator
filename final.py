import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile
import logging
from typing import Dict, List, Any, Optional, Tuple, Union
import re

# 設定 Log
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- 常數配置 ---

class AppConfig:
    PAGE_TITLE = "節能績效計劃書生成器"
    PAGE_ICON = "📊"
    LAYOUT = "wide"
    
    # 變數格式化規則 (針對 Sheet 1 變數設定，保持不變)
    FORMAT_RULES = {
        "me_prefix": {"description": "ME 類：千分位 + 保留原始小數"},
        "decimal_2": {"keywords": ["_rate", "elec_", "new_cop_std", "new_eff_std"], "description": "2 位小數"},
        "decimal_1": {"keywords": ["_year"], "description": "1 位小數"},
    }

    # 表格只有欄位名稱包含以下關鍵字的，才會進行數值格式化(千分位+小數點)
    TABLE_INCLUDE_KEYWORDS = ["kwh", "elecost", "eleccostperkwh"]

    # 識別關鍵字
    TARGET_NAMES = ["名稱", "name", "設備名稱"]
    TARGET_NOS = ["no", "編號", "設備編號", "冰水主機代號"]
    
    # 排序權重 (數字越小越前面)
    SORT_WEIGHTS = {
        "chiller": 1, "主機": 1,
        "pump": 2, "泵": 2,
        "tower": 3, "水塔": 3
    }

class DataFormatter:
    """清洗與格式化"""
    
    @staticmethod
    def clean_text(val: Any) -> str:
        if pd.isna(val): 
            return ""
        s = str(val).strip()
        if s.lower() in ["nan", "none", "nat", ""]: 
            return ""
        return s

    @staticmethod
    def format_variable_value(val: Any, key_name: str = "") -> str:
        """格式化邏輯"""
        val_str = DataFormatter.clean_text(val)
        if not val_str: 
            return ""
        

        if any(x in val_str for x in ["~", "CH", "CWP", "HP", "/", "New", "new"]): 
            return val_str
        
        try:
            float_val = float(val_str)
            key_lower = str(key_name).lower()
            
            # 1: ME 開頭
            if key_lower.startswith("me_"):
                if "." in val_str:
                    parts = val_str.split(".")
                    return f"{int(parts[0]):,}.{parts[1]}"
                return f"{int(float_val):,}"
            
            # 2: 兩位小數
            if any(k in key_lower for k in AppConfig.FORMAT_RULES["decimal_2"]["keywords"]):
                return f"{float_val:,.2f}"
            
            # 3: 一位小數
            if any(k in key_lower for k in AppConfig.FORMAT_RULES["decimal_1"]["keywords"]):
                return f"{float_val:,.1f}"
            
            # 整數
            return f"{float_val:,.0f}"

        except ValueError:
            return val_str

    @staticmethod
    def format_table_value(val: Any, col_name: str) -> str:
        """針對excel內數值的格式化邏輯"""
        val_str = DataFormatter.clean_text(val)
        if not val_str: 
            return ""

        col_lower = str(col_name).lower()

        is_target_col = any(k in col_lower for k in AppConfig.TABLE_INCLUDE_KEYWORDS)

        # 如果不是 kwh, elecost, eleccostperkwh，直接回傳原值
        if not is_target_col:
            
            return val_str

        # 2. 針對目標欄位進行數值格式化
        try:

            clean_num_str = val_str.replace(",", "")
            f_val = float(clean_num_str)
            

            # 若為整數，加千分位 (1,000)
            # 若為小數，加千分位 + 兩位小數 (1,000.50)
            if f_val.is_integer():
                return f"{int(f_val):,}"
            else:
                return f"{f_val:,.2f}"
                
        except ValueError:
            # 若目標欄位內容轉型失敗 (例如寫了 "N/A")，則回傳原值
            return val_str

class ExcelParser:
    """Excel 讀取"""
    
    @staticmethod
    def _find_header_row(df_preview: pd.DataFrame) -> Tuple[int, str]:

        #找前 20 列以尋找標題列與表格類型
        target_names = [x.lower() for x in AppConfig.TARGET_NAMES]
        target_nos = [x.lower() for x in AppConfig.TARGET_NOS]
        
        for i, row in df_preview.iterrows():
            row_clean = [str(x).strip().lower() for x in row.values if pd.notna(x) and str(x).strip() != ""]
            row_str = " ".join(row_clean)
            
            has_name = any(k in row_str for k in target_names)
            has_no = any(k in row_str for k in target_nos)
            
            if has_name and has_no:
                return i, "equipment"
            
        # 回傳第一個非空行作為普通表格
        for i, row in df_preview.iterrows():
             if any(pd.notna(x) and str(x).strip() != "" for x in row.values):
                 return i, "general"
                 
        return -1, "none"

    @staticmethod
    def parse_sheet(excel_file: Any, sheet_name: str) -> List[Dict[str, Any]]:
        try:
            df_preview = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, nrows=20, dtype=str)
            header_row, table_type = ExcelParser._find_header_row(df_preview)
            
            if header_row == -1:
                return []
            
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header_row, dtype=str)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed', na=False)]
            df = df.dropna(axis=1, how='all')
            df.columns = [str(c).strip() for c in df.columns]
            
            results = []
            
            if table_type == "equipment":
                results = ExcelParser._process_equipment_table(df)
            else:
                results = ExcelParser._process_general_table(df)
                
            return results
        except Exception as e:
            logger.error(f"Error parsing sheet {sheet_name}: {e}")
            return []

    @staticmethod
    def _process_equipment_table(df: pd.DataFrame) -> List[Dict[str, Any]]:

        col_map = {}
        target_names = [x.lower() for x in AppConfig.TARGET_NAMES]
        target_nos = [x.lower() for x in AppConfig.TARGET_NOS]

        for c in df.columns:
            c_low = c.lower()
            if c_low in target_names or any(t in c_low for t in target_names):
                if 'name' not in col_map: col_map['name'] = c
            if c_low in target_nos or any(t in c_low for t in target_nos):
                if 'no' not in col_map: col_map['no'] = c
        
        results = []
        if 'name' in col_map and 'no' in col_map:
            df['temp_name'] = df[col_map['name']]
            df['temp_no'] = df[col_map['no']]
            df = df.dropna(subset=['temp_name', 'temp_no'])
            df = df[~df['temp_name'].str.contains('名稱|Equipment|name', case=False, na=False)]
            
            for _, row in df.iterrows():
                row_dict = {}
                for col in df.columns:
                    if col in ['temp_name', 'temp_no']: continue
                    # 套用表格數值格式化邏輯
                    row_dict[col] = DataFormatter.format_table_value(row[col], col)
                

                row_dict['name'] = DataFormatter.clean_text(row[col_map['name']])
                row_dict['no'] = DataFormatter.clean_text(row[col_map['no']])
                results.append(row_dict)
        else:
            return ExcelParser._process_general_table(df)
            
        return results

    @staticmethod
    def _process_general_table(df: pd.DataFrame) -> List[Dict[str, Any]]:
        results = []
        for _, row in df.iterrows():
            if row.isna().all() or all(str(x).strip() == "" for x in row.values):
                continue

            row_dict = {col: DataFormatter.format_table_value(row[col], col) for col in df.columns}
            results.append(row_dict)
        return results

# --- main ---

class ContextBuilder:
    def __init__(self, excel_file: Any):
        self.excel_file = excel_file
        self.xls = pd.ExcelFile(excel_file)
        self.context: Dict[str, Any] = {}
        self.counters = {
            "pm": 1, 
            "fm": 1, 
            "t": 1
        }

    def build(self) -> Dict[str, Any]:
        self._load_variables()
        self._process_sheets()
        return self.context

    def _load_variables(self):
        """讀取單一變數設定頁籤"""
        sheet_name = "變數" if "變數" in self.xls.sheet_names else self.xls.sheet_names[0]
        try:
            df_var = self.xls.parse(sheet_name, header=None)
            for _, row in df_var.iterrows():
                if pd.isna(row[0]): continue
                key = str(row[0]).strip()
                val = row[1] if len(row) > 1 else ""
                self.context[key] = DataFormatter.format_variable_value(val, key)
        except Exception as e:
            logger.warning(f"變數讀取失敗或格式有誤: {e}")

    def _process_sheets(self):
        groups = {"before": [], "after": []}
        
        for sheet in self.xls.sheet_names:
            if sheet == "變數": continue
            
            data = ExcelParser.parse_sheet(self.excel_file, sheet)
            if not data: continue
            
            if "改善前" in sheet:
                groups["before"].append((sheet, data))
            elif "改善後" in sheet:
                groups["after"].append((sheet, data))
            else:
                self.context[sheet] = data
                if self._is_pump_sheet(sheet):
                    self._classify_pumps(sheet, data)

        self._process_group(groups["before"])
        self._process_group(groups["after"])

    def _process_group(self, sheet_list: List[Tuple[str, List[Dict]]]):
        sheet_list.sort(key=lambda x: self._get_sort_weight(x[0]))
        self._apply_numbering(sheet_list)
        
        for sheet_name, items in sheet_list:
            self.context[sheet_name] = items
            if self._is_pump_sheet(sheet_name):
                self._classify_pumps(sheet_name, items)

    def _get_sort_weight(self, name: str) -> int:
        name_lower = name.lower()
        for key, weight in AppConfig.SORT_WEIGHTS.items():
            if key in name_lower:
                return weight
        return 4

    def _is_pump_sheet(self, sheet_name: str) -> bool:
        return "泵" in sheet_name or "pump" in sheet_name.lower()

    def _classify_pumps(self, base_key: str, items: List[Dict]):
        categories = {
            "ice": [], "cool": [], "zone": [], "other": []
        }
        
        for item in items:
            name_str = str(item.get('name', ''))
            no_str = str(item.get('no', '')).upper()
            
            if 'ZP' in no_str or '區域' in name_str:
                categories["zone"].append(item)
            elif 'CWP' in no_str or '冷卻' in name_str:
                categories["cool"].append(item)
            elif 'CHP' in no_str or '冰水' in name_str:
                categories["ice"].append(item)
            else:
                categories["other"].append(item)
        
        self.context[f"{base_key}_冰水"] = categories["ice"]
        self.context[f"{base_key}_冷卻"] = categories["cool"]
        self.context[f"{base_key}_區域"] = categories["zone"]
        self.context[f"{base_key}_其他"] = categories["other"]

    def _apply_numbering(self, sheet_list: List[Tuple[str, List[Dict]]]):
        for _, items in sheet_list:
            for item in items:
                item['pm'] = f"PM{self.counters['pm']}"
                self.counters['pm'] += 1
        
        for sheet_name, items in sheet_list:
            if any(k in sheet_name.lower() for k in ["主機", "chiller", "冰水機"]):
                for item in items:
                    item['evap_fm'] = f"FM{self.counters['fm']}"
                    self.counters['fm'] += 1
                    item['evap_t_out'] = f"T{self.counters['t']}"
                    item['evap_t_in'] = f"T{self.counters['t']+1}"
                    self.counters['t'] += 2
                    
                    item['cond_fm'] = f"FM{self.counters['fm']}"
                    self.counters['fm'] += 1
                    item['cond_t_out'] = f"T{self.counters['t']}"
                    item['cond_t_in'] = f"T{self.counters['t']+1}"
                    self.counters['t'] += 2

# --- UI---

class ReportGeneratorUI:
    def __init__(self):
        self._setup_page()

    def _setup_page(self):
        try:
            st.set_page_config(
                page_title=AppConfig.PAGE_TITLE, 
                page_icon=AppConfig.PAGE_ICON, 
                layout=AppConfig.LAYOUT
            )
        except Exception:
            pass 
        
        st.title(f"{AppConfig.PAGE_ICON} {AppConfig.PAGE_TITLE}")
        self._render_instructions()

    def _render_instructions(self):
        st.markdown("""
        ### ⚠️ 重要使用說明
        1.  **Word 模板變數寫法：** `{{變數名稱}}` 
        2.  **Excel 設定：**
            * **Sheet 1**: 變數設定 (A欄名稱, B欄數值)。
            * **Sheet 2+**: 表格資料 (Sheet 名稱需對應 Word 標籤)。

        """)

    def run(self):
        col1, col2 = st.columns(2)
        with col1:
            uploaded_excel = st.file_uploader("1️⃣ 上傳 Excel", type="xlsx")
        with col2:
            uploaded_templates = st.file_uploader("2️⃣ 上傳 Word 模板", type="docx", accept_multiple_files=True)

        if uploaded_excel and uploaded_templates:
            if st.button("🚀 生成報告", type="primary"):
                self._generate_report(uploaded_excel, uploaded_templates)

    def _generate_report(self, excel_file, templates):
        try:
            with st.spinner("資料處理中，請稍候..."):
                builder = ContextBuilder(excel_file)
                context = builder.build()
                
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    for tpl in templates:
                        tpl.seek(0)
                        doc = DocxTemplate(tpl)
                        doc.render(context)
                        
                        out = io.BytesIO()
                        doc.save(out)
                        zf.writestr(f"Result_{tpl.name}", out.getvalue())
                
                st.success("✅ 報告生成成功！")
                st.download_button(
                    "📦 下載結果 (ZIP)", 
                    zip_buffer.getvalue(), 
                    "Reports.zip", 
                    "application/zip"
                )
                
        except Exception as e:
            logger.error(e, exc_info=True)
            st.error(f"發生錯誤: {str(e)}")

if __name__ == "__main__":
    app = ReportGeneratorUI()

    app.run()

