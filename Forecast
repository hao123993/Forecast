import streamlit as st
import pandas as pd
import os
from io import BytesIO

st.set_page_config(page_title="產地自動分配與生管檢核系統", layout="wide")

# ==========================================
# 🎨 專屬奶油莫蘭迪色系主題 (CSS 升級版)
# ==========================================
custom_css = """
<style>
    .stApp, .main { background-color: #F4EFE6 !important; }
    header { background-color: transparent !important; }
    .stMarkdown, .stText, h1, h2, h3, h4, h5, h6, p, span, label, li { color: #4A433D !important; }
    [data-testid="stFileUploadDropzone"] {
        background-color: #E2D4D4 !important;
        border: 2px dashed #9E8E8E !important;
        border-radius: 10px !important;
    }
    [data-testid="stFileUploadDropzone"] * { color: #4A433D !important; }
    [data-testid="stUploadedFile"] {
        background-color: #E8DCC4 !important;
        border: 1px solid #9E8E8E !important;
        border-radius: 8px !important;
    }
    [data-testid="stUploadedFile"] * { color: #4A433D !important; }
    .stButton > button, [data-testid="stDownloadButton"] > button, [data-testid="stBaseButton-secondary"] {
        background-color: #98A9BD !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: bold !important;
    }
    .stButton > button:hover, [data-testid="stDownloadButton"] > button:hover, [data-testid="stBaseButton-secondary"]:hover {
        background-color: #A9BA9D !important;
        color: #FFFFFF !important;
    }
    [data-testid="stAlert"] {
        background-color: #C0CEC3 !important;
        border: none !important;
        border-radius: 8px !important;
    }
    [data-testid="stAlert"] * { color: #2F3E35 !important; }
    [data-baseweb="tab"] {
        background-color: #E8DCC4 !important;
        border-radius: 6px 6px 0 0 !important;
        margin-right: 4px;
        padding: 0 10px;
    }
    [data-testid="stDataFrame"], [data-testid="stDataEditor"] {
        border: 2px solid #98A9BD !important;
        border-radius: 8px !important;
        padding: 5px !important;
        background-color: #FFFFFF !important;
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ==========================================
# 系統檔案與狀態初始化
# ==========================================
MAPPING_FILE = "mapping.xlsx"
ALLOCATED_FILE = "allocated_orders.xlsx"

SHEET_NAMES = [
    '(1)區域_PD_投產地', 
    '(2)區域_內外單_PD_投產地', 
    '(3)區域_客戶簡稱_PD_投產地', 
    '(4)區域_料號_PD_投產地'
]

if 'delivery_result_df' not in st.session_state:
    st.session_state['delivery_result_df'] = None
if 'warnings_df' not in st.session_state:
    st.session_state['warnings_df'] = None

@st.cache_data
def load_mapping():
    if os.path.exists(MAPPING_FILE):
        return {sheet: pd.read_excel(MAPPING_FILE, sheet_name=sheet) for sheet in SHEET_NAMES}
    return None

def save_mapping(mapping_dict):
    with pd.ExcelWriter(MAPPING_FILE, engine='openpyxl') as writer:
        for sheet, df in mapping_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet)

@st.cache_data
def load_allocated():
    if os.path.exists(ALLOCATED_FILE):
        return pd.read_excel(ALLOCATED_FILE)
    return None

def save_allocated(df):
    with pd.ExcelWriter(ALLOCATED_FILE, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='產地分配結果')

mapping_data = load_mapping()
allocated_data = load_allocated()

st.title("🏭 產地自動分配與生管檢核系統")

tab1, tab2, tab3 = st.tabs(["🚀 1. 訂單產地分配", "📊 2. 生管達交比對", "⚙️ 3. 對照表維護"])

# ==========================================
# 分頁 1: 訂單產地分配
# ==========================================
with tab1:
    st.header("訂單產地分配")
    
    if allocated_data is not None:
        st.success("✅ 系統已自動載入最新的【產地分配結果】！你可以直接前往「標籤 2」進行生管比對。")
        st.subheader("👀 目前系統中的分配結果預覽")
        
        preview_cols = ['年月', '業務地區別(TIPTOP)', '客戶簡稱', 'Product Name', '料號', '內/外單', '最終投產地', '起始日期', '備註', '系統匹配邏輯']
        display_cols = [col for col in preview_cols if col in allocated_data.columns]
        st.dataframe(allocated_data[display_cols], use_container_width=True)

        with open(ALLOCATED_FILE, "rb") as f:
            st.download_button("📥 下載目前系統分配結果 (Excel)", data=f, file_name="目前_訂單產地分配結果.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        st.divider()
        st.subheader("🔄 重新分配新訂單 (這將覆蓋上述舊資料)")

    if mapping_data is None:
        st.warning("⚠️ 請先至「對照表維護」分頁上傳對照表。")
    else:
        order_file = st.file_uploader("上傳新的【訂單明細】Excel", type=["xlsx"], key="order_upload")
        
        if order_file and st.button("🚀 開始自動分配產地", type="primary"):
            try:
                with st.spinner('讀取資料與比對中...'):
                    df_order = pd.read_excel(order_file).dropna(how='all')
                    
                    df_map1 = mapping_data[SHEET_NAMES[0]].copy()
                    df_map2 = mapping_data[SHEET_NAMES[1]].copy()
                    df_map3 = mapping_data[SHEET_NAMES[2]].copy()
                    df_map4 = mapping_data[SHEET_NAMES[3]].copy()

                    def get_best_match(matches, order_date):
                        if matches.empty: return None
                        matches = matches.copy()
                        matches['日期檢查'] = pd.to_datetime(matches['起始日期'], errors='coerce')
                        if pd.notna(order_date):
                            valid = matches[matches['日期檢查'].isna() | (matches['日期檢查'] <= order_date)]
                        else:
                            valid = matches[matches['日期檢查'].isna()]
                        if valid.empty: return None
                        return valid.sort_values(by='日期檢查', ascending=False, na_position='last').iloc[0]

                    def assign_production_site(row, df1, df2, df3, df4):
                        region = str(row.get('業務地區別(TIPTOP)', '')).strip()
                        cust = str(row.get('客戶簡稱', '')).strip()
                        pd_name = str(row.get('Product Name', '')).strip()
                        order_type = str(row.get('內/外單', '')).strip()
                        part_no = str(row.get('料號', '')).strip()
                        order_date = pd.to_datetime(row.get('年月', ''), errors='coerce')

                        # 🌟 第一優先級：匹配 (4)區域_料號_PD_投產地
                        df4_reg = df4['業務地區別'].astype(str).str.strip()
                        df4_cust = df4['客戶簡稱'].astype(str).str.strip().replace(['nan', 'None', '<NA>'], '')
                        df4_type = df4['內/外單'].astype(str).str.strip().replace(['nan', 'None', '<NA>'], '')
                        df4_pd = df4['Product Name'].astype(str).str.strip().replace(['nan', 'None', '<NA>'], '')
                        df4_part = df4['料號'].astype(str).str.strip().replace(['nan', 'None', '<NA>'], '')

                        m4_reg = (df4_reg == region)
                        m4_cust = (df4_cust == cust) | (df4_cust == '')
                        m4_type = (df4_type == order_type) | (df4_type == '')
                        m4_pd = (df4_pd == pd_name) | (df4_pd == '')
                        m4_part = (df4_part == part_no) | (df4_part == '')
                        m4_valid = (df4_pd != '') | (df4_part != '')

                        match4 = df4[m4_reg & m4_cust & m4_type & m4_pd & m4_part & m4_valid]
                        best4 = get_best_match(match4, order_date)
                        if best4 is not None: return pd.Series([best4['投產地'], best4['起始日期'], best4['備註'], "匹配 (4)"])

                        # 🌟 第二優先級：匹配 (3)區域_客戶簡稱_PD_投產地
                        match3 = df3[(df3['業務地區別'].astype(str).str.strip() == region) & (df3['客戶簡稱'].astype(str).str.strip() == cust) & (df3['Product Name'].astype(str).str.strip() == pd_name)]
                        best3 = get_best_match(match3, order_date)
                        if best3 is not None: return pd.Series([best3['投產地'], best3['起始日期'], best3['備註'], "匹配 (3)"])

                        # 🌟 第三優先級：匹配 (2)區域_內外單_PD_投產地
                        match2 = df2[(df2['業務地區別'].astype(str).str.strip() == region) & (df2['內/外單'].astype(str).str.strip() == order_type) & (df2['Product Name'].astype(str).str.strip() == pd_name)]
                        best2 = get_best_match(match2, order_date)
                        if best2 is not None: return pd.Series([best2['投產地'], best2['起始日期'], best2['備註'], "匹配 (2)"])

                        # 🌟 第四優先級：匹配 (1)區域_PD_投產地
                        match1 = df1[(df1['業務地區別'].astype(str).str.strip() == region) & (df1['Product Name'].astype(str).str.strip() == pd_name)]
                        best1 = get_best_match(match1, order_date)
                        if best1 is not None: return pd.Series([best1['投產地'], best1['起始日期'], best1['備註'], "匹配 (1)"])

                        # 🌟 最終預設
                        return pd.Series(["高雄", None, None, "預設 (高雄)"])

                    df_order[['最終投產地', '起始日期', '備註', '系統匹配邏輯']] = df_order.apply(
                        lambda row: assign_production_site(row, df_map1, df_map2, df_map3, df_map4), axis=1
                    )
                    
                    save_allocated(df_order)
                    
                    # 🚀 關鍵解法：強制清除 Streamlit 的記憶，讓它讀取最新檔案！
                    load_allocated.clear() 
                    
                    st.session_state['delivery_result_df'] = None
                    st.success("✅ 產地分配完成並已永久儲存！系統已自動更新為最新結果。")
                    st.rerun()
                    
            except Exception as e:
                st.error(f"執行發生錯誤：{e}")

# ==========================================
# 分頁 2: 生管達交比對
# ==========================================
with tab2:
    st.header("生管達交 v.s 預估投產地 比對")
    
    if allocated_data is None:
        st.warning("⚠️ 系統目前沒有【產地分配結果】資料。請先去「標籤 1」上傳訂單並執行產地分配。")
    else:
        st.info("💡 系統已自動使用「標籤 1」最新的產地分配資料作為比對基準。")
        
        # 🌟 新增：產生並提供空白範本下載
        template_cols = ['銷售組織', '客戶簡稱', '來源客戶簡稱', 'Product Name', '最終投產地']
        df_template = pd.DataFrame(columns=template_cols)
        template_io = BytesIO()
        with pd.ExcelWriter(template_io, engine='openpyxl') as writer:
            df_template.to_excel(writer, index=False, sheet_name='生管達交比對範本')
        
        st.download_button(
            label="📝 下載【生管達交】空白填寫範本 (Excel)",
            data=template_io.getvalue(),
            file_name="生管達交_空白範本.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.divider() # 加一條分隔線讓畫面更乾淨

        delivery_file = st.file_uploader("上傳填寫好的【生管達交v.s預估投產地】Excel", type=["xlsx"], key="delivery_upload")
        
        if delivery_file and st.button("🔍 開始比對與檢核", type="primary"):
            try:
                df_delivery = pd.read_excel(delivery_file).dropna(how='all')
                lookup_df = allocated_data.drop_duplicates(subset=['客戶簡稱', 'Product Name'])

                def get_delivery_site(row):
                    pd_name = str(row.get('Product Name', '')).strip()
                    src_cust = str(row.get('來源客戶簡稱', '')).strip()
                    cust = str(row.get('客戶簡稱', '')).strip()
                    
                    match1 = lookup_df[(lookup_df['客戶簡稱'].astype(str).str.strip() == src_cust) & (lookup_df['Product Name'].astype(str).str.strip() == pd_name)]
                    if not match1.empty: return match1.iloc[0]['最終投產地']
                        
                    match2 = lookup_df[(lookup_df['客戶簡稱'].astype(str).str.strip() == cust) & (lookup_df['Product Name'].astype(str).str.strip() == pd_name)]
                    if not match2.empty: return match2.iloc[0]['最終投產地']
                    
                    return ""

                with st.spinner('比對資料中...'):
                    df_delivery['最終投產地'] = df_delivery.apply(get_delivery_site, axis=1)
                    
                    org_mask = df_delivery['銷售組織'].astype(str).str.strip().isin(['CN10', 'TH10'])
                    site_mask = df_delivery['最終投產地'].astype(str).str.strip().isin(['', '高雄', 'nan'])
                    warnings_df = df_delivery[org_mask & site_mask]
                    
                    st.session_state['delivery_result_df'] = df_delivery
                    st.session_state['warnings_df'] = warnings_df

            except Exception as e:
                st.error(f"讀取或比對發生錯誤：{e}")

        if st.session_state['delivery_result_df'] is not None:
            st.success("✅ 達交比對完成！")
            
            if not st.session_state['warnings_df'].empty:
                st.error(f"🚨 警告！發現 {len(st.session_state['warnings_df'])} 筆異常紀錄 (銷售組織為 CN10/TH10，但投產地為高雄或空白)！")
                st.write("👇 **異常資料明細清單：**")
                warn_cols = ['銷售組織', '來源客戶簡稱', '客戶簡稱', 'Product Name', '最終投產地']
                display_warn_cols = [c for c in warn_cols if c in st.session_state['warnings_df'].columns]
                st.dataframe(st.session_state['warnings_df'][display_warn_cols], use_container_width=True)
            else:
                st.info("🎉 太棒了！本次檢核沒有發現任何異常的產地設定。")

            st.subheader("👀 生管達交比對結果預覽")
            st.dataframe(st.session_state['delivery_result_df'], use_container_width=True)

            output_delivery = BytesIO()
            with pd.ExcelWriter(output_delivery, engine='openpyxl') as writer:
                st.session_state['delivery_result_df'].to_excel(writer, index=False, sheet_name='比對後結果')
            
            st.download_button(
                label="📥 下載更新後的【生管達交v.s預估投產地】(Excel)",
                data=output_delivery.getvalue(),
                file_name="生管達交_比對後結果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
# ==========================================
# 分頁 3: 對照表維護
# ==========================================
with tab3:
    st.header("對照表維護")
    if mapping_data is None:
        st.info("💡 系統尚未建立對照表，請進行「首次上傳」。")
        uploaded_map = st.file_uploader("上傳初始【匹配對照表】Excel", type=["xlsx"], key="init_upload")
        if uploaded_map:
            try:
                init_mapping = {sheet: pd.read_excel(uploaded_map, sheet_name=sheet) for sheet in SHEET_NAMES}
                save_mapping(init_mapping)
                load_mapping.clear() # 🚀 清除快取
                st.success("✅ 初始對照表建立成功！請點擊下方按鈕或重新整理網頁。")
                st.rerun()
            except Exception as e:
                st.error(f"上傳失敗：請確認你的 Excel 檔案是否包含這四個分頁。詳細錯誤：{e}")
    else:
        st.write("✏️ **操作說明 1：直接修改** (在表格內點擊即可修改。滑動到最下方點擊 `+` 可新增)")
        edited_data = {}
        for sheet in SHEET_NAMES:
            st.subheader(f"📄 {sheet}")
            edited_df = st.data_editor(mapping_data[sheet], num_rows="dynamic", use_container_width=True, key=sheet)
            edited_data[sheet] = edited_df
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("💾 儲存網頁上的變更", type="primary"):
                save_mapping(edited_data)
                load_mapping.clear() # 🚀 清除快取
                st.success("✅ 變更已成功儲存至系統！")
                st.rerun()
        with col2:
            with open(MAPPING_FILE, "rb") as f:
                st.download_button("📥 下載目前系統對照表", data=f, file_name="目前系統_mapping.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        st.divider()
        st.subheader("🔄 操作說明 2：重新上傳 (整份覆蓋)")
        reupload_map = st.file_uploader("上傳整理好的【匹配對照表】Excel", type=["xlsx"], key="reupload")
        if reupload_map and st.button("🚨 確認覆蓋目前系統對照表", type="primary"):
            try:
                new_mapping = {sheet: pd.read_excel(reupload_map, sheet_name=sheet) for sheet in SHEET_NAMES}
                save_mapping(new_mapping)
                load_mapping.clear() # 🚀 清除快取
                st.success("✅ 對照表已更新！請重新整理網頁。")
                st.rerun()
            except Exception as e:
                st.error(f"上傳失敗：請確認你的 Excel 檔案是否包含這四個分頁。詳細錯誤：{e}")
