import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime
import io
import time

# ================= 設定區 =================
# 請確認這裡是您的 Google Sheet 網址
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1gpq9Cye25rmPgyOt508L1sBvlIpPis45R09vn0uy434/edit"

# 固定欄位順序 (確保資料結構一致)
REQUIRED_COLS = ['ID序號', '編號', '姓名(中文)', '姓名(英文)', '電話', '實習日數', '反思會', '反思表', '家長/監護人', 
                 'Collected', 'DocGeneratedDate', 'CollectedDate', 'ResponsibleStaff']

st.set_page_config(page_title="雲端實習津貼系統 (V51 重製穩定版)", layout="wide", page_icon="☁️")

# 建立連線 (使用 Reference App 的方式)
conn = st.connection("gsheets", type=GSheetsConnection)

# ================= 資料處理函式 (參考 Reference App) =================

def clean_dataframe(df):
    """
    清洗資料：轉字串、補空值、統一格式。
    這是穩定性的關鍵。
    """
    # 1. 確保所有欄位都存在
    for col in REQUIRED_COLS:
        if col not in df.columns:
            df[col] = ""
            
    # 2. 只保留需要的欄位，並照順序排好
    df = df[REQUIRED_COLS]
    
    # 3. 轉為字串並處理空值
    df = df.astype(str)
    for col in df.columns:
        df[col] = df[col].replace(['NaT', 'nan', 'None', '<NA>'], '')
        df[col] = df[col].str.strip() # 去除前後空白
        
    # 4. 特殊處理 ID (去除 .0)
    df['ID序號'] = df['ID序號'].apply(lambda x: x[:-2] if x.endswith('.0') else x)
    
    return df

def get_all_sheet_names():
    """取得所有工作表名稱"""
    try:
        # 透過底層 client 獲取
        return [ws.title for ws in conn.client.open_by_url(SPREADSHEET_URL).worksheets()]
    except Exception as e:
        st.error(f"連線錯誤: {e}")
        return []

def load_data(sheet_name):
    """讀取資料"""
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, ttl=0)
        return clean_dataframe(df)
    except:
        return pd.DataFrame(columns=REQUIRED_COLS)

def save_data(df, sheet_name):
    """將整張表寫回 Google Sheet (全覆蓋模式)"""
    try:
        clean_df = clean_dataframe(df)
        conn.update(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, data=clean_df)
        st.toast("✅ 資料已同步至雲端！", icon="☁️")
        return clean_df
    except Exception as e:
        if "429" in str(e):
            st.error("⚠️ 流量過大，請等待 1 分鐘後再試。")
        else:
            st.error(f"儲存失敗: {e}")
        return df

# ================= Session State 初始化 =================
if 'current_sheet' not in st.session_state: st.session_state.current_sheet = None
if 'df_main' not in st.session_state: st.session_state.df_main = None
if 'download_btn_key' not in st.session_state: st.session_state.download_btn_key = 0

# ================= 側邊欄 =================
with st.sidebar:
    st.header("🎛️ 控制台")
    staff_name = st.text_input("👤 負責職員姓名", key="staff_name")
    
    st.divider()
    
    # 1. 取得工作表清單
    sheet_names = get_all_sheet_names()
    if not sheet_names:
        st.stop()
        
    # 2. 選擇工作表 (鎖定 Index 防止跳頁)
    if st.session_state.current_sheet not in sheet_names:
        st.session_state.current_sheet = sheet_names[0]
        
    idx = sheet_names.index(st.session_state.current_sheet)
    selected_sheet = st.selectbox("📂 選擇工作表", sheet_names, index=idx)
    
    # 切換時重新讀取
    if selected_sheet != st.session_state.current_sheet:
        st.session_state.current_sheet = selected_sheet
        st.session_state.df_main = load_data(selected_sheet)
        st.rerun()

    if st.button("🔄 強制重新整理"):
        st.session_state.df_main = load_data(selected_sheet)
        st.rerun()

# 檢查登入
if not staff_name:
    st.warning("⚠️ 請先在左側輸入姓名。")
    st.stop()

# 確保有資料
if st.session_state.df_main is None:
    st.session_state.df_main = load_data(selected_sheet)

df = st.session_state.df_main

st.title(f"☁️ {selected_sheet}")

# ================= 主分頁 =================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📥 建立/上傳", 
    "📄 [1] 準備匯出", 
    "🔵 [2] 待領取", 
    "🟢 [3] 已取票", 
    "🚫 [4] 不符"
])

# ---------------- TAB 1: 建立新表 ----------------
with tab1:
    st.subheader("上傳 Excel 並建立新分頁")
    up_file = st.file_uploader("選擇 Excel", type=["xlsx", "xls"])
    new_name = st.text_input("新工作表名稱 (例如: 2024_Batch2)")
    
    if st.button("🚀 建立並上傳", type="primary"):
        if up_file and new_name:
            if new_name in sheet_names:
                st.error("名稱重複！")
            else:
                try:
                    new_df = pd.read_excel(up_file)
                    # 簡單欄位對應
                    if len(new_df.columns) >= 9:
                        mapping = {
                            new_df.columns[0]: 'ID序號', new_df.columns[1]: '編號',
                            new_df.columns[2]: '姓名(中文)', new_df.columns[3]: '姓名(英文)',
                            new_df.columns[4]: '電話', new_df.columns[5]: '實習日數',
                            new_df.columns[6]: '反思會', new_df.columns[7]: '反思表',
                            new_df.columns[8]: '家長/監護人'
                        }
                        new_df.rename(columns=mapping, inplace=True)
                        # 補齊系統欄位
                        for c in ['Collected', 'DocGeneratedDate', 'CollectedDate', 'ResponsibleStaff']:
                            new_df[c] = ""
                        
                        # 使用 gspread 建立新表
                        sh = conn.client.open_by_url(SPREADSHEET_URL)
                        ws = sh.add_worksheet(title=new_name, rows=len(new_df)+20, cols=15)
                        
                        # 清洗並寫入
                        clean_new = clean_dataframe(new_df)
                        conn.update(worksheet=new_name, data=clean_new)
                        
                        st.success(f"建立成功！")
                        st.session_state.current_sheet = new_name
                        time.sleep(2)
                        st.rerun()
                    else:
                        st.error("欄位不足")
                except Exception as e:
                    st.error(f"錯誤: {e}")

# ---------------- TAB 2: 準備匯出 ----------------
with tab2:
    st.subheader("步驟一：匯出資料")
    
    # 篩選邏輯：雙Y 且 無日期
    mask = (df['反思會'].str.upper() == 'Y') & (df['反思表'].str.upper() == 'Y') & (df['DocGeneratedDate'] == '')
    df_show = df[mask].copy()
    
    # 顯示編輯器
    df_show.insert(0, "選取", False)
    edited = st.data_editor(
        df_show, 
        column_config={"選取": st.column_config.CheckboxColumn(required=True)},
        disabled=[c for c in df_show.columns if c != "選取"],
        hide_index=True,
        key="editor_tab2"
    )
    
    if st.button("📤 匯出選取資料 & 更新狀態"):
        selected = edited[edited["選取"]]
        if selected.empty:
            st.warning("未選取")
        else:
            today = datetime.now().strftime("%Y-%m-%d")
            ids_to_update = selected['ID序號'].tolist()
            
            # --- 核心：在記憶體中更新 DataFrame ---
            # 直接用 Pandas 操作，保證不會找不到 ID
            df.loc[df['ID序號'].isin(ids_to_update), 'DocGeneratedDate'] = today
            df.loc[df['ID序號'].isin(ids_to_update), 'ResponsibleStaff'] = staff_name
            
            # --- 核心：整表寫回 ---
            save_data(df, selected_sheet)
            
            # 準備下載檔
            out_df = selected.drop(columns=['選取'])
            out_df['StaffName'] = staff_name
            out_df['TodayDate'] = today
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                out_df.to_excel(writer, index=False)
            
            # 存入 Session State 以顯示下載按鈕
            st.session_state.dl_data = buffer.getvalue()
            st.session_state.dl_name = "MailMerge_Source.xlsx"
            st.rerun()

    # 顯示下載按鈕 (持續存在直到切換)
    if 'dl_data' in st.session_state:
        st.success("✅ 更新成功！請下載檔案：")
        st.download_button(
            label="📥 下載 MailMerge Source",
            data=st.session_state.dl_data,
            file_name=st.session_state.dl_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

# ---------------- TAB 3: 待領取 ----------------
with tab3:
    st.subheader("步驟二：準備領取")
    
    # 篩選：有日期 且 未領取
    mask = (df['DocGeneratedDate'] != '') & (df['Collected'] != 'Y')
    df_show = df[mask].copy()
    
    df_show.insert(0, "確認", False)
    edited = st.data_editor(
        df_show, 
        column_config={"確認": st.column_config.CheckboxColumn(required=True)},
        disabled=[c for c in df_show.columns if c != "確認"],
        hide_index=True,
        key="editor_tab3"
    )
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ 確認已取票", type="primary"):
            ids = edited[edited["確認"]]['ID序號'].tolist()
            if ids:
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                # Pandas 更新
                df.loc[df['ID序號'].isin(ids), 'Collected'] = 'Y'
                df.loc[df['ID序號'].isin(ids), 'CollectedDate'] = now
                save_data(df, selected_sheet)
                st.rerun()
                
    with c2:
        if st.button("↩️ 退回至準備匯出"):
            ids = edited[edited["確認"]]['ID序號'].tolist()
            if ids:
                if st.checkbox("確定退回？"):
                    # Pandas 更新 (清空日期)
                    df.loc[df['ID序號'].isin(ids), 'DocGeneratedDate'] = ''
                    df.loc[df['ID序號'].isin(ids), 'ResponsibleStaff'] = ''
                    save_data(df, selected_sheet)
                    st.success("已退回")
                    st.rerun()

# ---------------- TAB 4: 已取票 ----------------
with tab4:
    st.subheader("已取票紀錄")
    mask = (df['Collected'] == 'Y')
    df_show = df[mask].copy()
    
    df_show.insert(0, "撤銷", False)
    edited = st.data_editor(
        df_show, 
        column_config={"撤銷": st.column_config.CheckboxColumn(required=True)},
        disabled=[c for c in df_show.columns if c != "撤銷"],
        hide_index=True,
        key="editor_tab4"
    )
    
    if st.button("↩️ 撤銷領取 (回到待領取)"):
        ids = edited[edited["撤銷"]]['ID序號'].tolist()
        if ids:
            if st.checkbox("確定撤銷？"):
                df.loc[df['ID序號'].isin(ids), 'Collected'] = ''
                df.loc[df['ID序號'].isin(ids), 'CollectedDate'] = ''
                save_data(df, selected_sheet)
                st.success("已撤銷")
                st.rerun()

# ---------------- TAB 5: 不符名單 ----------------
with tab5:
    st.subheader("不符資格名單")
    # 篩選：任一條件非Y 且 未處理
    mask = ((df['反思會'].str.upper() != 'Y') | (df['反思表'].str.upper() != 'Y')) & (df['DocGeneratedDate'] == '')
    df_show = df[mask].copy()
    
    df_show.insert(0, "放行", False)
    edited = st.data_editor(
        df_show, 
        column_config={"放行": st.column_config.CheckboxColumn(required=True)},
        disabled=[c for c in df_show.columns if c != "放行"],
        hide_index=True,
        key="editor_tab5"
    )
    
    if st.button("➡️ 強制放行 (改為 Y)"):
        ids = edited[edited["放行"]]['ID序號'].tolist()
        if ids:
            if st.checkbox("確認強制修改？"):
                df.loc[df['ID序號'].isin(ids), '反思會'] = 'Y'
                df.loc[df['ID序號'].isin(ids), '反思表'] = 'Y'
                save_data(df, selected_sheet)
                st.success("已放行，請至 [1] 匯出")
                st.rerun()
