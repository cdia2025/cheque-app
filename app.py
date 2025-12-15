import streamlit as st
import pandas as pd
import gspread
from streamlit_gsheets import GSheetsConnection
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import io
import time

# ================= 設定區 =================
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1gpq9Cye25rmPgyOt508L1sBvlIpPis45R09vn0uy434/edit"

SYSTEM_COLS = ['Collected', 'DocGeneratedDate', 'CollectedDate', 'ResponsibleStaff']
REQUIRED_COLS = ['ID序號', '編號', '姓名(中文)', '姓名(英文)', '電話', '實習日數', '反思會', '反思表', '家長/監護人']

st.set_page_config(page_title="雲端實習津貼系統", layout="wide", page_icon="☁️")

# ================= 初始化 Session State =================
if 'df_main' not in st.session_state: st.session_state.df_main = None
if 'current_sheet' not in st.session_state: st.session_state.current_sheet = None

# 用於控制刪除確認框的顯示狀態
if 'show_del_confirm' not in st.session_state: st.session_state.show_del_confirm = False
if 'show_clear_confirm' not in st.session_state: st.session_state.show_clear_confirm = False
if 'show_sheet_confirm' not in st.session_state: st.session_state.show_sheet_confirm = False

# ================= 連線設定 =================
@st.cache_resource
def get_write_client():
    try:
        creds_dict = dict(st.secrets["connections"]["gsheets"])
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        client = gspread.service_account_from_dict(creds_dict)
        return client
    except Exception as e:
        st.error(f"連線設定錯誤: {e}")
        st.stop()

conn = st.connection("gsheets", type=GSheetsConnection)

# ================= 核心函式 =================
def fetch_data_from_cloud(sheet_name):
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, ttl=0)
        if not df.empty:
            df.columns = df.columns.str.strip()
            if 'ID序號' in df.columns:
                df['ID序號'] = df['ID序號'].astype(str)
            else:
                if len(df.columns) > 0:
                    df.rename(columns={df.columns[0]: 'ID序號'}, inplace=True)
                    df['ID序號'] = df['ID序號'].astype(str)
            for col in SYSTEM_COLS:
                if col not in df.columns: df[col] = ''
            df = df.fillna('')
        else:
            df = pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)
        return df
    except:
        return pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)

# ================= 主程式 =================
st.title("☁️ 實習津貼管理系統 (V37 刪除修正版)")

# --- 側邊欄 ---
with st.sidebar:
    st.header("🎛️ 設定面板")
    staff_name = st.text_input("👤 負責職員姓名 (必填)", key="staff_input")
    st.divider()
    
    try:
        gc = get_write_client()
        sh = gc.open_by_url(SPREADSHEET_URL)
        sheet_names = [ws.title for ws in sh.worksheets()]
        
        idx = 0
        if st.session_state.current_sheet in sheet_names:
            idx = sheet_names.index(st.session_state.current_sheet)
        selected_sheet_name = st.selectbox("📂 選擇工作表", sheet_names, index=idx)
    except Exception as e:
        st.error(f"連線失敗: {e}")
        st.stop()

    if st.button("🔄 重新整理資料"):
        st.cache_data.clear()
        st.session_state.df_main = fetch_data_from_cloud(selected_sheet_name)
        st.session_state.current_sheet = selected_sheet_name
        # 切換或重整時，務必關閉所有確認框
        st.session_state.show_del_confirm = False
        st.session_state.show_clear_confirm = False
        st.session_state.show_sheet_confirm = False
        st.rerun()

    if st.session_state.df_main is None or st.session_state.current_sheet != selected_sheet_name:
        with st.spinner(f"讀取中..."):
            st.session_state.df_main = fetch_data_from_cloud(selected_sheet_name)
            st.session_state.current_sheet = selected_sheet_name

if not staff_name:
    st.warning("⚠️ 請先在左側輸入您的姓名。")
    st.stop()

df = st.session_state.df_main

try:
    worksheet = sh.worksheet(selected_sheet_name)
except:
    st.warning("工作表讀取中...")
    st.stop()

# ================= 分頁 =================
tab_upload, tab_prepare, tab_confirm, tab_manage = st.tabs([
    "📥 建立新表", "📄 [1] 匯出", "✅ [2] 領取", "🛠️ 刪除管理"
])

# ---------------- Tab 1: 建立新表 ----------------
with tab_upload:
    st.subheader("📥 上傳 Excel 並建立獨立工作表")
    uploaded_file = st.file_uploader("選擇 Excel 檔案", type=['xlsx', 'xls'], key="upl")
    new_sheet_name = st.text_input("輸入新工作表名稱", placeholder="2024_第一期", key="new_s_in")
    
    if st.button("🚀 建立並上傳", type="primary"):
        if uploaded_file and new_sheet_name:
            if new_sheet_name in sheet_names:
                st.error("名稱已存在！")
            else:
                try:
                    new_df = pd.read_excel(uploaded_file)
                    if len(new_df.columns) >= 9:
                        mapping = {
                            new_df.columns[0]: 'ID序號', new_df.columns[1]: '編號',
                            new_df.columns[2]: '姓名(中文)', new_df.columns[3]: '姓名(英文)',
                            new_df.columns[4]: '電話', new_df.columns[5]: '實習日數',
                            new_df.columns[6]: '反思會', new_df.columns[7]: '反思表',
                            new_df.columns[8]: '家長/監護人'
                        }
                        new_df.rename(columns=mapping, inplace=True)
                        new_df = new_df[REQUIRED_COLS]
                        for col in SYSTEM_COLS: new_df[col] = ''
                        new_df['ID序號'] = new_df['ID序號'].astype(str)
                        new_df = new_df.fillna('')
                        
                        with st.spinner("建立中..."):
                            new_ws = sh.add_worksheet(title=new_sheet_name, rows=len(new_df)+50, cols=20)
                            new_ws.update([new_df.columns.tolist()] + new_df.values.tolist())
                            st.success(f"成功建立「{new_sheet_name}」！")
                            time.sleep(1)
                            st.cache_data.clear()
                            st.rerun()
                    else: st.error("欄位不足")
                except Exception as e: st.error(f"錯誤: {e}")
        else: st.error("請填寫名稱並選擇檔案")

# ---------------- Tab 2: 準備匯出 ----------------
with tab_prepare:
    st.subheader(f"📄 準備匯出 ({selected_sheet_name})")
    if '反思會' in df.columns:
        mask_ready = ((df['反思會'].astype(str).str.upper() == 'Y') & (df['反思表'].astype(str).str.upper() == 'Y') & (df['DocGeneratedDate'] == ''))
        df_ready = df[mask_ready].copy()
        df_ready.insert(0, "選取", False)
        edited_ready = st.data_editor(df_ready, column_config={"選取": st.column_config.CheckboxColumn(required=True)}, disabled=[c for c in df.columns if c!="選取"], hide_index=True, key="ed_ready")
        
        if st.button("📤 匯出 & 更新狀態", type="primary"):
            sel = edited_ready[edited_ready["選取"]==True]
            if not sel.empty:
                today = datetime.now().strftime("%Y-%m-%d")
                head = worksheet.row_values(1)
                try:
                    c_doc = head.index('DocGeneratedDate')+1
                    c_staff = head.index('ResponsibleStaff')+1
                    prog = st.progress(0)
                    ex_list = []
                    for i, (idx, row) in enumerate(sel.iterrows()):
                        try:
                            cell = worksheet.find(row['ID序號'], in_column=1)
                            if cell:
                                worksheet.update_cell(cell.row, c_doc, today)
                                worksheet.update_cell(cell.row, c_staff, staff_name)
                                rec = row.to_dict(); del rec['選取']; rec.update({'StaffName':staff_name, 'TodayDate':today})
                                ex_list.append(rec)
                                st.session_state.df_main.loc[df['ID序號']==row['ID序號'], ['DocGeneratedDate','ResponsibleStaff']] = [today, staff_name]
                        except: pass
                        prog.progress((i+1)/len(sel))
                    
                    if ex_list:
                        out = io.BytesIO()
                        pd.DataFrame(ex_list).to_excel(out, index=False)
                        st.download_button("📥 下載 MailMerge Source", out.getvalue(), "MailMerge_Source.xlsx")
                        st.success("完成！")
                        time.sleep(1)
                        st.rerun()
                except: st.error("雲端欄位對應錯誤，請檢查標題列")

# ---------------- Tab 3: 確認領取 ----------------
with tab_confirm:
    st.subheader(f"✅ 確認領取 ({selected_sheet_name})")
    if 'Collected' in df.columns:
        mask_conf = ((df['DocGeneratedDate']!='') & (df['Collected']!='Y'))
        df_conf = df[mask_conf].copy()
        df_conf.insert(0, "確認", False)
        ed_conf = st.data_editor(df_conf, column_config={"確認": st.column_config.CheckboxColumn(required=True)}, disabled=[c for c in df.columns if c!="確認"], hide_index=True, key="ed_conf")
        
        if st.button("✅ 確認已取票", type="primary"):
            sel = ed_conf[ed_conf["確認"]==True]
            if not sel.empty:
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                head = worksheet.row_values(1)
                try:
                    c_col = head.index('Collected')+1
                    c_date = head.index('CollectedDate')+1
                    prog = st.progress(0)
                    for i, (idx, row) in enumerate(sel.iterrows()):
                        try:
                            cell = worksheet.find(row['ID序號'], in_column=1)
                            if cell:
                                worksheet.update_cell(cell.row, c_col, 'Y')
                                worksheet.update_cell(cell.row, c_date, now)
                                st.session_state.df_main.loc[df['ID序號']==row['ID序號'], ['Collected','CollectedDate']] = ['Y', now]
                        except: pass
                        prog.progress((i+1)/len(sel))
                    st.success("更新完成！")
                    st.rerun()
                except: st.error("雲端欄位對應錯誤")

# ---------------- Tab 4: 刪除管理 (關鍵修復區域) ----------------
with tab_manage:
    st.subheader(f"🛠️ 資料管理 - {selected_sheet_name}")
    st.error("⚠️ 危險操作區：請謹慎使用")
    
    # 資料選取區
    df_del = df.copy()
    df_del.insert(0, "刪除", False)
    ed_del = st.data_editor(df_del, column_config={"刪除": st.column_config.CheckboxColumn(required=True, label="選取")}, hide_index=True, key="ed_del")
    
    st.divider()
    
    # 這裡分為三個獨立區塊
    c1, c2, c3 = st.columns(3)
    
    # === 功能 1: 刪除選取列 (修復版) ===
    with c1:
        st.markdown("##### 🗑️ 刪除選取的列")
        
        # 按鈕 1：觸發確認框
        if st.button("請求刪除選取資料"):
            # 檢查是否有勾選
            sel_rows = ed_del[ed_del["刪除"]==True]
            if sel_rows.empty:
                st.toast("請先勾選上方的資料！", icon="⚠️")
            else:
                st.session_state.show_del_confirm = True
                st.rerun() # 強制刷新以顯示下方確認框
        
        # 確認框 (只有當 flag 為 True 時才顯示)
        if st.session_state.show_del_confirm:
            st.error("確定要刪除勾選的資料嗎？此動作無法復原。")
            
            col_confirm_1, col_cancel_1 = st.columns(2)
            with col_confirm_1:
                # 按鈕 2：真正執行
                if st.button("🔴 確認刪除", key="btn_confirm_del"):
                    with st.spinner("正在刪除..."):
                        # 重新取得勾選名單 (因為 rerun 後 ed_del 狀態還在)
                        sel_rows = ed_del[ed_del["刪除"]==True]
                        
                        rows_to_del = []
                        for idx, row in sel_rows.iterrows():
                            try:
                                cell = worksheet.find(row['ID序號'], in_column=1)
                                if cell: rows_to_del.append(cell.row)
                            except: pass
                        
                        # 倒序刪除
                        rows_to_del.sort(reverse=True)
                        for r in rows_to_del:
                            worksheet.delete_rows(r)
                        
                        st.success(f"已刪除 {len(rows_to_del)} 筆資料")
                        st.session_state.show_del_confirm = False
                        time.sleep(1)
                        st.cache_data.clear()
                        st.rerun()
            
            with col_cancel_1:
                if st.button("取消", key="btn_cancel_del"):
                    st.session_state.show_del_confirm = False
                    st.rerun()

    # === 功能 2: 清空整表 (修復版) ===
    with c2:
        st.markdown("##### 🧹 清空內容 (留標題)")
        
        if st.button("請求清空內容"):
            st.session_state.show_clear_confirm = True
            st.rerun()
            
        if st.session_state.show_clear_confirm:
            st.error("確定清空整張表？只會保留標題列。")
            
            col_confirm_2, col_cancel_2 = st.columns(2)
            with col_confirm_2:
                if st.button("🔴 確認清空", key="btn_confirm_clear"):
                    with st.spinner("清空中..."):
                        # 1. 取得目前的標題 (從 DataFrame 拿最保險)
                        headers = df.columns.tolist()
                        
                        # 2. 清空
                        worksheet.clear()
                        
                        # 3. 寫回標題
                        worksheet.update(range_name='A1', values=[headers])
                        
                        st.success("已清空內容！")
                        st.session_state.show_clear_confirm = False
                        time.sleep(1)
                        st.cache_data.clear()
                        st.rerun()
            
            with col_cancel_2:
                if st.button("取消", key="btn_cancel_clear"):
                    st.session_state.show_clear_confirm = False
                    st.rerun()

    # === 功能 3: 刪除工作表 ===
    with c3:
        st.markdown("##### 🔥 刪除本工作表")
        
        if st.button("請求刪除工作表"):
            if len(sheet_names) <= 1:
                st.error("這是最後一張表，無法刪除。")
            else:
                st.session_state.show_sheet_confirm = True
                st.rerun()
        
        if st.session_state.show_sheet_confirm:
            st.error(f"確定永久刪除「{selected_sheet_name}」？")
            
            col_confirm_3, col_cancel_3 = st.columns(2)
            with col_confirm_3:
                if st.button("🔴 確認刪除 Sheet", key="btn_confirm_sheet"):
                    with st.spinner("刪除中..."):
                        sh.del_worksheet(worksheet)
                        st.success("工作表已刪除")
                        st.session_state.show_sheet_confirm = False
                        st.session_state.current_sheet = None
                        time.sleep(1)
                        st.cache_data.clear()
                        st.rerun()
            
            with col_cancel_3:
                if st.button("取消", key="btn_cancel_sheet"):
                    st.session_state.show_sheet_confirm = False
                    st.rerun()
