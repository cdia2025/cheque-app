import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import io
import time

# ================= 設定區 =================
# 您的 Google Sheet 網址 (也可以設定在 secrets.toml 中自動讀取)
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/您的_GOOGLE_SHEET_ID/edit"

# 系統與必要欄位
SYSTEM_COLS = ['Collected', 'DocGeneratedDate', 'CollectedDate', 'ResponsibleStaff']
REQUIRED_COLS = ['ID序號', '編號', '姓名(中文)', '姓名(英文)', '電話', '實習日數', '反思會', '反思表', '家長/監護人']

st.set_page_config(page_title="雲端實習津貼系統", layout="wide", page_icon="☁️")

# ================= 連線設定 (使用 st-gsheets-connection) =================
# 建立連線物件
conn = st.connection("gsheets", type=GSheetsConnection)

# ================= 介面開始 =================
st.title("☁️ 實習津貼管理系統 (GSheets Connection 版)")

# --- 側邊欄 ---
with st.sidebar:
    st.header("🎛️ 設定面板")
    staff_name = st.text_input("👤 負責職員姓名 (必填)", key="staff_input")
    
    st.divider()
    
    # 取得工作表列表 (使用底層 gspread client)
    try:
        # conn.client 就是底層的 gspread client
        sh = conn.client.open_by_url(SPREADSHEET_URL)
        sheet_names = [ws.title for ws in sh.worksheets()]
        selected_sheet_name = st.selectbox("📂 選擇工作表", sheet_names)
    except Exception as e:
        st.error(f"連線失敗，請檢查 secrets 設定。\n錯誤: {e}")
        st.stop()

    if st.button("🔄 重新整理資料"):
        st.cache_data.clear()
        st.rerun()

if not staff_name:
    st.warning("⚠️ 請先在左側輸入您的姓名才能開始操作。")
    st.stop()

# --- 讀取資料 ---
try:
    # 使用 conn.read() 快速讀取資料為 DataFrame
    # ttl=0 代表不快取，每次都抓最新資料 (避免多人操作時看到舊資料)
    df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=selected_sheet_name, ttl=0)
    
    # 資料清理：確保 ID 為字串，並補齊欄位
    if not df.empty:
        # 強制轉字串避免 ID 變成數字
        df['ID序號'] = df['ID序號'].astype(str)
        # 補齊系統欄位
        for col in SYSTEM_COLS:
            if col not in df.columns:
                df[col] = ''
        # 補齊空值
        df = df.fillna('')
    else:
        df = pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)

except Exception as e:
    st.error(f"讀取資料失敗: {e}")
    st.stop()

# 取得底層 worksheet 物件 (用於精確寫入)
try:
    worksheet = sh.worksheet(selected_sheet_name)
except:
    st.error("無法取得工作表物件")
    st.stop()

# ================= 分頁功能 =================
tab_upload, tab_prepare, tab_confirm, tab_history = st.tabs([
    "📥 上載新資料", 
    "📄 [1] 準備匯出 (Mail Merge)", 
    "✅ [2] 確認領取", 
    "📜 資料總覽"
])

# -------------------------------------------
# TAB 1: 上載新資料
# -------------------------------------------
with tab_upload:
    st.subheader("📥 上傳 Excel 並附加到目前工作表")
    uploaded_file = st.file_uploader("選擇 Excel 檔案", type=['xlsx', 'xls'])
    
    if uploaded_file:
        try:
            new_df = pd.read_excel(uploaded_file)
            # 簡單檢查欄位數量
            if len(new_df.columns) >= 9:
                # 欄位對應 (假設順序固定)
                mapping = {
                    new_df.columns[0]: 'ID序號', new_df.columns[1]: '編號',
                    new_df.columns[2]: '姓名(中文)', new_df.columns[3]: '姓名(英文)',
                    new_df.columns[4]: '電話', new_df.columns[5]: '實習日數',
                    new_df.columns[6]: '反思會', new_df.columns[7]: '反思表',
                    new_df.columns[8]: '家長/監護人'
                }
                new_df.rename(columns=mapping, inplace=True)
                new_df = new_df[REQUIRED_COLS] # 只取需要的欄位
                
                # 補上系統欄位
                for col in SYSTEM_COLS: new_df[col] = ''
                new_df['ID序號'] = new_df['ID序號'].astype(str)
                
                st.write("預覽:", new_df.head())
                
                if st.button("🚀 確認上傳"):
                    with st.spinner("寫入中..."):
                        # 使用底層方法 append_rows
                        worksheet.append_rows(new_df.values.tolist())
                        st.success("成功新增資料！")
                        time.sleep(1)
                        st.rerun()
            else:
                st.error("欄位不足 9 欄")
        except Exception as e:
            st.error(f"錯誤: {e}")

# -------------------------------------------
# TAB 2: 準備匯出
# -------------------------------------------
with tab_prepare:
    st.subheader("📄 步驟一：匯出 Mail Merge 資料")
    
    # 篩選邏輯
    mask_ready = (
        (df['反思會'].astype(str).str.upper() == 'Y') & 
        (df['反思表'].astype(str).str.upper() == 'Y') & 
        (df['DocGeneratedDate'] == '')
    )
    df_ready = df[mask_ready].copy()
    
    df_ready.insert(0, "選取", False)
    edited_ready = st.data_editor(
        df_ready,
        column_config={"選取": st.column_config.CheckboxColumn(required=True)},
        disabled=[c for c in df.columns if c != "選取"],
        hide_index=True,
        key="editor_ready"
    )
    
    if st.button("📤 匯出 & 更新狀態", type="primary"):
        selected = edited_ready[edited_ready["選取"] == True]
        if selected.empty:
            st.warning("未選取")
        else:
            today = datetime.now().strftime("%Y-%m-%d")
            
            # 取得欄位 Index
            header = worksheet.row_values(1)
            try:
                col_doc_idx = header.index('DocGeneratedDate') + 1
                col_staff_idx = header.index('ResponsibleStaff') + 1
            except:
                st.error("雲端表格缺少系統欄位")
                st.stop()

            progress_bar = st.progress(0)
            export_list = []
            
            for i, (idx, row) in enumerate(selected.iterrows()):
                target_id = row['ID序號']
                try:
                    cell = worksheet.find(target_id, in_column=1)
                    if cell:
                        worksheet.update_cell(cell.row, col_doc_idx, today)
                        worksheet.update_cell(cell.row, col_staff_idx, staff_name)
                        
                        rec = row.to_dict()
                        del rec['選取']
                        rec['StaffName'] = staff_name
                        rec['TodayDate'] = today
                        export_list.append(rec)
                except:
                    pass # 若找不到ID則跳過
                progress_bar.progress((i + 1) / len(selected))
            
            if export_list:
                # 產生 Excel 下載
                out_df = pd.DataFrame(export_list)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    out_df.to_excel(writer, index=False)
                
                st.success(f"已更新 {len(export_list)} 筆！")
                st.download_button(
                    label="📥 下載 MailMerge_Source.xlsx",
                    data=buffer.getvalue(),
                    file_name="MailMerge_Source.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                time.sleep(2)
                st.rerun()

# -------------------------------------------
# TAB 3: 確認領取
# -------------------------------------------
with tab_confirm:
    st.subheader("✅ 步驟二：確認領取")
    
    mask_confirm = (
        (df['DocGeneratedDate'] != '') & 
        (df['Collected'] != 'Y')
    )
    df_confirm = df[mask_confirm].copy()
    
    df_confirm.insert(0, "確認", False)
    edited_confirm = st.data_editor(
        df_confirm,
        column_config={"確認": st.column_config.CheckboxColumn(required=True)},
        disabled=[c for c in df.columns if c != "確認"],
        hide_index=True,
        key="editor_confirm"
    )
    
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("✅ 確認已取票", type="primary"):
            selected = edited_confirm[edited_confirm["確認"] == True]
            if selected.empty:
                st.warning("未選取")
            else:
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                header = worksheet.row_values(1)
                col_col_idx = header.index('Collected') + 1
                col_date_idx = header.index('CollectedDate') + 1
                
                prog = st.progress(0)
                for i, (idx, row) in enumerate(selected.iterrows()):
                    cell = worksheet.find(row['ID序號'], in_column=1)
                    if cell:
                        worksheet.update_cell(cell.row, col_col_idx, 'Y')
                        worksheet.update_cell(cell.row, col_date_idx, now_str)
                    prog.progress((i + 1) / len(selected))
                
                st.success("更新完成！")
                time.sleep(1)
                st.rerun()
    
    with col2:
        if st.button("↩️ 退回至準備匯出"):
            selected = edited_confirm[edited_confirm["確認"] == True]
            if not selected.empty:
                header = worksheet.row_values(1)
                col_doc_idx = header.index('DocGeneratedDate') + 1
                col_staff_idx = header.index('ResponsibleStaff') + 1
                for idx, row in selected.iterrows():
                    cell = worksheet.find(row['ID序號'], in_column=1)
                    if cell:
                        worksheet.update_cell(cell.row, col_doc_idx, "")
                        worksheet.update_cell(cell.row, col_staff_idx, "")
                st.success("已退回")
                time.sleep(1)
                st.rerun()

# -------------------------------------------
# TAB 4: 總覽
# -------------------------------------------
with tab_history:
    st.subheader("📜 資料總覽")
    st.dataframe(df)
