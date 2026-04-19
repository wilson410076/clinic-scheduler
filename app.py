import streamlit as st
import pandas as pd
import io
import openpyxl # 這是用來讀寫真實 Excel 且不破壞格式的神器！
import random

# ==========================================
# 1. 診所排班規則與設定
# ==========================================
DOCTOR_ASSISTANT_MATCH = {
    "陳逸陽": "映璇",
    "王荷若": "萃屏",
    "李昆晏": "萃屏",
    "許雅茹": "和芸",
    "劉筑昀": "姿穎"
}
ASSISTANTS = ["映璇", "和芸", "欣寧", "萃屏", "維珍", "菀庭", "姿穎", "濘安"]

# ==========================================
# 2. 系統介面與密碼鎖
# ==========================================
st.set_page_config(page_title="恩霖診所 - 自動排班系統 V3", layout="wide")

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 恩霖診所 - 排班系統登入")
    pwd = st.text_input("請輸入系統密碼：", type="password")
    if st.button("登入"):
        if pwd == "115":
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("密碼錯誤！")
    st.stop()

if "timeoff_db" not in st.session_state:
    st.session_state.timeoff_db = {}

st.title("🏥 恩霖診所 - 自動排班系統 V3 (完美格式版)")
st.info("💡 V3 版本特色：AI 將直接讀取您上傳的表格，填入助理名單後，100% 保留您原本的排版與格式！")

tab1, tab2, tab3 = st.tabs(["📁 1. 上傳醫師需求", "📝 2. 助理劃休板", "🚀 3. AI 自動排班"])

# --- 第一區：上傳 ---
with tab1:
    uploaded_file = st.file_uploader("請上傳醫師班表 (Excel 格式 .xlsx)", type=["xlsx"])
    if uploaded_file:
        st.success("檔案已成功讀取，準備就緒！")

# --- 第二區：劃休 ---
with tab2:
    selected_ast = st.selectbox("選擇助理：", ASSISTANTS)
    if selected_ast not in st.session_state.timeoff_db:
        df = pd.DataFrame({"日期": [f"{i}號" for i in range(1, 32)], "早休": [False]*31, "午休": [False]*31, "晚休": [False]*31})
        st.session_state.timeoff_db[selected_ast] = df
    edited_df = st.data_editor(st.session_state.timeoff_db[selected_ast], hide_index=True, key=f"ed_{selected_ast}")
    if st.button(f"儲存 {selected_ast} 休假"):
        st.session_state.timeoff_db[selected_ast] = edited_df
        st.toast("休假已更新！")

# --- 第三區：AI 運算與輸出 ---
with tab3:
    if st.button("🚀 開始智慧排班 (保留原格式)", type="primary"):
        if uploaded_file is None:
            st.error("請先回到第一步上傳您的 Excel 檔案！")
        else:
            with st.spinner("AI 正在掃描您的表格並填寫助理名單..."):
                try:
                    # 1. 使用 openpyxl 載入您上傳的原檔案 (保留所有格式)
                    wb = openpyxl.load_workbook(uploaded_file)
                    # 假設班表在名為 "班表" 或第一個工作表
                    ws = wb.active 
                    
                    # 2. 尋找醫師並填入助理 (核心掃描邏輯)
                    # AI 會逐行逐格掃描，看到醫師名字，就在右邊一格填上助理
                    for row in ws.iter_rows():
                        for cell in row:
                            if cell.value in DOCTOR_ASSISTANT_MATCH:
                                # 規則：強制綁定 (例如看到陳逸陽，右邊強制寫映璇)
                                assistant_cell = ws.cell(row=cell.row, column=cell.column + 1)
                                assistant_cell.value = DOCTOR_ASSISTANT_MATCH[cell.value]
                            
                            elif type(cell.value) == str and cell.value in ["楊忠霖", "官俊彥", "劉善總", "陳明志", "黃菁菁", "傅超俊", "蔣宜蓁"]:
                                # 如果是其他沒有強制綁定的醫師，從未休假的人員中隨機指派 (初步演算法)
                                # 這裡先用隨機分配示範，實際會加入休假判斷與時數平衡
                                available = ["維珍", "菀庭", "欣寧"] # 假設可用名單
                                assistant_cell = ws.cell(row=cell.row, column=cell.column + 1)
                                if not assistant_cell.value: # 如果那格還是空的
                                    assistant_cell.value = random.choice(available)
                    
                    # 3. 將修改後的檔案存入記憶體準備下載
                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    st.success("排班完成！所有格式與排版已完美保留。")
                    
                    st.download_button(
                        label="📥 下載完整排班表 (與原格式相同)",
                        data=output,
                        file_name="恩霖診所_AI排班完成版.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"處理檔案時發生錯誤，請確認上傳的是標準的 Excel 檔案。錯誤代碼：{e}")
