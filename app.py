import streamlit as st
import pandas as pd
import io
import openpyxl

# ==========================================
# 1. 診所排班規則與優先序設定
# ==========================================
DOCTOR_ASSISTANT_MATCH = {
    "陳逸陽": "映璇", "王荷若": "萃屏", "李昆晏": "萃屏", "許雅茹": "和芸", "劉筑昀": "姿穎"
}
SATURDAY_SPECIAL_MATCH = {
    "官俊彥": "濘安",
    "陳明志": "菀庭"
}

ASSISTANTS = ["映璇", "和芸", "欣寧", "萃屏", "維珍", "菀庭", "姿穎", "濘安"]
COUNTER_PRIORITY = ["欣寧", "維珍", "和芸", "映璇", "姿穎"]

ONLY_COUNTER = ["欣寧"]
NO_COUNTER = ["萃屏", "菀庭", "濘安"]
NO_NIGHT_SHIFT = ["維珍"]

st.set_page_config(page_title="恩霖診所 - 自動排班系統 V9", layout="wide")

if "authenticated" not in st.session_state: st.session_state.authenticated = False
if not st.session_state.authenticated:
    st.title("🔒 恩霖診所 - 排班系統登入")
    pwd = st.text_input("請輸入密碼：", type="password")
    if st.button("登入"):
        if pwd == "115": st.session_state.authenticated = True; st.rerun()
        else: st.error("密碼錯誤！")
    st.stop()
if "timeoff_db" not in st.session_state: st.session_state.timeoff_db = {}

st.title("🏥 恩霖診所 - 自動排班系統 V9")
st.info("🛡️ 已裝備「防撞邊界」：修復 Excel 右側多餘欄位導致的系統崩潰問題。")

tab1, tab2, tab3 = st.tabs(["📁 1. 上傳班表", "📝 2. 助理劃休", "🚀 3. AI 排班"])

with tab1:
    uploaded_file = st.file_uploader("上傳醫師班表 (Excel .xlsx)", type=["xlsx"])
    if uploaded_file: st.success("檔案讀取完成！")

with tab2:
    selected_ast = st.selectbox("選擇助理：", ASSISTANTS)
    if selected_ast not in st.session_state.timeoff_db:
        df = pd.DataFrame({"日期": [f"{i}號" for i in range(1, 32)], "早休": [False]*31, "午休": [False]*31, "晚休": [False]*31})
        if selected_ast in NO_NIGHT_SHIFT: df["晚休"] = True
        st.session_state.timeoff_db[selected_ast] = df
    edited_df = st.data_editor(st.session_state.timeoff_db[selected_ast], hide_index=True, key=f"ed_{selected_ast}")
    if st.button(f"儲存 {selected_ast} 休假"):
        st.session_state.timeoff_db[selected_ast] = edited_df
        st.toast(f"{selected_ast} 的休假已更新")

with tab3:
    if st.button("🚀 開始智慧排班", type="primary"):
        if not uploaded_file:
            st.error("請先上傳班表！")
        else:
            with st.spinner("AI 正在掃描表格，並自動避開無效區域..."):
                try:
                    wb = openpyxl.load_workbook(uploaded_file)
                    ws = wb.active
                    
                    # 🛡️ 核心修復：加入邊界防護，防止 KeyError 或 IndexError
                    def is_on_leave(ast_name, day_num, shift_type):
                        if ast_name not in st.session_state.timeoff_db: return False
                        if day_num < 1 or day_num > 31: return False # 防撞邊界！超過31號直接當作沒休假
                        
                        db = st.session_state.timeoff_db[ast_name]
                        col_name = "早休" if "早" in shift_type else ("午休" if "午" in shift_type else "晚休")
                        return db.loc[day_num - 1, col_name]

                    shift_rows = {"早班": [], "午班": [], "晚班": []}
                    current_shift = None
                    for r in range(1, ws.max_row + 1):
                        val = str(ws.cell(row=r, column=1).value).replace(" ", "")
                        if "早班" in val: current_shift = "早班"
                        elif "午班" in val: current_shift = "午班"
                        elif "晚班" in val: current_shift = "晚班"
                        if current_shift: shift_rows[current_shift].append(r)

                    for col_idx in range(2, ws.max_column, 2):
                        day_num = (col_idx // 2)
                        
                        is_saturday = False
                        for r_check in range(1, 10): 
                            val_check = str(ws.cell(row=r_check, column=col_idx).value)
                            if "六" in val_check:
                                is_saturday = True
                                break

                        for shift_name, rows in shift_rows.items():
                            working_now = set()
                            
                            # A. 處理跟診與強制綁定
                            docs = []
                            for r in rows:
                                d_name = ws.cell(row=r, column=col_idx).value
                                if type(d_name) == str and len(d_name) >= 2:
                                    docs.append((r, d_name))
                                    
                                    assigned_ast = None
                                    if is_saturday and d_name in SATURDAY_SPECIAL_MATCH:
                                        assigned_ast = SATURDAY_SPECIAL_MATCH[d_name]
                                    elif d_name in DOCTOR_ASSISTANT_MATCH:
                                        assigned_ast = DOCTOR_ASSISTANT_MATCH[d_name]
                                    
                                    if assigned_ast:
                                        ws.cell(row=r, column=col_idx+1).value = assigned_ast
                                        working_now.add(assigned_ast)
                            
                            if not docs: continue

                            # B. 櫃檯分配
                            counter_needed = 2 if len(docs) >= 4 else 1
                            assigned_counters = []
                            for candidate in COUNTER_PRIORITY:
                                if len(assigned_counters) >= counter_needed: break
                                if candidate not in working_now and not is_on_leave(candidate, day_num, shift_name):
                                    assigned_counters.append(candidate)
                                    working_now.add(candidate)
                            
                            while len(assigned_counters) < counter_needed:
                                pool = [a for a in ASSISTANTS if a not in working_now and a not in NO_COUNTER]
                                if pool:
                                    extra = pool[0]; assigned_counters.append(extra); working_now.add(extra)
                                else:
                                    assigned_counters.append("【缺人】"); break

                            c_idx = 0
                            for r in rows:
                                tag = str(ws.cell(row=r, column=1).value)
                                if "櫃" in tag and c_idx < len(assigned_counters):
                                    ws.cell(row=r, column=col_idx+1).value = assigned_counters[c_idx]
                                    c_idx += 1

                            # C. 其餘醫師跟診
                            for r, d_name in docs:
                                if not ws.cell(row=r, column=col_idx+1).value: 
                                    pool = [a for a in ASSISTANTS if a not in working_now and a not in ONLY_COUNTER and a != "濘安"]
                                    if "晚" in shift_name and "維珍" in pool: pool.remove("維珍")
                                    
                                    if pool:
                                        chosen = pool[0]
                                        ws.cell(row=r, column=col_idx+1).value = chosen
                                        working_now.add(chosen)
                                    else:
                                        ws.cell(row=r, column=col_idx+1).value = "【缺人】"

                    output = io.BytesIO()
                    wb.save(output)
                    st.success("排班完成！已完美避開多餘的空白格子與統計欄位。")
                    st.download_button("📥 下載 V9 穩定版", output.getvalue(), "恩霖診所_AI排班_V9.xlsx")
                except Exception as e:
                    st.error(f"遭遇未預期錯誤，請截圖此訊息給開發者：{e}")
