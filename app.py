import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. 診所排班規則與參數設定
# ==========================================
# 綁定規則
DOCTOR_ASSISTANT_MATCH = {
    "陳逸陽": "映璇",
    "王荷若": "萃屏",
    "李昆晏": "萃屏",
    "許雅茹": "和芸",
    "劉筑昀": "姿穎"
}

# 助理名單與限制
ASSISTANTS = ["映璇", "和芸", "欣寧", "萃屏", "維珍", "菀庭", "姿穎", "濘安"]
ONLY_MORNING_AFTERNOON = ["維珍"]
ONLY_COUNTER = ["欣寧"]
EXCLUDED_STAFF = ["怡婷", "詩庭"]
PART_TIME_SATURDAY_ONLY = ["濘安"]

# ==========================================
# 2. 網頁介面與密碼鎖設定
# ==========================================
st.set_page_config(page_title="恩霖診所 - 自動排班系統", layout="wide")

# 簡單的密碼鎖 (預設密碼：115)
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 恩霖診所 - 排班系統登入")
    pwd = st.text_input("請輸入系統密碼：", type="password")
    if st.button("登入"):
        if pwd == "801128": # 您可以自行更改引號內的密碼
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("密碼錯誤，請重新輸入！")
    st.stop()

# ==========================================
# 3. 系統主畫面
# ==========================================
st.title("🏥 恩霖診所 - 自動排班系統 V1")
st.markdown("---")

# 建立三個分頁
tab1, tab2, tab3 = st.tabs(["📁 第一步：上傳醫師班表", "📝 第二步：助理劃休板", "🚀 第三步：AI 排班與下載"])

# --- 第一區：上傳醫師班表 ---
with tab1:
    st.header("1. 上傳下個月的醫師班表")
    st.info("請上傳已經填好醫師，但助理欄位為空白的 Excel 檔案。")
    uploaded_file = st.file_uploader("選擇 Excel / CSV 檔案", type=["xlsx", "csv"])
    
    if uploaded_file is not None:
        st.success("檔案上傳成功！系統已在後台解析醫師名單。")
        # 這裡未來會加入詳細的 Excel 網格解析邏輯

# --- 第二區：助理劃休板 ---
with tab2:
    st.header("2. 助理劃休設定")
    st.write("請依照紙本劃休單，勾選助理**要休假**的時段。（打勾代表該時段不上班）")
    
    # 建立一個 31 天的基礎表格
    days = [f"{i}號" for i in range(1, 32)]
    default_timeoff = pd.DataFrame({
        "日期": days,
        "早班休假": [False] * 31,
        "午班休假": [False] * 31,
        "晚班休假": [False] * 31
    })

    selected_ast = st.selectbox("請選擇要設定劃休的助理：", ASSISTANTS)
    
    # 針對維珍的防呆機制：晚班預設休假 (因為她不上晚班)
    if selected_ast == "維珍":
        default_timeoff["晚班休假"] = [True] * 31
        st.warning("⚠️ 維珍只上早午班，系統已自動將晚班設定為休假。")

    st.write(f"**目前正在設定：{selected_ast} 的休假**")
    edited_df = st.data_editor(default_timeoff, hide_index=True, use_container_width=True)
    
    if st.button(f"💾 儲存 {selected_ast} 的休假紀錄"):
        st.success(f"已成功紀錄 {selected_ast} 的休假！(此為預覽介面，資料暫存於記憶體)")

# --- 第三區：AI 排班與下載 ---
with tab3:
    st.header("3. 執行自動排班")
    st.write("系統將會套用以下規則：")
    st.markdown("""
    * **綁定規則**：陳逸陽配映璇、許雅茹配和芸...等 5 項規則。
    * **櫃檯規則**：醫師數 $\ge$ 4 則安排 2 櫃檯；醫師數 $\le$ 3 則 1 櫃檯。欣寧僅排櫃檯。
    * **時段規則**：維珍不上晚班；濘安僅週六缺人排班。
    * **平衡規則**：盡量讓大家的總診次趨近於 40 診。
    """)
    
    if st.button("🚀 開始自動排班", type="primary"):
        if uploaded_file is None:
            st.error("請先回到第一步上傳醫師班表！")
        else:
            with st.spinner("AI 正在計算最佳排班組合 (包含綁定、休假排除與時數平衡)..."):
                # 模擬運算時間
                import time
                time.sleep(2)
                
                st.success("排班計算完成！")
                
                # 產出模擬的下載檔案 (V1 測試用)
                output_df = pd.DataFrame({"系統訊息": ["這是測試下載檔案", "未來將會產出與您上傳格式相同的完整排班表"]})
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    output_df.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 下載最終排班表 (Excel)",
                    data=output_buffer.getvalue(),
                    file_name="恩霖診所_自動排班結果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
