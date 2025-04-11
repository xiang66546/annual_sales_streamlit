import streamlit as st
import tempfile
import os
from annual_sales import ReportCoordinator

st.set_page_config(page_title="營業報表自動產生系統", page_icon="📊")
st.title("📊 營業報表自動產生系統")
st.write("請依序上傳或填寫以下資料，然後點選「開始產生報表」")

# --- 用戶輸入區 ---

# 年份與月份
year = st.number_input("年份 (例如：114)", min_value=100, max_value=999, value=114)
month = st.number_input("月份 (1~12)", min_value=1, max_value=12, value=3)
company_name = st.text_input("公司名稱", value="斑鳩的窩")

# 各種路徑 (這裡用檔案上傳)
each_area_file = st.file_uploader("請上傳 各區域分店明細 Excel", type=['xlsx'])
path_one_folder = st.text_input("請輸入 日結單 資料夾路徑")
path_two_folder = st.text_input("請輸入 薪資檔 資料夾路徑")
last_year_file = st.file_uploader("請上傳 去年年度損益表 Excel", type=['xlsx'])
this_year_file = st.file_uploader("請上傳 今年年度損益表 Excel", type=['xlsx'])
path_four_file = st.file_uploader("請上傳 預算表 Excel", type=['xlsx'])
path_five_folder = st.text_input("請輸入 月報表 資料夾路徑")
output_folder = st.text_input("請輸入 輸出檔案儲存資料夾路徑", value="/tmp")

# --- 處理上傳的檔案 ---
temp_dir = tempfile.mkdtemp()

def save_uploaded_file(uploaded_file, save_name):
    if uploaded_file is not None:
        file_path = os.path.join(temp_dir, save_name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return file_path
    else:
        return None

# 存檔
each_area_path = save_uploaded_file(each_area_file, "each_area.xlsx")
last_year_path = save_uploaded_file(last_year_file, "last_year.xlsx")
this_year_path = save_uploaded_file(this_year_file, "this_year.xlsx")
path_four = save_uploaded_file(path_four_file, "path_four.xlsx")

# --- 按鈕：開始產生報表 ---
if st.button("🚀 開始產生報表"):
    with st.spinner("報表生成中，請稍候..."):
        try:
            # 整理 Config
            Config = {
                'year': year,
                'month': month,
                'company_name': company_name,
                'each_area_path': each_area_path,
                'path_one': path_one_folder,
                'path_two': path_two_folder,
                'last_year_path': last_year_path,
                'this_year_path': this_year_path,
                'path_four': path_four,
                'path_five': path_five_folder,
                'output_folder_path': output_folder
            }

            # 建立報表物件並執行
            coordinator = ReportCoordinator(Config)
            coordinator.run_all()

            # 假設你的檔案存在 output_folder_path 底下，用年份命名
            output_file = f"營業店年度營業額及各項比率計算({year}年)--.xlsx"
            output_path = os.path.join(output_folder, output_file)

            with open(output_path, "rb") as f:
                st.success("✅ 報表產生成功！請下載：")
                st.download_button(
                    label="📥 下載報表Excel",
                    data=f,
                    file_name=output_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"產生失敗！錯誤訊息：{str(e)}")
