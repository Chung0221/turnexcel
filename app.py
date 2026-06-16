python
import streamlit as st
import easyocr
import pandas as pd
from PIL import Image
import io

# 設定網頁標題
st.set_page_config(page_title="線上截圖轉 Excel", page_icon="📊", layout="centered")

st.title("📸 截圖轉 Excel 線上小工具")
st.write("上傳截圖，系統會自動辨識文字並轉換為 Excel 檔案供你下載。")

# 快取 OCR 引擎，避免重複載入
@st.cache_resource
def load_ocr():
    # 支援繁體中文 (ch_tra) 與英文 (en)
    return easyocr.Reader(['ch_tra', 'en'])

try:
    reader = load_ocr()
except Exception as e:
    st.error(f"OCR 引擎初始化失敗，請重新整理網頁。錯誤訊息: {e}")

# 上傳檔案區塊
uploaded_file = st.file_uploader("請選擇或拖曳截圖檔案 (PNG, JPG, JPEG)", type=["png", "jpg", "jpeg"])

if uploaded_file is not None:
    # 讀取並顯示圖片
    image = Image.open(uploaded_file)
    st.image(image, caption="已上傳的截圖預覽", use_container_width=True)
    
    # 辨識按鈕
    if st.button("🚀 開始辨識並轉換"):
        with st.spinner("雲端伺服器正在辨識圖片中的文字，請稍候..."):
            try:
                # 執行辨識
                result = reader.readtext(image, detail=0)
                
                if len(result) == 0:
                    st.warning("⚠️ 圖片中似乎沒有偵測到任何文字，請換一張清晰一點的截圖。")
                else:
                    st.success("🎉 辨識完成！")
                    
                    # 預覽辨識結果
                    st.subheader("📝 辨識到的文字預覽")
                    st.write(result)
                    
                    # 轉成 DataFrame
                    df = pd.DataFrame(result, columns=["辨識文字"])
                    
                    # 寫入記憶體
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='工作表1')
                    excel_data = excel_buffer.getvalue()
                    
                    # 下載按鈕
                    st.download_button(
                        label="📥 下載 Excel 檔案",
                        data=excel_data,
                        file_name="截圖辨識結果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"辨識過程中發生錯誤：{e}")
