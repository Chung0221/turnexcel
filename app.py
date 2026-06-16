import streamlit as st
import pandas as pd
from PIL import Image
import io
import json
import urllib.request
import base64

# 設定網頁標題與防翻譯
st.set_page_config(page_title="線上截圖轉 Excel", page_icon="📊", layout="centered")
st.markdown("<script>var meta = document.createElement('meta'); meta.name = 'google'; meta.content = 'notranslate'; document.getElementsByTagName('head')[0].appendChild(meta);</script>", unsafe_allow_html=True)

st.title("截圖轉 Excel 線上小工具")
st.write("上傳截圖，系統會自動辨識文字並轉換為 Excel 檔案供你下載。")

# 上傳檔案區塊
uploaded_file = st.file_uploader("請選擇或拖曳截圖檔案 (PNG, JPG, JPEG)", type=["png", "jpg", "jpeg"])

if uploaded_file is not None:
    image = Image.open(uploaded_file)
    st.image(image, caption="已上傳的截圖預覽", use_container_width=True)
    
    if st.button("開始辨識並轉換"):
        with st.spinner("免費雲端 API 正在努力辨識圖片中，請稍候..."):
            try:
                # 1. 將圖片轉為 base64 編碼
                img_byte_arr = io.BytesIO()
                # 統一轉成 JPEG 確保相容性
                image.convert("RGB").save(img_byte_arr, format='JPEG')
                encoded_string = base64.b64encode(img_byte_arr.getvalue()).decode()
                
                # 2. 呼叫免費、免申請金鑰的公共 OCR.space API
                url = "https://api.ocr.space/parse/image"
                # OCHT 代表繁體中文 (Traditional Chinese)
                payload = {
                    "base64Image": f"data:image/jpeg;base64,{encoded_string}",
                    "language": "cht",
                    "apikey": "helloworld"  # 公共免費免註冊金鑰
                }
                
                # 將推擠資料轉為 URL 編碼格式
                data = urllib.parse.urlencode(payload).encode('utf-8')
                req = urllib.request.Request(url, data=data)
                
                # 發送請求並讀取回應
                with urllib.request.urlopen(req, timeout=30) as response:
                    res_json = json.loads(response.read().decode('utf-8'))
                
                # 3. 解析 API 回傳的文字結果
                if res_json.get("OCRExitCode") == 1:
                    parsed_results = res_json.get("ParsedResults", [])
                    if parsed_results:
                        # 取得整張圖片辨識到的所有文字段落
                        text_lines = parsed_results[0].get("ParsedText", "").split("\r\n")
                        # 清除可能產生的空白行
                        text_lines = [line.strip() for line in text_lines if line.strip()]
                    else:
                        text_lines = []
                else:
                    st.error(f"API 辨識失敗，原因：{res_json.get('ErrorMessage')}")
                    text_lines = []

                # 4. 顯示與匯出結果
                if len(text_lines) == 0:
                    st.warning("⚠️ 圖片中似乎沒有偵測到任何繁體中文或英文文字，請換一張清晰一點的截圖。")
                else:
                    st.success("辨識完成！")
                    
                    # 轉成 DataFrame
                    df = pd.DataFrame(text_lines, columns=["辨識文字"])
                    
                    # 預覽辨識結果
                    st.subheader("辨識到的文字預覽")
                    st.dataframe(df, use_container_width=True)
                    
                    # 寫入記憶體
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='辨識結果')
                    excel_data = excel_buffer.getvalue()
                    
                    # 下載按鈕
                    st.download_button(
                        label="下載 Excel 檔案",
                        data=excel_data,
                        file_name="截圖辨識結果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"辨識過程中發生錯誤：{e}")
