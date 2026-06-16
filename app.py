import streamlit as st
import pandas as pd
from PIL import Image
import io
import json
import urllib.request
import urllib.parse
import base64

# 設定網頁標題與防翻譯
st.set_page_config(page_title="線上表格截圖轉 Excel", page_icon="📊", layout="centered")
st.markdown("<script>var meta = document.createElement('meta'); meta.name = 'google'; meta.content = 'notranslate'; document.getElementsByTagName('head')[0].appendChild(meta);</script>", unsafe_allow_html=True)

st.title("📸 智慧表格截圖轉 Excel")
st.write("上傳結構化表格（如薪資表、簽到表）截圖，系統會自動分析行列位置並進行表格編排。")

# 上傳檔案區塊
uploaded_file = st.file_uploader("請選擇或拖曳表格截圖 (PNG, JPG, JPEG)", type=["png", "jpg", "jpeg"])

if uploaded_file is not None:
    image = Image.open(uploaded_file)
    st.image(image, caption="已上傳的表格預覽", use_container_width=True)
    
    if st.button("🚀 開始智慧編排並轉換"):
        with st.spinner("正在分析表格結構並對齊行列，請稍候..."):
            try:
                # 1. 將圖片轉為 base64 編碼
                img_byte_arr = io.BytesIO()
                image.convert("RGB").save(img_byte_arr, format='JPEG')
                encoded_string = base64.b64encode(img_byte_arr.getvalue()).decode()
                
                # 2. 呼叫 OCR.space API，並啟用 isOverlayRequired（取得文字坐標資訊）
                url = "https://api.ocr.space/parse/image"
                payload = {
                    "base64Image": f"data:image/jpeg;base64,{encoded_string}",
                    "language": "cht",
                    "apikey": "helloworld",
                    "isOverlayRequired": "true"  # 關鍵：開啟文字坐標定位
                }
                
                data = urllib.parse.urlencode(payload).encode('utf-8')
                req = urllib.request.Request(url, data=data)
                
                with urllib.request.urlopen(req, timeout=30) as response:
                    res_json = json.loads(response.read().decode('utf-8'))
                
                # 3. 智慧行列對齊演算法
                if res_json.get("OCRExitCode") == 1:
                    parsed_results = res_json.get("ParsedResults", [])
                    if parsed_results and "TextOverlay" in parsed_results[0]:
                        lines = parsed_results[0]["TextOverlay"]["Lines"]
                        
                        raw_data = []
                        for line in lines:
                            for word in line.get("Words", []):
                                text = word.get("WordText", "").strip()
                                left = word.get("Left", 0)
                                top = word.get("Top", 0)
                                height = word.get("Height", 10)
                                
                                if text:
                                    raw_data.append({"text": text, "left": left, "top": top, "bottom": top + height})
                        
                        if raw_data:
                            # 依垂直位置(Top)進行初步分行
                            raw_data.sort(key=lambda x: x["top"])
                            rows = []
                            current_row = []
                            current_top = raw_data[0]["top"]
                            
                            # 容許在同一個水平線上的誤差值 (px)
                            row_threshold = 15 
                            
                            for item in raw_data:
                                if item["top"] - current_top > row_threshold:
                                    rows.append(current_row)
                                    current_row = [item]
                                    current_top = item["top"]
                                else:
                                    current_row.append(item)
                            if current_row:
                                rows.append(current_row)
                            
                            # 每一行內部再依水平位置(Left)排序
                            final_table = []
                            for r in rows:
                                r.sort(key=lambda x: x["left"])
                                final_table.append([item["text"] for item in r])
                            
                            # 找出最大欄位數，補齊格子防止錯位
                            max_cols = max(len(row) for row in final_table)
                            padded_table = [row + [""] * (max_cols - len(row)) for row in final_table]
                            
                            # 轉換為 DataFrame
                            df = pd.DataFrame(padded_table)
                            
                            st.success("🎉 智慧編排完成！")
                            st.subheader("📝 編排結果預覽")
                            st.dataframe(df, use_container_width=True)
                            
                            # 寫入 Excel
                            excel_buffer = io.BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                df.to_excel(writer, index=False, header=False, sheet_name='智慧編排結果')
                            excel_data = excel_buffer.getvalue()
                            
                            st.download_button(
                                label="📥 下載自動編排的 Excel 檔案",
                                data=excel_data,
                                file_name="表格編排結果.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.warning("⚠️ 沒能成功分析出表格文字。")
                    else:
                        st.warning("⚠️ 無法獲取圖片的結構化坐標，請確保截圖清晰。")
                else:
                    st.error(f"API 失敗：{res_json.get('ErrorMessage')}")
            except Exception as e:
                st.error(f"編排過程中發生錯誤：{e}")
