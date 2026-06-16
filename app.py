import streamlit as st
import pandas as pd
from PIL import Image
import io
import json
import urllib.request
import urllib.parse
import base64

# 設定網頁標題與防翻譯
st.set_page_config(page_title="智慧表格截圖轉 Excel", page_icon="📊", layout="centered")
st.markdown("<script>var meta = document.createElement('meta'); meta.name = 'google'; meta.content = 'notranslate'; document.getElementsByTagName('head')[0].appendChild(meta);</script>", unsafe_allow_html=True)

st.title("智慧表格截圖轉 Excel")

# 上傳檔案區塊
uploaded_file = st.file_uploader("請選擇或拖曳表格截圖 (PNG, JPG, JPEG)", type=["png", "jpg", "jpeg"])

if uploaded_file is not None:
    image = Image.open(uploaded_file)
    st.image(image, caption="已上傳的表格預覽", use_container_width=True)
    
    if st.button("開始編排並轉換"):
        with st.spinner("正在聚合文字並對齊欄位，請稍候..."):
            try:
                # 1. 將圖片轉為 base64 編碼
                img_byte_arr = io.BytesIO()
                image.convert("RGB").save(img_byte_arr, format='JPEG')
                encoded_string = base64.b64encode(img_byte_arr.getvalue()).decode()
                
                # 2. 呼叫 OCR.space API
                url = "https://api.ocr.space/parse/image"
                payload = {
                    "base64Image": f"data:image/jpeg;base64,{encoded_string}",
                    "language": "cht",
                    "apikey": "helloworld",
                    "isOverlayRequired": "true"
                }
                
                data = urllib.parse.urlencode(payload).encode('utf-8')
                req = urllib.request.Request(url, data=data)
                
                with urllib.request.urlopen(req, timeout=30) as response:
                    res_json = json.loads(response.read().decode('utf-8'))
                
                # 3. 智慧行列對齊與字串聚合演自动化
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
                                width = word.get("Width", 0)
                                height = word.get("Height", 10)
                                
                                # 排除網頁右下角自帶的「轉換為表格」浮水印按鈕文字
                                if text and "轉換" not in text and "表格" not in text:
                                    raw_data.append({
                                        "text": text, 
                                        "left": left, 
                                        "right": left + width,
                                        "top": top, 
                                        "bottom": top + height
                                    })
                        
                        if raw_data:
                            # 依垂直位置(Top)進行分行
                            raw_data.sort(key=lambda x: x["top"])
                            rows = []
                            current_row = []
                            current_top = raw_data[0]["top"]
                            
                            row_threshold = 18  # 同一行的垂直誤差容忍值
                            
                            for item in raw_data:
                                if item["top"] - current_top > row_threshold:
                                    rows.append(current_row)
                                    current_row = [item]
                                    current_top = item["top"]
                                else:
                                    current_row.append(item)
                            if current_row:
                                rows.append(current_row)
                            
                            final_table = []
                            for r in rows:
                                # 每一行內部先依水平位置(Left)排序
                                r.sort(key=lambda x: x["left"])
                                
                                # 【核心優化】智慧字串聚合：如果兩個字水平靠得很近，就合併成同一個儲存格
                                merged_row = []
                                if r:
                                    current_cell = r[0]
                                    col_threshold = 35  # 合併儲存格的水平距離極限 (px)
                                    
                                    for next_item in r[1:]:
                                        # 如果下一個字的起點跟上一個字的終點靠得很近，或者是像時間冒號等特殊符號
                                        if (next_item["left"] - current_cell["right"] < col_threshold) or \
                                           (next_item["text"] in [":", "-", "(", ")"]) or \
                                           (current_cell["text"] in [":", "-", "(", ")"]):
                                            # 合併文字並更新邊界
                                            current_cell["text"] += next_item["text"]
                                            current_cell["right"] = next_item["right"]
                                        else:
                                            merged_row.append(current_cell)
                                            current_cell = next_item
                                    merged_row.append(current_cell)
                                
                                final_table.append([item["text"] for item in merged_row])
                            
                            # 找出最大欄位數並對齊補白
                            max_cols = max(len(row) for row in final_table)
                            padded_table = [row + [""] * (max_cols - len(row)) for row in final_table]
                            
                            # 轉換為 DataFrame
                            df = pd.DataFrame(padded_table)
                            
                            st.success("智慧優化編排完成！")
                            st.subheader("修正後的表格預覽")
                            st.dataframe(df, use_container_width=True)
                            
                            # 寫入 Excel
                            excel_buffer = io.BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                df.to_excel(writer, index=False, header=False, sheet_name='精準辨識結果')
                            excel_data = excel_buffer.getvalue()
                            
                            st.download_button(
                                label="下載修正後的 Excel 檔案",
                                data=excel_data,
                                file_name="智慧精準表格結果.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.warning("沒能成功分析出表格文字。")
                    else:
                        st.warning("無法獲取圖片的結構化坐標，請確保截圖清晰。")
                else:
                    st.error(f"API 失敗：{res_json.get('ErrorMessage')}")
            except Exception as e:
                st.error(f"編排過程中發生錯誤：{e}")
