import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# --- 設定網頁標題與版面 ---
st.set_page_config(page_title="PDF 型錄轉 Excel 工具", page_icon="📂")

st.title("📂 PDF 型錄轉 Excel 工具")
st.markdown("""
此工具專為提取 **Mont-bell 型錄** (或其他類似格式) 設計。
上傳 PDF 後，程式將自動抓取 **產品名稱、Style#、價格、重量** 等資訊並整理成 Excel。
""")

# --- 核心解析函式 (與之前相同) ---
def parse_product_page(text, page_num):
    data = {}
    
    # 1. 擷取 Style Number (關鍵識別)
    style_match = re.search(r"Style#\s*(\d+)", text, re.IGNORECASE)
    if not style_match:
        return None
    
    data['Page'] = page_num
    data['Style#'] = style_match.group(1)

    # 2. 擷取價格 (MSRP)
    price_match = re.search(r"MSRP\s*[¥￥]([\d,]+)", text, re.IGNORECASE)
    if price_match:
        data['MSRP (JPY)'] = price_match.group(1).replace(',', '')
    else:
        data['MSRP (JPY)'] = ""

    # 3. 擷取重量
    weight_match = re.search(r"Estimated Average Weight\s*[\r\n]*\s*(\d+\.?\d*|TBA|ТВА)", text, re.IGNORECASE | re.MULTILINE)
    if weight_match:
        data['Weight (g)'] = weight_match.group(1).replace('ТВА', 'TBA')
    else:
        data['Weight (g)'] = ""

    # 4. 擷取產品名稱
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    product_name = ""
    for i, line in enumerate(lines):
        if "Style#" in line:
            for k in range(i - 1, -1, -1):
                current_line = lines[k]
                if "mont-bell" in current_line.lower() and "fall" in current_line.lower():
                    continue
                if "NEW" == current_line or "REVISED" == current_line:
                    continue
                product_name = current_line
                break
            break
    data['Product Name'] = product_name

    # 5. 擷取類別
    categories = [
        "ALPINE CLOTHING", "INSULATION", "THERMAL", "RAIN WEAR", 
        "WIND SHELL", "SOFT SHELL", "PANTS", "BASE LAYER", 
        "FIELD WEAR", "TRAVEL & COUNTRY", "CAP & HAT", "GLOVES",
        "SOCKS", "SLEEPING BAG", "FOOTWEAR", "BACKPACK", "BAG",
        "ACCESSORIES", "CYCLING", "SNOW GEAR", "CLIMBING", "FISHING",
        "PADDLE SPORTS", "DOG GEAR", "KIDS & BABY"
    ]
    data['Category'] = ""
    for cat in categories:
        if cat in text:
            data['Category'] = cat
            break

    return data

# --- 主程式邏輯 ---
uploaded_file = st.file_uploader("請上傳 PDF 檔案", type="pdf")

if uploaded_file is not None:
    st.success("檔案上傳成功！準備開始處理...")
    
    if st.button("開始擷取資料"):
        products = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            # 使用 pdfplumber 開啟上傳的檔案物件
            with pdfplumber.open(uploaded_file) as pdf:
                total_pages = len(pdf.pages)
                
                for i, page in enumerate(pdf.pages):
                    # 更新進度條
                    progress = (i + 1) / total_pages
                    progress_bar.progress(progress)
                    status_text.text(f"正在處理第 {i + 1}/{total_pages} 頁...")

                    text = page.extract_text()
                    if not text:
                        continue

                    product_data = parse_product_page(text, i + 1)
                    if product_data:
                        products.append(product_data)

            # 處理完成
            if products:
                df = pd.DataFrame(products)
                
                # 欄位排序
                cols = ['Page', 'Category', 'Product Name', 'Style#', 'MSRP (JPY)', 'Weight (g)']
                cols = [c for c in cols if c in df.columns]
                df = df[cols]

                st.success(f"成功擷取 {len(products)} 項產品！")
                
                # 顯示預覽
                st.dataframe(df.head(10))

                # --- 轉換為 Excel 供下載 (存入記憶體) ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Products')
                
                excel_data = output.getvalue()

                # 下載按鈕
                st.download_button(
                    label="📥 下載 Excel 報表",
                    data=excel_data,
                    file_name="Montbell_Product_List.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("警告：未在 PDF 中找到任何符合格式的產品資料。")

        except Exception as e:
            st.error(f"發生錯誤: {e}")