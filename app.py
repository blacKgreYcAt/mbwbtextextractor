import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import altair as alt

# --- 1. 頁面設定 ---
st.set_page_config(
    page_title="Mont-bell 萬用型錄解析器 Ver 8.0",
    page_icon="🏔️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. 核心解析邏輯 (Ver 8.0: 三層搜尋 + 暴力 7 碼數字匹配) ---
def parse_product_page(text, page_num):
    data = {}
    data['Page'] = page_num
    
    # 預處理：統一換行，移除 BOM 或是奇怪的隱形字元
    clean_text = text.replace('\r\n', '\n').strip()
    if not clean_text: return None

    lines = [l.strip() for l in clean_text.split('\n') if l.strip()]

    # --- A. Style Number 終極搜尋策略 ---
    primary_style_index = -1
    primary_style_num = ""

    # 定義不同的 Regex 模式 (優先級由高到低)
    patterns = [
        r"Style\s*#?\s*(\d{7})",  # 標準: Style# 1101002 (限制7碼更精準)
        r"Style\s*No\.?\s*(\d{7})", # 變體: Style No. 1101002
        r"Item\s*#?\s*(\d{7})",     # 變體: Item# 1101002
        r"(?<!\d)(\d{7})(?!\d)"     # 暴力: 任何獨立的 7 碼數字 (Mont-bell 邏輯)
    ]

    # 用來排除誤判的 7 碼數字 (例如電話、日期、KJ註記編號)
    # 這裡假設型號通常以 11, 23, 12 開頭 (根據你的型錄觀察)
    # 若你的型號範圍更廣，可以移除這個檢查
    valid_prefixes = ["11", "21", "23", "33", "04", "05", "12"] 

    found_match = False
    
    # 開始逐行掃描 (為了抓到正確的行號 primary_style_index)
    for i, line in enumerate(lines):
        # 忽略 Western Size 行
        if re.search(r"\(\s*style", line, re.IGNORECASE): continue
        if "KJ" in line: continue # 忽略註記行

        for pat in patterns:
            matches = list(re.finditer(pat, line, re.IGNORECASE))
            for m in matches:
                candidate = m.group(1)
                
                # 驗證候選人: 必須是 7 碼
                if len(candidate) != 7: continue
                
                # 驗證開頭 (可選: 增加準確度)
                # if not any(candidate.startswith(p) for p in valid_prefixes): continue
                
                # 排除像價格的數字 (雖然價格通常有逗號，但以防萬一)
                if "¥" in line or "MSRP" in line:
                    # 除非這行明確寫了 Style
                    if "Style" not in line and "Item" not in line:
                        continue

                primary_style_index = i
                primary_style_num = candidate
                found_match = True
                break
            if found_match: break
        if found_match: break

    # 如果逐行找不到，嘗試「全文跨行搜尋」 (針對 Style# 和數字斷行的情況)
    if not found_match:
        # 只用最寬鬆的 pattern 找
        candidates = list(re.finditer(r"Style\s*#?\s*(\d{7})", clean_text, re.IGNORECASE))
        if candidates:
            # 取第一個找到的
            match = candidates[0]
            primary_style_num = match.group(1)
            # 反查行號
            primary_style_index = clean_text.count('\n', 0, match.start())
            found_match = True

    if not found_match:
        return None # 真的找不到產品

    data['Style#'] = primary_style_num

    # --- B. 產品名稱 (基於 Style# 往上找) ---
    product_name = ""
    
    # 策略: 往上找 5 行以內，通常名稱都在附近
    if primary_style_index > 0:
        search_range = range(primary_style_index - 1, max(-1, primary_style_index - 6), -1)
        for k in search_range:
            curr = lines[k]
            
            # 排除雜訊
            skip_keywords = [
                "mont-bell", "Fall", "Winter", "Spring", "Summer", 
                "NEW", "REVISED", "MSRP", "¥", "CONFIDENTIAL", 
                "Western", "Available", "Fabric Sample", "KJ", "註記"
            ]
            is_noise = False
            for kw in skip_keywords:
                if kw.lower() in curr.lower(): is_noise = True; break
            
            if re.search(r"^[A-Z]{2,3}\(.*\)$", curr): is_noise = True
            
            # 排除純數字行 (可能是頁碼)
            if curr.isdigit(): is_noise = True

            if not is_noise:
                product_name = curr
                break
    
    # 如果往上找不到，試試看 Style# 同一行
    if not product_name and primary_style_index < len(lines):
        line = lines[primary_style_index]
        # 移除 Style# 及其後面的數字
        clean_line = re.sub(r"Style.*?(\d{7})", "", line, flags=re.IGNORECASE).strip()
        clean_line = re.sub(r"\d{7}", "", clean_line).strip() # 移除純數字
        if len(clean_line) > 3 and "MSRP" not in clean_line:
            product_name = clean_line

    data['Product Name'] = product_name

    # --- C. 價格與重量 ---
    price_match = re.search(r"MSRP\s*[¥￥]?\s*([\d,]+)", clean_text, re.IGNORECASE)
    alt_price = re.search(r"[¥￥]\s*([\d,]+)", clean_text)
    if price_match: data['MSRP'] = price_match.group(1).replace(',', '')
    elif alt_price: data['MSRP'] = alt_price.group(1).replace(',', '')
    else: data['MSRP'] = "0"

    weight_match = re.search(r"Estimated Average Weight\s*[\n]*\s*(\d+\.?\d*|TBA|ТВА)", clean_text, re.IGNORECASE)
    if weight_match: data['Weight (g)'] = weight_match.group(1).replace('ТВА', 'TBA')
    else: data['Weight (g)'] = ""

    # --- D. Features & Material (通用關鍵字搜尋) ---
    # 建立一個關鍵字映射表，應對不同年份的寫法
    headers = {
        "Features": ["Features", "Feature", "Functions", "Characteristics"],
        "Material": ["Material", "Materials", "Fabric", "Fabrics"]
    }
    
    stop_keywords = ["Size", "Estimated", "Last Updated", "CONFIDENTIAL", "MSRP"]
    
    # 輔助函式：抓取區塊
    def extract_block(target_headers):
        content = []
        is_collecting = False
        
        for line in lines:
            # 檢查是否為標題行
            if any(line.strip().startswith(h) for h in target_headers):
                is_collecting = True
                continue
            
            if is_collecting:
                # 檢查停止條件 (遇到其他大標題)
                # 檢查 Features 標題
                if any(line.strip().startswith(h) for h in headers["Features"]) and "Features" not in target_headers: break
                # 檢查 Material 標題
                if any(line.strip().startswith(h) for h in headers["Material"]) and "Material" not in target_headers: break
                # 檢查通用停止詞
                if any(line.startswith(kw) for kw in stop_keywords): break
                
                # 顏色與尺寸過濾
                if re.search(r"^[A-Z0-9]{2,4}\([A-Za-z0-9\s]+\)", line): continue # 顏色代碼
                if re.search(r"^[XSML\s,]+$", line) or "Size" in line: break      # 尺寸行

                content.append(line)
        return "\n".join(content)

    data['Features'] = extract_block(headers["Features"])
    data['Material'] = extract_block(headers["Material"])

    # --- E. Category ---
    categories = ["ALPINE CLOTHING", "INSULATION", "THERMAL", "RAIN WEAR", "SOFT SHELL", "PANTS", "BASE LAYER", "FIELD WEAR", "TRAVEL & COUNTRY", "CAP & HAT", "GLOVES", "SOCKS", "SLEEPING BAG", "FOOTWEAR", "BACKPACK", "BAG", "ACCESSORIES", "CYCLING", "SNOW GEAR", "CLIMBING", "FISHING", "PADDLE SPORTS", "DOG GEAR", "KIDS & BABY"]
    data['Category'] = "Uncategorized"
    for cat in categories:
        if cat in clean_text: data['Category'] = cat; break

    return data

# --- 3. 側邊欄 ---
with st.sidebar:
    st.header("步驟 1: 上傳檔案")
    uploaded_files = st.file_uploader("可多選上傳 PDF", type="pdf", accept_multiple_files=True)
    st.info("Ver 8.0 強力版：\n1. 支援多檔批次處理\n2. 暴力搜尋 7 碼型號 (解決格式跑版)\n3. 相容 KJ 註記與不同年份格式")

# --- 4. 主畫面 ---
st.title("🏔️ Mont-bell 萬用型錄解析器 (Ver 8.0)")

if uploaded_files:
    col1, col2 = st.columns([1, 5])
    with col1:
        start_btn = st.button("🚀 開始分析", type="primary", use_container_width=True)
    
    if start_btn:
        all_products = []
        
        # 建立進度條容器
        progress_text = st.empty()
        my_bar = st.progress(0)
        
        total_pdfs = len(uploaded_files)
        
        for file_idx, uploaded_file in enumerate(uploaded_files):
            try:
                with pdfplumber.open(uploaded_file) as pdf:
                    total_pages = len(pdf.pages)
                    filename = uploaded_file.name
                    
                    for i, page in enumerate(pdf.pages):
                        # 更新全域進度
                        current_progress = (file_idx + (i / total_pages)) / total_pdfs
                        my_bar.progress(current_progress)
                        progress_text.text(f"正在處理: {filename} (第 {i+1}/{total_pages} 頁)...")
                        
                        text = page.extract_text()
                        if not text: continue
                        
                        p_data = parse_product_page(text, i + 1)
                        if p_data:
                            p_data['Source File'] = filename # 標記來源檔案
                            all_products.append(p_data)
                            
            except Exception as e:
                st.error(f"檔案 {uploaded_file.name} 讀取失敗: {e}")

        my_bar.empty()
        progress_text.empty()

        if all_products:
            df = pd.DataFrame(all_products)
            df['MSRP'] = pd.to_numeric(df['MSRP'], errors='coerce').fillna(0)
            
            st.success(f"✅ 全部分析完成！共擷取 **{len(df)}** 項產品。")
            
            # 建立 Tabs
            tab1, tab2, tab3 = st.tabs(["📊 總表與下載", "📈 交叉分析", "🛠️ 除錯模式 (Debug)"])
            
            with tab1:
                st.subheader("📋 整合資料清單")
                display_cols = ['Source File', 'Page', 'Category', 'Product Name', 'Style#', 'MSRP', 'Features', 'Material']
                st.dataframe(
                    df[display_cols], 
                    use_container_width=True, 
                    column_config={
                        "Features": st.column_config.TextColumn("特點", width="medium"),
                        "Material": st.column_config.TextColumn("材質", width="medium")
                    }
                )
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='All_Products')
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 下載完整 Excel (含所有檔案)",
                    data=excel_data,
                    file_name="Montbell_Merged_Catalog.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            with tab2:
                st.subheader("年份/檔案交叉比對")
                chart = alt.Chart(df).mark_bar().encode(
                    x=alt.X('Source File', title='來源檔案'),
                    y=alt.Y('count()', title='產品數量'),
                    color='Category',
                    tooltip=['Source File', 'Category', 'count()']
                ).properties(height=400)
                st.altair_chart(chart, use_container_width=True)

            with tab3:
                st.subheader("🛠️ 原始資料檢視 (用於除錯)")
                st.markdown("如果你發現某頁資料抓錯，請在此查看該頁的「原始擷取文字」。")
                
                # 讓使用者選擇要檢查的檔案與頁數
                debug_file = st.selectbox("選擇檔案", [f.name for f in uploaded_files])
                debug_page = st.number_input("輸入頁碼 (1-based)", min_value=1, value=5)
                
                if st.button("檢視原始文字"):
                    # 重新讀取該頁 (為了顯示) - 這裡稍微沒效率但在 debug 模式可接受
                    target_file_obj = next(f for f in uploaded_files if f.name == debug_file)
                    # 需重置 pointer
                    target_file_obj.seek(0) 
                    with pdfplumber.open(target_file_obj) as dbg_pdf:
                        if debug_page <= len(dbg_pdf.pages):
                            raw_txt = dbg_pdf.pages[debug_page-1].extract_text()
                            st.text_area("PDF Raw Text Content:", raw_txt, height=400)
                        else:
                            st.error("頁碼超出範圍")

        else:
            st.warning("⚠️ 掃描了所有檔案，但未發現符合格式的資料。請切換到「除錯模式」檢查原始文字是否為亂碼或圖片。")