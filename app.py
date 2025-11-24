import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import altair as alt

# --- 1. 頁面設定 ---
st.set_page_config(
    page_title="Mont-bell 型錄解析器 Ver 9.0 (空間座標版)",
    page_icon="🏔️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. 核心解析邏輯 (Ver 9.0: 空間座標切割術) ---
def parse_product_page_spatial(page, page_num):
    """
    使用空間座標 (X, Y) 來解析頁面，解決左右欄位混合的問題。
    """
    data = {'Page': page_num, 'Category': 'Uncategorized', 'MSRP': '0', 'Weight (g)': '', 'Features': '', 'Material': '', 'Description': ''}
    
    # 1. 取得所有文字物件 (包含座標資訊)
    # x0: 左邊界, top: 上邊界, bottom: 下邊界, text: 文字內容
    words = page.extract_words(keep_blank_chars=True, x_tolerance=2, y_tolerance=2)
    full_text = page.extract_text() or ""
    
    if not words: return None

    # --- A. 尋找關鍵錨點 (Anchors) ---
    # 我們需要找到 "Features" 和 "Material" 的座標，以此作為切割畫面的基準
    features_anchor = None
    material_anchor = None
    
    for w in words:
        txt = w['text'].strip()
        if txt == "Features" and features_anchor is None:
            features_anchor = w
        elif txt == "Material" and material_anchor is None:
            material_anchor = w
    
    # 如果找不到這兩個錨點，退化回純文字搜尋 (Fallback)
    if not features_anchor or not material_anchor:
        # 這裡可以寫一個簡單的 fallback，或者直接回傳僅有基本資訊
        # 為了代碼簡潔，這裡做簡單處理
        pass 

    # --- B. 抓取 Style# (7碼暴力搜尋) ---
    # 優先使用全文正則，因為 Style# 可能在任何位置
    style_match = re.search(r"Style\s*#?\s*(\d{7})", full_text, re.IGNORECASE)
    if not style_match:
        # 嘗試找純數字 (排除電話等)
        candidates = list(re.finditer(r"(?<!\d)(\d{7})(?!\d)", full_text))
        valid_style = ""
        for m in candidates:
            # 簡單過濾: Montbell Style 通常以 11, 23, 04, 05, 12 等開頭
            # 這裡先不做嚴格過濾，取第一個看起來像的
            if "¥" not in full_text[max(0, m.start()-10):m.end()+10]: # 排除價格
                valid_style = m.group(1)
                break
        data['Style#'] = valid_style
    else:
        data['Style#'] = style_match.group(1)

    if not data.get('Style#'): return None # 無 Style# 則跳過此頁

    # --- C. 抓取產品名稱 & 敘述 (Description) ---
    # 定義區域：頁面頂部 ~ Features 標題上方
    limit_bottom = features_anchor['top'] if features_anchor else 600
    
    upper_lines = []
    # 簡單將文字依 Y 軸分組
    current_y = -1
    line_buffer = []
    sorted_words = sorted(words, key=lambda w: (w['top'], w['x0']))
    
    for w in sorted_words:
        if w['top'] > limit_bottom: continue # 超過下方界線
        
        # 換行判斷 (Y 軸差異 > 5 視為換行)
        if abs(w['top'] - current_y) > 5:
            if line_buffer: upper_lines.append(" ".join([x['text'] for x in line_buffer]))
            line_buffer = []
            current_y = w['top']
        line_buffer.append(w)
    if line_buffer: upper_lines.append(" ".join([x['text'] for x in line_buffer]))

    # 分析上半部文字
    desc_list = []
    product_name = ""
    
    for line in upper_lines:
        line_clean = line.strip()
        
        # 排除 Style# 行
        if data['Style#'] in line_clean: 
            # 嘗試從 Style# 同一行抓名稱 (例如 "Jacket Style# 1101002")
            pre_text = line_clean.split("Style")[0].strip()
            if len(pre_text) > 3 and "NEW" not in pre_text:
                product_name = pre_text
            continue
            
        # 排除價格行
        if "MSRP" in line_clean or "¥" in line_clean: 
            data['MSRP'] = re.search(r"[\d,]+", line_clean).group(0).replace(',', '') if re.search(r"[\d,]+", line_clean) else "0"
            continue

        # 排除雜訊
        if any(x in line_clean for x in ["mont-bell", "Fall", "Winter", "Spring", "Summer", "CONFIDENTIAL", "KJ"]): continue
        
        # 抓取敘述 (以 • 或 ● 開頭)
        if line_clean.startswith("•") or line_clean.startswith("●"):
            desc_list.append(line_clean)
        # 抓取產品名稱 (如果還沒找到，且是大寫字母為主，且長度夠)
        elif not product_name and len(line_clean) > 3 and not line_clean.isdigit():
             product_name = line_clean

    data['Product Name'] = product_name
    data['Description'] = "\n".join(desc_list)

    # --- D. 空間切割：Features vs Material ---
    if features_anchor and material_anchor:
        # 定義切割中線 (Split X)
        split_x = (features_anchor['x0'] + material_anchor['x0']) / 2
        header_bottom = max(features_anchor['bottom'], material_anchor['bottom'])
        
        # 定義底部停止線 (遇到 Size 或 Estimated Weight)
        footer_top = 10000
        for w in words:
            if w['text'] in ["Size", "Estimated"] and w['top'] > header_bottom:
                footer_top = min(footer_top, w['top'])
        
        features_txt = []
        material_txt = []
        
        # 重新掃描文字，這次針對下方區域
        # 這裡不使用 sorted_words，而是對 words 進行分類
        
        # 我們需要「逐行」組裝，才能保持句子完整
        # 所以先將 header_bottom ~ footer_top 之間的 words 分行
        body_words = [w for w in sorted_words if header_bottom < w['top'] < footer_top]
        
        # 分行邏輯
        curr_y = -1
        row_buffer = []
        rows = []
        
        for w in body_words:
            if abs(w['top'] - curr_y) > 5: # 新的一行
                if row_buffer: rows.append(row_buffer)
                row_buffer = []
                curr_y = w['top']
            row_buffer.append(w)
        if row_buffer: rows.append(row_buffer)
        
        # 判斷每一行屬於左邊 (Features) 還是右邊 (Material)
        for row in rows:
            # 計算這一行的平均 X 座標
            avg_x = sum([w['x0'] for w in row]) / len(row)
            line_str = " ".join([w['text'] for w in row])
            
            # 過濾顏色代碼
            if re.search(r"^[A-Z0-9]{2,4}\([A-Za-z0-9\s]+\)", line_str): continue
            
            if avg_x < split_x:
                features_txt.append(line_str)
            else:
                material_txt.append(line_str)
                
        data['Features'] = "\n".join(features_txt)
        data['Material'] = "\n".join(material_txt)

    # --- E. 補抓重量與 Category (全文搜尋) ---
    weight_match = re.search(r"Estimated Average Weight\s*[\n]*\s*(\d+\.?\d*|TBA|ТВА)", full_text, re.IGNORECASE)
    if weight_match: data['Weight (g)'] = weight_match.group(1).replace('ТВА', 'TBA')

    categories = ["ALPINE CLOTHING", "INSULATION", "THERMAL", "RAIN WEAR", "SOFT SHELL", "PANTS", "BASE LAYER", "FIELD WEAR", "TRAVEL & COUNTRY", "CAP & HAT", "GLOVES", "SOCKS", "SLEEPING BAG", "FOOTWEAR", "BACKPACK", "BAG", "ACCESSORIES", "CYCLING", "SNOW GEAR", "CLIMBING", "FISHING", "PADDLE SPORTS", "DOG GEAR", "KIDS & BABY"]
    for cat in categories:
        if cat in full_text: data['Category'] = cat; break
        
    return data

# --- 3. 側邊欄 ---
with st.sidebar:
    st.header("步驟 1: 上傳檔案")
    uploaded_files = st.file_uploader("可多選上傳 PDF", type="pdf", accept_multiple_files=True)
    st.info("Ver 9.0 空間座標版：\n1. 完美分離 Features 與 Material (不再混合)\n2. 找回遺失的產品敘述 (Description)")

# --- 4. 主畫面 ---
st.title("🏔️ Mont-bell 型錄解析器 Ver 9.0 (空間座標版)")

if uploaded_files:
    col1, col2 = st.columns([1, 5])
    with col1:
        start_btn = st.button("🚀 開始分析", type="primary", use_container_width=True)
    
    if start_btn:
        all_products = []
        progress_text = st.empty()
        my_bar = st.progress(0)
        total_pdfs = len(uploaded_files)
        
        for file_idx, uploaded_file in enumerate(uploaded_files):
            try:
                with pdfplumber.open(uploaded_file) as pdf:
                    total_pages = len(pdf.pages)
                    filename = uploaded_file.name
                    
                    for i, page in enumerate(pdf.pages):
                        current_progress = (file_idx + (i / total_pages)) / total_pdfs
                        my_bar.progress(current_progress)
                        progress_text.text(f"處理中: {filename} (頁面 {i+1}/{total_pages})...")
                        
                        # 傳遞 page 物件而非 text，以使用空間座標
                        p_data = parse_product_page_spatial(page, i + 1)
                        
                        if p_data:
                            p_data['Source File'] = filename
                            all_products.append(p_data)
                            
            except Exception as e:
                st.error(f"Error: {e}")

        my_bar.empty()
        progress_text.empty()

        if all_products:
            df = pd.DataFrame(all_products)
            
            st.success(f"✅ 完成！共擷取 {len(df)} 筆資料。")
            
            tab1, tab2 = st.tabs(["📊 資料總表", "🛠️ Debug"])
            
            with tab1:
                display_cols = ['Source File', 'Page', 'Category', 'Product Name', 'Style#', 'MSRP', 'Features', 'Material', 'Description']
                st.dataframe(
                    df[display_cols], 
                    use_container_width=True,
                    column_config={
                        "Features": st.column_config.TextColumn("特點 (Left)", width="medium"),
                        "Material": st.column_config.TextColumn("材質 (Right)", width="medium"),
                        "Description": st.column_config.TextColumn("產品敘述 (Top)", width="large"),
                    }
                )
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='All_Products')
                excel_data = output.getvalue()
                st.download_button("📥 下載 Excel", data=excel_data, file_name="Montbell_Ver9_Spatial.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

        else:
            st.warning("⚠️ 未擷取到資料。")