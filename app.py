import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import altair as alt

# --- 1. 頁面設定 ---
st.set_page_config(
    page_title="Mont-bell 型錄解析器 Ver 10.0 (絕對區域版)",
    page_icon="🏔️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. 核心解析邏輯 (Ver 10.0: 絕對區域鎖定) ---
def parse_product_page_v10(page, page_num):
    """
    使用物件座標進行絕對區域切割。
    """
    # 擷取頁面所有文字物件 (含座標)
    words = page.extract_words(keep_blank_chars=True, x_tolerance=2, y_tolerance=2)
    full_text = page.extract_text() or ""
    
    if not words: return None
    
    data = {
        'Page': page_num, 
        'Category': 'Uncategorized', 
        'Product Name': '', 
        'Style#': '', 
        'MSRP': '0', 
        'Weight (g)': '', 
        'Features': '', 
        'Material': '', 
        'Description': ''
    }

    # --- A. 尋找關鍵錨點 (Anchors) ---
    style_anchor = None
    features_anchor = None
    material_anchor = None
    
    # 掃描文字物件尋找地標
    for w in words:
        txt = w['text'].strip()
        
        # 找 Style# (抓取 Style 開頭的物件)
        if "Style" in txt and style_anchor is None:
            style_anchor = w
        
        # 找標題 (允許模糊匹配，例如 "Features" 或 "Feature")
        if txt.startswith("Feature") and features_anchor is None:
            features_anchor = w
        elif txt.startswith("Material") and material_anchor is None:
            material_anchor = w

    # --- B. 抓取 Style# (如果錨點沒找到，用 Regex 全文補抓) ---
    # 優先從全文抓取 7 碼數字，因為這最準確
    style_regex = re.search(r"Style\s*#?\s*(\d{7})", full_text, re.IGNORECASE)
    if style_regex:
        data['Style#'] = style_regex.group(1)
        # 如果前面沒找到錨點，嘗試反推錨點位置 (雖不精確但可用)
    else:
        # 暴力搜尋 7 碼
        candidates = list(re.finditer(r"(?<!\d)(\d{7})(?!\d)", full_text))
        for m in candidates:
            # 排除看起來像價格的
            if "¥" not in full_text[max(0, m.start()-10):m.end()+10]:
                data['Style#'] = m.group(1)
                break
    
    if not data['Style#']: return None # 沒有型號就跳過

    # --- C. 定義區域邊界 (Boundaries) ---
    # 上下分界線：預設為頁面中間，若有 Features 標題則以標題頂部為準
    split_y = features_anchor['top'] if features_anchor else (page.height / 2)
    
    # 左右分界線：Features 和 Material 的中間
    if features_anchor and material_anchor:
        split_x = (features_anchor['x0'] + material_anchor['x0']) / 2
    else:
        split_x = page.width / 2 # 預設切中線

    # --- D. 區域 1: 上半部 (Header Section) ---
    # 包含：Category, Product Name, MSRP, Description
    
    # 篩選出位於 split_y 之上的文字，並按 Y 軸排序
    upper_words = [w for w in words if w['bottom'] < split_y]
    upper_lines = words_to_lines(upper_words)

    # D-1. 抓取 Product Name (品名)
    # 策略：找到 Style# 那一行，然後往上找「最近的」一行非雜訊文字
    
    # 先定位 Style# 在哪一行
    style_line_idx = -1
    for i, line in enumerate(upper_lines):
        if data['Style#'] in line:
            style_line_idx = i
            break
            
    # 如果找不到 Style# 行 (可能 Style# 是 Regex 抓到的但 words 裡被拆開了)
    # 我們嘗試找 "Style" 字眼
    if style_line_idx == -1:
        for i, line in enumerate(upper_lines):
            if "Style" in line:
                style_line_idx = i
                break

    # 開始往上找品名
    potential_name = ""
    if style_line_idx > 0:
        # 往上檢查最多 5 行
        for k in range(style_line_idx - 1, max(-1, style_line_idx - 6), -1):
            curr_line = upper_lines[k].strip()
            
            # 雜訊過濾器
            skip_keywords = ["mont-bell", "Fall", "Winter", "Spring", "Summer", "CONFIDENTIAL", "KJ", "MSRP", "¥"]
            is_noise = False
            for kw in skip_keywords:
                if kw in curr_line: is_noise = True; break
            
            # 過濾純數字或頁碼
            if curr_line.isdigit(): is_noise = True
            
            # 過濾顏色 (Color) 代碼行 (例如 "BK(Black) RD(Red)")
            if re.search(r"[A-Z]{2,3}\([A-Za-z]+\)", curr_line): is_noise = True

            if not is_noise and len(curr_line) > 2:
                potential_name = curr_line
                break # 找到就停，因為最接近 Style# 的通常就是品名
    
    # 如果往上找不到，試試看 Style# 同一行前方 (例如 "Down Jacket Style# 1101...")
    if not potential_name and style_line_idx != -1:
        current_line = upper_lines[style_line_idx]
        if "Style" in current_line:
            pre_text = current_line.split("Style")[0].strip()
            if len(pre_text) > 3:
                potential_name = pre_text

    data['Product Name'] = potential_name

    # D-2. 抓取 Description (敘述)
    # 策略：在上半部區域中，抓取所有以 • 或 ● 開頭的行，或是位於標題與 Features 之間的長文字
    desc_lines = []
    for line in upper_lines:
        line = line.strip()
        if line.startswith("•") or line.startswith("●"):
            desc_lines.append(line)
        # 有些敘述沒有點點，但很長且不是品名
        elif len(line) > 40 and line != potential_name and "Style" not in line and "MSRP" not in line:
            # 再次確認不是雜訊
            if "mont-bell" not in line and "CONFIDENTIAL" not in line:
                desc_lines.append(line)
    
    data['Description'] = "\n".join(desc_lines)

    # --- E. 區域 2 & 3: 下半部 (Features & Material) ---
    # 篩選出位於 split_y 之下的文字
    # 設定一個底部邊界 (遇到 Size 或 Estimated Weight 停止)
    footer_y = page.height
    for w in words:
        if w['top'] > split_y and (w['text'] in ["Size", "Estimated", "Last"]):
            footer_y = min(footer_y, w['top'])
    
    lower_words = [w for w in words if w['top'] > split_y and w['bottom'] < footer_y]
    lower_lines = words_to_lines(lower_words) # 這裡先不轉行，因為要分左右

    # 針對 lower_words 進行左右分類
    feat_txt = []
    mat_txt = []
    
    # 我們需要將 lower_words 重新組裝成行，但這次要考慮 X 座標
    # 簡單做法：逐個 word 判斷
    # 進階做法(採用)：逐行組裝，然後看該行的重心在左邊還是右邊
    
    # 這裡我們重用 words_to_lines 的邏輯，但對每一行計算平均 X
    
    # 手動組裝行
    current_y = -1
    line_buffer = []
    sorted_lower = sorted(lower_words, key=lambda w: (w['top'], w['x0']))
    
    lines_with_pos = []
    for w in sorted_lower:
        if abs(w['top'] - current_y) > 5:
            if line_buffer: lines_with_pos.append(line_buffer)
            line_buffer = []
            current_y = w['top']
        line_buffer.append(w)
    if line_buffer: lines_with_pos.append(line_buffer)

    for row in lines_with_pos:
        # 計算這一行的中心點 X
        avg_x = sum([w['x0'] for w in row]) / len(row)
        line_str = " ".join([w['text'] for w in row])
        
        # 強力過濾顏色代碼 (這是你的痛點)
        if re.search(r"^[A-Z0-9]{2,4}\([A-Za-z0-9\s]+\)", line_str): continue
        if re.search(r"^[A-Z]{2}\s*$", line_str): continue # 單獨的兩個大寫字母
        
        # 分左右
        if avg_x < split_x:
            feat_txt.append(line_str)
        else:
            mat_txt.append(line_str)

    data['Features'] = "\n".join(feat_txt)
    data['Material'] = "\n".join(mat_txt)

    # --- F. 其他資訊補完 ---
    # MSRP
    msrp_match = re.search(r"MSRP\s*[¥￥]?\s*([\d,]+)", full_text, re.IGNORECASE)
    if msrp_match: data['MSRP'] = msrp_match.group(1).replace(',', '')
    
    # Weight
    weight_match = re.search(r"Estimated Average Weight\s*[\n]*\s*(\d+\.?\d*|TBA|ТВА)", full_text, re.IGNORECASE)
    if weight_match: data['Weight (g)'] = weight_match.group(1).replace('ТВА', 'TBA')
    
    # Category
    categories = ["ALPINE CLOTHING", "INSULATION", "THERMAL", "RAIN WEAR", "SOFT SHELL", "PANTS", "BASE LAYER", "FIELD WEAR", "TRAVEL & COUNTRY", "CAP & HAT", "GLOVES", "SOCKS", "SLEEPING BAG", "FOOTWEAR", "BACKPACK", "BAG", "ACCESSORIES", "CYCLING", "SNOW GEAR", "CLIMBING", "FISHING", "PADDLE SPORTS", "DOG GEAR", "KIDS & BABY"]
    for cat in categories:
        if cat in full_text: data['Category'] = cat; break

    return data

def words_to_lines(words):
    """將文字物件列表轉換為純文字行列表 (依 Y 軸分組)"""
    if not words: return []
    # 先按 Y 排序，再按 X 排序
    sorted_words = sorted(words, key=lambda w: (w['top'], w['x0']))
    lines = []
    current_y = -1
    line_buffer = []
    
    for w in sorted_words:
        # 如果 Y 軸差距超過 5，視為換行
        if abs(w['top'] - current_y) > 5:
            if line_buffer:
                lines.append(" ".join([x['text'] for x in line_buffer]))
            line_buffer = []
            current_y = w['top']
        line_buffer.append(w)
    
    if line_buffer:
        lines.append(" ".join([x['text'] for x in line_buffer]))
    return lines

# --- 3. 側邊欄 ---
with st.sidebar:
    st.header("步驟 1: 上傳檔案")
    uploaded_files = st.file_uploader("可多選上傳 PDF", type="pdf", accept_multiple_files=True)
    st.info("Ver 10.0 修正：\n1. 絕對區域鎖定 (Strict Zoning)\n2. 修正品名抓取邏輯 (Style# 上方搜尋)\n3. 徹底分離材質與敘述")

# --- 4. 主畫面 ---
st.title("🏔️ Mont-bell 型錄解析器 Ver 10.0 (絕對區域版)")

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
                        
                        p_data = parse_product_page_v10(page, i + 1)
                        
                        if p_data:
                            p_data['Source File'] = filename
                            all_products.append(p_data)
                            
            except Exception as e:
                st.error(f"Error processing {uploaded_file.name}: {e}")

        my_bar.empty()
        progress_text.empty()

        if all_products:
            df = pd.DataFrame(all_products)
            
            st.success(f"✅ 分析完成！共擷取 {len(df)} 筆資料。")
            
            tab1, tab2 = st.tabs(["📊 資料總表", "🛠️ 原始資料檢視"])
            
            with tab1:
                display_cols = ['Source File', 'Page', 'Category', 'Product Name', 'Style#', 'MSRP', 'Features', 'Material', 'Description']
                st.dataframe(
                    df[display_cols], 
                    use_container_width=True,
                    column_config={
                        "Features": st.column_config.TextColumn("Features (左下)", width="medium"),
                        "Material": st.column_config.TextColumn("Material (右下)", width="medium"),
                        "Description": st.column_config.TextColumn("Description (上方)", width="large"),
                        "Product Name": st.column_config.TextColumn("Product Name", width="medium"),
                    }
                )
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='All_Products')
                excel_data = output.getvalue()
                st.download_button("📥 下載 Excel", data=excel_data, file_name="Montbell_Ver10_Zoning.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

        else:
            st.warning("⚠️ 未擷取到資料。")