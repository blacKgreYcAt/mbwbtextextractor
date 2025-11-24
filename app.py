import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import altair as alt

# --- 1. 頁面設定 ---
st.set_page_config(
    page_title="Mont-bell 型錄解析器 Ver 13.0 (天際線分層版)",
    page_icon="🏔️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. 核心解析邏輯 (Ver 13.0: 橫線上下分層抓取) ---

def find_header_separator_y(page):
    """
    偵測頁面主橫線 (Header Line)。
    回傳 Y 座標。
    """
    try:
        edges = page.edges
        width = page.width
        # 篩選：水平線、夠長、在上半部
        candidates = [
            e for e in edges 
            if e['orientation'] == 'horizontal' 
            and (e['x1'] - e['x0']) > (width * 0.3)
            and e['top'] < (page.height / 2)
        ]
        if not candidates: return 0
        # 找最靠上面的一條 (通常標題下方那條)
        candidates.sort(key=lambda e: e['top'])
        return candidates[0]['bottom'] + 2
    except Exception:
        return 0

def parse_product_page_v13(page, page_num):
    # 1. 基礎資訊
    width = page.width
    height = page.height
    
    # 2. 找到分界橫線
    header_y = find_header_separator_y(page)
    # 如果沒找到線，預設一個頂部 buffer (避免抓到最上面的頁眉)
    if header_y == 0: header_y = height * 0.15 

    # 3. 取得所有文字物件
    all_words = page.extract_words(keep_blank_chars=True, x_tolerance=2, y_tolerance=2)
    
    # 4. 分層過濾 (關鍵步驟!)
    # 上層文字：找品名
    words_above = [w for w in all_words if w['bottom'] <= header_y]
    # 下層文字：找 Style#、特點、材質
    words_below = [w for w in all_words if w['top'] >= header_y]
    
    # 若下層沒字 (可能是空白頁)，跳過
    if not words_below: return None

    # 組裝下層文字供 Regex 搜尋
    text_below = " ".join([w['text'] for w in words_below])
    full_text_raw = page.extract_text() or "" # 用於備用搜尋

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

    # --- A. 抓取 Style# (嚴格限制在下層) ---
    # 策略：在 text_below 中搜尋 7 碼數字
    # 先找 "Style#" 關鍵字附近的
    style_match = re.search(r"Style\s*#?\s*(\d{7})", text_below, re.IGNORECASE)
    if style_match:
        data['Style#'] = style_match.group(1)
    else:
        # 暴力搜下層的 7 碼 (排除 MSRP 附近的)
        candidates = list(re.finditer(r"(?<!\d)(\d{7})(?!\d)", text_below))
        for m in candidates:
            # 簡單檢查周圍有沒有 ¥
            snippet = text_below[max(0, m.start()-10):m.end()+10]
            if "¥" not in snippet and "MSRP" not in snippet:
                data['Style#'] = m.group(1)
                break
    
    if not data['Style#']: return None

    # --- B. 抓取 Product Name (嚴格限制在上層) ---
    # 策略：分析 words_above，過濾掉固定雜訊，剩下的最後一行通常是品名
    
    # 將上層文字轉成行
    lines_above = words_to_lines(words_above)
    
    potential_name = ""
    # 倒敘搜尋 (因為品名通常最靠近橫線)
    for line in reversed(lines_above):
        line = line.strip()
        
        # 雜訊過濾器
        skip_keywords = [
            "mont-bell", "Fall", "Winter", "Spring", "Summer", 
            "CONFIDENTIAL", "KJ", "Item", "Workbook", "Distributor",
            "Page", "Last Updated"
        ]
        is_noise = False
        for kw in skip_keywords:
            if kw.lower() in line.lower(): is_noise = True; break
        
        # 過濾純數字或日期 (e.g. 2024-06-14)
        if re.search(r"\d{4}[-/]\d{2}[-/]\d{2}", line): is_noise = True
        if line.replace(" ", "").isdigit(): is_noise = True
        
        if not is_noise and len(line) > 2:
            potential_name = line
            break
            
    data['Product Name'] = potential_name

    # --- C. 定位下層錨點 (Features/Material) ---
    features_anchor = None
    material_anchor = None
    
    for w in words_below:
        txt = w['text'].strip()
        if txt.startswith("Feature") and features_anchor is None:
            features_anchor = w
        elif txt.startswith("Material") and material_anchor is None:
            material_anchor = w

    # --- D. 抓取 Description (敘述) ---
    # 區域：橫線下方 ~ Features 標題上方
    # 使用 .crop()
    try:
        desc_top = header_y
        desc_bottom = features_anchor['top'] if features_anchor else (height / 3)
        
        # 只有當空間足夠時才抓
        if desc_bottom > desc_top + 10:
            desc_crop = page.crop((0, desc_top, width, desc_bottom))
            desc_text = desc_crop.extract_text() or ""
            
            desc_lines = []
            for line in desc_text.split('\n'):
                line = line.strip()
                # 排除 Style# 行 (雖然它在下方，但有時候會被 crop 進來)
                if data['Style#'] in line: continue
                if "Style" in line: continue
                if "MSRP" in line: continue
                
                # 抓取敘述
                if line.startswith("•") or line.startswith("●"):
                    desc_lines.append(line)
                elif len(line) > 30 and "mont-bell" not in line:
                    desc_lines.append(line)
            data['Description'] = "\n".join(desc_lines)
    except Exception:
        pass

    # --- E. 抓取 Features & Material (Crop 分割) ---
    # 設定區域
    content_top = max(features_anchor['bottom'], material_anchor['bottom']) if (features_anchor and material_anchor) else desc_bottom + 10
    
    # 找底部邊界
    content_bottom = height
    for w in words_below:
        if w['top'] > content_top and w['text'] in ["Size", "Estimated", "Last"]:
            content_bottom = min(content_bottom, w['top'])
            
    # 設定左右分割線 (Material 標題左側)
    split_x = material_anchor['x0'] - 5 if material_anchor else (width / 2)

    try:
        # Features (左下)
        if split_x > 0 and content_bottom > content_top:
            feat_crop = page.crop((0, content_top, split_x, content_bottom))
            feat_raw = feat_crop.extract_text() or ""
            feat_clean = []
            for line in feat_raw.split('\n'):
                if re.search(r"^[A-Z0-9]{2,4}\([A-Za-z0-9\s]+\)", line): continue # 過濾顏色
                if "Material" in line: continue
                feat_clean.append(line.strip())
            data['Features'] = "\n".join(feat_clean)

        # Material (右下)
        if width > split_x and content_bottom > content_top:
            mat_crop = page.crop((split_x, content_top, width, content_bottom))
            mat_raw = mat_crop.extract_text() or ""
            mat_clean = []
            for line in mat_raw.split('\n'):
                if re.search(r"^[A-Z0-9]{2,4}\([A-Za-z0-9\s]+\)", line): continue # 過濾顏色
                if re.search(r"^[A-Z]{2}\s*$", line): continue
                if "Size" in line: break
                mat_clean.append(line.strip())
            data['Material'] = "\n".join(mat_clean)
    except Exception:
        pass

    # --- F. 其他資訊 (MSRP, Weight, Category) ---
    # MSRP, Weight 依然在下層文字找
    msrp_match = re.search(r"MSRP\s*[¥￥]?\s*([\d,]+)", text_below, re.IGNORECASE)
    if msrp_match: data['MSRP'] = msrp_match.group(1).replace(',', '')
    
    weight_match = re.search(r"Estimated Average Weight\s*[\n]*\s*(\d+\.?\d*|TBA|ТВА)", text_below, re.IGNORECASE)
    if weight_match: data['Weight (g)'] = weight_match.group(1).replace('ТВА', 'TBA')
    
    # Category 通常在最上面 (甚至在橫線上面)，所以用 full_text 找
    categories = ["ALPINE CLOTHING", "INSULATION", "THERMAL", "RAIN WEAR", "SOFT SHELL", "PANTS", "BASE LAYER", "FIELD WEAR", "TRAVEL & COUNTRY", "CAP & HAT", "GLOVES", "SOCKS", "SLEEPING BAG", "FOOTWEAR", "BACKPACK", "BAG", "ACCESSORIES", "CYCLING", "SNOW GEAR", "CLIMBING", "FISHING", "PADDLE SPORTS", "DOG GEAR", "KIDS & BABY"]
    for cat in categories:
        if cat in full_text_raw: data['Category'] = cat; break

    return data

def words_to_lines(words):
    if not words: return []
    sorted_words = sorted(words, key=lambda w: (w['top'], w['x0']))
    lines = []
    current_y = -1
    line_buffer = []
    for w in sorted_words:
        if abs(w['top'] - current_y) > 5:
            if line_buffer: lines.append(" ".join([x['text'] for x in line_buffer]))
            line_buffer = []
            current_y = w['top']
        line_buffer.append(w)
    if line_buffer: lines.append(" ".join([x['text'] for x in line_buffer]))
    return lines

# --- 3. 側邊欄 ---
with st.sidebar:
    st.header("步驟 1: 上傳檔案")
    uploaded_files = st.file_uploader("可多選上傳 PDF", type="pdf", accept_multiple_files=True)
    st.info("Ver 13.0 修正：\n1. 橫線上方：專找 Product Name\n2. 橫線下方：專找 Style# (避開圖標編號)\n3. 完美分離 Features/Material")

# --- 4. 主畫面 ---
st.title("🏔️ Mont-bell 型錄解析器 Ver 13.0 (天際線分層版)")

if uploaded_files:
    col1, col2 = st.columns([1, 5])
    with col1:
        start_btn = st.button("🚀 開始分析", type="primary", use_container_width=True)
    
    if start_btn:
        all_products = []
        my_bar = st.progress(0)
        total_pdfs = len(uploaded_files)
        
        for file_idx, uploaded_file in enumerate(uploaded_files):
            try:
                with pdfplumber.open(uploaded_file) as pdf:
                    total_pages = len(pdf.pages)
                    filename = uploaded_file.name
                    for i, page in enumerate(pdf.pages):
                        my_bar.progress((file_idx + (i / total_pages)) / total_pdfs)
                        p_data = parse_product_page_v13(page, i + 1)
                        if p_data:
                            p_data['Source File'] = filename
                            all_products.append(p_data)
            except Exception as e:
                st.error(f"Error: {e}")

        my_bar.empty()
        
        if all_products:
            df = pd.DataFrame(all_products)
            st.success(f"✅ 完成！共擷取 {len(df)} 筆資料。")
            
            tab1, tab2 = st.tabs(["📊 資料總表", "🛠️ 檢查區"])
            with tab1:
                display_cols = ['Source File', 'Page', 'Category', 'Product Name', 'Style#', 'MSRP', 'Features', 'Material', 'Description']
                st.dataframe(
                    df[display_cols], 
                    use_container_width=True,
                    column_config={
                        "Product Name": st.column_config.TextColumn("Product Name (上層)", width="medium"),
                        "Style#": st.column_config.TextColumn("Style# (下層)", width="small"),
                        "Features": st.column_config.TextColumn("Features (左下)", width="medium"),
                        "Material": st.column_config.TextColumn("Material (右下)", width="medium"),
                    }
                )
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='All_Products')
                st.download_button("📥 下載 Excel", data=output.getvalue(), file_name="Montbell_Ver13_Skyline.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        else:
            st.warning("⚠️ 未擷取到資料。")