import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import altair as alt

# --- 1. 頁面設定 ---
st.set_page_config(
    page_title="Mont-bell 型錄解析器 Ver 14.0 (相對定位寬容版)",
    page_icon="🏔️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. 核心解析邏輯 (Ver 14.0: 以 Style# 為錨點的相對定位) ---

def parse_product_page_v14(page, page_num):
    # 1. 基礎資訊
    width = page.width
    height = page.height
    
    # 取得所有文字物件
    words = page.extract_words(keep_blank_chars=True, x_tolerance=2, y_tolerance=2)
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

    # --- A. 抓取 Style# (全頁搜尋 + 高度過濾) ---
    # 策略：找出所有符合 7 碼數字的物件
    # 過濾條件：Y 座標必須大於頁面高度的 10% (避開 Header 圖標編號)
    
    candidate_styles = []
    
    # 1. 先用關鍵字 "Style" 定位
    style_anchor = None
    for w in words:
        if "Style" in w['text'] and w['top'] > (height * 0.1): 
            style_anchor = w
            break
            
    # 2. 如果有找到 "Style" 關鍵字，找它後面的數字
    if style_anchor:
        # 找同一行或附近的數字
        nearby_digits = [w for w in words if abs(w['top'] - style_anchor['top']) < 10 and w['x0'] > style_anchor['x0']]
        for w in nearby_digits:
            if re.match(r"^\d{7}$", w['text']):
                candidate_styles.append(w)
                break
    
    # 3. 如果沒找到，暴力搜全頁符合條件的數字
    if not candidate_styles:
        for w in words:
            # 條件：是 7 碼數字 AND 位置不在最頂端
            if re.match(r"^\d{7}$", w['text']) and w['top'] > (height * 0.1):
                # 排除疑似價格的 (周圍有 ¥)
                # 這裡簡單判斷：通常 Style# 不會太靠右邊 (價格通常靠右，或者在 Style# 下方)
                candidate_styles.append(w)
    
    if not candidate_styles: return None # 沒救了
    
    # 取第一個候選者當作 Style# (通常是最上面的那個)
    final_style_obj = candidate_styles[0]
    data['Style#'] = final_style_obj['text']
    
    # 設定 Style# 的 Y 座標為基準線
    style_y = final_style_obj['top']

    # --- B. 抓取 Product Name (往上找) ---
    # 策略：找出位於 Style# 上方 (bottom <= style_y) 的所有文字行
    # 倒敘排列 (離 Style# 最近的先檢查)
    
    # 篩選上方文字
    words_above = [w for w in words if w['bottom'] <= style_y + 5] # +5 容許同一行
    lines_above = words_to_lines(words_above)
    
    potential_name = ""
    
    for line in reversed(lines_above):
        line = line.strip()
        
        # 雜訊過濾
        if data['Style#'] in line: continue # 跳過 Style# 本身
        if "Style" in line: continue
        
        # 頁眉雜訊
        skip_keywords = [
            "mont-bell", "Fall", "Winter", "Spring", "Summer", 
            "CONFIDENTIAL", "KJ", "Item", "Workbook", "Distributor",
            "Page", "Last Updated", "MSRP", "¥"
        ]
        is_noise = False
        for kw in skip_keywords:
            if kw.lower() in line.lower(): is_noise = True; break
        
        # 顏色代碼過濾
        if re.search(r"^[A-Z]{2,3}\(.*\)", line): is_noise = True
        
        # 純數字/日期過濾
        if re.search(r"\d{4}[-/]\d{2}[-/]\d{2}", line): is_noise = True
        if line.replace(" ", "").isdigit(): is_noise = True
        
        if not is_noise and len(line) > 2:
            potential_name = line
            break
            
    # 如果找到的品名包含了 Style (例如 "Jacket Style#..."), 修剪它
    if "Style" in potential_name:
        potential_name = potential_name.split("Style")[0].strip()
        
    data['Product Name'] = potential_name

    # --- C. 定位 Features / Material 錨點 ---
    features_anchor = None
    material_anchor = None
    
    # 只在 Style# 下方找
    for w in words:
        if w['top'] > style_y:
            txt = w['text'].strip()
            if txt.startswith("Feature") and features_anchor is None:
                features_anchor = w
            elif txt.startswith("Material") and material_anchor is None:
                material_anchor = w

    # --- D. 抓取 Description (Style# 下方 ~ Features 上方) ---
    try:
        desc_top = style_y + 10 # Style# 下方一點點
        desc_bottom = features_anchor['top'] if features_anchor else (height / 2)
        
        if desc_bottom > desc_top:
            desc_crop = page.crop((0, desc_top, width, desc_bottom))
            desc_text = desc_crop.extract_text() or ""
            
            desc_lines = []
            for line in desc_text.split('\n'):
                line = line.strip()
                if "MSRP" in line: continue
                if "Style" in line: continue
                
                if line.startswith("•") or line.startswith("●"):
                    desc_lines.append(line)
                elif len(line) > 30 and "mont-bell" not in line:
                    desc_lines.append(line)
            data['Description'] = "\n".join(desc_lines)
    except Exception:
        pass

    # --- E. 抓取 Features & Material (左右分割) ---
    # 繼承之前的成功邏輯
    content_top = max(features_anchor['bottom'], material_anchor['bottom']) if (features_anchor and material_anchor) else desc_bottom + 10
    
    # 找底部邊界
    content_bottom = height
    for w in words:
        if w['top'] > content_top and w['text'] in ["Size", "Estimated", "Last"]:
            content_bottom = min(content_bottom, w['top'])
            
    split_x = material_anchor['x0'] - 5 if material_anchor else (width / 2)

    try:
        # Features (左)
        if split_x > 0 and content_bottom > content_top:
            feat_crop = page.crop((0, content_top, split_x, content_bottom))
            feat_raw = feat_crop.extract_text() or ""
            feat_clean = []
            for line in feat_raw.split('\n'):
                if re.search(r"^[A-Z0-9]{2,4}\([A-Za-z0-9\s]+\)", line): continue
                if "Material" in line: continue
                feat_clean.append(line.strip())
            data['Features'] = "\n".join(feat_clean)

        # Material (右)
        if width > split_x and content_bottom > content_top:
            mat_crop = page.crop((split_x, content_top, width, content_bottom))
            mat_raw = mat_crop.extract_text() or ""
            mat_clean = []
            for line in mat_raw.split('\n'):
                if re.search(r"^[A-Z0-9]{2,4}\([A-Za-z0-9\s]+\)", line): continue
                if re.search(r"^[A-Z]{2}\s*$", line): continue
                if "Size" in line: break
                mat_clean.append(line.strip())
            data['Material'] = "\n".join(mat_clean)
    except Exception:
        pass

    # --- F. 其他資訊 ---
    full_text = page.extract_text() or ""
    msrp_match = re.search(r"MSRP\s*[¥￥]?\s*([\d,]+)", full_text, re.IGNORECASE)
    if msrp_match: data['MSRP'] = msrp_match.group(1).replace(',', '')
    
    weight_match = re.search(r"Estimated Average Weight\s*[\n]*\s*(\d+\.?\d*|TBA|ТВА)", full_text, re.IGNORECASE)
    if weight_match: data['Weight (g)'] = weight_match.group(1).replace('ТВА', 'TBA')
    
    categories = ["ALPINE CLOTHING", "INSULATION", "THERMAL", "RAIN WEAR", "SOFT SHELL", "PANTS", "BASE LAYER", "FIELD WEAR", "TRAVEL & COUNTRY", "CAP & HAT", "GLOVES", "SOCKS", "SLEEPING BAG", "FOOTWEAR", "BACKPACK", "BAG", "ACCESSORIES", "CYCLING", "SNOW GEAR", "CLIMBING", "FISHING", "PADDLE SPORTS", "DOG GEAR", "KIDS & BABY"]
    for cat in categories:
        if cat in full_text: data['Category'] = cat; break

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
    st.info("Ver 14.0 寬容版：\n1. 放棄橫線強制偵測，改用相對位置\n2. Style# 定位：排除頁面頂端 10% 即可\n3. 品名定位：Style# 往上找最近的一行")

# --- 4. 主畫面 ---
st.title("🏔️ Mont-bell 型錄解析器 Ver 14.0 (寬容定位版)")

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
                        p_data = parse_product_page_v14(page, i + 1)
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
                        "Style#": st.column_config.TextColumn("Style#", width="small"),
                        "Features": st.column_config.TextColumn("Features (左下)", width="medium"),
                        "Material": st.column_config.TextColumn("Material (右下)", width="medium"),
                    }
                )
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='All_Products')
                st.download_button("📥 下載 Excel", data=output.getvalue(), file_name="Montbell_Ver14_Relative.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        else:
            st.warning("⚠️ 未擷取到資料。")