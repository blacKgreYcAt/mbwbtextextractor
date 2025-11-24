import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import altair as alt

# --- 1. 頁面設定 ---
st.set_page_config(
    page_title="Mont-bell 型錄解析器 Ver 11.0 (手術刀切割版)",
    page_icon="🏔️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. 核心解析邏輯 (Ver 11.0: 使用 .crop() 物理分離左右欄) ---
def parse_product_page_v11(page, page_num):
    """
    使用 page.crop() 針對特定區域進行獨立文字萃取，徹底解決左右欄混合問題。
    """
    # 1. 取得頁面基礎資訊
    width = page.width
    height = page.height
    full_text = page.extract_text() or ""
    
    # 擷取單字物件用於定位
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

    # --- A. 定位關鍵錨點 (Anchors) ---
    style_anchor = None
    features_anchor = None
    material_anchor = None
    
    for w in words:
        txt = w['text'].strip()
        # 找 Style#
        if "Style" in txt and style_anchor is None:
            style_anchor = w
        # 找標題
        if txt.startswith("Feature") and features_anchor is None:
            features_anchor = w
        elif txt.startswith("Material") and material_anchor is None:
            material_anchor = w

    # --- B. 抓取 Style# (優先使用全文 Regex) ---
    # 這是最穩的方法
    style_regex = re.search(r"Style\s*#?\s*(\d{7})", full_text, re.IGNORECASE)
    if style_regex:
        data['Style#'] = style_regex.group(1)
    else:
        # 暴力搜尋 7 碼
        candidates = list(re.finditer(r"(?<!\d)(\d{7})(?!\d)", full_text))
        for m in candidates:
            if "¥" not in full_text[max(0, m.start()-10):m.end()+10]:
                data['Style#'] = m.group(1)
                break
    
    if not data['Style#']: return None

    # --- C. 定義「上方區域」 (Header Section) ---
    # 分界線：Features 標題的上方 (如果沒找到，就抓頁面 1/3 處)
    split_y_top = features_anchor['top'] if features_anchor else (height / 3)
    
    # C-1. 抓取 Product Name (品名)
    # 策略：鎖定 Style# 座標，往上找
    if style_anchor:
        style_top = style_anchor['top']
        # 篩選出位於 Style# 上方且在同一區塊的文字
        potential_lines = [w for w in words if w['bottom'] <= style_top + 5] # +5 容許誤差
        # 轉成行
        header_lines = words_to_lines(potential_lines)
        
        # 倒敘搜尋 (離 Style# 最近的)
        found_name = ""
        for line in reversed(header_lines):
            line = line.strip()
            # 雜訊過濾
            if "Style" in line: continue # 跳過 Style# 本身行
            if any(x in line for x in ["mont-bell", "Fall", "Winter", "CONFIDENTIAL", "KJ", "MSRP", "¥"]): continue
            if re.search(r"^[A-Z]{2,3}\(.*\)", line): continue # 顏色代碼
            if line.isdigit(): continue
            
            if len(line) > 2:
                found_name = line
                break
        data['Product Name'] = found_name
    
    # 若上方找不到，試試看 Style# 同一行
    if not data['Product Name'] and style_anchor:
        # 找出與 Style# 差不多高度的文字
        same_line_words = [w['text'] for w in words if abs(w['top'] - style_anchor['top']) < 5]
        line_str = " ".join(same_line_words)
        if "Style" in line_str:
            pre_text = line_str.split("Style")[0].strip()
            if len(pre_text) > 3: data['Product Name'] = pre_text

    # C-2. 抓取 Description (敘述)
    # 範圍：Page Top ~ Features Header Top
    # 使用 .crop() 抓取上方純文字，避免格式干擾
    try:
        header_box = (0, 0, width, split_y_top)
        header_crop = page.crop(header_box)
        header_text = header_crop.extract_text() or ""
        
        desc_lines = []
        for line in header_text.split('\n'):
            line = line.strip()
            if line.startswith("•") or line.startswith("●"):
                desc_lines.append(line)
            # 補抓長敘述
            elif len(line) > 40 and "Style" not in line and "MSRP" not in line and data['Product Name'] not in line:
                if "mont-bell" not in line:
                    desc_lines.append(line)
        data['Description'] = "\n".join(desc_lines)
    except Exception:
        pass # Crop 失敗就跳過

    # --- D. 定義「下方區域」 (Features & Material) - 手術刀切割 ---
    
    # 1. 確定 Y 軸範圍
    # 上界：標題底部
    top_y = max(features_anchor['bottom'], material_anchor['bottom']) if (features_anchor and material_anchor) else split_y_top + 10
    
    # 下界：找到 "Size" 或 "Estimated" 的位置
    bottom_y = height
    for w in words:
        if w['top'] > top_y and w['text'] in ["Size", "Estimated", "Last"]:
            bottom_y = min(bottom_y, w['top'])
    
    # 2. 確定 X 軸切割線
    # 以 Material 標題的左邊界為準，稍微往左留一點 buffer (例如 5px)
    split_x = material_anchor['x0'] - 5 if material_anchor else (width / 2)

    # 3. 執行切割與萃取 (Crucial Step!)
    try:
        # --- 左邊：Features ---
        # 範圍：(0, top_y, split_x, bottom_y)
        # 檢查座標合法性
        if split_x > 0 and bottom_y > top_y:
            feat_box = (0, top_y, split_x, bottom_y)
            feat_crop = page.crop(feat_box)
            # 使用 layout=True 嘗試保持格式，或預設
            feat_raw = feat_crop.extract_text() or ""
            
            # 清洗 Features 文字
            feat_clean = []
            for line in feat_raw.split('\n'):
                if re.search(r"^[A-Z0-9]{2,4}\([A-Za-z0-9\s]+\)", line): continue # 顏色代碼
                if "Material" in line: continue # 標題誤入
                feat_clean.append(line.strip())
            data['Features'] = "\n".join(feat_clean)

        # --- 右邊：Material ---
        # 範圍：(split_x, top_y, width, bottom_y)
        if width > split_x and bottom_y > top_y:
            mat_box = (split_x, top_y, width, bottom_y)
            mat_crop = page.crop(mat_box)
            mat_raw = mat_crop.extract_text() or ""
            
            # 清洗 Material 文字
            mat_clean = []
            for line in mat_raw.split('\n'):
                if re.search(r"^[A-Z0-9]{2,4}\([A-Za-z0-9\s]+\)", line): continue # 顏色代碼
                if re.search(r"^[A-Z]{2}\s*$", line): continue
                if "Size" in line: break # 碰到 Size 停止
                mat_clean.append(line.strip())
            data['Material'] = "\n".join(mat_clean)

    except Exception as e:
        # 如果 crop 失敗 (例如座標錯誤)，不讓程式崩潰，保留空白
        print(f"Crop error: {e}")

    # --- E. 其他資訊 ---
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
    st.info("Ver 11.0 修正：\n使用 .crop() 技術物理分割左右欄位，保證特點與材質資料絕不混合。")

# --- 4. 主畫面 ---
st.title("🏔️ Mont-bell 型錄解析器 Ver 11.0 (手術刀切割版)")

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
                        p_data = parse_product_page_v11(page, i + 1)
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
                        "Features": st.column_config.TextColumn("Features (左)", width="medium"),
                        "Material": st.column_config.TextColumn("Material (右)", width="medium"),
                        "Description": st.column_config.TextColumn("Description (上)", width="large"),
                    }
                )
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='All_Products')
                st.download_button("📥 下載 Excel", data=output.getvalue(), file_name="Montbell_Ver11_Crop.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        else:
            st.warning("⚠️ 未擷取到資料。")