import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import altair as alt

# --- 1. 頁面設定 (iOS 風格配置) ---
st.set_page_config(
    page_title="Mont-bell原廠工作本文字解析精靈",
    page_icon="🏔️",
    layout="wide",
    initial_sidebar_state="collapsed" # 隱藏側邊欄
)

# --- 2. iOS 風格 CSS 注入 ---
st.markdown("""
<style>
    /* 全局背景色 - iOS Light Gray */
    .stApp {
        background-color: #F2F2F7;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    
    /* 標題樣式 */
    h1 {
        color: #1C1C1E;
        font-weight: 700;
        text-align: center;
        padding-top: 20px;
        padding-bottom: 10px;
    }
    
    /* 卡片容器樣式 */
    .ios-card {
        background-color: #FFFFFF;
        padding: 30px;
        border-radius: 20px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
        margin-bottom: 25px;
    }

    /* 上傳區塊優化 */
    .stFileUploader {
        padding: 10px;
    }

    /* 按鈕樣式 - iOS Blue */
    div.stButton > button {
        background-color: #007AFF;
        color: white;
        border: none;
        border-radius: 14px;
        padding: 12px 24px;
        font-size: 16px;
        font-weight: 600;
        width: 100%;
        transition: all 0.2s ease;
    }
    div.stButton > button:hover {
        background-color: #005BB5;
        box-shadow: 0 4px 12px rgba(0, 122, 255, 0.3);
        border: none;
        color: white;
    }
    div.stButton > button:active {
        transform: scale(0.98);
    }

    /* 進度條顏色 */
    .stProgress > div > div > div > div {
        background-color: #007AFF;
    }
    
    /* 下載按鈕 (綠色) */
    .download-btn {
        background-color: #34C759 !important;
    }
    
    /* 隱藏預設的主選單漢堡 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
</style>
""", unsafe_allow_html=True)

# --- 3. 核心解析邏輯 (Ver 14.0 邏輯保持不變) ---

def parse_product_page_v14(page, page_num):
    width = page.width
    height = page.height
    words = page.extract_words(keep_blank_chars=True, x_tolerance=2, y_tolerance=2)
    if not words: return None

    data = {
        'Page': page_num, 'Category': 'Uncategorized', 'Product Name': '', 
        'Style#': '', 'MSRP': '0', 'Weight (g)': '', 
        'Features': '', 'Material': '', 'Description': ''
    }

    # A. Style#
    candidate_styles = []
    style_anchor = None
    for w in words:
        if "Style" in w['text'] and w['top'] > (height * 0.1): 
            style_anchor = w; break
            
    if style_anchor:
        nearby = [w for w in words if abs(w['top'] - style_anchor['top']) < 10 and w['x0'] > style_anchor['x0']]
        for w in nearby:
            if re.match(r"^\d{7}$", w['text']): candidate_styles.append(w); break
    
    if not candidate_styles:
        for w in words:
            if re.match(r"^\d{7}$", w['text']) and w['top'] > (height * 0.1):
                candidate_styles.append(w)
    
    if not candidate_styles: return None
    final_style_obj = candidate_styles[0]
    data['Style#'] = final_style_obj['text']
    style_y = final_style_obj['top']

    # B. Product Name
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

    words_above = [w for w in words if w['bottom'] <= style_y + 5]
    lines_above = words_to_lines(words_above)
    potential_name = ""
    for line in reversed(lines_above):
        line = line.strip()
        if data['Style#'] in line or "Style" in line: continue
        skip_keywords = ["mont-bell", "Fall", "Winter", "Spring", "Summer", "CONFIDENTIAL", "KJ", "Item", "Workbook", "Distributor", "Page", "Last Updated", "MSRP", "¥"]
        is_noise = False
        for kw in skip_keywords:
            if kw.lower() in line.lower(): is_noise = True; break
        if re.search(r"^[A-Z]{2,3}\(.*\)", line): is_noise = True
        if re.search(r"\d{4}[-/]\d{2}[-/]\d{2}", line): is_noise = True
        if line.replace(" ", "").isdigit(): is_noise = True
        
        if not is_noise and len(line) > 2:
            potential_name = line; break
            
    if "Style" in potential_name: potential_name = potential_name.split("Style")[0].strip()
    data['Product Name'] = potential_name

    # C. Anchors
    features_anchor = None
    material_anchor = None
    for w in words:
        if w['top'] > style_y:
            txt = w['text'].strip()
            if txt.startswith("Feature") and features_anchor is None: features_anchor = w
            elif txt.startswith("Material") and material_anchor is None: material_anchor = w

    # D. Description
    try:
        desc_top = style_y + 10
        desc_bottom = features_anchor['top'] if features_anchor else (height / 2)
        if desc_bottom > desc_top:
            desc_crop = page.crop((0, desc_top, width, desc_bottom))
            desc_text = desc_crop.extract_text() or ""
            desc_lines = []
            for line in desc_text.split('\n'):
                line = line.strip()
                if "MSRP" in line or "Style" in line: continue
                if line.startswith("•") or line.startswith("●"): desc_lines.append(line)
                elif len(line) > 30 and "mont-bell" not in line: desc_lines.append(line)
            data['Description'] = "\n".join(desc_lines)
    except: pass

    # E. Features & Material
    content_top = max(features_anchor['bottom'], material_anchor['bottom']) if (features_anchor and material_anchor) else desc_bottom + 10
    content_bottom = height
    for w in words:
        if w['top'] > content_top and w['text'] in ["Size", "Estimated", "Last"]:
            content_bottom = min(content_bottom, w['top'])
    split_x = material_anchor['x0'] - 5 if material_anchor else (width / 2)

    try:
        if split_x > 0 and content_bottom > content_top:
            feat_crop = page.crop((0, content_top, split_x, content_bottom))
            feat_raw = feat_crop.extract_text() or ""
            feat_clean = []
            for line in feat_raw.split('\n'):
                if re.search(r"^[A-Z0-9]{2,4}\([A-Za-z0-9\s]+\)", line): continue
                if "Material" in line: continue
                feat_clean.append(line.strip())
            data['Features'] = "\n".join(feat_clean)

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
    except: pass

    # F. Others
    full_text = page.extract_text() or ""
    msrp_match = re.search(r"MSRP\s*[¥￥]?\s*([\d,]+)", full_text, re.IGNORECASE)
    if msrp_match: data['MSRP'] = msrp_match.group(1).replace(',', '')
    weight_match = re.search(r"Estimated Average Weight\s*[\n]*\s*(\d+\.?\d*|TBA|ТВА)", full_text, re.IGNORECASE)
    if weight_match: data['Weight (g)'] = weight_match.group(1).replace('ТВА', 'TBA')
    categories = ["ALPINE CLOTHING", "INSULATION", "THERMAL", "RAIN WEAR", "SOFT SHELL", "PANTS", "BASE LAYER", "FIELD WEAR", "TRAVEL & COUNTRY", "CAP & HAT", "GLOVES", "SOCKS", "SLEEPING BAG", "FOOTWEAR", "BACKPACK", "BAG", "ACCESSORIES", "CYCLING", "SNOW GEAR", "CLIMBING", "FISHING", "PADDLE SPORTS", "DOG GEAR", "KIDS & BABY"]
    for cat in categories:
        if cat in full_text: data['Category'] = cat; break

    return data

# --- 4. 前端介面佈局 (UI Layout) ---

# 標題區
st.title("Mont-bell原廠工作本文字解析精靈")
st.markdown("<div style='text-align: center; color: #8E8E93; margin-bottom: 20px;'>Ver 14.0 Stable • iOS Design</div>", unsafe_allow_html=True)

# 內容容器
col_center_layout, = st.columns([1]) # 使用單一列容器

with col_center_layout:
    # --- 卡片 1: 上傳與操作 ---
    st.markdown('<div class="ios-card">', unsafe_allow_html=True)
    st.subheader("📂 匯入資料")
    
    # 檔案上傳元件
    uploaded_files = st.file_uploader("上傳檔案", type="pdf", accept_multiple_files=True, label_visibility="visible")
    
    # 開始按鈕 (如果有檔案才顯示)
    start_btn = False
    if uploaded_files:
        st.write("") # Spacer
        start_btn = st.button("開始解析 (Start Analysis)", type="primary")
    
    st.markdown('</div>', unsafe_allow_html=True)

    # --- 邏輯處理與結果顯示 ---
    if start_btn and uploaded_files:
        
        # --- 卡片 2: 進度狀態 ---
        st.markdown('<div class="ios-card">', unsafe_allow_html=True)
        st.subheader("⏳ 處理進度")
        
        all_products = []
        my_bar = st.progress(0)
        status_text = st.empty()
        
        total_pdfs = len(uploaded_files)
        
        for file_idx, uploaded_file in enumerate(uploaded_files):
            try:
                with pdfplumber.open(uploaded_file) as pdf:
                    total_pages = len(pdf.pages)
                    filename = uploaded_file.name
                    for i, page in enumerate(pdf.pages):
                        # 更新進度
                        progress = (file_idx + (i / total_pages)) / total_pdfs
                        my_bar.progress(progress)
                        status_text.caption(f"正在分析: {filename} ... 第 {i+1} / {total_pages} 頁")
                        
                        p_data = parse_product_page_v14(page, i + 1)
                        if p_data:
                            p_data['Source File'] = filename
                            all_products.append(p_data)
            except Exception as e:
                st.error(f"檔案 {filename} 解析錯誤: {e}")

        my_bar.progress(100)
        status_text.success("✅ 解析完成！")
        st.markdown('</div>', unsafe_allow_html=True)

        # --- 卡片 3: 結果與下載 ---
        if all_products:
            df = pd.DataFrame(all_products)
            
            st.markdown('<div class="ios-card">', unsafe_allow_html=True)
            st.subheader("📊 解析結果")
            
            # 建立 Tab 分頁
            tab1, tab2 = st.tabs(["資料總覽", "統計圖表"])
            
            with tab1:
                display_cols = ['Source File', 'Page', 'Category', 'Product Name', 'Style#', 'MSRP', 'Features', 'Material', 'Description']
                st.dataframe(
                    df[display_cols], 
                    use_container_width=True,
                    column_config={
                        "Product Name": st.column_config.TextColumn("品名", width="medium"),
                        "Style#": st.column_config.TextColumn("型號", width="small"),
                        "Features": st.column_config.TextColumn("產品特點", width="medium"),
                        "Material": st.column_config.TextColumn("材質", width="medium"),
                    }
                )
                
                # Excel 下載
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='All_Products')
                
                st.write("")
                st.download_button(
                    label="📥 下載 Excel 報表",
                    data=output.getvalue(),
                    file_name="Montbell_Export_iOS.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary" # 在 CSS 中會被綠色覆蓋
                )

            with tab2:
                # 簡單圖表
                df['MSRP'] = pd.to_numeric(df['MSRP'], errors='coerce').fillna(0)
                chart = alt.Chart(df).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(
                    x=alt.X('Category', sort='-y', title='產品類別'),
                    y=alt.Y('count()', title='數量'),
                    color=alt.value("#007AFF")
                ).properties(height=300)
                st.altair_chart(chart, use_container_width=True)
            
            st.markdown('</div>', unsafe_allow_html=True)
            
        else:
            st.warning("⚠️ 未在檔案中發現符合格式的資料。")

# 頁尾
st.markdown("<div style='text-align: center; color: #C7C7CC; margin-top: 50px; font-size: 12px;'>Designed for Mont-bell Workbook Automation</div>", unsafe_allow_html=True)