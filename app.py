import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import altair as alt

# --- 1. 頁面設定 ---
st.set_page_config(
    page_title="Mont-bell 型錄數位化儀表板 Ver 6.0",
    page_icon="🏔️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. 核心解析邏輯 (Ver 6.0: 修復誤殺 Bug + 同行名稱偵測) ---
def parse_product_page(text, page_num):
    data = {}
    data['Page'] = page_num
    
    # 預處理：移除多餘空白，統一換行
    clean_text = text.replace('\r\n', '\n')
    lines = [l.strip() for l in clean_text.split('\n') if l.strip()]

    # --- A. 精準定位主 Style# ---
    primary_style_index = -1
    primary_style_num = ""

    for i, line in enumerate(lines):
        # 尋找 Style#
        match = re.search(r"Style#\s*(\d+)", line, re.IGNORECASE)
        if match:
            # 【Ver 6.0 修正】: 只有當 "(style#" 緊接著出現時才視為 Western Size
            # 避免誤殺像 "Style# 1101 (Men's)" 這樣的正常產品
            if re.search(r"\(\s*style#", line, re.IGNORECASE):
                continue
                
            primary_style_index = i
            primary_style_num = match.group(1)
            break # 找到第一個合格的就鎖定
    
    if primary_style_index == -1:
        return None # 這頁真的沒有產品
    
    data['Style#'] = primary_style_num

    # --- B. 產品名稱 (雙重策略) ---
    product_name = ""
    
    # 策略 1: 檢查 Style# 同一行前方是否有字 (e.g. "Alpine Jacket Style# 1101")
    style_line = lines[primary_style_index]
    # 移除 Style# 後面的部分，看剩下什麼
    pre_style_text = re.split(r"Style#", style_line, flags=re.IGNORECASE)[0].strip()
    
    # 過濾掉常見雜訊 (如 "NEW", "REVISED")
    if pre_style_text and len(pre_style_text) > 3 and pre_style_text not in ["NEW", "REVISED"]:
        product_name = pre_style_text
    
    # 策略 2: 如果同一行沒東西，才往上找 (Ver 5.0 的邏輯)
    if not product_name and primary_style_index > 0:
        for k in range(primary_style_index - 1, -1, -1):
            curr = lines[k]
            
            # 排除雜訊清單
            skip_keywords = [
                "mont-bell", "Fall", "Winter", "NEW", "REVISED", 
                "MSRP", "¥", "CONFIDENTIAL", "Western", "Available",
                "Fabric Sample", "Men's", "Women's", "Kid's", "Baby's"
            ]
            
            is_noise = False
            for kw in skip_keywords:
                if kw.lower() in curr.lower(): is_noise = True; break
            
            # 排除顏色代碼行 (e.g. "BK(Black)")
            if re.search(r"^[A-Z]{2,3}\(.*\)$", curr): is_noise = True
            
            if not is_noise:
                product_name = curr
                break
                
    data['Product Name'] = product_name

    # --- C. 價格與重量 ---
    price_match = re.search(r"MSRP\s*[¥￥]?\s*([\d,]+)", clean_text, re.IGNORECASE)
    alt_price = re.search(r"[¥￥]\s*([\d,]+)", clean_text) # 備用：抓取單獨的價格
    
    if price_match:
        data['MSRP'] = price_match.group(1).replace(',', '')
    elif alt_price:
        data['MSRP'] = alt_price.group(1).replace(',', '')
    else:
        data['MSRP'] = "0"

    weight_match = re.search(r"Estimated Average Weight\s*[\n]*\s*(\d+\.?\d*|TBA|ТВА)", clean_text, re.IGNORECASE)
    if weight_match:
        data['Weight (g)'] = weight_match.group(1).replace('ТВА', 'TBA')
    else:
        data['Weight (g)'] = ""

    # --- D. Features (區塊抓取) ---
    features_list = []
    is_collecting_features = False
    # 這些關鍵字出現代表 Features 區塊結束
    stop_keywords = ["Material", "Size", "Estimated", "Last Updated", "CONFIDENTIAL"]

    for line in lines:
        if line.strip() == "Features":
            is_collecting_features = True
            continue
        
        if is_collecting_features:
            # 檢查是否撞到停止詞
            if any(line.startswith(kw) for kw in stop_keywords): break
            
            # 過濾顏色代碼
            if re.search(r"^[A-Z0-9]{2,4}\([A-Za-z0-9\s]+\)", line): continue
            
            features_list.append(line)

    data['Features'] = "\n".join(features_list)

    # --- E. Material (區塊抓取 + 強力過濾) ---
    material_list = []
    is_collecting_material = False
    
    for line in lines:
        if line.strip() == "Material":
            is_collecting_material = True
            continue
        
        if is_collecting_material:
            # 1. 檢查結束條件
            if any(line.startswith(kw) for kw in ["Size", "Estimated", "Last Updated", "CONFIDENTIAL"]):
                break
            
            # 2. 檢查是否為尺寸列表 (強烈訊號)
            if re.search(r"^[XSML\s,]+$", line) or "Size" in line: 
                break

            # 3. 過濾顏色代碼 (例如 BK(Black), NV(Navy))
            # 邏輯：開頭是大寫英文(2-4碼)緊接左括號
            if re.search(r"^[A-Z0-9]{2,4}\(", line): 
                continue
            
            # 4. 過濾多個顏色併排 (例如 "BL(Blue) RD(Red)")
            if re.search(r"\)\s+[A-Z]{2,3}\(", line):
                continue

            material_list.append(line)

    data['Material'] = "\n".join(material_list)

    # --- F. Category ---
    categories = ["ALPINE CLOTHING", "INSULATION", "THERMAL", "RAIN WEAR", "SOFT SHELL", "PANTS", "BASE LAYER", "FIELD WEAR", "TRAVEL & COUNTRY", "CAP & HAT", "GLOVES", "SOCKS", "SLEEPING BAG", "FOOTWEAR", "BACKPACK", "BAG", "ACCESSORIES", "CYCLING", "SNOW GEAR", "CLIMBING", "FISHING", "PADDLE SPORTS", "DOG GEAR", "KIDS & BABY"]
    data['Category'] = "Uncategorized"
    for cat in categories:
        if cat in clean_text: data['Category'] = cat; break

    return data

# --- 3. 側邊欄介面 ---
with st.sidebar:
    st.header("步驟 1: 上傳檔案")
    uploaded_file = st.file_uploader("請選擇 Mont-bell PDF 型錄", type="pdf")
    st.info("Ver 6.0 修正重點：\n1. 修復誤刪資料 Bug\n2. 修正 Style# 抓取順序\n3. 強化顏色過濾")

# --- 4. 主畫面介面 ---
st.title("🏔️ Mont-bell 產品型錄數位化儀表板 (Ver 6.0 終極版)")

if uploaded_file is not None:
    col1, col2 = st.columns([1, 5])
    with col1:
        start_btn = st.button("🚀 開始分析 PDF", type="primary", use_container_width=True)
    
    if start_btn:
        products = []
        progress_text = "正在啟動 PDF 引擎..."
        my_bar = st.progress(0, text=progress_text)
        
        try:
            with pdfplumber.open(uploaded_file) as pdf:
                total_pages = len(pdf.pages)
                
                for i, page in enumerate(pdf.pages):
                    percent = int((i + 1) / total_pages * 100)
                    my_bar.progress(percent, text=f"正在分析第 {i+1}/{total_pages} 頁... (已擷取 {len(products)} 項產品)")
                    
                    text = page.extract_text()
                    if not text: continue
                    
                    p_data = parse_product_page(text, i + 1)
                    if p_data:
                        products.append(p_data)
            
            my_bar.empty()

            if products:
                df = pd.DataFrame(products)
                df['MSRP'] = pd.to_numeric(df['MSRP'], errors='coerce').fillna(0)
                
                st.success(f"✅ 分析完成！共擷取 **{len(products)}** 項產品資料。")
                
                kpi1, kpi2, kpi3, kpi4 = st.columns(4)
                kpi1.metric("總產品數", f"{len(df)} 件")
                kpi2.metric("產品類別數", f"{df['Category'].nunique()} 類")
                avg_price = df[df['MSRP'] > 0]['MSRP'].mean()
                kpi3.metric("平均單價", f"¥{avg_price:,.0f}")
                kpi4.metric("資料來源", f"{total_pages} 頁")
                
                st.markdown("---")

                tab1, tab2 = st.tabs(["📊 視覺化分析", "📋 詳細資料表 & 下載"])
                
                with tab1:
                    st.subheader("📦 產品類別分佈")
                    chart_data = df['Category'].value_counts().reset_index()
                    chart_data.columns = ['Category', 'Count']
                    bar_chart = alt.Chart(chart_data).mark_bar().encode(
                        x=alt.X('Category', sort='-y', title='產品類別'),
                        y=alt.Y('Count', title='產品數量'),
                        color=alt.Color('Category', legend=None, scale=alt.Scale(scheme='tableau20')),
                        tooltip=['Category', 'Count']
                    ).properties(height=400)
                    st.altair_chart(bar_chart, use_container_width=True)

                with tab2:
                    st.subheader("詳細資料清單")
                    display_cols = ['Page', 'Category', 'Product Name', 'Style#', 'MSRP', 'Weight (g)', 'Material', 'Features']
                    st.dataframe(
                        df[display_cols], 
                        use_container_width=True,
                        column_config={
                            "Features": st.column_config.TextColumn("產品特點", width="large"),
                            "Material": st.column_config.TextColumn("材質", width="medium"),
                        }
                    )
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='Products')
                    excel_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 下載 Excel 報表",
                        data=excel_data,
                        file_name="Montbell_Product_List_v6.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
            else:
                st.warning("⚠️ 未發現符合格式的資料。")

        except Exception as e:
            st.error(f"❌ 發生錯誤: {e}")