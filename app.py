import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import altair as alt

# --- 1. 頁面設定 (必須在程式最開頭) ---
st.set_page_config(
    page_title="Mont-bell 型錄數位化儀表板",
    page_icon="🏔️",
    layout="wide",  # 使用寬版面，讓表格和圖表更清楚
    initial_sidebar_state="expanded"
)

# --- 2. 核心解析邏輯 (維持 Ver 3.0 的精準度) ---
def parse_product_page(text, page_num):
    data = {}
    data['Page'] = page_num
    
    # 預處理
    clean_text = text.replace('\r\n', '\n')

    # 基礎資訊
    style_match = re.search(r"Style#\s*(\d+)", clean_text, re.IGNORECASE)
    if not style_match: return None
    data['Style#'] = style_match.group(1)

    price_match = re.search(r"MSRP\s*[¥￥]?\s*([\d,]+)", clean_text, re.IGNORECASE)
    alt_price = re.search(r"[¥￥]\s*([\d,]+)", clean_text)
    if price_match: data['MSRP'] = price_match.group(1).replace(',', '')
    elif alt_price: data['MSRP'] = alt_price.group(1).replace(',', '')
    else: data['MSRP'] = "0" # 預設為 0 以便統計

    weight_match = re.search(r"Estimated Average Weight\s*[\n]*\s*(\d+\.?\d*|TBA|ТВА)", clean_text, re.IGNORECASE)
    if weight_match: data['Weight (g)'] = weight_match.group(1).replace('ТВА', 'TBA')
    else: data['Weight (g)'] = ""

    # 產品名稱
    lines = [l.strip() for l in clean_text.split('\n') if l.strip()]
    product_name = ""
    style_idx = -1
    for i, line in enumerate(lines):
        if "Style#" in line:
            style_idx = i
            break
    if style_idx > 0:
        for k in range(style_idx - 1, -1, -1):
            curr = lines[k]
            skip_keywords = ["mont-bell", "Fall", "Winter", "NEW", "REVISED", "MSRP", "¥", "CONFIDENTIAL", "Western", "Available"]
            is_noise = False
            for kw in skip_keywords:
                if kw.lower() in curr.lower(): is_noise = True; break
            if re.search(r"^[A-Z]{2,3}\(.*\)$", curr): is_noise = True
            if not is_noise:
                product_name = curr
                break
    data['Product Name'] = product_name

    # Features & Material
    features_match = re.search(r"Features\s*\n(.*?)(?=\n\s*Material)", clean_text, re.DOTALL | re.IGNORECASE)
    data['Features'] = features_match.group(1).strip() if features_match else ""

    material_match = re.search(r"Material\s*\n(.*?)(?=\n\s*(Size|Estimated Average Weight))", clean_text, re.DOTALL | re.IGNORECASE)
    data['Material'] = material_match.group(1).strip() if material_match else ""

    # Description
    desc_content = []
    for line in lines:
        if "Features" in line: break
        if line.startswith("•") or line.startswith("●"):
            desc_content.append(line.replace("•", "").replace("●", "").strip())
    data['Description'] = "\n".join(desc_content)

    # Category
    categories = ["ALPINE CLOTHING", "INSULATION", "THERMAL", "RAIN WEAR", "SOFT SHELL", "PANTS", "BASE LAYER", "FIELD WEAR", "TRAVEL & COUNTRY", "CAP & HAT", "GLOVES", "SOCKS", "SLEEPING BAG", "FOOTWEAR", "BACKPACK", "BAG", "ACCESSORIES", "CYCLING", "SNOW GEAR", "CLIMBING", "FISHING", "PADDLE SPORTS", "DOG GEAR", "KIDS & BABY"]
    data['Category'] = "Uncategorized"
    for cat in categories:
        if cat in clean_text: data['Category'] = cat; break

    return data

# --- 3. 側邊欄介面 (Sidebar) ---
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/8/87/PDF_file_icon.svg/1667px-PDF_file_icon.svg.png", width=50)
    st.header("步驟 1: 上傳檔案")
    uploaded_file = st.file_uploader("請選擇 Mont-bell PDF 型錄", type="pdf")
    
    st.markdown("---")
    st.info("💡 **提示：** \n此工具會自動識別產品頁面，並忽略目錄或封面頁。")

# --- 4. 主畫面介面 (Main) ---
st.title("🏔️ Mont-bell 產品型錄數位化儀表板")

if uploaded_file is None:
    # 尚未上傳檔案時的歡迎畫面
    st.markdown("""
    ### 👋 歡迎使用
    這個工具能將 PDF 型錄轉換為 **視覺化數據** 與 **Excel 報表**。
    
    **功能特色：**
    * ✅ **智慧擷取**：自動抓取 Style#, 價格, 重量, 材質, 特色。
    * ✅ **數據清洗**：自動移除頁眉、頁碼等雜訊。
    * ✅ **視覺分析**：自動生成分類統計圖表。
    
    👈 請從左側上傳檔案以開始。
    """)

else:
    # 檔案已上傳，顯示操作按鈕
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
                    # 更新進度條
                    percent = int((i + 1) / total_pages * 100)
                    my_bar.progress(percent, text=f"正在分析第 {i+1}/{total_pages} 頁... (已擷取 {len(products)} 項產品)")
                    
                    text = page.extract_text()
                    if not text: continue
                    
                    p_data = parse_product_page(text, i + 1)
                    if p_data:
                        products.append(p_data)
            
            my_bar.empty() # 清除進度條

            if products:
                # 資料處理
                df = pd.DataFrame(products)
                
                # 數值轉換 (方便做圖表)
                df['MSRP (JPY)'] = pd.to_numeric(df['MSRP'], errors='coerce').fillna(0)
                
                # --- 5. 視覺化儀表板呈現 ---
                
                st.success(f"✅ 分析完成！共擷取 **{len(products)}** 項產品資料。")
                
                # 頂部關鍵指標 (KPIs)
                kpi1, kpi2, kpi3, kpi4 = st.columns(4)
                kpi1.metric("總產品數", f"{len(df)} 件")
                kpi2.metric("產品類別數", f"{df['Category'].nunique()} 類")
                avg_price = df[df['MSRP (JPY)'] > 0]['MSRP (JPY)'].mean()
                kpi3.metric("平均單價 (MSRP)", f"¥{avg_price:,.0f}")
                kpi4.metric("資料來源頁數", f"{total_pages} 頁")
                
                st.markdown("---")

                # 分頁內容
                tab1, tab2 = st.tabs(["📊 視覺化分析", "📋 詳細資料表 & 下載"])
                
                with tab1:
                    # 圖表：各類別產品數量
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
                    
                    # 圖表：價格分佈 (Histogram)
                    st.subheader("💰 價格分佈區間 (JPY)")
                    price_chart = alt.Chart(df[df['MSRP (JPY)'] > 0]).mark_bar().encode(
                        x=alt.X('MSRP (JPY)', bin=alt.Bin(maxbins=20), title='價格區間 (JPY)'),
                        y=alt.Y('count()', title='產品數量'),
                        color=alt.value('#ff7f0e')
                    ).properties(height=300)
                    st.altair_chart(price_chart, use_container_width=True)

                with tab2:
                    # 資料預覽與下載
                    st.subheader("詳細資料清單")
                    
                    # 欄位篩選顯示
                    display_cols = ['Page', 'Category', 'Product Name', 'Style#', 'MSRP', 'Weight (g)', 'Features']
                    st.dataframe(
                        df[display_cols], 
                        use_container_width=True,
                        column_config={
                            "Page": st.column_config.NumberColumn("頁碼", width="small"),
                            "MSRP": st.column_config.TextColumn("價格 (JPY)", width="small"),
                            "Features": st.column_config.TextColumn("產品特點", width="large"),
                        }
                    )
                    
                    # Excel 下載按鈕
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='Products')
                    excel_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 下載完整 Excel 報表",
                        data=excel_data,
                        file_name="Montbell_Product_List.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )

            else:
                st.warning("⚠️ PDF 讀取完畢，但未發現符合「產品頁面格式」的資料。請確認上傳檔案是否正確。")

        except Exception as e:
            st.error(f"❌ 發生錯誤: {e}")