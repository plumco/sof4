import streamlit as st
import pandas as pd
from io import BytesIO

# --- 1. CONFIG & STYLING ---
st.set_page_config(page_title="Huliot Sales Order", layout="wide")

# CSS to mimic the blue buttons and clean look from your screenshot
st.markdown("""
    <style>
    /* Header styling */
    .header-box { background-color: #11468F; padding: 25px; border-radius: 15px; color: white; margin-bottom: 20px; }
    
    /* Button & Control styling */
    div[data-testid="stSegmentedControl"] button {
        background-color: #2b65b0;
        color: white !important;
        border: none;
        padding: 10px 20px;
        font-weight: 500;
    }
    div[data-testid="stSegmentedControl"] button[aria-checked="true"] {
        background-color: #11468F !important;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.3);
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. CORE LOGIC ---
@st.cache_data
def load_data(file):
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip() # Prevent KeyErrors from trailing spaces
    # Clean data to prevent TypeErrors during filtering
    df['sub_group'] = df['sub_group'].astype(str).str.strip()
    df['category'] = df['category'].astype(str).str.strip()
    return df

def get_excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# --- 3. UI LAYOUT ---
st.markdown("""
    <div class="header-box">
        <h1 style='margin:0;'>📋 Huliot Pipes & Fittings — Sales Order Form</h1>
        <p style='margin:0; opacity:0.8;'>Format No: HPF-01 | Rev. 01 | Issue Date: 10.04.2026</p>
    </div>
    """, unsafe_allow_html=True)

uploaded_file = st.sidebar.file_uploader("Upload Master Excel", type=["xlsx"])

if uploaded_file:
    df = load_data(uploaded_file)
    
    tab_details, tab_products, tab_summary = st.tabs(["📝 Order Details", "🔍 Add Products", "📥 Summary & Download"])

    with tab_products:
        st.subheader("🛒 Product Catalog")
        
        # --- SIZE SELECTION (With Dynamic Counts) ---
        st.write("**SELECT SIZE (DN)**")
        
        # Generate labels like "DN40 64"
        size_counts = df['sub_group'].value_counts()
        raw_sizes = sorted(df['sub_group'].unique().tolist())
        
        # Create a mapping for display names vs real values
        size_map = {f"ALL {len(df)}": "All"}
        for s in raw_sizes:
            size_map[f"{s} {size_counts[s]}"] = s
            
        selected_size_label = st.segmented_control(
            "Sizes", options=list(size_map.keys()), default=list(size_map.keys())[0], label_visibility="collapsed"
        )
        selected_size = size_map[selected_size_label]

        # --- CATEGORY SELECTION ---
        st.write("**SELECT TYPE**")
        categories = ["All"] + sorted(df['category'].unique().tolist())
        selected_cat = st.segmented_control(
            "Categories", options=categories, default="All", label_visibility="collapsed"
        )

        # --- FILTERING ---
        filtered_df = df.copy()
        if selected_size != "All":
            filtered_df = filtered_df[filtered_df['sub_group'] == selected_size]
        if selected_cat != "All":
            filtered_df = filtered_df[filtered_df['category'] == selected_cat]

        # Search functionality
        search = st.text_input("🔍 Search item code, description, size...", "")
        if search:
            filtered_df = filtered_df[filtered_df.apply(lambda row: search.lower() in row.astype(str).str.lower().values, axis=1)]

        # --- RESULTS ---
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)

    with tab_summary:
        st.success(f"Successfully filtered {len(filtered_df)} items.")
        st.download_button(
            label="📥 Download Exported Excel",
            data=get_excel_download(filtered_df),
            file_name="Huliot_Order_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("👋 Please upload your Excel sheet in the sidebar to populate the catalog.")
