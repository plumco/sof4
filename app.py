import streamlit as st
import pandas as pd
from io import BytesIO

# --- PAGE CONFIG ---
st.set_page_config(page_title="Huliot Sales Order Form", layout="wide")

# --- CUSTOM CSS FOR STYLING ---
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stHeader { background-color: #1d4e89; color: white; padding: 20px; border-radius: 10px; }
    .filter-container { background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# --- HELPER FUNCTIONS ---
def load_data(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file)
        # Clean column names to prevent KeyErrors (removes leading/trailing spaces)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Error loading Excel: {e}")
        return None

def to_excel(df):
    """Converts dataframe back to Excel bytes for download."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sales Order')
    return output.getvalue()

# --- APP HEADER ---
st.markdown("""
    <div style="background-color:#11468F; padding:20px; border-radius:10px; margin-bottom:20px;">
        <h1 style="color:white; margin:0;">📋 Huliot Pipes & Fittings — Sales Order Form</h1>
        <p style="color:#d1d1d1; margin:0;">Format No: HPF-01 | Rev. 01 | Issue Date: 10.04.2026</p>
    </div>
    """, unsafe_allow_html=True)

# --- FILE UPLOADER ---
uploaded_file = st.sidebar.file_uploader("Upload Product Excel", type=["xlsx"])

if uploaded_file:
    df = load_data(uploaded_file)
    
    if df is not None:
        # 1. TABS FOR NAVIGATION
        tab_details, tab_products, tab_summary = st.tabs(["📝 Order Details", "🔍 Add Products", "📥 Summary & Download"])

        with tab_products:
            st.subheader("🛒 Product Catalog")
            
            # 2. SELECT SIZE (DN) - Horizontal Selection
            st.write("**SELECT SIZE (DN)**")
            # Get unique sub_groups, handling NaN and sorting
            sizes = ["All"] + sorted([str(x) for x in df['sub_group'].unique() if pd.notna(x)])
            selected_size = st.segmented_control(
                "Sub Group Filter", options=sizes, default="All", label_visibility="collapsed"
            )

            st.write("") # Spacer

            # 3. SELECT CATEGORY - Horizontal Selection
            st.write("**SELECT TYPE**")
            categories = ["All"] + sorted([str(x) for x in df['category'].unique() if pd.notna(x)])
            selected_cat = st.segmented_control(
                "Category Filter", options=categories, default="All", label_visibility="collapsed"
            )

            # --- FILTER LOGIC (Addressing the TypeError/KeyError) ---
            mask = pd.Series([True] * len(df))
            
            if selected_size != "All":
                mask &= df['sub_group'].astype(str) == selected_size
            
            if selected_cat != "All":
                mask &= df['category'].astype(str) == selected_cat

            filtered_df = df[mask]

            # 4. SEARCH BAR
            search_query = st.text_input("🔍 Search item code, description, size...", "").lower()
            if search_query:
                # Search across all columns
                search_mask = df.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)
                filtered_df = filtered_df[filtered_df.index.isin(df[search_mask].index)]

            # 5. DISPLAY RESULTS
            st.dataframe(filtered_df, use_container_width=True, hide_index=True)

        with tab_summary:
            st.subheader("Finalize Order")
            st.info("You can download the filtered view below in the original Excel format.")
            
            # Download Logic
            excel_data = to_excel(filtered_df)
            st.download_button(
                label="📥 Download Sales Order (Excel)",
                data=excel_data,
                file_name="Huliot_Order_Export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

else:
    st.info("Please upload an Excel file in the sidebar to start.")
