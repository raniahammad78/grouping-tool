import streamlit as st
import pandas as pd
import io

# Set page to wide mode for a better dashboard feel
st.set_page_config(page_title="Thickness Grouping Tool", layout="wide")

st.title("📊 Grouping Dashboard")


# --- IMPROVEMENT 1: CACHING ---
# This tells Streamlit to remember the file so it doesn't reload it every second
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    # Clean hidden spaces
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.replace('\xa0', '').str.strip()
            df[col] = df[col].replace('nan', '')
    return df


# --- IMPROVEMENT 2: SIDEBAR LAYOUT ---
with st.sidebar:
    st.header("⚙️ Controls")
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Use our new caching function
    df = load_data(uploaded_file)

    with st.sidebar:
        column = st.selectbox("Select thickness column", df.columns)
        run_button = st.button("Run Grouping", type="primary")

    # Initialize session state
    if "grouped_df" not in st.session_state:
        st.session_state.grouped_df = None

    # Only run the heavy processing when the button is clicked
    if run_button:
        clean_df = df.dropna(subset=[column]).copy()
        clean_df[column] = clean_df[column].astype(str).str.replace(r'[^\d.]', '', regex=True).astype(float).round(2)
        clean_df = clean_df.loc[:, ~clean_df.columns.str.contains('Unnamed')]
        clean_df = clean_df.sort_values(by=column)

        st.session_state.grouped_df = clean_df
        st.sidebar.success("Grouping completed!")

    # Display results
    if st.session_state.grouped_df is not None:
        base_df = st.session_state.grouped_df

        # --- IMPROVEMENT 3: DASHBOARD METRICS ---
        total_items = len(base_df)
        unique_groups = base_df[column].nunique()

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Items", total_items)
        col2.metric("Unique Thickness Groups", unique_groups)
        col3.metric("Data Status", "Cleaned & Sorted")
        st.divider()  # Adds a nice horizontal line

        # --- CHARTS AND TABLES ---
        st.subheader("📈 Items per Thickness")

        chart_data = base_df[column].value_counts().reset_index()
        chart_data.columns = [column, 'Item Count']
        chart_data = chart_data.sort_values(by=column)
        chart_data[column] = chart_data[column].astype(str) + " mm"

        st.bar_chart(data=chart_data, x=column, y='Item Count', use_container_width=True)

        st.subheader("📊 Detailed Organized Report")

        report_rows = []
        for thickness_val, group in base_df.groupby(column):
            is_first = True
            item_count = 0

            for _, row in group.iterrows():
                new_row = {}
                if is_first:
                    new_row["Group Name"] = f"{thickness_val} mm"
                    is_first = False
                else:
                    new_row["Group Name"] = ""

                for col in base_df.columns:
                    new_row[col] = row[col]

                report_rows.append(new_row)
                item_count += 1

            total_row = {col: "" for col in base_df.columns}
            total_row["Group Name"] = f"TOTAL: {thickness_val} mm"
            total_row[column] = f"{item_count} items"

            report_rows.append(total_row)
            report_rows.append({"Group Name": ""} | {col: "" for col in base_df.columns})

        report_df = pd.DataFrame(report_rows)


        def style_report(row):
            group_text = str(row['Group Name'])
            if group_text.startswith('TOTAL:'):
                return ['background-color: #374b61; color: white; font-weight: bold'] * len(row)
            elif group_text != "" and not group_text.startswith('TOTAL:'):
                return ['background-color: #f0f2f6; font-weight: bold'] * len(row)
            else:
                return [''] * len(row)


        styled_report = report_df.style.apply(style_report, axis=1)
        st.dataframe(styled_report, use_container_width=True)

        # Download Logic
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            st.session_state.grouped_df.to_excel(writer, index=False, sheet_name='All Raw Data')
            styled_report.to_excel(writer, index=False, sheet_name='Detailed Report')

            worksheet = writer.sheets['Detailed Report']
            for i, col_name in enumerate(report_df.columns):
                worksheet.set_column(i, i, 20)

        st.download_button(
            label="📥 Download Detailed Report",
            data=buffer.getvalue(),
            file_name="exact_thickness_report.xlsx",
            mime="application/vnd.ms-excel"
        )
else:
    # If no file is uploaded, show a friendly welcome message in the center
    st.info("👈 Please upload an Excel file in the sidebar to get started.")
