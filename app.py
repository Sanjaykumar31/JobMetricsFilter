import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta

def process_excel(file, start_date, end_date):
    # Load Excel file
    xls = pd.ExcelFile(file)
    sheet_names = xls.sheet_names[:2]  # Only take the first two sheets

    all_results = {}

    for idx, sheet in enumerate(sheet_names, start=1):
        df = pd.read_excel(xls, sheet_name=sheet)

        # Standardize column names
        df.columns = [col.strip().title() for col in df.columns]

        # Convert Completed At to datetime with UTC
        try:
            df['Completed At'] = pd.to_datetime(df['Completed At'], utc=True, errors='coerce')
        except Exception as e:
            raise ValueError(f"Error parsing 'Completed At' column: {str(e)}")

        # Convert start_date and end_date to timezone-aware UTC timestamps
        start_date_utc = pd.to_datetime(start_date).tz_localize('UTC')
        end_date_utc = pd.to_datetime(end_date).tz_localize('UTC') + timedelta(days=1)

        # Filter by selected date range
        mask = (df['Completed At'] >= start_date_utc) & (df['Completed At'] <= end_date_utc)
        df = df[mask]

        if df.empty:
            raise ValueError(f"No data found between {start_date} and {end_date} in sheet '{sheet}'.")

        # ----- 1. Job Type Classification -----
        job_type_summary = df['System Job'].str.lower().value_counts().reindex(['yes', 'no'], fill_value=0)
        job_type_df = pd.DataFrame({
            'Job Type': ['System Jobs', 'User-Defined Jobs'],
            'Count': [job_type_summary['yes'], job_type_summary['no']]
        })
        job_type_df.loc['Total'] = ['Total', job_type_df['Count'].sum()]

        # ----- 2. Trigger Type Count -----
        trigger_type_summary = df['Trigger Type'].str.lower().value_counts().reindex(['ad-hoc', 'scheduled'], fill_value=0)
        trigger_type_df = pd.DataFrame({
            'Trigger Type': ['Adhoc', 'Scheduled'],
            'Count': [trigger_type_summary['ad-hoc'], trigger_type_summary['scheduled']]
        })
        trigger_type_df.loc['Total'] = ['Total', trigger_type_df['Count'].sum()]

        # ----- 3. Tenant-wise Job Volume -----
        tenant_job_count = df['Tenant'].value_counts().reset_index()
        tenant_job_count.columns = ['Tenant', 'Job Count']
        tenant_total = tenant_job_count['Job Count'].sum()
        tenant_job_count.loc['Total'] = ['Total', tenant_total]

        # Store results
        all_results[f"Environment_{idx}_Job_Type"] = job_type_df
        all_results[f"Environment_{idx}_Trigger_Type"] = trigger_type_df
        all_results[f"Environment_{idx}_Tenant_Job_Count"] = tenant_job_count

    # Save to Excel in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in all_results.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output, all_results

# --- Streamlit UI ---
st.set_page_config(page_title="Comparative Job Analysis", layout="centered")
st.title("📊 Job Report Generator")

# 📘 README Instructions
with st.expander("📖 README - Instructions for Uploading Excel File"):
    st.markdown("""
**Please upload an Excel (.xlsx) file that meets the following requirements:**

✅ The file must contain **exactly two sheets** named:
- `1`
- `2`

✅ Each sheet must contain the following columns in **this exact spelling (case-sensitive):**

1. Tenant  
2. System Job  
3. Trigger Type  
4. Completed At

⚠️ Column names must be spelled **exactly** as listed above for the app to work properly.
""")

# 📅 Date Range Picker
st.markdown("### 📅 Select Date Range to Filter the Metrics")
date_range = st.date_input(
    "Filter jobs completed between these dates (inclusive):",
    value=(datetime.today() - timedelta(days=30), datetime.today())
)

if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    st.error("Please select a valid date range.")
    st.stop()

# 📤 File Upload
uploaded_file = st.file_uploader("Upload your Excel file Here::", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Processing file..."):
        try:
            output_excel, extracted_dataframes = process_excel(uploaded_file, start_date, end_date)
            st.success("✅ Report generated successfully!")

            # 📥 Download button
            st.download_button(
                label="📥 Download Processed File",
                data=output_excel,
                file_name="Job_Metrics.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # 👀 Show extracted tables below
            st.markdown("### 📂 Extracted Report Preview")
            for sheet_name, df in extracted_dataframes.items():
                with st.expander(f"📄 {sheet_name}"):
                    st.dataframe(df, use_container_width=True)

        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
else:
    st.info("Please upload an Excel (.xlsx) file with exactly two sheets named '1' and '2'.")
