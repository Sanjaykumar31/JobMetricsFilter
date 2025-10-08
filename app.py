import streamlit as st
import pandas as pd
import io

def process_excel(file):
    # Load Excel file
    xls = pd.ExcelFile(file)
    sheet_names = xls.sheet_names[:2]  # Only take the first two sheets

    all_results = {}

    for idx, sheet in enumerate(sheet_names, start=1):
        df = pd.read_excel(xls, sheet_name=sheet)

        # Ensure column names are standardized
        df.columns = [col.strip().title() for col in df.columns]

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
    return output

# --- Streamlit UI ---
st.set_page_config(page_title="Comparative Job Analysis", layout="centered")
st.title("📊 Comparative Job Analysis Report Generator")

uploaded_file = st.file_uploader("Upload your Excel file with two environment sheets", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Processing file..."):
        try:
            output_excel = process_excel(uploaded_file)
            st.success("✅ Report generated successfully!")

            st.download_button(
                label="📥 Download Comparative_Analysis_Report.xlsx",
                data=output_excel,
                file_name="Comparative_Analysis_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
else:
    st.info("Please upload an Excel (.xlsx) file with at least two sheets.")
