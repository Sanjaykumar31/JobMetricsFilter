import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta

def process_excel(file, start_date, end_date):
    xls = pd.ExcelFile(file)
    all_results = {}

    # Convert start_date and end_date to timezone-aware UTC timestamps
    start_date_utc = pd.to_datetime(start_date).tz_localize('UTC')
    end_date_utc = pd.to_datetime(end_date).tz_localize('UTC') + timedelta(days=1)

    for idx, sheet in enumerate(xls.sheet_names, start=1):
        df = pd.read_excel(xls, sheet_name=sheet)

        # Standardize column names
        df.columns = [col.strip().title() for col in df.columns]

        required_columns = {'Tenant', 'System Job', 'Trigger Type', 'Completed At'}
        if not required_columns.issubset(df.columns):
            raise ValueError(f"Sheet '{sheet}' is missing required columns.")

        # Convert Completed At to datetime with UTC
        df['Completed At'] = pd.to_datetime(df['Completed At'], utc=True, errors='coerce')
        df = df.dropna(subset=['Completed At'])

        # Filter by selected date range
        mask = (df['Completed At'] >= start_date_utc) & (df['Completed At'] <= end_date_utc)
        df_filtered = df[mask]

        if df_filtered.empty:
            continue  # Skip if no data in the range

        # ----- 1. Job Type Classification -----
        job_type_summary = df_filtered['System Job'].str.lower().value_counts().reindex(['yes', 'no'], fill_value=0)
        job_type_total = job_type_summary.sum()
        job_type_df = pd.DataFrame({
            'Job Type': ['System Jobs', 'User-Defined Jobs'],
            'Count': [job_type_summary['yes'], job_type_summary['no']]
        })
        job_type_df['Percentage'] = (job_type_df['Count'] / job_type_total * 100).round(2).astype(str) + '%'
        job_type_df.loc[len(job_type_df.index)] = ['Total', job_type_total, '100%']

        # ----- 2. Trigger Type Count -----
        trigger_summary = df_filtered['Trigger Type'].str.lower().value_counts().reindex(['ad-hoc', 'scheduled'], fill_value=0)
        trigger_total = trigger_summary.sum()
        trigger_df = pd.DataFrame({
            'Trigger Type': ['Adhoc', 'Scheduled'],
            'Count': [trigger_summary['ad-hoc'], trigger_summary['scheduled']]
        })
        trigger_df['Percentage'] = (trigger_df['Count'] / trigger_total * 100).round(2).astype(str) + '%'
        trigger_df.loc[len(trigger_df.index)] = ['Total', trigger_total, '100%']

        # ----- 3. Tenant-wise Job Volume -----
        tenant_df = df_filtered['Tenant'].value_counts().reset_index()
        tenant_df.columns = ['Tenant', 'Job Count']
        tenant_total = tenant_df['Job Count'].sum()
        tenant_df['Percentage'] = (tenant_df['Job Count'] / tenant_total * 100).round(2).astype(str) + '%'
        tenant_df.loc[len(tenant_df.index)] = ['Total', tenant_total, '100%']

        # ----- 4. Tenant-wise System Job and Trigger Type Count with Global Percentages -----
        tenant_sysjob = df_filtered.groupby(['Tenant', df_filtered['System Job'].str.lower()]).size().unstack(fill_value=0)
        tenant_sysjob = tenant_sysjob.rename(columns={'yes': 'System Jobs', 'no': 'User-Defined Jobs'})

        tenant_trigger = df_filtered.groupby(['Tenant', df_filtered['Trigger Type'].str.lower()]).size().unstack(fill_value=0)
        tenant_trigger = tenant_trigger.rename(columns={'ad-hoc': 'Adhoc', 'scheduled': 'Scheduled'})

        # Merge both on Tenant
        tenant_metrics = tenant_sysjob.join(tenant_trigger, how='outer').fillna(0).reset_index()

        # Ensure all expected columns exist
        for col in ['System Jobs', 'User-Defined Jobs', 'Adhoc', 'Scheduled']:
            if col not in tenant_metrics.columns:
                tenant_metrics[col] = 0

        # Convert to integer
        tenant_metrics[['System Jobs', 'User-Defined Jobs', 'Adhoc', 'Scheduled']] = tenant_metrics[
            ['System Jobs', 'User-Defined Jobs', 'Adhoc', 'Scheduled']
        ].astype(int)

        # Total sums across all tenants (for global percentages)
        total_yes = tenant_metrics['System Jobs'].sum()
        total_no = tenant_metrics['User-Defined Jobs'].sum()
        total_adhoc = tenant_metrics['Adhoc'].sum()
        total_scheduled = tenant_metrics['Scheduled'].sum()

        # Global percentages per column, avoiding division by zero
        tenant_metrics['System Jobs %'] = (
            tenant_metrics['System Jobs'] / (total_yes if total_yes != 0 else 1) * 100
        ).round(2).astype(str) + '%'
        tenant_metrics['User-Defined Jobs %'] = (
            tenant_metrics['User-Defined Jobs'] / (total_no if total_no != 0 else 1) * 100
        ).round(2).astype(str) + '%'
        tenant_metrics['Adhoc %'] = (
            tenant_metrics['Adhoc'] / (total_adhoc if total_adhoc != 0 else 1) * 100
        ).round(2).astype(str) + '%'
        tenant_metrics['Scheduled %'] = (
            tenant_metrics['Scheduled'] / (total_scheduled if total_scheduled != 0 else 1) * 100
        ).round(2).astype(str) + '%'

        # Reorder columns as requested
        ordered_columns = [
            'Tenant',
            'System Jobs', 'System Jobs %',
            'User-Defined Jobs', 'User-Defined Jobs %',
            'Adhoc', 'Adhoc %',
            'Scheduled', 'Scheduled %'
        ]

        existing_cols = [col for col in ordered_columns if col in tenant_metrics.columns]

        tenant_metrics = tenant_metrics[existing_cols]

        # Store results
        all_results[f"Env_{idx}_Job_Type"] = job_type_df
        all_results[f"Env_{idx}_Trigger_Type"] = trigger_df
        all_results[f"Env_{idx}_TenantWise_Job_Count"] = tenant_df
        all_results[f"Env_{idx}_TenantWise_System_Trigger_Count"] = tenant_metrics

    if not all_results:
        raise ValueError("No data found within the selected date range in any sheet.")

    # Save to Excel in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in all_results.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)

    return output, all_results


# --- Streamlit UI ---
st.set_page_config(page_title="Jobs Metrics Generator", layout="wide", page_icon=":biohazard:")
st.title("📊 Job Metrics Generator")

# 📘 README Instructions
with st.expander("📖 README - Instructions for Uploading Excel File"):
    st.markdown("""
**Please upload an Excel (.xlsx) file that meets the following requirements:**

✅ One or more sheets, each with **at least** the following columns:
- `Tenant`  
- `System Job`  
- `Trigger Type`  
- `Completed At`

⚠️ Column names must be spelled **exactly** as listed above (case-insensitive is now handled).
""")

# 📤 File Upload
uploaded_file = st.file_uploader("📂 Upload your file here ::", type=["xlsx"])

# 📅 Date Range Picker
st.markdown("### 📅 Select Date Range to Filter the Metrics")
date_range = st.date_input(
    "Filter jobs completed between these dates (inclusive):",
    value=(datetime.today() - timedelta(days=30), datetime.today())
)

if uploaded_file and isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
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

elif not uploaded_file:
    st.info("Please upload a valid Excel (.xlsx) file.")
