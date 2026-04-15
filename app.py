import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime, timedelta


# --- Column name mapping (lowercase key -> standardized display name) ---
COLUMN_MAP = {
    'tenant': 'Tenant',
    'issystemjob': 'System Job',
    'system job': 'System Job',
    'systemjob': 'System Job',
    'is_system_job': 'System Job',
    'triggertype': 'Trigger Type',
    'trigger type': 'Trigger Type',
    'trigger_type': 'Trigger Type',
    'completedat': 'Completed At',
    'completed at': 'Completed At',
    'completed_at': 'Completed At',
    'timetaken_uiformat': 'TimeTaken_UIFormat',
    'timetaken_dbformat': 'TimeTaken_DBFormat',
    'time taken ui format': 'TimeTaken_UIFormat',
    'time taken db format': 'TimeTaken_DBFormat',
    'timetaken_ui_format': 'TimeTaken_UIFormat',
    'timetaken_db_format': 'TimeTaken_DBFormat',
    'timtaken_uiformat': 'TimeTaken_UIFormat',
    'timtaken_dbformat': 'TimeTaken_DBFormat',
}

REQUIRED_COLUMNS = {'Tenant', 'System Job', 'Trigger Type', 'Completed At'}
TIME_COLUMNS = {'TimeTaken_UIFormat', 'TimeTaken_DBFormat'}


def normalize_columns(df):
    """Map raw column headers to standardized names using case-insensitive matching."""
    new_columns = {}
    for col in df.columns:
        stripped = col.strip()
        lookup = stripped.lower().replace(' ', '')
        # Try direct lookup with no spaces
        if lookup in {k.replace(' ', ''): k for k in COLUMN_MAP}:
            for key, val in COLUMN_MAP.items():
                if key.replace(' ', '') == lookup:
                    new_columns[col] = val
                    break
        else:
            # Fallback: keep title-cased version
            new_columns[col] = stripped.title()
    df = df.rename(columns=new_columns)
    # Drop duplicate columns (keep first)
    df = df.loc[:, ~df.columns.duplicated()]
    return df


def parse_ui_time_to_ms(val):
    """Parse UI format like '2m 17s 357ms' or '1h 5m 3s 100ms' into total milliseconds."""
    if pd.isna(val) or not isinstance(val, str) or val.strip() == '':
        return 0
    val = val.strip().lower()
    total_ms = 0
    h = re.search(r'(\d+)\s*h', val)
    m = re.search(r'(\d+)\s*m(?!s)', val)
    s = re.search(r'(\d+)\s*s(?!$|\s*\d)', val)
    # more lenient 's' match
    s = re.search(r'(\d+)\s*s', val)
    ms = re.search(r'(\d+)\s*ms', val)
    if h:
        total_ms += int(h.group(1)) * 3600000
    if m:
        total_ms += int(m.group(1)) * 60000
    if s and not ms:
        # If there's no ms, the 's' match is fine
        total_ms += int(s.group(1)) * 1000
    elif s and ms:
        # Need to be careful: '17s 357ms' — the s match might grab '357' if greedy
        # Re-parse more carefully
        s_match = re.search(r'(\d+)\s*s(?:\s|$|\s*\d+\s*ms)', val)
        if s_match:
            total_ms += int(s_match.group(1)) * 1000
    if ms:
        total_ms += int(ms.group(1))
    return total_ms


def parse_db_time_to_ms(val):
    """Parse DB format like '00:02:17:357' (HH:MM:SS:mmm) into total milliseconds."""
    if pd.isna(val) or not isinstance(val, str) or val.strip() == '':
        return 0
    val = val.strip()
    parts = val.split(':')
    if len(parts) == 4:
        h, m, s, ms = parts
        return int(h) * 3600000 + int(m) * 60000 + int(s) * 1000 + int(ms)
    elif len(parts) == 3:
        h, m, s = parts
        return int(h) * 3600000 + int(m) * 60000 + int(float(s) * 1000)
    return 0


def ms_to_display(total_ms):
    """Convert milliseconds to a human-readable string like '1h 5m 17s 357ms'."""
    if total_ms <= 0:
        return '0s'
    total_ms = int(total_ms)
    h = total_ms // 3600000
    remainder = total_ms % 3600000
    m = remainder // 60000
    remainder = remainder % 60000
    s = remainder // 1000
    ms = remainder % 1000
    parts = []
    if h > 0:
        parts.append(f"{h}h")
    if m > 0:
        parts.append(f"{m}m")
    if s > 0:
        parts.append(f"{s}s")
    if ms > 0:
        parts.append(f"{ms}ms")
    return ' '.join(parts) if parts else '0s'


def read_uploaded_file(file):
    """Read CSV or XLSX file and return a dict of {sheet_name: DataFrame}."""
    name = file.name.lower()
    if name.endswith('.csv'):
        df = pd.read_csv(file)
        return {'Sheet1': df}
    elif name.endswith('.xlsx') or name.endswith('.xls'):
        xls = pd.ExcelFile(file)
        sheets = {}
        for sheet in xls.sheet_names:
            sheets[sheet] = pd.read_excel(xls, sheet_name=sheet)
        return sheets
    else:
        raise ValueError("Unsupported file type. Please upload a .csv or .xlsx file.")


def compute_tenant_runtime(df_filtered):
    """Compute tenant-wise total run time from TimeTaken columns."""
    has_ui = 'TimeTaken_UIFormat' in df_filtered.columns
    has_db = 'TimeTaken_DBFormat' in df_filtered.columns

    if not has_ui and not has_db:
        return None

    if has_ui:
        df_filtered = df_filtered.copy()
        df_filtered['_runtime_ms'] = df_filtered['TimeTaken_UIFormat'].apply(parse_ui_time_to_ms)
    elif has_db:
        df_filtered = df_filtered.copy()
        df_filtered['_runtime_ms'] = df_filtered['TimeTaken_DBFormat'].apply(parse_db_time_to_ms)

    tenant_runtime = df_filtered.groupby('Tenant')['_runtime_ms'].sum().reset_index()
    tenant_runtime.columns = ['Tenant', 'Total Runtime (ms)']
    tenant_runtime['Total Runtime'] = tenant_runtime['Total Runtime (ms)'].apply(ms_to_display)
    total_ms = tenant_runtime['Total Runtime (ms)'].sum()
    # Add total row
    tenant_runtime.loc[len(tenant_runtime)] = ['Total', total_ms, ms_to_display(total_ms)]
    tenant_runtime = tenant_runtime[['Tenant', 'Total Runtime']]
    return tenant_runtime


def process_file(file, start_date, end_date):
    sheets = read_uploaded_file(file)
    all_results = {}

    start_date_utc = pd.to_datetime(start_date).tz_localize('UTC')
    end_date_utc = pd.to_datetime(end_date).tz_localize('UTC') + timedelta(days=1)

    for idx, (sheet_name, df) in enumerate(sheets.items(), start=1):
        df = normalize_columns(df)

        missing = REQUIRED_COLUMNS - set(df.columns)
        if missing:
            raise ValueError(
                f"Sheet '{sheet_name}' is missing required columns: {', '.join(missing)}.\n"
                f"Found columns: {', '.join(df.columns)}"
            )

        # Convert Completed At to datetime with UTC
        df['Completed At'] = pd.to_datetime(df['Completed At'], utc=True, errors='coerce')
        df = df.dropna(subset=['Completed At'])

        # Filter by date range
        mask = (df['Completed At'] >= start_date_utc) & (df['Completed At'] <= end_date_utc)
        df_filtered = df[mask]

        if df_filtered.empty:
            continue

        # ----- 1. Job Type Classification -----
        job_type_summary = df_filtered['System Job'].astype(str).str.lower().value_counts().reindex(
            ['yes', 'no'], fill_value=0
        )
        job_type_total = job_type_summary.sum()
        job_type_df = pd.DataFrame({
            'Job Type': ['System Jobs', 'User-Defined Jobs'],
            'Count': [job_type_summary['yes'], job_type_summary['no']]
        })
        job_type_df['Percentage'] = (job_type_df['Count'] / job_type_total * 100).round(2).astype(str) + '%'
        job_type_df.loc[len(job_type_df)] = ['Total', job_type_total, '100%']

        # ----- 2. Trigger Type Count -----
        trigger_summary = df_filtered['Trigger Type'].astype(str).str.lower().value_counts().reindex(
            ['ad-hoc', 'scheduled'], fill_value=0
        )
        trigger_total = trigger_summary.sum()
        trigger_df = pd.DataFrame({
            'Trigger Type': ['Adhoc', 'Scheduled'],
            'Count': [trigger_summary['ad-hoc'], trigger_summary['scheduled']]
        })
        trigger_df['Percentage'] = (trigger_df['Count'] / trigger_total * 100).round(2).astype(str) + '%'
        trigger_df.loc[len(trigger_df)] = ['Total', trigger_total, '100%']

        # ----- 3. Tenant-wise Job Volume -----
        tenant_df = df_filtered['Tenant'].value_counts().reset_index()
        tenant_df.columns = ['Tenant', 'Job Count']
        tenant_total = tenant_df['Job Count'].sum()
        tenant_df['Percentage'] = (tenant_df['Job Count'] / tenant_total * 100).round(2).astype(str) + '%'

        # Compute tenant runtime and merge
        runtime_df = compute_tenant_runtime(df_filtered)
        if runtime_df is not None:
            # Remove the Total row from runtime_df before merging (we'll add our own total)
            runtime_data = runtime_df[runtime_df['Tenant'] != 'Total']
            tenant_df = tenant_df.merge(runtime_data, on='Tenant', how='left')
            tenant_df['Total Runtime'] = tenant_df['Total Runtime'].fillna('0s')
            # Total row
            total_runtime_row = runtime_df[runtime_df['Tenant'] == 'Total']['Total Runtime'].values
            total_runtime_str = total_runtime_row[0] if len(total_runtime_row) > 0 else '0s'
            tenant_df.loc[len(tenant_df)] = ['Total', tenant_total, '100%', total_runtime_str]
        else:
            tenant_df.loc[len(tenant_df)] = ['Total', tenant_total, '100%']

        # ----- 4. Tenant-wise System Job and Trigger Type Count -----
        tenant_sysjob = df_filtered.groupby(
            ['Tenant', df_filtered['System Job'].astype(str).str.lower()]
        ).size().unstack(fill_value=0)
        tenant_sysjob = tenant_sysjob.rename(columns={'yes': 'System Jobs', 'no': 'User-Defined Jobs'})

        tenant_trigger = df_filtered.groupby(
            ['Tenant', df_filtered['Trigger Type'].astype(str).str.lower()]
        ).size().unstack(fill_value=0)
        tenant_trigger = tenant_trigger.rename(columns={'ad-hoc': 'Adhoc', 'scheduled': 'Scheduled'})

        tenant_metrics = tenant_sysjob.join(tenant_trigger, how='outer').fillna(0).reset_index()

        for col in ['System Jobs', 'User-Defined Jobs', 'Adhoc', 'Scheduled']:
            if col not in tenant_metrics.columns:
                tenant_metrics[col] = 0

        tenant_metrics[['System Jobs', 'User-Defined Jobs', 'Adhoc', 'Scheduled']] = tenant_metrics[
            ['System Jobs', 'User-Defined Jobs', 'Adhoc', 'Scheduled']
        ].astype(int)

        total_yes = tenant_metrics['System Jobs'].sum()
        total_no = tenant_metrics['User-Defined Jobs'].sum()
        total_adhoc = tenant_metrics['Adhoc'].sum()
        total_scheduled = tenant_metrics['Scheduled'].sum()

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

with st.expander("📖 README - Instructions"):
    st.markdown("""
**Upload a CSV or Excel (.xlsx) file with job execution data.**

The app reads the header row and maps columns automatically (case-insensitive). 
It looks for these required columns:
- `Tenant`
- `System Job` (also accepts: `isSystemJob`, `Is_System_Job`)
- `Trigger Type` (also accepts: `triggerType`, `Trigger_Type`)
- `Completed At` (also accepts: `completedAt`, `Completed_At`)

Optional columns for runtime analysis:
- `timeTaken_uiFormat` — e.g. `2m 17s 357ms`
- `timeTaken_dbFormat` — e.g. `00:02:17:357`

If either time column is present, tenant-wise total runtime will be included in the report.

The header names are matched **case-insensitively** and extra whitespace is trimmed.
""")

uploaded_file = st.file_uploader("📂 Upload your file here:", type=["csv", "xlsx", "xls"])

st.markdown("### 📅 Select Date Range to Filter the Metrics")
date_range = st.date_input(
    "Filter jobs completed between these dates (inclusive):",
    value=(datetime.today() - timedelta(days=30), datetime.today())
)

if uploaded_file and isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
    with st.spinner("Processing file..."):
        try:
            output_excel, extracted_dataframes = process_file(uploaded_file, start_date, end_date)
            st.success("✅ Report generated successfully!")

            st.download_button(
                label="📥 Download Processed File",
                data=output_excel,
                file_name="Job_Metrics.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.markdown("### 📂 Extracted Report Preview")
            for sheet_name, df in extracted_dataframes.items():
                with st.expander(f"📄 {sheet_name}"):
                    st.dataframe(df, use_container_width=True)

        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")

elif not uploaded_file:
    st.info("Please upload a CSV or Excel (.xlsx) file.")
