import io
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title='Excel & CSV Data Cleaner', page_icon='🧹', layout='wide')
st.title('🧹 Excel & CSV Data Cleaner')
st.caption('Clean spreadsheets, preview every change, and download a formatted Excel report.')


def read_data(file):
    if file.name.lower().endswith('.csv'):
        return pd.read_csv(file, dtype='string', keep_default_na=True, encoding='utf-8-sig')
    return pd.read_excel(file, dtype='string', engine='openpyxl')


def clean_data(df, trim=True, remove_empty=True, remove_duplicates=True, normalize_case=False):
    result = df.copy()
    original_rows = len(result)
    if trim:
        for col in result.columns:
            result[col] = result[col].str.strip()
    # Treat blank/whitespace-only cells as missing (never replace with invented values).
    result = result.replace(r'^\s*$', pd.NA, regex=True)
    if normalize_case:
        for col in result.select_dtypes(include=['string', 'object']).columns:
            result[col] = result[col].str.lower()
    empty_removed = 0
    if remove_empty:
        before = len(result)
        result = result.dropna(how='all')
        empty_removed = before - len(result)
    duplicates_removed = 0
    if remove_duplicates:
        before = len(result)
        result = result.drop_duplicates()
        duplicates_removed = before - len(result)
    result = result.reset_index(drop=True)
    stats = {'Input rows': original_rows, 'Empty rows removed': empty_removed,
             'Duplicate rows removed': duplicates_removed, 'Output rows': len(result),
             'Missing cells remaining': int(result.isna().sum().sum())}
    return result, stats


def excel_bytes(cleaned, stats):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        cleaned.to_excel(writer, index=False, sheet_name='Cleaned Data')
        pd.DataFrame(list(stats.items()), columns=['Metric', 'Value']).to_excel(
            writer, index=False, sheet_name='Cleaning Summary')
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = 'A2'
            sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', fgColor='17365D')
                cell.alignment = Alignment(horizontal='center')
            for column in sheet.columns:
                width = max((len(str(c.value)) if c.value is not None else 0) for c in column)
                sheet.column_dimensions[get_column_letter(column[0].column)].width = min(max(width + 3, 15), 45)
    return output.getvalue()


uploaded = st.file_uploader('Upload an Excel or CSV file', type=['xlsx', 'csv'])
if uploaded is not None:
    try:
        original = read_data(uploaded)
    except Exception as exc:
        st.error(f'Unable to read this file: {exc}')
        st.stop()
    if len(original.columns) == 0:
        st.error('The file has no columns.')
        st.stop()
    if original.empty:
        st.warning('The file has headers but no data rows.')
        st.stop()
    st.subheader('Cleaning options')
    c1, c2, c3, c4 = st.columns(4)
    trim = c1.checkbox('Trim extra spaces', value=True)
    remove_empty = c2.checkbox('Remove fully empty rows', value=True)
    remove_duplicates = c3.checkbox('Remove duplicate rows', value=True)
    normalize_case = c4.checkbox('Lowercase text (optional)', value=False)
    cleaned, stats = clean_data(original, trim, remove_empty, remove_duplicates, normalize_case)
    st.subheader('Cleaning dashboard')
    cols = st.columns(5)
    for col, (label, value) in zip(cols, stats.items()):
        col.metric(label, f'{value:,}')
    st.subheader('Before & after')
    left, right = st.columns(2)
    with left:
        st.markdown('**Original data**')
        st.dataframe(original, use_container_width=True)
    with right:
        st.markdown('**Cleaned data**')
        st.dataframe(cleaned, use_container_width=True)
    st.info('Missing values are kept as blanks. Review the output before using it in production.')
    st.download_button('📥 Download cleaned Excel report', excel_bytes(cleaned, stats),
                       file_name='cleaned_data_report.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
