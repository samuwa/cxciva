import streamlit as st
import pandas as pd
import os
from datetime import datetime
from io import BytesIO

def process_excel(df):
    
    
    # Calculate days overdue
    if 'Due Date' in df.columns:
        today = pd.to_datetime('today').normalize()  # Normalize to remove time component
        df['Due Date'] = pd.to_datetime(df['Due Date'], dayfirst=True)  # Ensure due date is treated as day-first format
        df['Dias Vencidos'] = (today - df['Due Date']).dt.days

    # Format Document Date and Due Date to DD/MM/YYYY
    if 'Document Date' in df.columns:
        df['Document Date'] = pd.to_datetime(df['Document Date']).dt.strftime('%d/%m/%Y')
    if 'Due Date' in df.columns:
        df['Due Date'] = pd.to_datetime(df['Due Date']).dt.strftime('%d/%m/%Y')

    # Perform the calculations
    df['IVA BS'] = df['Sales Amount'] * 0.16
    df['IVA BS 75%'] = df['IVA BS'] * 0.75
    df['IVA BS 25%'] = df['IVA BS'] * 0.25
    df['TOTAL BS'] = df['Sales Amount'] + df['IVA BS']
    df['SUBTOTAL $'] = df['Sales Amount'] / df['Exchange Rate']
    df['IVA $'] = df['SUBTOTAL $'] * 0.16
    df['TOTAL $'] = df['SUBTOTAL $'] + df['IVA $']
    df['75% IVA'] = df['IVA $'] * 0.75
    df['25% IVA'] = df['IVA $'] * 0.25

    # Round Exchange Rate to four decimal places first
    if 'Exchange Rate' in df.columns:
        df['Exchange Rate'] = df['Exchange Rate'].round(4)

    # Round all numerical columns to two decimal places, excluding 'Exchange Rate'
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    numeric_cols.remove('Exchange Rate')  # Remove 'Exchange Rate' from the list
    df[numeric_cols] = df[numeric_cols].round(2)

    # Drop specified columns
    columns_to_drop = ['Sales Amount'] # 'Current Trx Amount', 'Original Trx Amount']
    df = df.drop(columns=columns_to_drop, errors='ignore')

    df = df.rename(columns={
    'Current Trx Amount': 'Saldo Original BS',
    'Original Trx Amount': 'Saldo Restante BS'})

    dfs = {
        'BRILUX_no_Automercados': df[(df['COMPAÑIA'] == 'FABRICA BRILUX C.A.') & (~df['Customer Name'].str.contains('AUTOMERCADOS PLAZA', na=False))],
        'BRILUX_Automercados': df[(df['COMPAÑIA'] == 'FABRICA BRILUX C.A.') & (df['Customer Name'].str.contains('AUTOMERCADOS PLAZA', na=False))],
        'EXTRUVENSO_no_Specified': df[(df['COMPAÑIA'] == 'FABRICA EXTRUVENSO C.A.') & (~df['Customer Name'].str.contains('FERRETOTAL|CENTROBECO|FERRETERIA EPA', na=False))],
        'EXTRUVENSO_Specified': df[(df['COMPAÑIA'] == 'FABRICA EXTRUVENSO C.A.') & (df['Customer Name'].str.contains('FERRETOTAL|CENTROBECO|FERRETERIA EPA', na=False))]
    }

    return dfs

st.title('Excel Processing App')

uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    if st.button('Process and Download Excel'):
        processed_dfs = process_excel(df)  # Process the Excel file and get DataFrames

        # Create a BytesIO buffer to save the Excel file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet_name, df_processed in processed_dfs.items():
                df_processed.to_excel(writer, sheet_name=sheet_name, index=False)

        # Seek to the beginning of the stream
        output.seek(0)

        # Download button to download the combined Excel file
        st.download_button(label="Download Processed Excel File",
                           data=output,
                           file_name='combined_processed_excel.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        # Reset the buffer
        output.close()
