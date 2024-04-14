import streamlit as st
import pandas as pd
import os
from datetime import datetime

def process_excel(df):
    # Perform the calculations
    df['IVA BS'] = df['Sales Amount'] * 0.16
    df['TOTAL BS'] = df['Sales Amount'] + df['IVA BS']
    df['SUBTOTAL $'] = df['Sales Amount'] / df['Exchange Rate']
    df['IVA $'] = df['SUBTOTAL $'] * 0.16
    df['TOTAL $'] = df['SUBTOTAL $'] + df['IVA $']
    df['75% IVA'] = df['IVA $'] * 0.75
    df['25% IVA'] = df['IVA $'] * 0.25

    # Format date columns to remove time (assuming column names are 'Date Column')
    df['Date Column'] = pd.to_datetime(df['Date Column']).dt.date

    # Ensure 4 decimal places for 'Exchange Rate'
    df['Exchange Rate'] = df['Exchange Rate'].round(4)

    # Calculate "Días Vencimiento"
    today = datetime.now().date()
    df['Due Date'] = pd.to_datetime(df['Due Date']).dt.date  # Ensure 'Due Date' is only the date
    df['Días Vencimiento'] = (df['Due Date'] - today).dt.days

    # Round all numerical columns to two decimal places
    df = df.round(2)

    return df


# Streamlit app
st.title('Excel Processing App')

uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    if st.button('Process'):
        processed_df = process_excel(df)

        # Save the processed DataFrame to a new Excel file
        output_file = 'processed_excel_file.xlsx'
        processed_df.to_excel(output_file, index=False)

        # Let the user download the processed file
        with open(output_file, "rb") as file:
            st.download_button(label="Download Processed Excel File",
                               data=file,
                               file_name=output_file,
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        # Optionally, clean up the directory by removing the file after download
        os.remove(output_file)
