import streamlit as st
import pandas as pd
import os
from datetime import datetime

def process_excel(df):
    # Convert date columns to dates only
    df['Document Date'] = pd.to_datetime(df['Document Date']).dt.date
    df['Due Date'] = pd.to_datetime(df['Due Date']).dt.date

    # Perform the calculations
    df['IVA BS'] = df['Sales Amount'] * 0.16
    df['TOTAL BS'] = df['Sales Amount'] + df['IVA BS']
    df['SUBTOTAL $'] = df['Sales Amount'] / df['Exchange Rate']
    df['IVA $'] = df['SUBTOTAL $'] * 0.16
    df['TOTAL $'] = df['SUBTOTAL $'] + df['IVA $']
    df['75% IVA'] = df['IVA $'] * 0.75
    df['25% IVA'] = df['IVA $'] * 0.25

    # Calculate 'Días Vencimiento'
    today = datetime.now().date()
    df['Días Vencimiento'] = (pd.to_datetime(df['Due Date']) - pd.to_datetime(today)).dt.days

    # Selective rounding
    numerical_cols = df.select_dtypes(include=['number']).columns.tolist()
    numerical_cols.remove('Exchange Rate')  # Exclude Exchange Rate from rounding
    df[numerical_cols] = df[numerical_cols].round(2)
    df['Exchange Rate'] = df['Exchange Rate'].round(4)  # Round Exchange Rate to four decimal places

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
