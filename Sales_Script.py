import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(layout="wide")

st.title("📊 Sales Consolidator App")

# Upload all required Excel files
uploaded_files = {
    "SAP-8235": st.file_uploader("Upload SAP-8235 Excel", type=["xlsx"]),
    "M3-8223": st.file_uploader("Upload M3-8223 Excel", type=["xlsx"]),
    "SAP-8224": st.file_uploader("Upload SAP-8224 Excel", type=["xlsx"]),
    "SAP-8225": st.file_uploader("Upload SAP-8225 Excel", type=["xlsx"]),
    "SAP-8229": st.file_uploader("Upload SAP-8229 Excel", type=["xlsx"]),
    "M3-8236": st.file_uploader("Upload M3-8236 Excel", type=["xlsx"]),
    "Aurora-8226-8297": st.file_uploader("Upload Aurora-8226-8297 Excel", type=["xlsx"]),
}

if all(uploaded_files.values()):
    # Load files into DataFrames
    dfs = {}
    for name, file in uploaded_files.items():
        dfs[name] = pd.read_excel(file, sheet_name="Sheet1")
        st.success(f"{name} loaded with {dfs[name].shape[0]} rows")
    
    # Assigning DataFrames for easier access
    df_8223 = dfs['M3-8223']
    df_8224 = dfs['SAP-8224']
    df_8235 = dfs['SAP-8235']
    df_8225 = dfs['SAP-8225']
    df_8229 = dfs['SAP-8229']
    df_8236 = dfs['M3-8236']
    df_Aurora = dfs['Aurora-8226-8297']

    # Assuming df_Aurora is your DataFrame

    # Convert the 'Document Date' and 'Accounting Date' columns to datetime
    df_Aurora['Document Date'] = pd.to_datetime(df_Aurora['Document Date'])
    df_Aurora['Accounting Date'] = pd.to_datetime(df_Aurora['Accounting Date'])

    # Format the datetime columns to only include the date part
    df_Aurora['Document Date'] = df_Aurora['Document Date'].dt.strftime('%d-%m-%Y')
    df_Aurora['Accounting Date'] = df_Aurora['Accounting Date'].dt.strftime('%d-%m-%Y')

    # Perform the remaining operations
    df_Aurora['Entity Code'] = df_Aurora['Company'].map({'IN': '8226'}).fillna('8297')
    df_Aurora['Goods/Service'] = df_Aurora['Service Indicator'].map({'1': 'Service'}).fillna('Goods')
    df_Aurora['Document Type'] = df_Aurora['Transaction Type'].map({'CreditMemo': 'Credit Note'}).fillna('Invoice')
    df_Aurora.loc[df_Aurora['Entity Code'].isin(['8297', '8226']), 'ERP'] = 'AURORA'
    df_Aurora['Party State'] = df_Aurora['Party State'].astype(str)
    df_Aurora['Country Code'] = df_Aurora['Party State'].str[:2]
    #df_Aurora = df_Aurora.drop(["Goods/Service"],axis=1)

    import pandas as pd

    df_8236['Entity'] = df_8236['Company Code'].map({'PU1': '8223'}).fillna('8236')
    df_8236['ERP'] = 'M3'
    df_8236['Invoice Date'] = df_8236['Invoice Date'].astype(str)
    df_8236['Accounting Date'] = df_8236['Accounting Date'].astype(str)
    df_8236['Invoice Date'] = pd.to_datetime(df_8236['Invoice Date'], format='%Y%m%d')
    df_8236['Accounting Date'] = pd.to_datetime(df_8236['Accounting Date'], format='%Y%m%d')
    df_8236['Invoice Date'] = df_8236['Invoice Date'].dt.strftime('%d-%m-%Y')
    df_8236['Accounting Date'] = df_8236['Accounting Date'].dt.strftime('%d-%m-%Y')
    df_8236['Month'] = df_8236['Invoice Date'].str[3:5]
    df_8236['Fiscal Period'] = df_8236['Month']
    df_8236['Year'] = df_8236['Invoice Date'].str[6:10]
    df_8236['Document Type'] = df_8236['Document type'].apply(lambda x: 'Credit Note' if x == 'CRN' else 'Invoice')
    df_8236['Goods/Service'] = df_8236['New - Goods/Service'].map({'GOODS': 'Goods'}).fillna('Service')
    df_8236 = df_8236.drop(['New - Goods/Service', 'Customer Bill City'], axis=1)

    import pandas as pd

    df_8229['ERP'] = 'SAP SMART'
    df_8229['Accounting Pos Date'] = pd.to_datetime(df_8229['Accounting Pos Date'])
    df_8229['Accounting Pos Date'] = df_8229['Accounting Pos Date'].dt.strftime('%d-%m-%Y')
    df_8229['Billing Date'] = df_8229['Billing Date'].astype(str)
    df_8229['Month'] = df_8229['Billing Date'].str[5:7]
    df_8229['Fiscal Year'] = df_8229['Billing Date'].str[:4]
    df_8229['Document type'] = df_8229['Document Type'].map(lambda x: 'Credit Note' if x == 'O' else 'Invoice')
    df_8229['Goods/Service'] = df_8229['Goods/Service'].apply(lambda x: 'Goods' if x == 'G' else 'Service')
    df_8229['Base Value'] = df_8229['Base Value'] * df_8229['Document type'].map({'Credit Note': -1}).fillna(1)
    df_8229['CGST AMT'] = df_8229['CGST AMT'] * df_8229['Document type'].map({'Credit Note': -1}).fillna(1)
    df_8229['SGST AMT'] = df_8229['SGST AMT'] * df_8229['Document type'].map({'Credit Note': -1}).fillna(1)
    df_8229['IGST AMT'] = df_8229['IGST AMT'] * df_8229['Document type'].map({'Credit Note': -1}).fillna(1)
    df_8229['TCS AMT'] = df_8229['TCS AMT'] * df_8229['Document type'].map({'Credit Note': -1}).fillna(1)
    df_8229['TCSG AMT'] = df_8229['TCSG AMT'] * df_8229['Document type'].map({'Credit Note': -1}).fillna(1)
    df_8229 = df_8229.drop('Document Type', axis=1)

    import pandas as pd

    df_8225['ERP'] = 'SAP SMART'
    df_8225['Billing Date'] = pd.to_datetime(df_8225['Billing Date'])
    df_8225['Billing Date'] = df_8225['Billing Date'].astype(str)
    df_8225['Month'] = df_8225['Billing Date'].str[5:7]
    df_8225['Fiscal Year'] = df_8225['Billing Date'].str[:4]
    df_8225['Document type'] = df_8225['Document Type'].map(lambda x: 'Credit Note' if x == 'O' else 'Invoice')
    df_8225['Goods/Service'] = df_8225['Goods/Service'].map(lambda x: 'Goods' if x == 'G' else 'Service')

    def adjust_amounts(row, column_name):
        return row[column_name] * -1 if row['Document type'] == 'Credit Note' else row[column_name]

    df_8225['Base Value'] = df_8225.apply(lambda row: adjust_amounts(row, 'Base Value'), axis=1)
    df_8225['CGST AMT'] = df_8225.apply(lambda row: adjust_amounts(row, 'CGST AMT'), axis=1)
    df_8225['SGST AMT'] = df_8225.apply(lambda row: adjust_amounts(row, 'SGST AMT'), axis=1)
    df_8225['IGST AMT'] = df_8225.apply(lambda row: adjust_amounts(row, 'IGST AMT'), axis=1)
    df_8225['TCS AMT'] = df_8225.apply(lambda row: adjust_amounts(row, 'TCS AMT'), axis=1)
    df_8225['TCSG AMT'] = df_8225.apply(lambda row: adjust_amounts(row, 'TCSG AMT'), axis=1)

    df_8225 = df_8225.drop('Document Type', axis=1)

    # Process 8224 DataFrame
    df_8224['ERP'] = 'SAP SMART'
    df_8224['PU to SU']=df_8224['Unnamed: 38']
    df_8224['Billing Date'] = pd.to_datetime(df_8224['Billing Date'])
    df_8224['Billing Date'] = df_8224['Billing Date'].dt.strftime('%d-%m-%Y')
    df_8224['Billing Date'] = df_8224['Billing Date'].astype(str)
    df_8224['Month'] = df_8224['Billing Date'].str[3:5]
    df_8224['Fiscal Year'] = df_8224['Billing Date'].str[6:]
    df_8224['Document type'] = np.where(df_8224['Document Type'] == 'O', 'Credit Note', 'Invoice')
    df_8224['Goods/Service'] = np.where(df_8224['Goods/Service'] == 'G', 'Goods', 'Service')
    df_8224['Base Value'] = np.where(df_8224['Document type'] == 'Credit Note', 
                                    df_8224['Base Value'] * -1,
                                    df_8224['Base Value'] * 1)
    df_8224['CGST AMT'] = np.where(df_8224['Document type'] == 'Credit Note', 
                                    df_8224['CGST AMT'] * -1,
                                    df_8224['CGST AMT'] * 1)
    df_8224['SGST AMT'] = np.where(df_8224['Document type'] == 'Credit Note', 
                                    df_8224['SGST AMT'] * -1,
                                    df_8224['SGST AMT'] * 1)
    df_8224['IGST AMT'] = np.where(df_8224['Document type'] == 'Credit Note', 
                                    df_8224['IGST AMT'] * -1,
                                    df_8224['IGST AMT'] * 1)
    df_8224['TCS AMT'] = np.where(df_8224['Document type'] == 'Credit Note', 
                                    df_8224['TCS AMT'] * -1,
                                    df_8224['TCS AMT'] * 1)
    df_8224['TCSG AMT'] = np.where(df_8224['Document type'] == 'Credit Note', 
                                    df_8224['TCSG AMT'] * -1,
                                    df_8224['TCSG AMT'] * 1)
    df_8224 = df_8224.drop('Document Type', axis=1)
    df_8224['PU to SU']=df_8224['PU to SU'].map(lambda x: 'PU to SU' if x == 'X' else None)

    df_8235['Accounting Pos Date'] = pd.to_datetime(df_8235['Accounting Pos Date'])
    df_8235['Fiscal Period'] = df_8235['Accounting Pos Date'].dt.month.astype(str).str.zfill(2)
    df_8235['Fiscal Year'] = df_8235['Accounting Pos Date'].dt.year
    df_8235['Accounting Pos Date'] = df_8235['Accounting Pos Date'].dt.strftime('%Y-%m-%d')
    df_8235['Document type'] = df_8235['Document Type'].map({'O': 'Credit Note'}).fillna('Invoice')
    df_8235['Goods/Service'] = df_8235['Goods/Service'].map({'G': 'Goods'}).fillna('Services')
    df_8235['ERP'] = 'SAP UGD'
    credit_note_mask = df_8235['Document type'] == 'Credit Note'
    df_8235.loc[credit_note_mask, ['Base Value', 'CGST AMT', 'SGST AMT', 'IGST AMT', 'TCS AMT', 'TCSG AMT']] *= -1
    df_8235.rename(columns={
        'CGST AMT': 'CGST Amount',
        'SGST AMT': 'SGST Amount',
        'IGST AMT': 'IGST Amount',
        'TCS AMT': 'TCS Amount',
        'TCSG AMT': 'TCSG Amount'
    }, inplace=True)
    df_8235.drop(columns='Document Type', inplace=True)

    #8223

    df_8223['Entity'] = np.where(df_8223['Company Code'] == 'PU1', '8223', '8236')
    df_8223['ERP'] = 'M3'
    df_8223['Invoice Date'] = df_8223['Invoice Date'].astype(str)
    df_8223['Accounting Date'] = df_8223['Accounting Date'].astype(str)
    df_8223['Invoice Date'] = pd.to_datetime(df_8223['Invoice Date'], format='%Y%m%d')
    df_8223['Accounting Date'] = pd.to_datetime(df_8223['Accounting Date'], format='%Y%m%d')
    df_8223['Invoice Date'] = df_8223['Invoice Date'].dt.strftime('%d-%m-%Y')
    df_8223['Accounting Date'] = df_8223['Accounting Date'].dt.strftime('%d-%m-%Y')
    df_8223['Month'] = df_8223['Accounting Date'].str[3:5]
    df_8223["Fiscal Period"] = df_8223["Month"]
    df_8223['Year'] = df_8223['Accounting Date'].str[6:10]
    df_8223['Document Type'] = df_8223['Document type'].apply(lambda x: 'Credit Note' if x == "CRN" else 'Invoice')
    df_8223['Goods/Service'] = np.where(df_8223['New - Goods/Service'] == 'GOODS', 'Goods', 'Service')
    df_8223 = df_8223.drop('New - Goods/Service',axis=1)
    df_8223 = df_8223.drop('Customer Bill City',axis=1)

    df_8235['Billing document copy'] = df_8235["Billing document"]
    mapping_8235 = {
        "Billing document copy": "SAP Key/ Invoice ID/ Voucher no.",
        "Document type" : "Document Type",
        "Customer Bill GSTIN" : "Customer Bill GSTIN",
        "Company Code": "Entity Code",
        "Billing document": "Invoice Number",
        "Accounting Pos Date": "Invoice Date",
        "Billing Date": "Accounting Date",
        "Orignal Invoice No.": "Original Invoice Number",
        "Orignal Inv. Date": "Original Invoice Date",
        "Sold to Party Name": "Customer Name",
        "Item Code": "Item ID",
        "ITEM Quantity": "Quantity",
        "Base Value": "Taxable Value",
        "TCSGRate": "TCSG Rate",
        "NotionalSales": "NON GST Flag",
        "Stock ID": "Customer Number",
        "Customer Bill Address": "Customer Bill Addr",
        "Bill Of Supply": "Customer State Code",
        "ITEMDISAmt": "Party Country",
        "Remarks": "Sales Register Remarks"
    }


    df_8236["Amt.in loc.cur"] = df_8236["Amt.in loc.cur."]
    df_8223["Amt.in loc.cur"] = df_8223["Amt.in loc.cur."]
    mapping_8223_8236 = {
        "Year" : "Fiscal Year",
        "Voucher No": "SAP Key/ Invoice ID/ Voucher no.",
        "Entity": "Entity Code",
        "Invoice Number": "Invoice Number",
        "Invoice Date": "Invoice Date",
        "Accounting Date": "Accounting Date",
        "Original Invoice No": "Original Invoice Number",
        "Bill to Party GSTN": "Customer Bill GSTIN",
        "Customer Name": "Customer Name",
        "Material": "Item ID",
        "Material description": "Item Description",
        "Billing qty in SKU": "Quantity",
        "HSN/SAC Code": "HSN/SAC",
        "Amt.in loc.cur.": "Taxable Value",
        "CGST Amount": "CGST Amount",
        "SGST Amount": "SGST Amount",
        "IGST Amount": "IGST Amount",
        "TCS Rate": "TCS Rate",
        "TCS Amount": "TCS Amount",
        "Bill Of Supply": "NON GST Flag",
        "Customer SAFIR Code": "Customer Number",
        "Customer Bill Address": "Customer Bill Addr",
        "City": "Customer Bill City",
        "Customer State Code": "Customer State Code",
        "Country Key": "Party Country",
        "Supplier GSTIN": "Supplier GSTIN",
        "Shipping Bill Number": "Shipping Bill Number",
        "Shipping Bill Date": "Shipping Bill Date",
        "Shipping Port Code": "Shipping Port Code",
        "Revenue GL Acnt": "Revenue GL Account",
        "Remarks": "Sales Register Remarks"
    }

    df_8224["Customer Bill City copy"] = df_8224["Customer Bill City"]
    df_8225["Customer Bill City copy"] = df_8225["Customer Bill City"]
    df_8229["Customer Bill City copy"] = df_8229["Customer Bill City"]
    mapping_8224_8225_8229 = {
        "Accounting Doc Number": "SAP Key/ Invoice ID/ Voucher no.",
        "ERP": "ERP",
        "Company Code": "Entity Code",
        "Document type": "Document Type",
        "Month": "Fiscal Period",
        "Fiscal Year": "Fiscal Year",
        "Billing document": "Invoice Number",
        "Billing Date": "Invoice Date",
        "Accounting Pos Date": "Accounting Date",
        "Orignal Invoice No.": "Original Invoice Number",
        "Orignal Inv. Date": "Original Invoice Date",
        "Customer Bill GSTIN": "Customer Bill GSTIN",
        "Sold to Party Name": "Customer Name",
        "Item Code": "Item ID",
        "Item Description": "Item Description",
        "ITEM Quantity": "Quantity",
        "Goods/Service": "Goods/Service",
        "Control code": "HSN/SAC",
        "Base Value": "Taxable Value",
        "CGST AMT": "CGST Amount",
        "SGST AMT": "SGST Amount",
        "IGST AMT": "IGST Amount",
        "TCS RATE": "TCS Rate",
        "TCS AMT": "TCS Amount",
        "TCSG RATE": "TCSG Rate",
        "TCSG AMT": "TCSG Amount",
        "FlagRelevantGST": "NON GST Flag",
        "Payer Number": "Customer Number",
        "Customer Bill Address": "Customer Bill Addr",
        "Customer Bill City": "Customer Bill City",
        "ITD State Code": "Customer State Code",
        "Customer Bill City copy": "Party Country",
        "Supplier GSTIN .": "Supplier GSTIN",
        "Remarks": "Sales Register Remarks"
    }

    df_Aurora["Invoice ID copy"] = df_Aurora["Invoice ID"]
    df_Aurora["Party Name copy"] = df_Aurora["Party Name"]
    mapping_Aurora_8226_8297 = {
        "Invoice ID copy": "SAP Key/ Invoice ID/ Voucher no.",
        "ERP": "ERP",
        "Entity Code": "Entity Code",
        "Document Type": "Document Type",
        "Fiscal Period": "Fiscal Period",
        "Fiscal Year": "Fiscal Year",
        "Invoice ID": "Invoice Number",
        "Document Date": "Invoice Date",
        "Accounting Date": "Accounting Date",
        "Original Invoice Number": "Original Invoice Number",
        "Original Invoice Date": "Original Invoice Date",
        "Adjustment Reason": "Adjustment Reason",
        "Party Tax ID": "Customer Bill GSTIN",
        "Party Name copy": "Customer Name",
        "Item ID": "Item ID",
        "Item Description": "Item Description",
        "Quantity": "Quantity",
        "Goods/Service": "Goods/Service",
        "HSN/SAC": "HSN/SAC",
        "Taxable Value": "Taxable Value",
        "CGST": "CGST Amount",
        "SGST": "SGST Amount",
        "IGST": "IGST Amount",
        "TCS Rate": "TCS Rate",
        "TCS Amount": "TCS Amount",
        "Nil_Exempt_Non-GST": "NON GST Flag",
        "Party ID": "Customer Number",
        "Party Name": "Customer Bill Addr",
        "Party City": "Customer Bill City",
        "Country Code": "Customer State Code",
        "Party Country": "Party Country",
        "Party Country Tax ID": "Party Country Tax ID",
        "Source Location": "Source Location",
        "GSTIN": "Supplier GSTIN",
        "Tax Type": "Sales Register Remarks",
    }

    # Rename columns
    df_8235.rename(columns=mapping_8235, inplace=True)
    df_8223.rename(columns=mapping_8223_8236, inplace=True)
    df_8224.rename(columns=mapping_8224_8225_8229, inplace=True)
    df_8225.rename(columns=mapping_8224_8225_8229, inplace=True)
    df_8229.rename(columns=mapping_8224_8225_8229, inplace=True)
    df_8236.rename(columns=mapping_8223_8236, inplace=True)
    df_Aurora.rename(columns=mapping_Aurora_8226_8297, inplace=True)
    def create_customer_pan(gstin):
        if pd.notna(gstin) and len(gstin) >= 10:  # Ensure GSTIN is not NaN and has enough length
            return gstin[2:12]
        return None
    df_8235['Customer PAN'] = df_8235['Customer Bill GSTIN'].apply(create_customer_pan)
    df_8236['Customer PAN'] = df_8236['Customer Bill GSTIN'].apply(create_customer_pan)
    df_8223['Customer PAN'] = df_8223['Customer Bill GSTIN'].apply(create_customer_pan)
    df_8224['Customer PAN'] = df_8224['Customer Bill GSTIN'].apply(create_customer_pan)
    df_8225['Customer PAN'] = df_8225['Customer Bill GSTIN'].apply(create_customer_pan)
    df_8229['Customer PAN'] = df_8229['Customer Bill GSTIN'].apply(create_customer_pan)
    df_Aurora['Customer PAN'] = df_Aurora['Customer Bill GSTIN'].apply(create_customer_pan)
    df_8235['Stock Transfer'] = np.where(df_8235['Customer PAN'] == "ABBCS6927M", 'Y', 'N')
    df_8236['Stock Transfer'] = np.where(df_8236['Customer PAN'] == "ABBCS6927M", 'Y', 'N')
    df_8223['Stock Transfer'] = np.where(df_8223['Customer PAN'] == "ABBCS6927M", 'Y', 'N')
    df_8224['Stock Transfer'] = np.where(df_8224['Customer PAN'] == "ABBCS6927M", 'Y', 'N')
    df_8225['Stock Transfer'] = np.where(df_8225['Customer PAN'] == "ABBCS6927M", 'Y', 'N')
    df_8229['Stock Transfer'] = np.where(df_8229['Customer PAN'] == "ABBCS6927M", 'Y', 'N')
    df_Aurora['Stock Transfer'] = np.where(df_Aurora['Customer PAN'] == "ABBCS6927M", 'Y', 'N')
    required_columns = [
        "SAP Key/ Invoice ID/ Voucher no.",
        "ERP",
        "Entity Code",
        "Document Type",
        "Fiscal Period",
        "Fiscal Year",
        "Invoice Number",
        "Invoice Date",
        "Accounting Date",
        "Original Invoice Number",
        "Original Invoice Date",
        "Adjustment Reason",
        "Customer Bill GSTIN",
        "Customer Name",
        "Item ID",
        "Item Description",
        "Quantity",
        "Goods/Service",
        "HSN/SAC",
        "Taxable Value",
        "CGST Rate",
        "CGST Amount",
        "SGST Rate",
        "SGST Amount",
        "IGST Rate",
        "IGST Amount",
        "Total Tax",
        "Invoice Total",
        "TCS Rate",
        "TCS Amount",
        "TCSG Rate",
        "TCSG Amount",
        "NON GST Flag",
        "Customer Number",
        "Customer Bill Addr",
        "Customer Bill City",
        "Customer State Code",
        "Party Country",
        "Party Country Tax ID",
        "Source Location",
        "Supplier GSTIN",
        "Shipping Bill Number",
        "Shipping Bill Date",
        "Shipping Port Code",
        "Customer PAN",
        "Stock Transfer",
        "Revenue GL Account",
        "Sales Register Remarks",
        "GSTR 1 Retrun Remarks - for  Monthly Summary",
        "Comments - for adjustment",
        "Reconciliation Remarks - for  Monthly Summary",
        "Month",
        "Reconciliation Remarks - for  YTD Reco",
        "PU to SU",
        "Source Location",
        "Party ID",
        "Internal Ref. No.",
        "External Ref. No."    
    ]
    dataframes = [df_8235, df_8236, df_8223, df_8224, df_8225, df_8229, df_Aurora]

    # Ensure all DataFrames have the same columns
    processed_dfs = []

    for df in dataframes:
        # Ensure columns match required_columns
        columns_to_drop = [col for col in df.columns if col not in required_columns]
        if columns_to_drop:
            df.drop(columns=columns_to_drop, inplace=True)
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            for col in missing_columns:
                df[col] = pd.NA  # Or another default value if preferred

        # Reorder columns to match required_columns order
        df = df[required_columns]
        
        # Append processed DataFrame to the list
        processed_dfs.append(df)

    # Concatenate all DataFrames
    df_Aurora = processed_dfs[6]
    df_8235 = processed_dfs[0]
    df_8224 = processed_dfs[3]
    df_8223 = processed_dfs[2]
    df_8225 = processed_dfs[4]
    df_8229 = processed_dfs[5]
    df_8236 = processed_dfs[1]
    import pandas as pd
    import streamlit as st
    from io import BytesIO

    # Simulate final DataFrames after transformation
    dfs = {
        '8235': df_8235,
        '8236': df_8236,
        '8223': df_8223,
        '8224': df_8224,
        '8225': df_8225,
        '8229': df_8229,
        'Aurora': df_Aurora
    }

    # Show message while transforming
    st.info("Transforming data...")

    # Combine all into one DataFrame
    all_data = pd.concat(dfs.values(), ignore_index=True)

    # Preview top 10 rows
    st.dataframe(all_data.head(10))

    # Function to export multiple sheets to Excel in memory
    def to_excel(df_dict):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in df_dict.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name)
        output.seek(0)
        return output

    # Get Excel binary
    excel_data = to_excel(dfs)

    # Download button
    st.download_button(
        label="📥 Download Consolidated Excel",
        data=excel_data,
        file_name="consolidated_sales_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
