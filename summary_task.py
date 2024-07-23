import pandas as pd
import os
import csv
from enum import Enum

import helper

class Columns(Enum):
    SALES_ORDER = 'Sales Order #'
    CHARGE_DESC = 'Charge Description'
    CHARGE_AMOUNT_EX_TAX = 'Charge Amount (ex Tax)'
    SITE_ID = 'Site Id'
    LLDGCODE = 'LLDGCODE'
    LNARR1 = 'LNARR1'




def lnarr1_exp(df):
    # Orro | SDWan Charge | June 2024
    df['From'] = pd.to_datetime(df['From'], format='%d/%m/%Y')

    # Format the date to 'Month Year'
    df['formatted_date'] = df['From'].dt.strftime('%B %Y')

    # Extract a single unique formatted date
    unique_date = df['formatted_date'].iloc[0]
    return f'Orro | SDWan Charge | {unique_date}'

#folder = '/Users/lucycai/Downloads/Orro_Bills'
src_excel = os.path.join(helper.folder, '1071219.XLSX')
sheet_name = "Bill Charge Detail"

def transformed_df(src_excel, sheet_name):
    df = helper.excel_2_df(src_excel,sheet_name)
    lnarr1 = lnarr1_exp(df)

    selected_cols = [Columns.SALES_ORDER.value, Columns.CHARGE_DESC.value,
                 Columns.CHARGE_AMOUNT_EX_TAX.value,
                 Columns.SITE_ID.value]

    hardcoded_cols = {Columns.LLDGCODE.value:'GL',
                   Columns.LNARR1.value:lnarr1}

    new_col_order = [Columns.LLDGCODE.value,
                 Columns.SITE_ID.value,
                 Columns.CHARGE_AMOUNT_EX_TAX.value,
                 Columns.LNARR1.value,
                 Columns.SALES_ORDER.value,
                 Columns.CHARGE_DESC.value,
                 ]
    selected_df = df[selected_cols]
    for col,value in hardcoded_cols.items():
        selected_df[col] = value

    reordered_df = selected_df[new_col_order]
    return reordered_df



def add_summary(df_grouped):
    result = []
    for name, group in df_grouped:
        summary = group[[Columns.CHARGE_AMOUNT_EX_TAX.value]].sum()
        summary[Columns.SALES_ORDER.value] = 'Charge Back Journal'

        keep = group[[Columns.LLDGCODE.value, Columns.SITE_ID.value,
                     Columns.LNARR1.value]].iloc[0]
        summary = pd.concat([summary,keep] )
        summary_df = pd.DataFrame([summary],columns=group.columns)
        result.append(summary_df)
        result.append(group)
        
    final_df = pd.concat(result, ignore_index=True)
    return final_df

df_grouped = transformed_df(src_excel, sheet_name).groupby(Columns.SITE_ID.value)
f_df = add_summary(df_grouped)
output_file_path = os.path.join(helper.folder,'new_file.xlsx')
#f_df.to_csv(output_file_path, index=False)
f_df.to_excel(output_file_path, index=False, engine='openpyxl')



helper.highlight_excel(output_file_path)
 

