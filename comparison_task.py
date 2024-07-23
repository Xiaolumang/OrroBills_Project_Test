import pandas as pd
import os
import csv
from enum import Enum
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import helper
from summary_task import Columns as Summary_Columns
from openpyxl.styles import NamedStyle
import numpy as np

class Columns(Enum):
    SITE = 'Site'
    SITE_ID = 'Site ID'
    COST_CENTER = 'Cost Centre'
    EXPECTED_MONTHLY_COST = 'Expected Monthly Cost'
    LAST_MONTHS_COST= 'Last Months Cost'
    THIS_MONTHS_COST = 'This months cost'
    BILLING_COMMENT = 'Billing Comment'
    DIFF = 'charged - expected'



fname = 'Orro Monthly Billing Review Aug 2024.xlsx'
sheet_name = 'Carriage Reconcilliation'
src_excel = os.path.join(helper.folder, fname)

fname2 = 'highlighted.xlsx'
sheet_name2 = 'summary'
src_excel2 = os.path.join(helper.folder, fname2)

def get_merged_df(src1_excel,src1_sheet, src2_excel,src2_sheet):
    df = helper.excel_2_df(src1_excel,src1_sheet)
    df = df.drop(columns=[Columns.LAST_MONTHS_COST.value,
                          Columns.THIS_MONTHS_COST.value])
    
    df2 = helper.excel_2_df(src2_excel,src2_sheet)
    extracted_df2 = df2[df2[Summary_Columns.SALES_ORDER.value]=='Charge Back Journal']
    extracted_df2 = extracted_df2[[Summary_Columns.SITE_ID.value,Summary_Columns.CHARGE_AMOUNT_EX_TAX.value]]
    #print(extracted_df2)
    merged_df = pd.merge(df, extracted_df2,right_on = Summary_Columns.SITE_ID.value,left_on= Columns.SITE_ID.value,how='right')

    merged_df.loc[merged_df[Columns.SITE_ID.value].notna(), Columns.DIFF.value] \
= round(merged_df[Summary_Columns.CHARGE_AMOUNT_EX_TAX.value] - merged_df[Columns.EXPECTED_MONTHLY_COST.value],2)
    merged_df[Columns.COST_CENTER.value] = merged_df[Columns.COST_CENTER.value].astype(object)
    return merged_df

def custom_sort_key(row):
    v = row[Columns.DIFF.value]
    if pd.isna(v):
        charge_amount = row[Summary_Columns.CHARGE_AMOUNT_EX_TAX.value]
        return (0,(0,-charge_amount) if charge_amount>=0 else (1,charge_amount))
    elif v > 0:
        return (1, -v)
    elif v < 0:
        return (2, v)
    elif v ==0:
        return (3, 0)



def adjust_columns(df):
    columns = list(df.columns)
    columns.insert(0,columns.pop(columns.index(Summary_Columns.SITE_ID.value)) )
    columns.pop(columns.index(Columns.SITE_ID.value))
    columns.remove(Columns.BILLING_COMMENT.value)
    columns.append(Columns.BILLING_COMMENT.value)
    df = df[columns]
    return df


merged_df = get_merged_df(src_excel, sheet_name,src_excel2, sheet_name2)

merged_df['sort_key'] = merged_df.apply(custom_sort_key, axis=1)
merged_df_sorted = merged_df.sort_values(by = ['sort_key',Summary_Columns.SITE_ID.value],
                                         ascending=[True, True])
merged_df_sorted = merged_df_sorted.drop(columns=['sort_key'])

merged_df_ajusted = adjust_columns(merged_df_sorted)


path = os.path.join(helper.folder, 'comp.xlsx')
helper.export_2_excel(path, merged_df_ajusted,Columns.COST_CENTER.value)
#merged_df_ajusted.to_excel(path, index=False, engine='openpyxl')


