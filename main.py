#!/usr/local/bin/python3
import pandas as pd
import os
import numpy as np
import openpyxl


# Function to create omschrijving based on the row data
def get_omschrijving(row):
    # Define the transaction type based on the Type of the row
    if row[0] == 'Verkoopprijs artikel(en), ontvangen van kopers en door bol.com door te storten':
        transaction_type = 'SALE'
    elif row[0] == 'Correctie verkoopprijs artikel(en)':
        transaction_type = 'REFUND'
    else:
        print('transaction_type unknown in get_omschrijving')

    # Construct the omschrijving string
    omschrijving = f"{row['Bestelnummer']}_{transaction_type}_NL_{row['Land van verzending']}"
    return omschrijving


# Function to process the dataframe
def process_data(df):
    # Used accounting accounts in e-boekhouden
    relatie_bol = '130010'
    debtor_account = '130010'
    nl_sales_tax_high = '800000'
    sales_oss_nl_be = '804001'

    # Convert the datatypes
    df = df.convert_dtypes()

    # Replace '-' with '.' in the 'Datum' field
    df['Datum'] = df['Datum'].astype(str).str.replace('-', '.')

    # Define the type mapping logic
    type_mapping = {
        'Verkoopprijs artikel(en), ontvangen van kopers en door bol.com door te storten': (debtor_account, ''),
        'Correctie verkoopprijs artikel(en)': (debtor_account, 'r')
    }

    # Only keep the desired types
    types_to_keep = list(type_mapping.keys())
    df = df[df['Type'].isin(types_to_keep)]

    # Determine if the tax applies for NL
    is_tax_nl = df['Land van verzending'] == 'NL'
    is_tax_nl = is_tax_nl.fillna(False)  # Replace NA values with False

    # Convert the 'Bestelnummer' column to a string
    df['Bestelnummer'] = df['Bestelnummer'].astype(str)

    # Fill the dataframe with required values based on conditions
    # We are creating new columns and filling them with data based on conditions
    df.insert(loc=1, column='Soort', value=np.where(is_tax_nl, 'Factuur verstuurd', 'Memoriaal'))
    df.insert(loc=2, column='Rekening', value=df.apply(lambda row: debtor_account if row['Type'] == 'Verkoopprijs artikel(en), ontvangen van kopers en door bol.com door te storten' else (nl_sales_tax_high if is_tax_nl[row.name] else sales_oss_nl_be), axis=1))
    df.insert(loc=3, column='Omschrijving', value=df.apply(get_omschrijving, axis=1))
    df.insert(loc=4, column='Boekstuk', value='')
    df.insert(loc=5, column='Bedrag excl', value=np.where(is_tax_nl, '', df['Bedrag'].astype(str).str.lstrip('-€')))
    df.insert(loc=6, column='BTW-code', value=np.where(is_tax_nl, 'HOOG', ''))
    df.insert(loc=7, column='BTW-bedrag', value='')
    df.insert(loc=8, column='Bedrag incl', value=np.where(is_tax_nl, df['Bedrag'].astype(str).str.lstrip('-€'), ''))
    df.insert(loc=9, column='Tegenrekening', value=df.apply(lambda row: debtor_account if row['Type'] == 'Correctie verkoopprijs artikel(en)' else (nl_sales_tax_high if is_tax_nl[row.name] else sales_oss_nl_be), axis=1))
    df.insert(loc=10, column='Relatie', value=relatie_bol)
    df.insert(loc=11, column='Factuurnummer', value=np.where(is_tax_nl, df['Type'].map(type_mapping).str[1] + df['Bestelnummer'], ''))
    df.insert(loc=12, column='Betalingstermijn', value='30')

    return df


# Main script execution starts from here
directory = './'
# Iterate over all files in the directory
for filename in os.listdir(directory):
    f = os.path.join(directory, filename)
    # Check if it is a file and if it is an xlsx file
    if os.path.isfile(f) and filename.endswith('.xlsx'):
        # Read the xlsx file, skip first 7 rows
        df = pd.read_excel(f, skiprows=7)
        # Process the dataframe
        df = process_data(df)
        # Write the processed dataframe to an xlsx file with a new filename
        csv_filename = filename.replace('.xlsx', '.csv')
        df.to_csv('./export/exp_' + csv_filename, index=False,
                  columns=['Datum', 'Soort', 'Rekening', 'Omschrijving', 'Boekstuk',
                           'Bedrag excl', 'BTW-code', 'BTW-bedrag', 'Bedrag incl',
                           'Tegenrekening', 'Relatie', 'Factuurnummer', 'Betalingstermijn'])