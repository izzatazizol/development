from pandas import DataFrame
from tkinter import filedialog, messagebox
from datetime import datetime, timedelta
import pandas as pd
import re
import tkinter as tk


def process_shopee(filepath, uom_filepath):
    df = pd.read_excel(filepath, sheet_name='orders')
    uom_df = pd.read_excel(uom_filepath, sheet_name='Sheet1')

    df = df[~df['Order Status'].isin(['Cancelled', 'Unpaid', 'Package Return', '(blank)'])]

    empty_values = df['Parent SKU Reference No.'].isna() | (df['Parent SKU Reference No.'] == '')

    df['Deal Price'] = pd.to_numeric(df['Deal Price'], errors='coerce')
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')
    df['unit price'] = df['Deal Price'] * df['Quantity']

    if df['Variation Name'].notna().any():
        df['multiplier'] = df['Variation Name'].apply(
            lambda x: int(re.search(r'(\d+)\s*bottl(es)?', str(x)).group(1)) if re.search(r'(\d+)\s*bottl(es)?',
                                                                                          str(x)) else 1
        )
        df['Quantity'] *= df['multiplier']

    mask1 = df['Parent SKU Reference No.'].isin(['NTV828-CARTON', 'NTV830-CARTON'])
    mask2 = df['Parent SKU Reference No.'].isin(['NTV832-CARTON', 'NTV834-CARTON'])
    df.loc[mask1, 'Quantity'] *= 12
    df.loc[mask2, 'Quantity'] *= 24

    df['Parent SKU Reference No.'] = df['Parent SKU Reference No.'].str.replace('-CARTON', '')

    def determine_uom(sku_code):
        if 'NDT' in sku_code:
            if 'Purchase UoM/ID' in uom_df.columns:
                uom_series = uom_df.set_index('Internal Reference')['Purchase UoM/ID']
            else:
                uom_series = uom_df.set_index('Internal Reference')['Purchase UoM/External ID']
        else:
            if 'Unit of Measure/ID' in uom_df.columns:
                uom_series = uom_df.set_index('Internal Reference')['Unit of Measure/ID']
            else:
                uom_series = uom_df.set_index('Internal Reference')['Unit of Measure/External ID']
        return uom_series.get(sku_code)

    df['UoM'] = df['Parent SKU Reference No.'].apply(determine_uom)

    df['Order Lines/Product'] = df['Parent SKU Reference No.']
    df['Order Lines/Quantity'] = df['Quantity']
    df['Order Lines/Unit Price'] = df['unit price']

    pivot_table = df.pivot_table(index='Order Lines/Product',
                                 values=['Order Lines/Quantity', 'Order Lines/Unit Price'],
                                 aggfunc='sum')
    uom_mapping = df.drop_duplicates('Order Lines/Product').set_index('Order Lines/Product')['UoM']
    pivot_table['Order Lines/Unit of Measure/External ID'] = pivot_table.index.map(uom_mapping)

    pivot_table['Order Lines/Unit Price'] = pivot_table['Order Lines/Unit Price'] / pivot_table['Order Lines/Quantity']
    total_unit_price_1 = df['unit price'].sum()
    total_unit_price_1 = round(total_unit_price_1, 2)
    total_unit_price = pivot_table['Order Lines/Unit Price'] * pivot_table['Order Lines/Quantity']
    total_unit_price = total_unit_price.sum()
    total_unit_price = round(total_unit_price, 2)
    is_tally = total_unit_price_1 == total_unit_price

    return pivot_table.reset_index(), empty_values, is_tally


def process_lazada(filepath, uom_filepath):
    df = pd.read_excel(filepath, sheet_name='sheet1')
    uom_df = pd.read_excel(uom_filepath, sheet_name='Sheet1')

    df = df[~df['status'].isin([
        'canceled', 'In Transit: Returning to seller', 'Package Returned', 'returned', '(blank)'])]

    def extract_sku_code(sku_code):
        start_index = 1
        end_index = sku_code.find(']')
        return sku_code[start_index:end_index]

    df['SKU code'] = df['sellerSku'].apply(extract_sku_code)
    empty_values = df['SKU code'].isna().any()

    df['quantity'] = 1
    df['quantity'] = pd.to_numeric(df['quantity'], errors='coerce')

    df['multiplier'] = 1

    df.loc[df['sellerSku'].str.contains('BUNDLE'), 'multiplier'] = df['sellerSku'].apply(
        lambda x: int(re.search(r'\(BUNDLE (\d+)', str(x)).group(1)) if re.search(r'\(BUNDLE (\d+)', str(x)) else 1
    )
    df.loc[df['sellerSku'].str.contains('BUNDLE'), 'quantity'] *= df['multiplier']

    mask1 = df['SKU code'].isin(['NTV828-CARTON', 'NTV830-CARTON'])
    mask2 = df['SKU code'].isin(['NTV832-CARTON', 'NTV834-CARTON'])
    df.loc[mask1, 'quantity'] *= 12
    df.loc[mask2, 'quantity'] *= 24

    df['SKU code'] = df['SKU code'].str.replace('-CARTON', '')

    def determine_uom(sku_code):
        if 'NDT' in sku_code:
            if 'Purchase UoM/ID' in uom_df.columns:
                uom_series = uom_df.set_index('Internal Reference')['Purchase UoM/ID']
            else:
                uom_series = uom_df.set_index('Internal Reference')['Purchase UoM/External ID']
        else:
            if 'Unit of Measure/ID' in uom_df.columns:
                uom_series = uom_df.set_index('Internal Reference')['Unit of Measure/ID']
            else:
                uom_series = uom_df.set_index('Internal Reference')['Unit of Measure/External ID']
        return uom_series.get(sku_code)

    df['UoM'] = df['SKU code'].apply(determine_uom)

    df['Order Lines/Product'] = df['SKU code']
    df['Order Lines/Quantity'] = df['quantity']
    df['Order Lines/Unit Price'] = df['unitPrice']

    pivot = df.pivot_table(index='Order Lines/Product',
                           values=['Order Lines/Quantity', 'Order Lines/Unit Price'],
                           aggfunc='sum')
    pivot['Order Lines/Unit Price'] = pivot['Order Lines/Unit Price'] / pivot['Order Lines/Quantity']
    uom_mapping = df.drop_duplicates('Order Lines/Product').set_index('Order Lines/Product')['UoM']
    pivot['Order Lines/Unit of Measure/External ID'] = pivot.index.map(uom_mapping)

    total_nett_price = df['unitPrice'].sum()
    total_nett_price = round(total_nett_price, 2)
    total_unit_price = pivot['Order Lines/Unit Price'] * pivot['Order Lines/Quantity']
    total_unit_price = total_unit_price.sum()
    total_unit_price = round(total_unit_price, 2)
    is_tally = total_nett_price == total_unit_price

    return pivot.reset_index(), empty_values, is_tally


def create_csv(pivot_table, customer_info, empty_values, is_tally):
    num_empty = empty_values.sum()
    if num_empty > 0:
        messagebox.showinfo(
            "Information", f"There are {num_empty} missing SKU codes. Please fill in the SKU codes")
        return

    metadata = {
        'Company': 'Neutrovis Sdn. Bhd.',
        'Customer': customer_info,
        'Invoice Address': customer_info,
        'Delivery Address': customer_info,
        'Order Date': datetime.now().strftime('%Y-%m-%d'),
        'Expiration': (datetime.now()+timedelta(days=7)).strftime('%Y-%m-%d'),
        'Pricelist': 'Public Pricelist',
        'Payment Terms': 'Immediate Payment',
        'Salesperson': 'Callie Leong',
        'Sales Team': 'eCommerce',
        'Warehouse': 'Online Warehouse',
        'Delivery Date': (datetime.now()+timedelta(days=7)).strftime('%Y-%m-%d'),
    }

    metadata_df = pd.DataFrame([metadata])

    if not isinstance(pivot_table, pd.DataFrame):
        raise ValueError("pivot_table should be a pandas DataFrame")
    result_df = pd.concat([metadata_df, pivot_table], axis=1)

    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*csv")],
                                             title="Save the file")
    if save_path:
        result_df.to_csv(save_path, index=False, encoding='utf-8-sig')
        messagebox.showinfo("Information", f"The data file is {"tally" if is_tally else 
        "not tally. Do not import to Odoo. Please find developer"}")
    else:
        messagebox.showinfo("Information", "File saving cancelled")


def main():
    root = tk.Tk()
    root.title("Select an Option")
    root.geometry("300x150")

    def select_option(option):
        root.destroy()
        if option == "Shopee":
            filepath = filedialog.askopenfilename(title="Select the Shopee excel file")
            uom_filepath = filedialog.askopenfilename(title="Select the uom excel file")
            if filepath:
                pivot_table, empty_values, is_tally = process_shopee(filepath, uom_filepath)
                customer_info = "SHOPEE MALL - NEUTROVIS"
                create_csv(pivot_table, customer_info, empty_values, is_tally)
            else:
                messagebox.showinfo("Information", "File selection cancelled.")
        elif option == "Lazada":
            filepath = filedialog.askopenfilename(title="Select the Lazada excel file")
            uom_filepath = filedialog.askopenfilename(title="Select the UoM excel file")
            if filepath and uom_filepath:
                pivot_table, empty_values, is_tally = process_lazada(filepath, uom_filepath)
                customer_info = "LAZADA -"
                create_csv(pivot_table, customer_info, empty_values, is_tally)
            else:
                messagebox.showinfo("Information", "File selection cancelled.")
        else:
            messagebox.showinfo("Information", "Selected option is not yet supported.")

    options = ["Shopee", "Lazada", "TikTok"]
    for option in options:
        button = tk.Button(root, text=option, command=lambda opt=option: select_option(opt))
        button.pack(pady=5)

    root.mainloop()


if __name__ == "__main__":
    main()