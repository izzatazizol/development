import re
import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox



# mapping = {
#     'CARTON 40': 'Ctn (50/40)',
#     'CARTON 12': 'Ctn (20/12)',
#     'CARTON 1000': 'Ctn (1/50/20)',
#     'CARTON 12': 'Ctn (25/12)',
#     'CARTON 24': 'Ctn (7/24)',
#     'CARTON 6': 'Ctn (6)',
#     'CARTON 20': 'Ctn (20)',
#     'CARTON 4': 'Ctn (4)',
#     'CARTON 20': 'Ctn (30/20)',
#     'CARTON 32': 'Ctn (50/32)',
#     'CARTON 144': 'Ctn (10/180)',
#     'CARTON 30': 'Ctn (30/40)',
# }



def process_aeon_big(aeon_big_filepath):
    aeon_big_df = pd.read_csv(aeon_big_filepath)
    empty_values = aeon_big_df['Supplier Item Sub Code'].isna() | (aeon_big_df['Supplier Item Sub Code'] == '')

    delivery_date = datetime.strptime(str(aeon_big_df['Delivery Date/Time'].mode()[0]), "%Y%m%d").strftime("%d/%m/%Y")

    aeon_big_df['Order Lines/Product'] = aeon_big_df['Supplier Item Sub Code'].fillna(aeon_big_df['Item Description'])
    aeon_big_df['Order Lines/Quantity'] = aeon_big_df['Qty/ Pack']
    aeon_big_df['Order Lines/Unit Price'] = aeon_big_df['Unit Price']
    aeon_big_df['Order Lines/Discount'] = ''

    all_dfs = []

    for po_number in aeon_big_df['Order No'].unique():
        po_df = aeon_big_df[aeon_big_df['Order No'] == po_number]
        metadata = {
            'Company': 'Neutrovis Sdn. Bhd.',
            'Customer': 'AEON BIG (M) SDN BHD',
            'Invoice Address': 'AEON BIG (M) SDN BHD',
            'Delivery Address': 'AEON BIG (M) SDN BHD, AEON KLRDC',
            'Order Date': datetime.now().strftime('%Y-%m-%d'),
            'Expiration': (datetime.now() + timedelta(days=7)).strftime('%Y-%m-%d'),
            'Pricelist': 'Public Pricelist',
            'Payment Terms': '60 Days',
            'Customer PO': po_number,
            'Salesperson': 'Candice Wong',
            'Sales Team': 'Modern Trade',
            'Warehouse': 'North Port Warehouse',
            'Delivery Date': delivery_date,
        }
        metadata_df = pd.DataFrame([metadata])
        selected_columns_df = po_df[['Order Lines/Product', 'Order Lines/Quantity',
                                     'Order Lines/Unit Price', 'Order Lines/Discount']].reset_index()
        # combine metadata with dynamic data
        combined_df = pd.concat([metadata_df, selected_columns_df], axis=1)
        all_dfs.append(combined_df)

    result_df = pd.concat(all_dfs, ignore_index=True)
    result_df = result_df.drop('index', axis=1)

    # save the csv file in a user-selected location
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*csv")],
                                             title="Save the file")
    if save_path:
        result_df.to_csv(save_path, index=False, encoding='utf-8-sig')
        num_empty = empty_values.sum()
        if num_empty > 0:
            messagebox.showinfo(
                "Information", f"There are {num_empty} with no SKU code. Please fill in the SKU code in the export "
                               f"file and do this again")
        messagebox.showinfo("Information", f"File saved successfully at {save_path}")
    else:
        messagebox.showinfo("Information", "File saving cancelled")


def process_aeon_gms_maxvalu_super(aeon_filepath):
    aeon_df = pd.read_csv(aeon_filepath, skiprows=1, skipfooter=1, engine='python')
    aeon_df['PO Number'] = aeon_df['PO Number'].apply(lambda x: '{:.0f}'.format(x) if pd.notna(x) else x)
    aeon_df = aeon_df[aeon_df['Line Type'].isin(['D'])]

    delivery_date = datetime.strptime(str(aeon_df['Delivery Date'].mode()[0]), "%Y%m%d").strftime("%d/%m/%Y")

    aeon_df['Customer'] = aeon_df.apply(lambda row: 'AEON WELLNESS'
    if 'WELLNESS' in str(row['Store Name'])
    else 'AEON CO.(M) BHD', axis=1)

    def determine_delivery_address(row):
        if row['Delivery To'] == 'DC 8010 DCXD':
            return 'AEON CO.(M) BHD, AEON KLRDC'
        elif row['Delivery To'] == 'STORE 1043 PTJ':
            return 'AEON CO.(M) BHD, AEON Putrajaya @ IOI City Mall (1043)'
        elif row['Delivery To'] == 'STORE 5005 WNMV':
            return 'AEON WELLNESS, AEON WELLNESS MIDVALLEY (5005)'
        elif row['Delivery To'] == 'DC 8015 XDWN':
            return 'AEON WELLNESS, AEON WELLNESS KUALA LUMPUR REGIONAL DISTRIBUTION CENTRE (KL RDC)'
        else:
            return ''

    aeon_df['Delivery Address'] = aeon_df.apply(determine_delivery_address, axis=1)

    empty_values = aeon_df['Supplier Item No'].isna() | (aeon_df['Supplier Item No'] == '')

    #aeon_df['UOM'] = aeon_df['UOM'].map(mapping)

    aeon_df['Customer PO'] = aeon_df['PO Number']
    aeon_df['Order Lines/Product'] = aeon_df['Supplier Item No'].fillna(aeon_df['Item Description'])
    aeon_df['Order Lines/Quantity'] = aeon_df['Order Qty']
    #aeon_df['Order Lines/Unit of Measure'] = aeon_df['UOM']
    aeon_df['Order Lines/Unit Price'] = aeon_df['Cost Unit Price']
    aeon_df['Order Lines/Discount'] = aeon_df['Total Discount']

    all_dfs = []

    for po_number in aeon_df['Customer PO'].unique():
        po_df = aeon_df[aeon_df['Customer PO'] == po_number]
        metadata = {
            'Company': 'Neutrovis Sdn. Bhd.',
            'Customer': aeon_df['Customer'].mode()[0],
            'Invoice Address': aeon_df['Customer'].mode()[0],
            'Delivery Address': aeon_df['Delivery Address'].mode()[0],
            'Order Date': datetime.now().strftime('%Y-%m-%d'),
            'Expiration': (datetime.now() + timedelta(days=7)).strftime('%Y-%m-%d'),
            'Pricelist': 'Public Pricelist',
            'Payment Terms': '60 Days',
            'Customer PO': po_number,
            'Salesperson': 'Chew Mei Feng',
            'Sales Team': 'Modern Trade',
            'Warehouse': 'North Port Warehouse',
            'Delivery Date': delivery_date,
        }
        # convert metadata to dataframe with only one row
        metadata_df = pd.DataFrame([metadata])

        selected_columns_df = po_df[['Order Lines/Product', 'Order Lines/Quantity',
                                     'Order Lines/Unit Price', 'Order Lines/Discount']].reset_index()
        # combine metadata with dynamic data
        combined_df = pd.concat([metadata_df, selected_columns_df], axis=1)
        all_dfs.append(combined_df)

    result_df = pd.concat(all_dfs, ignore_index=True)
    result_df = result_df.drop('index', axis=1)

    # save the csv file in a user-selected location
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*csv")],
                                             title="Save the file")
    if save_path:
        result_df.to_csv(save_path, index=False, encoding='utf-8-sig')
        num_empty = empty_values.sum()
        if num_empty > 0:
            messagebox.showinfo(
                "Information", f"There are {num_empty} with no SKU code. Please fill in the SKU code in the export "
                               f"file and do this again")
        messagebox.showinfo("Information", f"File saved successfully at {save_path}")
    else:
        messagebox.showinfo("Information", "File saving cancelled")


def process_caring(caring_filepath):
    # None means the file does not have a header row
    caring_df = pd.read_excel(caring_filepath, header=None, skiprows=10)
    # get value from 2nd row and 2nd column
    purchase_order = caring_df.at[1, 1]
    delivery_date = caring_df.at[4, 1]
    delivery_date = datetime.strptime(str(delivery_date), "%Y%m%d").strftime("%d/%m/%Y")
    # Load data, skipping the first 10 rows and using the 11th row as the header
    #caring_df = pd.read_excel(caring_filepath, skiprows=10)
    empty_values = caring_df['SKU code'].isna() | (caring_df['SKU code'] == '')
    # copy other columns directly
    caring_df['Order Lines/Product'] = caring_df['SKU code'].fillna(caring_df['Description'])
    caring_df['Order Lines/Quantity'] = caring_df['Delivery Quantity']
    caring_df['Order Lines/Unit Price'] = caring_df['Unit Price (MYR)']
    caring_df['Order Lines/Discount'] = ''

    # metadata for the first row
    metadata = {
        'Company': 'Neutrovis Sdn. Bhd.',
        'Customer': 'CARING PHARMACY RETAIL MANAGEMENT SDN BHD',
        'Invoice Address': 'CARING PHARMACY RETAIL MANAGEMENT SDN BHD',
        'Delivery Address': 'CARING PHARMACY RETAIL MANAGEMENT SDN BHD, CARING PHARMACY RETAIL MANAGEMENT SDN BHD',
        'Order Date': datetime.now().strftime('%Y-%m-%d'),
        'Expiration': (datetime.now()+timedelta(days=7)).strftime('%Y-%m-%d'),
        'Pricelist': 'Public Pricelist',
        'Payment Terms': '60 Days',
        'Customer PO': purchase_order,
        'Salesperson': 'Kenny Tan',
        'Sales Team': 'H&B Team',
        'Warehouse': 'North Port Warehouse',
        'Delivery Date': delivery_date,
    }
    #convert metadata to dataframe with only one row
    metadata_df = pd.DataFrame([metadata])
    selected_columns_df = caring_df[['Order Lines/Product', 'Order Lines/Quantity',
                                 'Order Lines/Unit Price', 'Order Lines/Discount']].reset_index()
    #combine metadata with dynamic data
    result_df = pd.concat([metadata_df, selected_columns_df], axis=1)

    # save the csv file in a user-selected location
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*csv")],
                                             title="Save the file")
    if save_path:
        result_df.to_csv(save_path, index=False, encoding='utf-8-sig')
        num_empty = empty_values.sum()
        if num_empty > 0:
            messagebox.showinfo(
                "Information", f"There are {num_empty} with no SKU code. Please fill in the SKU code in the export "
                               f"file and do this again")
        messagebox.showinfo("Information", f"File saved successfully at {save_path}")
    else:
        messagebox.showinfo("Information", "File saving cancelled")


def process_giant(giant_filepath):
    giant_df = pd.read_csv(giant_filepath, skiprows=1)
    giant_df = giant_df.iloc[:-1]

    empty_values = giant_df['SKU code'].isna() | (giant_df['SKU code'] == '')

    delivery_date = datetime.strptime(str(giant_df['DELIVERY_DATE'].mode()[0]), "%Y%m%d").strftime("%d/%m/%Y")

    giant_df['Order Lines/Product'] = giant_df['SKU code'].fillna(giant_df['PRD_DESC'])
    giant_df['Order Lines/Quantity'] = giant_df['ORDER_QTY']
    giant_df['Order Lines/Unit Price'] = giant_df['UNIT_PRICE']
    giant_df['Order Lines/Discount'] = ''

    all_dfs = []

    for po_number in giant_df['PO_NUMBER'].unique():
        po_df = giant_df[giant_df['PO_NUMBER'] == po_number]

        metadata = {
            'Company': 'Neutrovis Sdn. Bhd.',
            'Customer': 'GCH RETAIL (MALAYSIA) SDN BHD',
            'Invoice Address': 'GCH RETAIL (MALAYSIA) SDN BHD',
            'Delivery Address': 'GCH RETAIL (MALAYSIA) SDN BHD, DC SEPANG FLOW THROUGH',
            'Order Date': datetime.now().strftime('%Y-%m-%d'),
            'Expiration': (datetime.now()+timedelta(days=7)).strftime('%Y-%m-%d'),
            'Pricelist': 'Public Pricelist',
            'Payment Terms': '60 Days',
            'Customer PO': po_number,
            'Salesperson': 'Loh Mei Mei',
            'Sales Team': 'Modern Trade',
            'Warehouse': 'North Port Warehouse',
            'Delivery Date': delivery_date,
        }
        metadata_df = pd.DataFrame([metadata])
        selected_columns_df = po_df[['Order Lines/Product', 'Order Lines/Quantity',
                                     'Order Lines/Unit Price', 'Order Lines/Discount']].reset_index()
        combined_df = pd.concat([metadata_df, selected_columns_df], axis=1)
        all_dfs.append(combined_df)

    result_df = pd.concat(all_dfs, ignore_index=True)
    result_df = result_df.drop('index', axis=1)

    #save the csv file in a user-selected location
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*csv")],
                                             title="Save the file")
    if save_path:
        result_df.to_csv(save_path, index=False, encoding='utf-8-sig')
        num_empty = empty_values.sum()
        if num_empty > 0:
            messagebox.showinfo(
                "Information", f"There are {num_empty} with no SKU code. Please fill in the SKU code in the export "
                               f"file and do this again")
        messagebox.showinfo("Information", f"File saved successfully at {save_path}")
    else:
        messagebox.showinfo("Information", "File saving cancelled")


def process_guardian(guardian_filepath):
    guardian_df = pd.read_csv(guardian_filepath)
    empty_values = guardian_df['VENDOR_PART_NO'].isna() | (guardian_df['VENDOR_PART_NO'] == '')

    delivery_date = datetime.strptime(str(guardian_df['DELIVERY_DATE'].mode()[0]), "%Y%m%d").strftime("%d/%m/%Y")

    guardian_df['Customer PO'] = guardian_df['PO_NUMBER']
    guardian_df['Order Lines/Product'] = guardian_df['VENDOR_PART_NO'].fillna(guardian_df['PRD_DESC'])
    guardian_df['Order Lines/Quantity'] = guardian_df['ORDER_QTY']
    guardian_df['Order Lines/Unit Price'] = guardian_df['UNIT_PRICE']
    guardian_df['Order Lines/Discount'] = ''

    all_dfs = []

    for po_number in guardian_df['PO_NUMBER'].unique():
        po_df = guardian_df[guardian_df['PO_NUMBER'] == po_number]
        metadata = {
            'Company': 'Neutrovis Sdn. Bhd.',
            'Customer': 'GUARDIAN HEALTH AND BEAUTY SDN BHD',
            'Invoice Address': 'GUARDIAN HEALTH AND BEAUTY SDN BHD',
            'Delivery Address': 'GUARDIAN HEALTH AND BEAUTY SDN BHD, GUARDIAN HEALTH AND BEAUTY SDN BHD',
            'Order Date': datetime.now().strftime('%Y-%m-%d'),
            'Expiration': (datetime.now()+timedelta(days=7)).strftime('%Y-%m-%d'),
            'Pricelist': 'Public Pricelist',
            'Payment Terms': '60 Days',
            'Customer PO': po_number,
            'Salesperson': 'Kenny Tan',
            'Sales Team': 'H&B Team',
            'Warehouse': 'North Port Warehouse',
            'Delivery Date': delivery_date,
        }
        metadata_df = pd.DataFrame([metadata])
        selected_columns_df = po_df[['Order Lines/Product', 'Order Lines/Quantity',
                                     'Order Lines/Unit Price', 'Order Lines/Discount']].reset_index()
        combined_df = pd.concat([metadata_df,selected_columns_df], axis=1)
        all_dfs.append(combined_df)

    result_df = pd.concat(all_dfs, ignore_index=True)
    result_df = result_df.drop('index', axis=1)

    # save the csv file in a user-selected location
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*csv")],
                                             title="Save the file")
    if save_path:
        result_df.to_csv(save_path, index=False, encoding='utf-8-sig')
        num_empty = empty_values.sum()
        if num_empty > 0:
            messagebox.showinfo(
                "Information", f"There are {num_empty} with no SKU code. Please fill in the SKU code in the export "
                               f"file and do this again")
        messagebox.showinfo("Information", f"File saved successfully at {save_path}")
    else:
        messagebox.showinfo("Information", "File saving cancelled")


def process_jayagrocer(jayagrocer_filepath):
    jayagrocer_df = pd.read_excel(jayagrocer_filepath)
    empty_values = jayagrocer_df['Article Code'].isna() | (jayagrocer_df['Article Code'] == '')
    jayagrocer_df['Customer PO'] = jayagrocer_df['PO No']
    jayagrocer_df['Order Lines/Product'] = jayagrocer_df['Article Code'].fillna(jayagrocer_df['Item Description'])
    jayagrocer_df['Order Lines/Quantity'] = jayagrocer_df['Ordered Qty']
    jayagrocer_df['Order Lines/Unit Price'] = jayagrocer_df['Unit Cost']
    jayagrocer_df['Order Lines/Discount'] = ''

    all_dfs = []

    for po_number in jayagrocer_df['Customer PO'].unique():
        po_df = jayagrocer_df[jayagrocer_df['Customer PO'] == po_number]
        metadata = {
            'Company': 'Neutrovis Sdn. Bhd.',
            'Customer': 'TRENDCELL SDN BHD',
            'Invoice Address': 'TRENDCELL SDN BHD',
            'Delivery Address': 'TRENDCELL SDN BHD, (544047-T) - DC5',
            'Order Date': datetime.now().strftime('%Y-%m-%d'),
            'Expiration': (datetime.now()+timedelta(days=7)).strftime('%Y-%m-%d'),
            'Pricelist': 'Public Pricelist',
            'Payment Terms': '60 Days',
            'Customer PO': po_number,
            'Salesperson': 'Loh Mei Mei',
            'Sales Team': 'Modern Trade',
            'Warehouse': 'North Port Warehouse',
            'Delivery Date': datetime.now().strftime('%Y-%m-%d'),
        }
        metadata_df = pd.DataFrame([metadata])
        selected_columns_df = po_df[['Order Lines/Product', 'Order Lines/Quantity',
                                     'Order Lines/Unit Price', 'Order Lines/Discount']].reset_index()
        combined_df = pd.concat([metadata_df,selected_columns_df], axis=1)
        all_dfs.append(combined_df)

    result_df = pd.concat(all_dfs, ignore_index=True)
    result_df = result_df.drop('index', axis=1)

    # save the csv file in a user-selected location
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*csv")],
                                             title="Save the file")
    if save_path:
        result_df.to_csv(save_path, index=False, encoding='utf-8-sig')
        num_empty = empty_values.sum()
        if num_empty > 0:
            messagebox.showinfo(
                "Information", f"There are {num_empty} with no SKU code. Please fill in the SKU code in the export "
                               f"file and do this again")
        messagebox.showinfo("Information", f"File saved successfully at {save_path}")
    else:
        messagebox.showinfo("Information", "File saving cancelled")


def process_lotus(lotus_filepath):
    lotus_df = pd.read_csv(lotus_filepath)
    empty_values = lotus_df['Supplier Item Sub Code'].isna() | (lotus_df['Supplier Item Sub Code'] == '')

    delivery_date = datetime.strptime(str(lotus_df['Delivery Date/Time'].mode()[0]), "%Y%m%d").strftime("%d/%m/%Y")

    lotus_df['Customer PO'] = lotus_df['Order No']
    lotus_df['Order Lines/Product'] = lotus_df['Supplier Item Sub Code'].fillna(lotus_df['Item Description'])
    lotus_df['Order Lines/Quantity'] = lotus_df['Total Qty']
    lotus_df['Order Lines/Unit Price'] = lotus_df['Order Unit Price']
    lotus_df['Order Lines/Discount'] = ''

    all_dfs = []

    for po_number in lotus_df['Customer PO'].unique():
        po_df = lotus_df[lotus_df['Customer PO'] == po_number]
        metadata = {
            'Company': 'Neutrovis Sdn. Bhd.',
            'Customer': 'LOTUSS STORES (MALAYSIA) SDN BHD',
            'Invoice Address': 'LOTUSS STORES (MALAYSIA) SDN BHD',
            'Delivery Address': "LOTUS'S AMBIENT DISTRIBUTION CENTRE (LADC)",
            'Order Date': datetime.now().strftime('%Y-%m-%d'),
            'Expiration': (datetime.now()+timedelta(days=7)).strftime('%Y-%m-%d'),
            'Pricelist': 'Public Pricelist',
            'Payment Terms': '60 Days',
            'Customer PO': po_number,
            'Salesperson': 'Candice Wong',
            'Sales Team': 'Modern Trade',
            'Warehouse': 'North Port Warehouse',
            'Delivery Date': delivery_date,
        }
        metadata_df = pd.DataFrame([metadata])
        selected_columns_df = po_df[['Order Lines/Product', 'Order Lines/Quantity',
                                     'Order Lines/Unit Price', 'Order Lines/Discount']].reset_index()
        combined_df = pd.concat([metadata_df, selected_columns_df], axis=1)
        all_dfs.append(combined_df)

    result_df = pd.concat(all_dfs, axis=1)
    result_df = result_df.drop('index', axis=1)

    #save the csv file in a user-selected location
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*csv")],
                                             title="Save the file")
    if save_path:
        result_df.to_csv(save_path, index=False, encoding='utf-8-sig')
        num_empty = empty_values.sum()
        if num_empty > 0:
            messagebox.showinfo(
                "Information", f"There are {num_empty} with no SKU code. Please fill in the SKU code in the export "
                               f"file and do this again")
        messagebox.showinfo("Information", f"File saved successfully at {save_path}")
    else:
        messagebox.showinfo("Information", "File saving cancelled")


def process_manjaku(manjaku_filepath):
    manjaku_df = pd.read_excel(manjaku_filepath)
    empty_values = manjaku_df['Article Code'].isna() | (manjaku_df['Article Code'] == '')

    delivery_date = datetime.strptime(str(manjaku_df['Delivery Date'].mode()[0]), "%Y%m%d").strftime("%d/%m/%Y")

    manjaku_df['Customer PO'] = manjaku_df['PO No']
    manjaku_df['Order Lines/Product'] = manjaku_df['Article Code'].fillna(manjaku_df['Description'])
    manjaku_df['Order Lines/Quantity'] = manjaku_df['Ordered Qty']
    manjaku_df['Order Lines/Unit Price'] = manjaku_df['Unit Cost']
    manjaku_df['Order Lines/Discount'] = ''

    all_dfs = []

    for po_number in manjaku_df['Customer PO'].unique():
        po_df = manjaku_df[manjaku_df['Customer PO'] == po_number]
        metadata = {
            'Company': 'Neutrovis Sdn. Bhd.',
            'Customer': 'MANJAKU HOLDINGS SDN BHD',
            'Invoice Address': 'MANJAKU HOLDINGS SDN BHD',
            'Delivery Address': 'MANJAKU HOLDINGS SDN BHD, MANJAKU PRIMA SAUJANA',
            'Order Date': datetime.now().strftime('%Y-%m-%d'),
            'Expiration': (datetime.now()+timedelta(days=7)).strftime('%Y-%m-%d'),
            'Pricelist': 'Public Pricelist',
            'Payment Terms': '60 Days',
            'Customer PO': po_number,
            'Salesperson': 'Chew Mei Feng',
            'Sales Team': 'Modern Trade',
            'Warehouse': 'North Port Warehouse',
            'Delivery Date': delivery_date,
        }
        metadata_df = pd.DataFrame([metadata])
        selected_columns_df = po_df[['Order Lines/Product', 'Order Lines/Quantity',
                                     'Order Lines/Unit Price', 'Order Lines/Discount']].reset_index()
        combined_df = pd.concat([metadata_df, selected_columns_df], axis=1)
        all_dfs.append(combined_df)

    result_df = pd.concat(all_dfs, axis=1)
    result_df = result_df.drop('index', axis=1)

    #save the csv file in a user-selected location
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*csv")],
                                             title="Save the file")
    if save_path:
        result_df.to_csv(save_path, index=False, encoding='utf-8-sig')
        num_empty = empty_values.sum()
        if num_empty > 0:
            messagebox.showinfo(
                "Information", f"There are {num_empty} with no SKU code. Please fill in the SKU code in the export "
                               f"file and do this again")
        messagebox.showinfo("Information", f"File saved successfully at {save_path}")
    else:
        messagebox.showinfo("Information", "File saving cancelled")


def process_mynews(mynews_filepath):
    mynews_df = pd.read_csv(mynews_filepath)
    empty_values = mynews_df['Article Code'].isna() | (mynews_df['Article Code'] == '')

    mynews_df['Customer PO'] = mynews_df['PO No.']
    mynews_df['Order Lines/Product'] = mynews_df['Article Code'].fillna(mynews_df['Description'])
    mynews_df['Order Lines/Quantity'] = mynews_df['Ordered Qty']
    mynews_df['Order Lines/Unit Price'] = mynews_df['Unit Cost']
    mynews_df['Order Lines/Discount'] = ''

    all_dfs = []

    for po_number in mynews_df['Customer PO'].unique():
        po_df = mynews_df[mynews_df['Customer PO'] == po_number]
        metadata = {
            'Company': 'Neutrovis Sdn. Bhd.',
            'Customer': 'MANJAKU HOLDINGS SDN BHD',
            'Invoice Address': 'MANJAKU HOLDINGS SDN BHD',
            'Delivery Address': 'MANJAKU HOLDINGS SDN BHD, MANJAKU PRIMA SAUJANA',
            'Order Date': datetime.now().strftime('%Y-%m-%d'),
            'Expiration': (datetime.now()+timedelta(days=7)).strftime('%Y-%m-%d'),
            'Pricelist': 'Public Pricelist',
            'Payment Terms': '60 Days',
            'Customer PO': po_number,
            'Salesperson': 'Chew Mei Feng',
            'Sales Team': 'Modern Trade',
            'Warehouse': 'North Port Warehouse',
            'Delivery Date': datetime.now().strftime('%Y-%m-%d'),
        }
        metadata_df = pd.DataFrame([metadata])
        selected_columns_df = po_df[['Order Lines/Product', 'Order Lines/Quantity',
                                     'Order Lines/Unit Price', 'Order Lines/Discount']].reset_index()
        combined_df = pd.concat([metadata_df, selected_columns_df], axis=1)
        all_dfs.append(combined_df)

    result_df = pd.concat(all_dfs, axis=1)
    result_df = result_df.drop('index', axis=1)

    #save the csv file in a user-selected location
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*csv")],
                                             title="Save the file")
    if save_path:
        result_df.to_csv(save_path, index=False, encoding='utf-8-sig')
        num_empty = empty_values.sum()
        if num_empty > 0:
            messagebox.showinfo(
                "Information", f"There are {num_empty} with no SKU code. Please fill in the SKU code in the export "
                               f"file and do this again")
        messagebox.showinfo("Information", f"File saved successfully at {save_path}")
    else:
        messagebox.showinfo("Information", "File saving cancelled")


def process_watson(watson_filepath):
    watson_df = pd.read_csv(watson_filepath)
    # find the index of the first empty row - this assumes the gap has NaNs or empty strings
    # get the first index of the row where all columns are NaN
    gap_index = watson_df.index[watson_df.isnull().all(axis=1)][0]
    # keep only the rows before the gap
    watson_df = watson_df.loc[:gap_index - 1]

    empty_values = watson_df['Supplier Item Code'].isna() | (watson_df['Supplier Item Code'] == '')

    watson_df['Customer PO'] = watson_df['Order No']
    watson_df['Order Lines/Product'] = watson_df['Supplier Item Code'].fillna(watson_df['Item Description'])
    watson_df['Order Lines/Quantity'] = watson_df['Accepted Qty']
    watson_df['Order Lines/Unit Price'] = watson_df['Unit Price']
    watson_df['Order Lines/Discount'] = ''

    all_dfs = []

    for po_number in watson_df['Customer PO'].unique():
        po_df = watson_df[watson_df['Customer PO'] == po_number]
        metadata = {
            'Company': 'Neutrovis Sdn. Bhd.',
            'Customer': "WATSON 'S PERSONAL CARE STORES SDN BHD",
            'Invoice Address': "WATSON 'S PERSONAL CARE STORES SDN BHD",
            'Delivery Address': "WATSON 'S PERSONAL CARE STORES SDN BHD, 961 961 DC2",
            'Order Date': datetime.now().strftime('%Y-%m-%d'),
            'Expiration': (datetime.now() + timedelta(days=7)).strftime('%Y-%m-%d'),
            'Pricelist': 'Public Pricelist',
            'Payment Terms': '60 Days',
            'Customer PO': po_number,
            'Salesperson': 'Samantha Poon',
            'Sales Team': 'H&B Team',
            'Warehouse': 'North Port Warehouse',
            'Delivery Date': datetime.now().strftime('%Y-%m-%d'),
        }
        # convert metadata to dataframe with only one row
        metadata_df = pd.DataFrame([metadata])

        selected_columns_df = po_df[['Order Lines/Product', 'Order Lines/Quantity',
                                     'Order Lines/Unit Price', 'Order Lines/Discount']].reset_index()
        # combine metadata with dynamic data
        combined_df = pd.concat([metadata_df, selected_columns_df], axis=1)
        all_dfs.append(combined_df)

    result_df = pd.concat(all_dfs, ignore_index=True)
    result_df = result_df.drop('index', axis=1)

    # save the csv file in a user-selected location
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*csv")],
                                             title="Save the file")
    if save_path:
        result_df.to_csv(save_path, index=False, encoding='utf-8-sig')
        num_empty = empty_values.sum()
        if num_empty > 0:
            messagebox.showinfo(
                "Information", f"There are {num_empty} with no SKU code. Please fill in the SKU code in the export "
                               f"file and do this again")
        messagebox.showinfo("Information", f"File saved successfully at {save_path}")
    else:
        messagebox.showinfo("Information", "File saving cancelled")

def main():
    root = tk.Tk()
    root.title("Select an Option")
    root.geometry("300x370")

    def select_option(option):
        root.destroy()
        if option == "Aeon Big":
            aeon_big_filepath = filedialog.askopenfilename(title="Select the Aeon Big csv file")
            if aeon_big_filepath:
                process_aeon_big(aeon_big_filepath)
            else:
                messagebox.showinfo("Information", "File selection cancelled.")
        elif option == "Aeon":
            aeon_filepath = filedialog.askopenfilename(title="Select the Aeon csv file")
            if aeon_filepath:
                process_aeon_gms_maxvalu_super(aeon_filepath)
            else:
                messagebox.showinfo("Information", "File selection cancelled.")
        elif option == "Caring":
            caring_filepath = filedialog.askopenfilename(title="Select the Caring excel file")
            if caring_filepath:
                process_caring(caring_filepath)
            else:
                messagebox.showinfo("Information", "File selection cancelled.")
        elif option == "Giant":
            giant_filepath = filedialog.askopenfilename(title="Select the Giant csv file")
            if giant_filepath:
                process_giant(giant_filepath)
            else:
                messagebox.showinfo("Information", "File selection cancelled.")
        elif option == "Guardian":
            guardian_filepath = filedialog.askopenfilename(title="Select the Guardian csv file")
            if guardian_filepath:
                process_guardian(guardian_filepath)
            else:
                messagebox.showinfo("Information", "File selection cancelled.")
        elif option == "Jaya Grocer":
            jayagrocer_filepath = filedialog.askopenfilename(title="Select the Jaya Grocer excel file")
            if jayagrocer_filepath:
                process_jayagrocer(jayagrocer_filepath)
            else:
                messagebox.showinfo("Information", "File selection cancelled.")
        elif option == "Lotus":
            lotus_filepath = filedialog.askopenfilename(title="Select the Lotus csv file")
            if lotus_filepath:
                process_lotus(lotus_filepath)
            else:
                messagebox.showinfo("Information", "File selection cancelled.")
        elif option == "Manjaku":
            manjaku_filepath = filedialog.askopenfilename(title="Select the Manjaku excel file")
            if manjaku_filepath:
                process_manjaku(manjaku_filepath)
            else:
                messagebox.showinfo("Information", "File selection cancelled.")
        elif option == "MyNews":
            mynews_filepath = filedialog.askopenfilename(title="Select the MyNews csv file")
            if mynews_filepath:
                process_mynews(mynews_filepath)
            else:
                messagebox.showinfo("Information", "File selection cancelled.")
        elif option == "Watson":
            watson_filepath = filedialog.askopenfilename(title="Select the Watson csv file")
            if watson_filepath:
                process_watson(watson_filepath)
            else:
                messagebox.showinfo("Information", "File selection cancelled.")
        else:
            messagebox.showinfo("Information", "Selected option is not yet supported.")

    # Create buttons for each option
    options = ["Aeon Big", "Aeon", "Caring", "Giant", "Guardian", "Jaya Grocer", "Lotus", "Manjaku", "MyNews", "Watson"]
    for option in options:
        button = tk.Button(root, text=option, command=lambda opt=option: select_option(opt))
        button.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()