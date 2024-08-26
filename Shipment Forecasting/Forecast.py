import pandas as pd
import re
from pathlib import Path

# Read Files
# [Shipping Agent Report]
# [Master File]
# [Product Family]

DESKTOP_DIR = Path.home() / "OneDrive" / "Desktop" / "Automation" 

forecast_df = pd.read_excel(rf"{str(DESKTOP_DIR)}\Outbound Forecast\Reports\New Forecast Report.xlsx", sheet_name='Sheet1')
masterfile_df = pd.read_csv(r"C:\Users\Teams Folder\Forecast\Dataset")
ProdFamily_df = pd.read_excel(rf"{str(DESKTOP_DIR)}\\Outbound Forecast\Family Table\Product Family.xlsx", sheet_name= 'ALL')

# To change [Ship Out Date] to date
dtypes = {
    '[Ship Out Date]':'datetime64[ns]'
}

masterfile_df = masterfile_df.astype(dtypes)

masterfile_df.columns = masterfile_df.columns.str.replace(r'[\[\]]','',regex = True)

masterfile_df['Prod Family'] = masterfile_df['Prod Family'].astype(object)
masterfile_df['Config Type'] = masterfile_df['Config Type'].astype(object)

### Data Cleaning: forecast_df[Prod Family], [Part Number w/o SO]

# Add in [Prod Family] to the forecast_df based on Config Type
merged_df = forecast_df.merge(
    ProdFamily_df,  # CHANGE TO MERGED PRODUCT FAMILY LIST
    how = 'left',
    left_on = 'Config Type',
    right_on = 'CONFIG TYPE'
)

forecast_df['Prod Family'] = merged_df['Prod Family']


# Function to remove the Sales Order part from the PO Part
def remove_sales_order(po_part, sale_order):
    if f"-{sale_order}" in po_part:
        return po_part.replace(f"-{sale_order}", "")
    return po_part

# Create the new column 'Part Number w/o SO'
forecast_df['PO Part'] = forecast_df['PO Part'].astype(str)
forecast_df['Part Number w/o SO'] = forecast_df.apply(lambda row: remove_sales_order(row['PO Part'], row['Sale order']), axis=1)

# Function to create the PO Part

def create_po_part(row):
    po_part_wo_so = row['Part Number w/o SO']
    sale_order = row['Sale order']

    # Handle NaN or unexpected None values early
    if pd.isna(po_part_wo_so):
        return None

    po_part_wo_so = str(po_part_wo_so)  # Ensure it's a string

    # Check if there is an alphabet at the end
    if re.match(r'.*[A-Za-z]$', po_part_wo_so):
        result = re.sub(r'([A-Za-z])$', f'-{sale_order}\\1', po_part_wo_so)
        return result
    else:
        result = f'{po_part_wo_so}-{sale_order}'
        return result



# Function to find matching parts and add rows from master_df
def find_master(row, master_df, forecast_df):
    customer_name = row['Customer Name']
    product_family = row['Prod Family']
    config_type = row['Config Type']
    main_part = row['Part Number w/o SO']
    sale_order = row['Sale order']
    # print(f"SO: {sale_order}")
    def safe_lower(x):
        if isinstance(x, str):
            return x.lower()
        return ""

    def filter_master_df(master_df, customer_name, key, value, main_part):
        return master_df[
            (master_df['Customer Name'].apply(safe_lower) == safe_lower(customer_name)) &
            (master_df[key].apply(safe_lower) == safe_lower(value)) &
            (master_df['Part Number w/o SO'] == main_part)
        ]

    # Escape special characters in customer_name_lower
    customer_name_lower = safe_lower(customer_name)
    escaped_customer_name_lower = re.escape(customer_name_lower)

    # Check if the customer name exists in master_df
    if customer_name_lower and not master_df['Customer Name'].apply(safe_lower).str.contains(escaped_customer_name_lower).any():
        row['Search Master'] = 'N/A - Customer Not Found'
        return pd.DataFrame([row]), False

    # First attempt: match by Config Type
    filtered_df = filter_master_df(master_df, customer_name, 'Config Type', config_type, main_part)

    # Second attempt: if no match by Config Type, match by Prod Family
    if filtered_df.empty:
        product_family_lower = safe_lower(product_family)
        if product_family_lower:
            escaped_product_family_lower = re.escape(product_family_lower)
            filtered_df = filter_master_df(master_df, customer_name, 'Prod Family', escaped_product_family_lower, main_part)

    if filtered_df.empty:
        # Determine why no match was found
        config_type_lower = safe_lower(config_type)
        escaped_config_type_lower = re.escape(config_type_lower)
        if config_type_lower and not master_df['Config Type'].apply(safe_lower).str.contains(escaped_config_type_lower).any():
            row['Search Master'] = 'N/A - Config Type Not Found'
        elif product_family_lower and not master_df['Prod Family'].apply(safe_lower).str.contains(escaped_product_family_lower).any():
            row['Search Master'] = 'N/A - Prod Family Not Found'
        else:
            row['Search Master'] = 'N/A - No Matching Part'
        return pd.DataFrame([row]), False
    
    if filtered_df["Part Type"].iloc[0]=="Sub Part":
        return pd.DataFrame([row]), True
    # print("Filtered DataFrame after filtering:", filtered_df[['Part Number w/o SO']])

    # Add in the latest SO when a match is found
    row['Previous SO'] = filtered_df['Sale order'].iloc[0]  # Use the first match found

    # Regular expression to split and search Part Number w/o SO
    alph_search = re.match(r'^(.*?)([A-Z][0-9]*)?$', main_part)
    alph=""
    alph_id = alph_search.start(2) if alph_search.group(2) else len(main_part)
    if alph_id<9:
        alph_id=len(main_part)
    alph = main_part[alph_id:]
    main_part = main_part[:alph_id]
    
    # print("main_part: ",main_part, "alph: " ,alph)
    
    # df to store the subparts for this mother part
    if alph:
        subparts_df = master_df[
            (master_df['Customer Name'].apply(safe_lower) == safe_lower(customer_name)) &
            (master_df['Config Type'].apply(safe_lower) == safe_lower(config_type)) & 
            (master_df['Prod Family'].apply(safe_lower) == safe_lower(product_family)) &
            (master_df['Part Number w/o SO'].str.contains(main_part, na=False)) &
            (master_df['Part Number w/o SO'].str.endswith(alph, na=False)) &
            (master_df['Part Type'] == 'Sub Part')
        ].reset_index(drop=True)
    else:
        subparts_df=master_df[
            (master_df['Customer Name'].apply(safe_lower) == safe_lower(customer_name)) &
            (master_df['Config Type'].apply(safe_lower) == safe_lower(config_type)) & 
            (master_df['Prod Family'].apply(safe_lower) == safe_lower(product_family)) &
            (master_df['Part Number w/o SO'].str.contains(main_part, na=False)) &
            (master_df['Part Number w/o SO'].apply(lambda x: x.split('-')[-1].isdigit())) &
            (master_df['Part Type'] == 'Sub Part')
        ].reset_index(drop=True)

    # Check if any subparts already exist within the same Sale order to avoid duplicates
    existing_parts = forecast_df[forecast_df['Sale order'] == sale_order]['Part Number w/o SO'].unique()
    subparts_df = subparts_df[~subparts_df['Part Number w/o SO'].isin(existing_parts)].reset_index(drop=True)

    if subparts_df.empty:
        return pd.DataFrame([row]), True


    # s_df to store the data for each subpart that matches the final report
    
    s_df = pd.DataFrame([row] * len(subparts_df)).reset_index(drop=True) # To create rows by the number of subparts found
    # print(f"part number without SO: {row['Part Number w/o SO']}")
    # print(len(subparts_df))
    
    s_df['Part Number w/o SO'] = subparts_df['Part Number w/o SO'] # To take the Part Number from Master
    s_df['Module type'] = 'SHIP WITH'
    s_df['Part description'] = 'SHIP WITH KITS'
    s_df['Part Type'] = subparts_df['Part Type']
    s_df['Previous SO'] = subparts_df['Sale order']  # Copy 'Sale order' from master_df

    #print("Row before applying create_po_part:")

    s_df['PO Part'] = s_df.apply(create_po_part, axis=1)  # type mismatch -float

    s_df['Search Master'] = 'New'  # For those found in Master, mark as 'New'
    ### Dimension
    s_df['Length (M)'] = ''
    s_df['Width (M)'] = ''
    s_df['Height (M)'] = ''
    s_df['Weight (KG)'] = ''

    result_df = pd.concat([pd.DataFrame([row]), s_df])

    return result_df, True

all_rows = []

# Search master data row by row
for _, row in forecast_df.iterrows():
    # if row["Sale order"]=='K3620':
    #     print("current row:" , row['Part Number w/o SO'])
    matched_parts, is_matched = find_master(row, masterfile_df, forecast_df)
    all_rows.append(matched_parts)

final_df = pd.concat(all_rows, ignore_index=True)

# Export the dataframes to excel
report_path = r"C:\Users\OneDrive\Desktop\Automation\Outbound Forecast\Forecast Report\New Forecast Report_Final.xlsx"

with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
    final_df.to_excel(writer, sheet_name="Forecast Report", index=False)

print(f"Forecast Report Updated. \nPlease check: {report_path}")

 