import os
import sys
import pandas as pd

def clean_sales_data(input_csv_path, output_excel_path):
    print(f"[*] Reading raw CSV data from: {input_csv_path}")
    
    if not os.path.exists(input_csv_path):
        print(f"[!] File not found: {input_csv_path}")
        return

    try:
        # dataframe
        df = pd.read_csv(
            input_csv_path, 
            sep=',', 
            skipinitialspace=True, 
            encoding='utf-8-sig', 
            on_bad_lines='skip'
        )
    except Exception as e:
        print(f"[!] Error loading CSV file: {e}")
        return

    print(f"[*] Initial dataset shape: {df.shape[0]} rows, {df.shape[1]} columns")

    # Standardize column headers
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')

    #Remove fully empty rows and dupe entries
    df = df.dropna(how='all')
    df = df.drop_duplicates()

    # cleaning and formatting steps
    if 'customer_name' in df.columns:
        df['customer_name'] = df['customer_name'].astype(str).str.strip().str.title()
        df['customer_name'] = df['customer_name'].replace(['Nan', 'None', ''], 'Unknown Customer')

    # cleaning and formatting unit_price column and converting to float
    if 'unit_price' in df.columns:
        df['unit_price'] = (
            df['unit_price']
            .astype(str)
            .str.replace('$', '', regex=False)
            .str.replace(',', '', regex=False)
            .str.replace('"', '', regex=False)  
            .str.replace("'", '', regex=False)  
            .str.strip()
        )
        df['unit_price'] = pd.to_numeric(df['unit_price'], errors='coerce').fillna(0.0)

    # cleaning and formatting quantity column and converting to int
    if 'quantity' in df.columns:
        df['quantity'] = pd.to_numeric(df['quantity'], errors='coerce').fillna(0).astype(int)

    # calculate total_amount column if both quantity and unit_price exist
    if 'quantity' in df.columns and 'unit_price' in df.columns:
        df['total_amount'] = (df['quantity'] * df['unit_price']).round(2)

    # Parse and standardize dates to YYYY-MM-DD
    if 'order_date' in df.columns:
        df['order_date'] = pd.to_datetime(df['order_date'], errors='coerce').dt.strftime('%Y-%m-%d')
        df['order_date'] = df['order_date'].fillna('N/A')

    # build summary sheet grouped by customer_name
    if 'customer_name' in df.columns and 'total_amount' in df.columns:
        summary_df = (
            df.groupby('customer_name')
            .agg(
                Total_Orders=('quantity', 'count'),
                Total_Units_Bought=('quantity', 'sum'),
                Total_Revenue_USD=('total_amount', 'sum')
            )
            .reset_index()
            .sort_values(by='Total_Revenue_USD', ascending=False)
        )
    else:
        print("[!] Warning: 'customer_name' column not found. Creating empty summary sheet.")
        summary_df = pd.DataFrame(columns=['customer_name', 'Total_Orders', 'Total_Units_Bought', 'Total_Revenue_USD'])

    # Export to Excel
    print(f"[*] Exporting formatted Excel workbook to: {output_excel_path}")
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Cleaned_Data', index=False)
        summary_df.to_excel(writer, sheet_name='Customer_Summary', index=False)

    print("[*] Processing complete. Excel report generated successfully.")

if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # CLI Argument or fallback file path
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        input_file = os.path.join(script_dir, "raw_sales_data.csv")
        print("[!] No input CSV specified. Looking for 'raw_sales_data.csv' in script directory.")

    output_file = os.path.join(script_dir, "cleaned_sales_report.xlsx")
    clean_sales_data(input_file, output_file)