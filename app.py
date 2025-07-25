from flask import Flask, request, send_file, render_template
import pandas as pd
import tempfile
import os
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)

def process_excel(input_path):
    # Read the 'New Data' sheet
    df = pd.read_excel(input_path, sheet_name="New Data")

    # Filter for commercial units
    df_commercial = df[df['Unit type'].isin(['Commercial', 'Office', 'Retail'])].dropna(subset=['Unit Reference', 'Fund type'])

    # Group by Unit and Fund type
    unit_summary = (
        df_commercial.groupby(['Fund type', 'Unit Reference', 'Name.1'], dropna=False)
        .agg(
            Total_Gross_Demanded=('Gross Demanded', 'sum'),
            Total_Settled=('Settled', 'sum')
        )
        .reset_index()
    )

    unit_summary.rename(columns={'Name.1': 'Name'}, inplace=True)

    # Calculate outstanding amount
    unit_summary['Outstanding'] = unit_summary['Total_Gross_Demanded'] - unit_summary['Total_Settled']

    # Filter for units with outstanding balances
    units_with_outstanding = unit_summary[unit_summary['Outstanding'] > 0]

    # Aggregate totals per fund type
    fund_type_summary = (
        units_with_outstanding.groupby('Fund type')
        .agg(
            Total_Outstanding=('Outstanding', 'sum'),
            Number_of_Units_With_Outstanding=('Unit Reference', 'nunique')
        )
        .reset_index()
    )

    # Calculate grand total for all fund types
    grand_total_data = {
        'Name': 'Grand Total',
        'Total_Gross_Demanded': units_with_outstanding['Total_Gross_Demanded'].sum(),
        'Total_Settled': units_with_outstanding['Total_Settled'].sum(),
        'Outstanding': units_with_outstanding['Outstanding'].sum()
    }

    # Save result to a temporary output file
    output_path = input_path.replace(".xlsx", "_analysis_report.xlsx")
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        # Write the fund type summary sheet
        fund_type_summary.to_excel(writer, sheet_name='Fund Type Summary', index=False)
        
        # Get the workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Fund Type Summary']
        
        # Add some formatting to the fund type summary sheet
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D9E1F2',
            'border': 1
        })
        
        # Apply the header format
        for col_num, value in enumerate(fund_type_summary.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Create the Outstanding Units sheet with separate tables for each fund type
        worksheet_units = workbook.add_worksheet('Outstanding Units')
        
        # Define formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D9E1F2',
            'border': 1
        })
        
        total_format = workbook.add_format({
            'bold': True,
            'bg_color': '#E2EFDA',
            'border': 1
        })
        
        grand_total_format = workbook.add_format({
            'bold': True,
            'bg_color': '#FFC7CE',
            'border': 1
        })
        
        normal_format = workbook.add_format({
            'border': 1
        })
        
        # Get unique fund types
        fund_types = units_with_outstanding['Fund type'].unique()
        
        # Column headers
        columns = ['Fund type', 'Unit Reference', 'Name', 'Total_Gross_Demanded', 'Total_Settled', 'Outstanding']
        
        # Start at row 0
        current_row = 0
        
        # For each fund type, create a separate table
        for fund_type in fund_types:
            # Filter data for this fund type
            fund_data = units_with_outstanding[units_with_outstanding['Fund type'] == fund_type]
            
            # Write the fund type header (merged across all columns)
            worksheet_units.merge_range(current_row, 0, current_row, len(columns)-1, f"Fund Type: {fund_type}", title_format)
            current_row += 1
            
            # Write column headers
            for col_num, column in enumerate(columns):
                worksheet_units.write(current_row, col_num, column, header_format)
            current_row += 1
            
            # Write data rows
            for _, row in fund_data.iterrows():
                for col_num, column in enumerate(columns):
                    worksheet_units.write(current_row, col_num, row.get(column, ''), normal_format)
                current_row += 1
            
            # Calculate and write fund type total
            fund_total = {
                'Fund type': '',
                'Unit Reference': '',
                'Name': f'Total for {fund_type}',
                'Total_Gross_Demanded': fund_data['Total_Gross_Demanded'].sum(),
                'Total_Settled': fund_data['Total_Settled'].sum(),
                'Outstanding': fund_data['Outstanding'].sum()
            }
            
            for col_num, column in enumerate(columns):
                worksheet_units.write(current_row, col_num, fund_total.get(column, ''), total_format)
            current_row += 2  # Add an empty row between fund types
        
        # Write grand total at the end
        worksheet_units.merge_range(current_row, 0, current_row, len(columns)-1, "Grand Total", title_format)
        current_row += 1
        
        for col_num, column in enumerate(columns):
            if column in grand_total_data:
                worksheet_units.write(current_row, col_num, grand_total_data.get(column, ''), grand_total_format)
            else:
                worksheet_units.write(current_row, col_num, '', grand_total_format)
        
        # Auto-adjust column widths
        for i, col in enumerate(columns):
            # Find the maximum length in the column
            max_len = max(len(str(units_with_outstanding[col].iloc[i])) if i < len(units_with_outstanding) else 0 
                          for i in range(min(10, len(units_with_outstanding))))  # Check only first 10 rows for performance
            max_len = max(max_len, len(col))  # Consider header length too
            worksheet_units.set_column(i, i, max_len + 2)  # Add some padding
    
    return output_path

# def process_excel(input_path):
#     # Read the 'New Data' sheet
#     df = pd.read_excel(input_path, sheet_name="New Data")

#     # Filter for commercial units
#     df_commercial = df[df['Unit type'] == 'Commercial'].dropna(subset=['Unit Reference', 'Fund type'])

#     # Group by Unit and Fund type
#     unit_summary = (
#         df_commercial.groupby(['Fund type', 'Unit Reference', 'Name.1'], dropna=False)
#         .agg(
#             Total_Gross_Demanded=('Gross Demanded', 'sum'),
#             Total_Settled=('Settled', 'sum')
#         )
#         .reset_index()
#     )

#     unit_summary.rename(columns={'Name.1': 'Name'}, inplace=True)

#     # Calculate outstanding amount
#     unit_summary['Outstanding'] = unit_summary['Total_Gross_Demanded'] - unit_summary['Total_Settled']

#     # Filter for units with outstanding balances
#     units_with_outstanding = unit_summary[unit_summary['Outstanding'] > 0]

#     # Aggregate totals per fund type
#     fund_type_summary = (
#         units_with_outstanding.groupby('Fund type')
#         .agg(
#             Total_Outstanding=('Outstanding', 'sum'),
#             Number_of_Units_With_Outstanding=('Unit Reference', 'nunique')
#         )
#         .reset_index()
#     )

#     # Append total row
#     total_row_data = {
#         'Name': 'Total',
#         'Total_Gross_Demanded': units_with_outstanding['Total_Gross_Demanded'].sum(),
#         'Total_Settled': units_with_outstanding['Total_Settled'].sum(),
#         'Outstanding': units_with_outstanding['Outstanding'].sum()
#     }
#     units_with_outstanding = pd.concat([
#         units_with_outstanding,
#         pd.DataFrame([total_row_data])
#     ], ignore_index=True).fillna('')

#     # Save result to a temporary output file
#     output_path = input_path.replace(".xlsx", "_analysis_report.xlsx")
#     with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
#         fund_type_summary.to_excel(writer, sheet_name='Fund Type Summary', index=False)
#         units_with_outstanding.to_excel(writer, sheet_name='Outstanding Units', index=False)

#     return output_path

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file uploaded', 400
    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, file.filename)
        file.save(input_path)

        try:
            output_path = process_excel(input_path)
            return send_file(output_path, as_attachment=True)
        except Exception as e:
            return f"Error processing file: {str(e)}", 500


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)