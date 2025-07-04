from flask import Flask, request, send_file, render_template
import pandas as pd
import tempfile
import os

app = Flask(__name__)

def process_excel(input_path):
    # Read the 'New Data' sheet
    df = pd.read_excel(input_path, sheet_name="New Data")

    # Filter for commercial units
    df_commercial = df[df['Unit type'] == 'Commercial'].dropna(subset=['Unit Reference', 'Fund type'])

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

    # Append total row
    total_row_data = {
        'Name': 'Total',
        'Total_Gross_Demanded': units_with_outstanding['Total_Gross_Demanded'].sum(),
        'Total_Settled': units_with_outstanding['Total_Settled'].sum(),
        'Outstanding': units_with_outstanding['Outstanding'].sum()
    }
    units_with_outstanding = pd.concat([
        units_with_outstanding,
        pd.DataFrame([total_row_data])
    ], ignore_index=True).fillna('')

    # Save result to a temporary output file
    output_path = input_path.replace(".xlsx", "_analysis_report.xlsx")
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        fund_type_summary.to_excel(writer, sheet_name='Fund Type Summary', index=False)
        units_with_outstanding.to_excel(writer, sheet_name='Outstanding Units', index=False)

    return output_path

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
    app.run(debug=True)