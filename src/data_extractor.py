import openpyxl
import csv

# raw_input_file = "/Users/omkar/data/delivery-time-prediction/Feb2024.xlsx"
raw_inputs = [("~/data/delivery-time-prediction/Dec2023.xlsx", "Dec2023"), ("~/data/delivery-time-prediction/Jan2024.xlsx", "Jan2024")]

output_tsv_file = "~/data/delivery-time-prediction/dec_jan_trimmed.tsv"

def extract_data_by_columns():
    columns_of_interest = [
        'ITEM_ID',
        'ORDERED_QTY',
        'SHIP_MODE',
        'SCAC_2',
        'CARRIER_SERVICE_CODE_1',
        'CARRIER_SERVICE_CODE_2',
        'ZIP_CODE',
        'ZIP_CODE_1',
        'ORDER_DATE',
        'ACTUAL_SHIPMENT_DATE',
        'SELLER_ORGANIZATION_CODE_1',
        'ITEM_WEIGHT',
        'SHIPNODE_KEY',
        'SHIPNODE_KEY_1'
        ]
    with open(output_tsv_file, 'w', encoding='utf-8') as ofile:
        writer = csv.DictWriter(ofile, fieldnames=columns_of_interest, delimiter='\t')
        writer.writeheader()
        for filepath, sheetname in raw_inputs:
            workbook = openpyxl.load_workbook(filename=filepath)
            dataset = workbook[sheetname]
            columns = [cell.value for cell in dataset[1]]
            for row in dataset.iter_rows(min_row=2, values_only=True):
                row_data = dict(zip(columns, row))
                row_to_write = {col: row_data[col] for col in columns_of_interest}
                writer.writerow(row_to_write)
            workbook.close()
    print('done')



if __name__ == "__main__":
    extract_data_by_columns()