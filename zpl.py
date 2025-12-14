import openpyxl


def xlsx_to_zpl(xlsx_file, zpl_file):
    """
    Convert Excel file with barcode data to ZPL format for Zebra printers.

    Parameters:
    -----------
    xlsx_file : str
        Path to input Excel file containing barcode data
        Expected format: Column A contains barcode data
    zpl_file : str
        Path to output ZPL file

    The function reads the first column (Barcode Data) from the Excel sheet
    and generates ZPL commands for each barcode entry.
    """
    # Load the workbook and select the active sheet
    wb = openpyxl.load_workbook(xlsx_file, data_only=True)
    ws = wb.active

    # Open the ZPL output file
    with open(zpl_file, "w", encoding="utf-8") as f:
        # Iterate through rows in the Excel sheet (skip header row)
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            if row and row[0]:  # Ensure first column has data
                barcode_data = str(row[0]).strip()  # Remove leading/trailing spaces

                # Generate ZPL content for the barcode
                zpl_content = (
                    "^XA\n"
                    "^FO50,50^BY3^BCN,100,Y,N,N\n"
                    f"^FD{barcode_data}^FS\n"
                    "^FO10,200\n"
                    "^FB300,1,0,C,0\n"
                    "^A0N,30,20\n"
                    f"^FD{barcode_data}^FS\n"
                    "^XZ\n"
                )
                f.write(zpl_content)

    print(f"ZPL file '{zpl_file}' has been generated from '{xlsx_file}'.")

# Example usage
xlsx_file = "/home/dataopske/Desktop/FINDSTD BARCODES/ICF_Barcodes.xlsx"  # Input Excel file
zpl_file = "001-350ICF.zpl"  # Output ZPL file


xlsx_to_zpl(xlsx_file, zpl_file)
