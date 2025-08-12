# ==============================================================================
# Module: ReadExcel.py
# Author: Parsa Shahi
# Date: 2025-08-10
# Description:
#     This module contains the function to read product pricing data from an Excel file,
#     normalize the data structure, and export it as a JSON file.
#     It also supports optional deletion of the source Excel file after processing.
# ==============================================================================

import openpyxl
import json
import os

def read_excel_to_json_gui(file_path, delete_excel=False):
    """
    Reads data from an Excel file and converts it into a normalized JSON format.

    Parameters:
    -----------
    file_path : str
        The full path to the input Excel file.
    delete_excel : bool, optional (default=False)
        Flag indicating whether to delete the source Excel file after successful conversion.

    Returns:
    --------
    tuple (bool, str)
        A tuple where the first element indicates success status,
        and the second element contains a success message or error details.

    Raises:
    -------
    Exception:
        Captures any exception that occurs during file reading or writing,
        returning an error message without stopping program execution.
    """

    try:
        # Load the Excel workbook and activate the default worksheet
        workbook = openpyxl.load_workbook(file_path)
        worksheet = workbook.active

        # Initialize container for structured data
        data = []

        # Extract header row to use as dictionary keys
        headers = [cell.value for cell in worksheet[1]]

        # Iterate over data rows, starting from the second row (to skip headers)
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            # Map headers to corresponding cell values in current row
            product = dict(zip(headers, row))

            # Normalize the data keys and provide default fallback values if necessary
            normalized = {
                "Product": product.get("Product", "") if isinstance(product, dict) else product[0] if product else "",
                "Current Price": product.get("Current Price", "") if isinstance(product, dict) else (product[1] if len(product) > 1 else ""),
                "Forecast Months": product.get("Forecast Months", "0") if isinstance(product, dict) else (product[2] if len(product) > 2 else "0"),
                "Country": product.get("Country", "") if isinstance(product, dict) else (product[3] if len(product) > 3 else ""),
                "Currency": product.get("Currency", "") if isinstance(product, dict) else (product[4] if len(product) > 4 else "")
            }

            # Append normalized entry to data list
            data.append(normalized)

        # Write the normalized data list to a JSON file with pretty formatting and UTF-8 encoding
        with open("data.json", "w", encoding="utf-8") as json_file:
            json.dump(data, json_file, indent=2, ensure_ascii=False)

        # Optionally delete the original Excel file after successful JSON creation
        if delete_excel:
            os.remove(file_path)

        return True, "Excel file converted to JSON successfully."

    except Exception as e:
        # Return failure status and error message without raising exception
        return False, f"Error reading Excel file: {e}"
