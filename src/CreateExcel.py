# ==============================================================================
# Module: CreateExcel.py
# Author: Parsa Shahi
# Date: 2025-08-10
# Description:
#     This module provides functionality to save a list of product data into an Excel file.
#     It also creates a JSON backup of the product list for data persistence.
# ==============================================================================

import openpyxl
import json
import os
from datetime import datetime


def save_to_excel(product_list, file_path="products.xlsx"):
    """
    Save a list of product dictionaries to an Excel file with a header row.
    Additionally, create a JSON backup file alongside the Excel file for app state persistence.

    Args:
        product_list (list of dict): List of product data dictionaries, each containing keys:
            "Product", "Current Price", "Forecast Months", "Country", "Currency".
        file_path (str, optional): Path to save the Excel file. Defaults to "products.xlsx".

    Returns:
        tuple: (success (bool), message (str))
            success: True if file saved successfully, False otherwise.
            message: Description of the result or error.
    """
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Products"

        # Append header row for Excel sheet
        headers = ["Product", "Current Price", "Forecast Months", "Country", "Currency"]
        ws.append(headers)

        # Append product data rows
        for product in product_list:
            ws.append([
                product.get("Product", ""),
                product.get("Current Price", ""),
                product.get("Forecast Months", ""),
                product.get("Country", ""),
                product.get("Currency", "")
            ])

        # Save the Excel workbook to disk
        wb.save(file_path)

        # Create a JSON backup for app state persistence
        json_path = os.path.splitext(file_path)[0] + "_backup.json"
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(product_list, f, ensure_ascii=False, indent=2)

        return True, f"Excel file saved successfully to {file_path} (backup: {json_path})"

    except Exception as e:
        return False, f"Error saving Excel file: {e}"
