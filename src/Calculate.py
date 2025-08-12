# ==============================================================================
# Module: Calculate.py
# Author: Parsa Shahi
# Date: 2025-08-11
# Description:
#     This module provides functions to fetch inflation data from the World Bank API,
#     parse and forecast product prices over time considering inflation, and generate
#     output reports in Excel and PDF formats with graphical price projections.
#     
#     Main functionalities include:
#       - Fetching latest annual inflation rates by country code with caching
#       - Parsing price strings to extract numeric value and currency symbol
#       - Projecting price changes month-by-month with compounding inflation
#       - Forecasting future prices over a given number of months
#       - Generating an Excel file summarizing forecasts
#       - Creating a PDF report with textual summaries and price projection charts
#
# Usage:
#     Call main(product_list, output_excel, output_pdf) where product_list is a list
#     of dictionaries containing product info and forecast parameters.
# ==============================================================================

import requests
import re
import math
import json
import os
from datetime import datetime
from openpyxl import Workbook
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
import matplotlib.pyplot as plt

# In-memory cache for storing inflation data keyed by country code to avoid repeated API calls
_inflation_cache = {}

# Mapping common currency codes to their respective symbols (can be extended as needed)
CURRENCY_SYMBOLS = {
    "USD": "$",
    "EUR": "€",
    "GBP": "£",
    "JPY": "¥",
    "TRY": "₺",
    "IRR": "﷼",
    "SEK": "kr",
    "" : ""
}

# Default output file names
OUTPUT_EXCEL_FILE = "forecast_result.xlsx"
REPORT_PDF_FILE = "forecast_report.pdf"


def get_inflation_from_worldbank(country_code):
    """
    Fetch the latest available annual inflation rate (%) for a given country code from the World Bank API.

    Args:
        country_code (str): ISO country code (e.g., 'US', 'IR', 'TR').

    Returns:
        tuple: (inflation_rate_percent (float), year (str)) or (None, None) if data is unavailable.
    
    Uses caching to minimize repeated network calls and includes error handling for network or data issues.
    """
    if not country_code:
        return None, None

    code = country_code.strip().lower()
    if code in _inflation_cache:
        return _inflation_cache[code]

    try:
        url = f"https://api.worldbank.org/v2/country/{code}/indicator/FP.CPI.TOTL.ZG?format=json&per_page=100"
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        if not isinstance(data, list) or len(data) < 2:
            _inflation_cache[code] = (None, None)
            return None, None

        # Iterate to find the latest non-null inflation value
        latest = None
        for entry in data[1]:
            if entry.get("value") is not None:
                latest = (entry["value"], entry["date"])  # (inflation rate, year)
                break

        if latest:
            _inflation_cache[code] = latest
            return latest

    except requests.RequestException:
        # Network or HTTP error: cache None and return
        _inflation_cache[code] = (None, None)
        return None, None
    except Exception:
        # Any unexpected error: cache None and return
        _inflation_cache[code] = (None, None)
        return None, None

    _inflation_cache[code] = (None, None)
    return None, None


# Regular expression pattern to parse price strings with optional currency symbols and codes
_price_pattern = re.compile(
    r"([A-Za-z]{3})?\s*([\$¥£€₺﷼]?)([-+\d,.]+)\s*([A-Za-z]{3})?",
    re.UNICODE
)


def _parse_price(price_str, default_currency="USD"):
    """
    Parse a price string to extract the numeric value and currency symbol/code.

    Args:
        price_str (str): Input price string (e.g. "USD 123.45", "$123", "123.45 EUR").
        default_currency (str): Default currency code if none found in string.

    Returns:
        tuple: (price_value (float), currency_symbol (str), currency_code (str))
               or (None, None, None) if parsing fails.
    """
    if price_str is None:
        return None, None, None
    s = str(price_str).strip()
    s = s.replace('\u00A0', ' ')  # Replace non-breaking spaces if any

    m = _price_pattern.search(s)
    if not m:
        # Fallback: remove non-numeric characters except dot and minus sign
        cleaned = re.sub(r"[^0-9.-]", "", s)
        try:
            return float(cleaned), CURRENCY_SYMBOLS.get(default_currency, ""), default_currency
        except Exception:
            return None, None, None

    gcode1, gsym, gnum, gcode2 = m.groups()
    currency_code = gcode1 or gcode2 or default_currency
    num = gnum.replace(',', '')  # Remove thousand separators

    try:
        number = float(num)
    except Exception:
        try:
            number = float(num.replace(' ', ''))
        except Exception:
            return None, None, None

    symbol = CURRENCY_SYMBOLS.get(currency_code.upper(), gsym or CURRENCY_SYMBOLS.get(default_currency, ''))
    return number, symbol, currency_code.upper()


def _project_price_over_months(start_price, annual_inflation_pct, months):
    """
    Compute projected prices month-by-month with monthly compounding inflation.

    Args:
        start_price (float): Initial price.
        annual_inflation_pct (float): Annual inflation rate in percent.
        months (int): Number of months to project.

    Returns:
        list: Projected prices for months 0 to `months` inclusive, rounded to 2 decimals.
        None if inflation data is not provided.
    """
    if annual_inflation_pct is None:
        return None

    monthly_rate = (1 + annual_inflation_pct / 100.0) ** (1 / 12.0) - 1
    prices = [round(start_price * ((1 + monthly_rate) ** m), 2) for m in range(months + 1)]
    return prices


def forecast_price(current_price, inflation_rate, months, default_currency='USD'):
    """
    Forecast future price after a given number of months based on current price and inflation rate.

    Args:
        current_price (str or float): Current price with optional currency symbol.
        inflation_rate (float): Annual inflation rate in percent.
        months (int or str): Number of months to forecast ahead.
        default_currency (str): Currency code to use if not specified in price.

    Returns:
        tuple: (final_forecasted_price (float), price_series (list of floats), currency_symbol (str))
               or (None, None, None) if input price is invalid.
    """
    number, symbol, code = _parse_price(current_price, default_currency)
    if number is None:
        return None, None, None

    try:
        months = int(months)
    except Exception:
        months = 0

    prices = _project_price_over_months(number, inflation_rate, months)
    if prices is None:
        return None, None, None

    final_price = prices[-1]
    return final_price, prices, (symbol or CURRENCY_SYMBOLS.get(code, ''))


def main(product_list, output_excel=OUTPUT_EXCEL_FILE, output_pdf=REPORT_PDF_FILE):
    """
    Process a list of products with price forecasts, save results to Excel and generate PDF report.

    Args:
        product_list (list of dict): Each dict must contain keys like:
            - "Product": Product name (str)
            - "Current Price": Current price string or number
            - "Forecast Months": Number of months to forecast (int or str)
            - "Country": Country code for inflation lookup (str)
            - "Currency": Currency code (str)
        output_excel (str): Path to save Excel output.
        output_pdf (str): Path to save PDF report.

    Returns:
        tuple: (success (bool), message (str))
    """
    forecasts = []
    items = list(product_list)  # Defensive copy

    for item in items:
        country_code = (item.get("Country", "") or "").strip()
        currency = (item.get("Currency", "") or "").upper() or 'USD'
        inflation, year = get_inflation_from_worldbank(country_code)

        if inflation is None:
            forecasts.append({
                "Product": item.get("Product", ""),
                "Current Price": item.get("Current Price", "N/A"),
                "Forecast Months": item.get("Forecast Months", "0"),
                "Country": country_code,
                "Currency": currency,
                "Inflation Rate (Year)": None,
                "Inflation Year": None,
                "Forecasted Price": "Inflation data not found",
                "Price Series": None
            })
            continue

        price_result = forecast_price(item.get("Current Price", "0"), inflation, int(item.get("Forecast Months", 0)), default_currency=currency)
        if price_result[0] is None:
            forecasted_price = "Invalid price format"
            price_series = None
            symbol = CURRENCY_SYMBOLS.get(currency, '')
        else:
            final_price, price_series, symbol = price_result
            forecasted_price = f"{symbol}{final_price}" if symbol else f"{final_price} {currency}"

        forecasts.append({
            "Product": item.get("Product", ""),
            "Current Price": item.get("Current Price", "N/A"),
            "Forecast Months": item.get("Forecast Months", "0"),
            "Country": country_code,
            "Currency": currency,
            "Inflation Rate (Year)": inflation,
            "Inflation Year": year,
            "Forecasted Price": forecasted_price,
            "Price Series": price_series
        })

    # Save forecast results to Excel
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Forecast Results"

        headers = ["Product", "Current Price", "Forecast Months", "Country", "Currency",
                   "Inflation Rate (Year)", "Inflation Year", "Forecasted Price"]
        ws.append(headers)

        for f in forecasts:
            ws.append([
                f.get("Product"),
                f.get("Current Price"),
                f.get("Forecast Months"),
                f.get("Country"),
                f.get("Currency"),
                f.get("Inflation Rate (Year)", "N/A"),
                f.get("Inflation Year", "N/A"),
                f.get("Forecasted Price")
            ])

        wb.save(output_excel)
    except Exception as e:
        return False, f"Error saving excel: {e}"

    # Build PDF report including price projection charts for each product
    try:
        doc = SimpleDocTemplate(output_pdf, pagesize=A4)
        styles = getSampleStyleSheet()
        flowables = []

        flowables.append(Paragraph("Forecast Report", styles["Title"]))
        flowables.append(Spacer(1, 12))

        for idx, fcast in enumerate(forecasts, 1):
            flowables.append(Paragraph(f"{idx}. Product: {fcast.get('Product', '')}", styles["Heading3"]))

            if fcast.get("Forecasted Price") in ["Inflation data not found", "Invalid price format"]:
                flowables.append(Paragraph(f"Error: {fcast.get('Forecasted Price')}", styles["Normal"]))
            else:
                flowables.append(Paragraph(f"Country: {fcast.get('Country', '')}", styles["Normal"]))
                flowables.append(Paragraph(f"Inflation Rate (Year {fcast.get('Inflation Year', 'N/A')}): "
                                           f"{fcast.get('Inflation Rate (Year)', 'N/A')}%", styles["Normal"]))
                flowables.append(Paragraph(f"Current Price: {fcast.get('Current Price', '')}", styles["Normal"]))
                flowables.append(Paragraph(f"Forecast Months: {fcast.get('Forecast Months', '')}", styles["Normal"]))
                flowables.append(Paragraph(f"Forecasted Price: {fcast.get('Forecasted Price')}", styles["Normal"]))

                # Add price projection chart if available
                series = fcast.get('Price Series')
                if series:
                    months = list(range(len(series)))
                    plt.figure(figsize=(6, 3))
                    plt.plot(months, series)
                    plt.title(f"Price projection: {fcast.get('Product')}")
                    plt.xlabel('Months')
                    plt.ylabel(f"Price ({fcast.get('Currency')})")
                    plt.tight_layout()

                    img_path = f"_tmp_chart_{idx}.png"
                    plt.savefig(img_path, dpi=150)
                    plt.close()

                    # Insert the chart image into PDF document
                    flowables.append(Spacer(1, 6))
                    flowables.append(Image(img_path, width=450, height=200))
                    flowables.append(Spacer(1, 6))

        doc.build(flowables)

        # Clean up temporary chart image files
        for idx in range(1, len(forecasts) + 1):
            p = f"_tmp_chart_{idx}.png"
            if os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass

    except Exception as e:
        return False, f"Error building PDF: {e}"

    return True, f"Forecast completed. Output saved to {output_excel} and {output_pdf}"
