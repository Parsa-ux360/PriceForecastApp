# Price Forecasting Application

## Description

This desktop application forecasts product prices based on inflation rates. It allows users to manage product data, perform price predictions, and generate detailed reports including Excel and PDF outputs with charts. The program is designed for easy data handling and insightful financial forecasting.

## Features

* Price prediction based on inflation data
* Add, edit, and delete products
* Save and load product data from Excel files
* Display logs and textual reports within the app
* Export results and charts as PDF and Excel files
* User-friendly graphical interface for desktop use

## Installation

### Prerequisites

* Python 3.8 or higher
* Windows operating system (desktop environment)

### Install required Python libraries

```bash
pip install -r requirements.txt
```

### Run the program

```bash
python main.py
```

Or use the precompiled executable `PriceForecastApp.exe` (Windows only).

## Usage

* Use the interface to add new products and input their data
* Save your data into Excel files for future use
* Load existing Excel data into the application
* Run the price forecasting function to generate updated price predictions
* View logs and reports within the app for detailed information
* Export results to Excel and PDF files with charts for presentations or analysis

## Modules

* **CreateExcel.py**: Creates initial Excel files containing product data and inflation information
* **ReadExcel.py**: Reads Excel files and converts the data into JSON format for processing
* **Calculate.py**: Performs price forecasting calculations and generates output files and reports

## Configuration

* The program expects the Excel and icon files to be located in the same directory as the executable or script
* Adjust paths inside the configuration variables if needed before packaging or deployment

## Limitations

* The program may not handle very large datasets efficiently (e.g., over 1000 months of inflation data)
* PDF reports may not correctly render Persian or Arabic characters due to font compatibility issues
* The application is designed specifically for desktop environments and is not optimized for mobile devices

## Contributing

Contributions and improvements are welcome! Feel free to fork the repository on GitHub, submit issues or pull requests. Your feedback helps enhance the project.

## Screenshots

### Application Main Interface  
![App Main Interface](images/app_screenshot1.png)  
*The main window where you load data, run forecasts, and view logs.*

### PDF Report Preview  
![PDF Report View](images/app_report_pdf.png)  
*Sample of the generated detailed PDF report.*

### Excel Output Sample  
![Excel Output View](images/app_excel_output.png)  
*Excel file showing predicted prices and inflation calculations.*

### Application Logo  
![App Logo](images/app_logo.png)  
*Brand identity of PriceForecastApp.*

## Download Executable

You can download the Windows executable file (`PriceForecastApp.exe`) here:

[Download PriceForecastApp.exe](https://drive.google.com/file/d/1ZxJQipgcM0ip01E3A6D-Nwv67gGJBPXl/view?usp=drive_link)

**Note:** This application runs on Windows OS only.

### How to Use

1. Download the executable from the link above.
2. Run `PriceForecastApp.exe` on your Windows machine.
3. Follow the on-screen instructions to load your data and generate forecasts.

## Contact

**Parsa Shahi**

📧 Email: [parsashahi404@gmail.com](mailto:parsashahi404@gmail.com)  
💼 LinkedIn: [Parsa Shahi](https://www.linkedin.com/in/parsa-shahi-266b2a372)  
🐙 GitHub: [Parsa-ux360](https://github.com/Parsa-ux360)
