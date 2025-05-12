# eBay Terapeak Scraper

A specialized web scraping tool designed to extract sold item data from eBay's Terapeak Research tool for product trend analysis.

> **DISCLAIMER:** This project is shared for **EDUCATIONAL PURPOSES ONLY**. Please read the [Legal Disclaimer](#legal-disclaimer) section before using this code.

## Overview

This application automates the process of searching for items on eBay's Terapeak Research platform and extracts detailed information about sold items, including pricing, shipping costs, sales volume, and more. The scraper exports all the collected data into formatted Excel spreadsheets for easy analysis across different time periods.

## Features

- **Keyword-Based Searching**: Search for items using specific part numbers or keywords from a text file
- **Multithreaded Processing**: Efficiently scrapes data using concurrent processing with a thread pool
- **Multiple Time Period Analysis**: Collects data for both 30-day and 90-day periods for comprehensive trend analysis
- **Comprehensive Data Extraction**: Captures detailed item information including:
  - Item title with clickable links
  - Average sold price
  - Average shipping cost
  - Total quantity sold
  - Total sales value
  - Date last sold
  - Product images
- **Organized Output**: Exports data to well-formatted Excel spreadsheets with separate sheets for different time periods
- **Screenshot Capture**: Takes screenshots of search results for reference
- **Robust Error Handling**: Implements comprehensive exception handling, CAPTCHA detection, and logging
- **Automated eBay Login**: Handles eBay login sessions and CAPTCHA challenges

## How to Use

### Prerequisites

- Windows OS
- Python 3.x (if running from source)
- Python UV (for virtual environment management)
- Chrome browser
- eBay account with access to Terapeak Research

### Running the Application

#### From Source Code

1. Create a virtual environment using UV:

   ```bash
   uv venv .venv
   ```

2. Activate the virtual environment:
   - On Windows:

     ```bash
     .venv\\Scripts\\activate
     ```

3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

4. Run the main script:

   ```bash
   python main.py
   ```

### Setup Instructions

1. When first run, the application will create a `Keywords.txt` file
2. Add your part numbers or keywords to this file, one per line
   - Lines starting with `#` will be ignored (can be used for comments)
   - Example format:

     ```text
     # Enter your keywords below
     iPhone 13 Pro
     Sony WH-1000XM4
     ```

3. Run the application again to start the scraping process
4. Results will be saved to Excel files in the date-stamped output directory

## Output

The scraper generates the following outputs:

1. **Excel Spreadsheets**:
   - Individual spreadsheets for each keyword
   - A combined spreadsheet with all scraped data
   - Each spreadsheet contains sheets for both 30-day and 90-day periods
   - Columns include:
     - Image
     - Keyword
     - Title (with hyperlink)
     - Avg Sold Price
     - Avg Shipping Cost
     - Total Sold
     - Total Sale
     - Last Sold Date

2. **Screenshots**: Captures of search result pages are saved in the "Terapeak Screenshots" folder

3. **Log Files**: Detailed logs are stored in the "logs" folder for troubleshooting

## Project Structure

- `main.py`: Entry point of the application
- `my_libs/`: Contains the core functionality
  - `logging_config.py`: Configures application logging
  - `utils.py`: Utility functions used throughout the application
  - `web_driver.py`: Manages Chrome WebDriver and eBay session handling
  - `xlsxwriter_formats.py`: Excel formatting helpers
  - `terapeak/`: Terapeak-specific scraping modules
    - `terapeak_data_extraction.py`: Core data extraction logic
    - `terapeak_scrape.py`: Main scraping orchestration
    - `terapeak_xlsx_writer.py`: Formats and writes data to Excel

## Troubleshooting

- If the application fails to run, check the log files in the `logs` directory
- Ensure Chrome is installed on your system
- For login issues, the application will open a browser window for manual login when needed
- If the Excel file fails to save, make sure it's not already open in another application

## Legal Considerations

This tool is for personal use only. Please respect eBay's terms of service and use the tool responsibly with appropriate rate limiting to avoid overwhelming their servers.

## Legal Disclaimer

### IMPORTANT: READ BEFORE DOWNLOADING, COPYING, INSTALLING, OR USING

This software project is shared for **EDUCATIONAL PURPOSES ONLY** to demonstrate programming techniques for web automation and data extraction. By using, modifying, or distributing this code, you acknowledge and agree to the following:

1. **Terms of Service Compliance**: Most websites, including eBay, have Terms of Service that may prohibit automated data collection. Using this code to scrape websites may violate these terms.

2. **Personal Responsibility**: You are solely responsible for how you use this code. The author(s) of this project cannot be held liable for any misuse or legal consequences resulting from your use of this code.

3. **Rate Limiting**: If you choose to use this code, implement appropriate rate limiting to avoid overloading target websites' servers.

4. **Alternative API Usage**: Where available, consider using official APIs instead of web scraping.

5. **No Warranty**: This software is provided "AS IS" without warranty of any kind, express or implied.

6. **No Legal Advice**: This disclaimer is not legal advice. Consult with a legal professional if you have questions about the legality of web scraping in your jurisdiction.

Before using this code for any purpose, ensure you have the right to collect data from your target website, preferably by obtaining explicit permission.

The author(s) of this project disclaim any responsibility for how this code is used and any consequences that may arise from its use.
