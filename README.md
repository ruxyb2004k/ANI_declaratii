# Automated Romanian Asset Declarations Downloader

This project is a Python automated data downloader for the Romanian asset declarations website (https://declaratii.integritate.eu/). It allows you to search for people and download their asset declarations.

## Features

- Automated web scraping using Selenium and undetected-chromedriver
- Handles Cloudflare protection
- Downloads PDF declarations
- Processes multiple names from Excel file
- Saves all data to Excel file
- Includes error handling and logging

## Requirements

- Python 3.8+ (in my installation 3.9.13)
- Chrome browser
- Required Python packages (see requirements.txt)

## Installation

1. Clone the repository:
```bash
git clone <your-repository-url>
cd <repository-name>
```

2. Create and activate a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Prepare your Excel file with names in the "Nume" column
2. Run the script with your Excel file as an argument:
```bash
python scraper.py your_excel_file.xlsx
```

For example:
```bash
python scraper.py "Baza de date - Cautare ANI.xlsx"
```

The script will:
- Read names from the specified Excel file
- Search for each person's declarations
- Download available PDF files
- Save all data to an Excel file

## Output

- Downloaded PDF files are saved in the `downloads` directory
- All scraped data is saved to `all_declarations_data.xlsx`
- Error screenshots are saved if any issues occur

## Notes
- The script requires Chrome browser to be installed
- Although Cloudflare is most of the time handled automatically, you might sometime need to it manually
- The script includes random delays to avoid being blocked
- Files are renamed with meaningful names based on declaration data
- If multiple declarations would have the same filename, numbers are added to make them unique