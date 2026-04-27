[![en](https://img.shields.io/badge/中文-Chinese-red.svg)](https://github.com/jonatasemidio/multilanguage-readme-pattern/blob/master/README.md)

# selenium_sncs

## Features

1. **Automated ChromeDriver management**: Automatically downloads and configures the correct ChromeDriver version based on the installed Google Chrome browser version.
2. **SNCS web automation**: Automatically logs in to the SNCS platform, selects a specific lot number, and downloads the required CSV files.
3. **Excel report generation**: Processes the downloaded CSV files into an integrated Excel report with formatted tables, charts, and statistical data, including mean and standard deviation.
4. **Real-time progress display**: Displays the current step during the download and processing workflow using progress bars.
5. **Advanced Excel formatting**: Applies custom cell formatting, borders, and conditional formatting to improve report readability.

## Required Packages

- selenium
- pandas
- openpyxl
- tqdm
- tkinter

Install all dependencies with the following command:

```bash
pip install -r requirements.txt
```

## Usage

### I. `sncs.py`

#### Step 1: ChromeDriver setup

Make sure Google Chrome is installed. The script will automatically download and configure ChromeDriver according to your installed browser version.

#### Step 2: Configure environment variables

Create a `.env` file in the same directory as the script. The file should contain the login account and password for [XQC](https://sncs-web.com/quality/login).

The `.env` variable names are as follows:

```env
SNCS_ACCOUNT='ur_account'
SNCS_PASSWORD='ur_password'
```

#### Step 3: Run the script

To start SNCS data processing, run the following command:

```bash
python sncs.py
```

This script performs the following actions:

1. Opens a GUI for selecting the download folder.
2. Logs in to the SNCS system using the configured account.
3. Downloads the selected CSV files.
4. Generates an Excel file named `merge.xlsx` from the downloaded data, including detailed formatting and statistical information.

#### Step 4: Excel report output

After the workflow is completed, the Excel report will be saved in the folder you selected.

### II. `sncs_lot.py`

#### Step 1: ChromeDriver setup

Make sure Google Chrome is installed. The script will automatically download and configure ChromeDriver according to your installed browser version.

#### Step 2: Configure environment variables

Create a `.env` file in the same directory as the script. The file should contain the login account and password for [Sysmex Academy](https://academy.sysmex.com.tw/user/login).

The `.env` variable names are as follows:

```env
LOT_SNCS_ACCOUNT='ur_account'
LOT_SNCS_PASSWORD='ur_password'
```

#### Step 3: Run the script

To start downloading SNCS lot data, run the following command:

```bash
python sncs_lot.py
```

This script performs the following actions:

1. Opens a GUI for selecting the download folder.
2. Prompts the user in the terminal to enter the first four characters of the lot number.
3. Logs in to the Academy system using the configured account.
4. Searches using the first four characters entered by the user.
5. Displays the search results in the terminal and asks the user to select which document to download.
6. Downloads the corresponding XQN file according to the document number entered by the user, then extracts the downloaded file.

## Overview

- `sncs.py`: The main script responsible for SNCS automation, CSV file processing, and Excel report generation.
- `sncs_lot.py`: Downloads SNCS new-lot configuration files in XQN format.
- `drivertester.py`: Contains functions for automatically downloading, extracting, and managing ChromeDriver versions.

## License

This project is licensed under the MIT License.
