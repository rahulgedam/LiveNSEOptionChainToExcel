
---

# Live NSE Option Chain to Excel

This project retrieves live option chain data from the National Stock Exchange of India (NSE) and stores it in an Excel file. It updates the file at regular intervals with new data, making it easier for traders and analysts to track option chain metrics in real-time.

## Features

- Fetches option chain data for specified indices from the NSE.
- Processes and stores data in an Excel file using `xlwings`.
- Updates the Excel file with new data at regular intervals.
- Calculates At-The-Money (ATM) strike prices based on the last traded price.
- Reads configuration from a JSON file for option and futures expiry dates.

## Requirements

- Python 3.x
- `requests` - For HTTP requests
- `xlwings` - For interacting with Excel
- `json` - For handling JSON data
- `math` - For mathematical operations

You can install the required Python packages using pip:

```sh
pip install requests xlwings
```

## Setup

1. **Clone the Repository**

   ```sh
   git clone https://github.com/rahulgedam/LiveNSEOptionChainToExcel.git
   cd LiveNSEOptionChainToExcel
   ```

2. **Prepare Configuration Files**

   Ensure you have a `config.json` file in the project directory. This file should include expiry dates for options and futures. 

   Example `config.json`:

   ```json
   {
     "data": [
       {
         "expiryDateOption": "2024-07-31",
         "expiryDateFuture": "2024-08-31"
       }
     ]
   }
   ```

3. **Run the Script**

   Execute the script to start fetching and updating option chain data:

   ```sh
   python script_name.py
   ```

   Replace `script_name.py` with the actual name of your script file.

## Script Overview

- **`addOptionDataInSheet(wb, currTime, strikePrice, ...)`**: Adds option data to the specified Excel sheet.
- **`createOptionChain(wb, optionData, ATM, lotSize, instrName, optExpDate, FUT_SPOT)`**: Creates the option chain data in the Excel file.
- **`makeOptionChainFile(wb, opt, currTime, ATM, instrName, optExpDate, FUT_SPOT)`**: Initializes the Excel sheet and calls `createOptionChain`.
- **`putOptionChainData(opt, currTime, ATM, instrName, optExpDate, FUT_SPOT)`**: Saves the option chain data into the Excel file.
- **`findATM(lastTraded, instrName)`**: Calculates the ATM strike price based on the last traded price.
- **`getOptionChain(currTime, isFifteenMin, instrName, wb)`**: Fetches option chain data and updates the Excel file.
- **`createExcelSheet(instrName)`**: Creates a new Excel file for storing option chain data.
- **`createSheetsInExcel(optionExpDate, wb)`**: Creates sheets in the Excel file for each expiry date.
- **`downloadFuturesJson(instrName, futExpDate)`**: Downloads and parses futures data.
- **`readFuturesJson(instrName, futExpDate)`**: Reads futures data from a JSON file.
- **`readConfigJson()`**: Reads configuration data from `config.json`.
- **`createAndInitFiles(currTime, instrName)`**: Initializes the Excel file and updates it with initial data.

## Usage

- The script runs continuously from market opening to closing time, updating the Excel file with option chain data every 5 minutes.
- Modify the `while` loop in the script to adjust the data fetching frequency or add specific time conditions if needed.

## Notes

- Ensure the Excel file is not open while the script is running to avoid file access issues.
- This script is designed for educational purposes and may need adjustments for production use.

## License

Public
