# ThousandEyes Data Collection Script

## Overview
This project provides a Python-based script to collect, process, and export data from ThousandEyes APIs into a single Excel file. It fetches enterprise agents, endpoint agents, scheduled tests, and usage statistics for specified account groups, then formats the data for easy analysis.

## Features
- Retrieves data for multiple account groups from the ThousandEyes API.
- Consolidates data from the following endpoints:
  - **Enterprise Agents**
  - **Endpoint Agents**
  - **Enterprise Tests**
  - **Scheduled Tests**
  - **Labels**
  - **Usage Summary** (Single API call for all accounts)
- Correlates test-to-agent mappings.
- Exports data to an Excel file with multiple sheets.
- Applies formatting and styling to the Excel output.

## Table of Contents
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Excel Output](#excel-output)
- [Logging](#logging)
- [License](#license)

---

## Prerequisites
Ensure you have the following installed on your system:
- Python 3.8 or higher
- A ThousandEyes account with API access enabled

## Installation
1. Clone this repository:
    ```bash
    git clone https://github.com/yourusername/thousandeyes-data-collection.git
    cd thousandeyes-data-collection
    ```
2. Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```
3. Create a `.env` file to store your ThousandEyes API credentials:
    ```env
    TE_API_KEY=your_api_key
    ```

## Usage
1. Prepare an Excel file containing your account groups. Save it as `account_ids.xlsx` in the same directory. The file should have the following columns:
    - `accountGroupName`
    - `aid`

2. Run the script:
    ```bash
    python main.py
    ```

3. The script will:
    - Fetch data for each account group specified in `account_ids.xlsx`.
    - Fetch usage data (using the first account group in the list).
    - Export all consolidated data into a timestamped Excel file.

4. Output example:
    ```
    thousandeyes_data-1672567890.xlsx
    ```

## Configuration
The script behavior can be customized by modifying the following files:
- **`account_ids.xlsx`**: List of account groups to fetch data for.
- **`.env`**: API credentials for ThousandEyes.

## Excel Output
The output Excel file includes the following sheets:

1. **Account Groups**: Input list of account groups.
2. **Agents**: Data on enterprise agents.
3. **Endpoint Agents**: Data on endpoint agents.
4. **Enterprise Tests**: Details about enterprise tests.
5. **Scheduled Test Endpoint Agent**: Scheduled tests and associated endpoint agents.
6. **Test ↔ Agent Assignments**: Mappings of tests to assigned agents.
7. **Labels**: Details about agent labels.
8. **Usage Summary**: Aggregated usage data.
9. **Usage Tests**: Detailed usage data for tests.
10. **Usage Endpoint Agents**: Endpoint agent usage statistics.
11. **Usage Enterprise Agents**: Enterprise agent usage statistics.

## Logging
The script uses structured logging to provide status updates. Key stages of execution (e.g., API calls, data processing, and file creation) are logged to the console. Success messages indicate script completion.
