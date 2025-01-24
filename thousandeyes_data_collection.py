import os
import time
import sys
import requests
import pandas as pd
from dotenv import load_dotenv
from loguru import logger

# For coloring cells / applying formats
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ------------------------------------------------------------------------------------
# Configure the logger (Loguru)
# ------------------------------------------------------------------------------------
logger.remove()  # Remove Loguru's default handler
logger.add(
    sys.stderr,
    format=(
        "<green>{time:YYYY-MM-DD HH:mm:ss}</green> | "
        "<level>{level: <8}</level> | "
        "<cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - "
        "<level>{message}</level>"
    ),
    colorize=True,
    level="INFO"
)

# ------------------------------------------------------------------------------------
# Load environment variables from .env
# ------------------------------------------------------------------------------------
load_dotenv()

API_KEY = os.getenv("API_KEY")
BASE_URL = os.getenv("BASE_URL", "https://api.thousandeyes.com/v7")

# Common headers for most calls
HEADERS = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}

# Special headers for the Endpoint Agents call (accepting HAL+JSON)
HEADERS_HAL = HEADERS.copy()
HEADERS_HAL["Accept"] = "application/hal+json"


def _apply_color_fills(excel_file: str, sheet_name: str = "Labels",
                       color_col: str = "color"):
    """
    Open the given Excel file with openpyxl and apply cell background
    colors in the given sheet, based on the hex codes in `color_col`.

    - If color_col has #93249F or 93249F, we fill the cell with that color.
    - Hex must be exactly 6 characters after removing the optional '#' or we skip.
    """
    logger.info(
        f"Applying cell color fills in '{sheet_name}' using column '{color_col}'...")
    wb = load_workbook(excel_file)

    if sheet_name not in wb.sheetnames:
        logger.warning(f"Sheet '{sheet_name}' not found; skipping color fills.")
        wb.close()
        return

    ws = wb[sheet_name]

    # Find the column index for color_col.
    col_index = None
    for cell in ws[1]:  # row 1 => headers
        if cell.value == color_col:
            col_index = cell.column  # 1-based index
            break

    if col_index is None:
        logger.warning(
            f"Column '{color_col}' not found in '{sheet_name}' sheet; skipping color fills.")
        wb.close()
        return

    # Iterate from row 2 downward because row 1 is the header
    for row_idx in range(2, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=col_index)
        color_val = str(cell.value).strip() if cell.value else ""
        if not color_val:
            continue

        # Remove a leading "#" if present
        if color_val.startswith("#"):
            color_val = color_val[1:]

        # Must be exactly 6 hex digits (case-insensitive)
        color_val = color_val.upper()
        if len(color_val) == 6:
            fill = PatternFill(start_color=color_val, end_color=color_val,
                               fill_type="solid")
            cell.fill = fill
        else:
            logger.debug(f"Skipping invalid color code '{cell.value}' in row {row_idx}.")

    wb.save(excel_file)
    wb.close()
    logger.info(f"Cell coloring applied; file re-saved as '{excel_file}'.")


def _auto_format_numbers(excel_file: str):
    """
    After writing the DataFrame to Excel, open it with openpyxl and
    detect columns that are entirely numeric. For those columns,
    set 'number_format' to 'General' (or '0', '#,##0.00', etc.).

    This ensures Excel interprets them as numbers rather than text.

    Rules:
    - We skip the header row (row 1).
    - If every non-empty cell in a column can be cast to float, we mark that column numeric.
    - Then we apply 'number_format' to each cell in that column from row 2 downward.
    """
    logger.info("Auto-formatting numeric columns in all sheets...")
    wb = load_workbook(excel_file)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        logger.info(f"  Checking numeric columns in sheet '{sheet_name}'...")

        # We'll find how many columns we have by checking row 1
        max_col = ws.max_column
        # row 1 is the header
        for col_idx in range(1, max_col + 1):
            is_numeric = True
            for row_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                val = cell.value
                if val is None or val == "":
                    # skip empty
                    continue
                # If it's already a number in Python, that's good
                if isinstance(val, (int, float)):
                    continue
                # If it's a string, attempt to parse float
                if isinstance(val, str):
                    try:
                        float(val.replace(",", ""))  # remove comma if present
                    except ValueError:
                        is_numeric = False
                        break
                else:
                    # Not int, not float, not string => not numeric
                    is_numeric = False
                    break

            # If still numeric, apply a number format
            if is_numeric:
                # We'll do it from row 2 to last row
                for row_idx in range(2, ws.max_row + 1):
                    num_cell = ws.cell(row=row_idx, column=col_idx)
                    # Set a number format
                    # Could use '0', '#,##0', '#,##0.00', etc.
                    num_cell.number_format = "General"

    wb.save(excel_file)
    wb.close()
    logger.info(f"Numeric columns formatted in '{excel_file}'.")


def get_account_ids(file_path: str) -> pd.DataFrame:
    """
    Reads an Excel file containing account IDs in one tab.
    Expected columns:
        - accountGroupName
        - aid
    Returns a DataFrame with columns: [accountGroupName, aid]
    """
    logger.info(f"Reading account IDs from '{file_path}'...")
    df = pd.read_excel(file_path, sheet_name=0)  # read the first sheet
    # Ensure required columns are present
    required_cols = {"accountGroupName", "aid"}
    if not required_cols.issubset(df.columns):
        logger.error("Excel file must contain 'accountGroupName' and 'aid' columns.")
        raise ValueError("Excel must contain 'accountGroupName' and 'aid' columns.")
    logger.success("Account IDs read successfully.")
    return df[["accountGroupName", "aid"]]


def fetch_agents(aid: str) -> pd.DataFrame:
    """
    Fetch Enterprise Agents for a given Account ID.
    Endpoint: /agents?aid=XXXXXXX
    """
    url = f"{BASE_URL}/agents?aid={aid}"
    logger.info(f"Fetching Agents for AID={aid} from '{url}'...")
    resp = requests.get(url, headers=HEADERS)
    resp.raise_for_status()

    data = resp.json().get("agents", [])
    logger.info(f"Received {len(data)} Agents for AID={aid}.")
    records = []
    for agent in data:
        records.append({
            "OrgId": agent.get("orgId"),
            "agentId": agent.get("agentId"),
            "agentName": agent.get("agentName"),
            "agentType": agent.get("agentType"),
            "agentState": agent.get("agentState"),
            "lastSeen": agent.get("lastSeen"),
            "createdDate": agent.get("createdDate"),
            "utilization": agent.get("utilization"),
            "location": agent.get("location"),
            "enabled": agent.get("enabled"),
            "hostname": agent.get("hostname"),
            "ipAddresses": ", ".join(agent.get("ipAddresses", []))
            if isinstance(agent.get("ipAddresses"), list) else ""
        })
    return pd.DataFrame(records)


def fetch_endpoint_agents(aid: str) -> pd.DataFrame:
    """
    Fetch the list of Endpoint Agents for a given Account ID (using the "agents" array),
    following pagination links (the "next" link) until all agents are retrieved.
    Endpoint: /endpoint/agents?aid=XXXXXXX

    Returns columns:
      [id, name, computerName, osVersion, platform, lastSeen, status,
       deleted, version, createdAt, numberOfClients, locationName, agentType, licenseType]
    """
    next_url = f"{BASE_URL}/endpoint/agents?aid={aid}"
    logger.info(
        f"Fetching Endpoint Agents for AID={aid} from '{next_url}' using HAL+JSON headers...")

    all_records = []
    page_count = 1

    while next_url:
        resp = requests.get(next_url, headers=HEADERS_HAL)
        resp.raise_for_status()

        data = resp.json()
        agents_list = data.get("agents", [])

        links = data.get("_links", {})
        next_link = links.get("next", {}).get("href")

        logger.info(
            f"Page {page_count}: Received {len(agents_list)} Endpoint Agents. "
            f"{'Continuing...' if next_link else 'No more pages.'}"
        )

        for agent in agents_list:
            location_obj = agent.get("location", {}) or {}
            location_name = location_obj.get("locationName", "")

            all_records.append({
                "id": agent.get("id"),
                "name": agent.get("name"),
                "computerName": agent.get("computerName"),
                "osVersion": agent.get("osVersion"),
                "platform": agent.get("platform"),
                "lastSeen": agent.get("lastSeen"),
                "status": agent.get("status"),
                "deleted": agent.get("deleted"),
                "version": agent.get("version"),
                "createdAt": agent.get("createdAt"),
                "numberOfClients": agent.get("numberOfClients"),
                "locationName": location_name,
                "agentType": agent.get("agentType"),
                "licenseType": agent.get("licenseType"),
            })

        next_url = next_link
        page_count += 1

    return pd.DataFrame(all_records)


def fetch_enterprise_tests(aid: str) -> pd.DataFrame:
    """
    Fetch the list of Enterprise Tests for a given Account ID.
    Endpoint: /tests?aid=XXXXXXX
    """
    url = f"{BASE_URL}/tests?aid={aid}"
    logger.info(f"Fetching Enterprise Tests for AID={aid} from '{url}'...")
    resp = requests.get(url, headers=HEADERS)
    resp.raise_for_status()

    data = resp.json().get("tests", [])
    logger.info(f"Received {len(data)} Enterprise Tests for AID={aid}.")
    records = []
    for test in data:
        records.append({
            "testID": test.get("testId"),
            "testName": test.get("testName"),
            "createdBy": test.get("createdBy"),
            "createdDate": test.get("createdDate"),
            "modifiedBy": test.get("modifiedBy"),
            "modifiedDate": test.get("modifiedDate"),
            "type": test.get("type"),
            "alertsEnabled": test.get("alertsEnabled"),
            "enabled": test.get("enabled"),
            "direction": test.get("direction"),
            "targetAgentID": test.get("targetAgentId")
        })
    return pd.DataFrame(records)


def fetch_scheduled_tests(aid: str) -> pd.DataFrame:
    """
    Fetch the scheduled endpoint tests for a given Account ID.
    Endpoint: /endpoint/tests/scheduled-tests?aid=XXXXXXX

    NOTE: Based on the raw JSON you provided, the top-level key is "tests",
    and the boolean flag is "isEnabled" rather than "enabled".
    """
    url = f"{BASE_URL}/endpoint/tests/scheduled-tests?aid={aid}"
    logger.info(f"Fetching Scheduled Tests for AID={aid} from '{url}'...")
    resp = requests.get(url, headers=HEADERS)
    resp.raise_for_status()

    data = resp.json().get("tests", [])
    logger.info(f"Received {len(data)} Scheduled Tests for AID={aid}.")

    records = []
    for test in data:
        records.append({
            "testID": test.get("testId"),
            "testName": test.get("testName"),
            "server": test.get("server"),
            "createdDate": test.get("createdDate"),
            "type": test.get("type"),
            "isEnabled": test.get("isEnabled"),
        })

    return pd.DataFrame(records)


def fetch_scheduled_test_assigned_agents(aid: str, test_id: str) -> pd.DataFrame:
    """
    Fetch the assigned agents for a specific scheduled test ID, handling pagination.
    Endpoint:
      GET /endpoint/test-results/scheduled-tests/<testID>/http-server?aid=<AID>
    """
    url = f"{BASE_URL}/endpoint/test-results/scheduled-tests/{test_id}/http-server"
    logger.info(
        f"Fetching Assigned Agents for Scheduled Test (testID={test_id}, AID={aid}) from '{url}'..."
    )

    records = []
    params = {"aid": aid}

    while True:
        resp = requests.get(url, headers=HEADERS, params=params)
        # Clear params for subsequent pages
        params = None

        if resp.status_code != 200:
            logger.warning(
                f"Could not fetch assigned agents for testID={test_id} (status code {resp.status_code}). "
                "Returning the data fetched so far (if any)."
            )
            break

        data = resp.json()
        page_results = data.get("results", [])
        logger.info(
            f"Received {len(page_results)} records of Assigned Agents for testID={test_id} on this page."
        )

        for item in page_results:
            records.append({
                "testID": item.get("testId"),
                "serverIP": item.get("serverIp"),
                "agentID": item.get("agentId")
            })

        links = data.get("_links", {})
        next_link = links.get("next", {}).get("href")
        if not next_link:
            logger.info("No more pages found. Pagination completed.")
            break

        url = next_link
        logger.info(f"Found more pages... continuing to '{url}'...")

    if records:
        df = pd.DataFrame(records)
    else:
        df = pd.DataFrame(columns=["testID", "serverIP", "agentID"])

    return df


def fetch_labels(aid: str) -> pd.DataFrame:
    """
    Fetch endpoint labels for a given Account ID.
    Endpoint: /endpoint/labels?aid=XXXXXXX
    """
    url = f"{BASE_URL}/endpoint/labels?aid={aid}"
    logger.info(f"Fetching Labels for AID={aid} from '{url}'...")
    resp = requests.get(url, headers=HEADERS)
    resp.raise_for_status()

    data = resp.json().get("labels", [])
    logger.info(f"Received {len(data)} Labels for AID={aid}.")
    records = []
    for label in data:
        records.append({
            "id": label.get("id"),
            "name": label.get("name"),
            "color": label.get("color"),  # e.g. "#93249F" or "93249F"
            "matchType": label.get("matchType")
        })
    return pd.DataFrame(records)


def main():
    logger.info("Starting ThousandEyes data collection script...")

    # 1) Read account groups from local Excel file
    account_groups_df = get_account_ids("account_ids.xlsx")
    account_groups_tab = account_groups_df.copy()

    # We'll store final results in a dict of DataFrames (sheet_name -> DataFrame)
    sheets = {"Account Groups": account_groups_tab}

    # Prepare lists to aggregate data
    agents_records = []
    endpoint_agents_records = []
    enterprise_tests_records = []
    scheduled_tests_records = []
    scheduled_test_assigned_agents_records = []
    labels_records = []

    # 2) For each account group, fetch data
    for idx, row in account_groups_df.iterrows():
        account_group_name = row["accountGroupName"]
        aid = str(row["aid"])
        logger.info(f"Processing accountGroupName='{account_group_name}', AID={aid}...")

        # ---- A) Enterprise Agents ----
        agents_df = fetch_agents(aid)
        if not agents_df.empty:
            agents_df.insert(0, "aid", aid)
            agents_df.insert(0, "accountGroupName", account_group_name)
            agents_records.append(agents_df)

        # ---- B) Endpoint Agents ----
        endpoint_agents_df = fetch_endpoint_agents(aid)
        if not endpoint_agents_df.empty:
            endpoint_agents_df.insert(0, "aid", aid)
            endpoint_agents_df.insert(0, "accountGroupName", account_group_name)
            endpoint_agents_records.append(endpoint_agents_df)

        # ---- C) Enterprise Tests ----
        enterprise_tests_df = fetch_enterprise_tests(aid)
        if not enterprise_tests_df.empty:
            enterprise_tests_df.insert(0, "aid", aid)
            enterprise_tests_df.insert(0, "accountGroupName", account_group_name)
            enterprise_tests_records.append(enterprise_tests_df)

        # ---- D) Scheduled Test Endpoint Agent ----
        scheduled_tests_df = fetch_scheduled_tests(aid)
        if not scheduled_tests_df.empty:
            scheduled_tests_df.insert(0, "aid", aid)
            scheduled_tests_df.insert(0, "accountGroupName", account_group_name)
            scheduled_tests_records.append(scheduled_tests_df)

            # For each scheduled test ID, fetch assigned agents
            for test_id in scheduled_tests_df["testID"]:
                assigned_agents_df = fetch_scheduled_test_assigned_agents(aid,
                                                                          str(test_id))
                if not assigned_agents_df.empty:
                    assigned_agents_df.insert(0, "aid", aid)
                    assigned_agents_df.insert(0, "accountGroupName", account_group_name)
                    scheduled_test_assigned_agents_records.append(assigned_agents_df)

        # ---- E) Labels ----
        labels_df = fetch_labels(aid)
        if not labels_df.empty:
            labels_df.insert(0, "aid", aid)
            labels_df.insert(0, "accountGroupName", account_group_name)
            labels_records.append(labels_df)

    logger.info("Finished fetching data from all accounts. Consolidating data...")

    # 3) Concatenate or create empty dataframes for each category:

    # ---- A) Enterprise Agents (Tab: Agents) ----
    if agents_records:
        agents_final_df = pd.concat(agents_records, ignore_index=True)
    else:
        agents_final_df = pd.DataFrame(
            columns=[
                "accountGroupName", "aid", "OrgId", "agentId", "agentName",
                "agentType", "agentState", "lastSeen", "createdDate",
                "utilization", "location", "enabled", "hostname", "ipAddresses"
            ]
        )
    sheets["Agents"] = agents_final_df

    # ---- B) Endpoint Agents (Tab: Endpoint Agents) ----
    if endpoint_agents_records:
        endpoint_agents_final_df = pd.concat(endpoint_agents_records, ignore_index=True)
    else:
        endpoint_agents_final_df = pd.DataFrame(
            columns=[
                "accountGroupName", "id", "aid", "name", "computerName", "osVersion",
                "platform", "lastSeen", "status", "deleted", "version", "createdAt",
                "numberOfClients", "locationName", "agentType", "licenseType"
            ]
        )
    desired_order = [
        "accountGroupName",
        "id",
        "aid",
        "name",
        "computerName",
        "osVersion",
        "platform",
        "lastSeen",
        "status",
        "deleted",
        "version",
        "createdAt",
        "numberOfClients",
        "locationName",
        "agentType",
        "licenseType"
    ]
    endpoint_agents_final_df = endpoint_agents_final_df.loc[
                               :, [c for c in desired_order if
                                   c in endpoint_agents_final_df.columns]
                               ]
    sheets["Endpoint Agents"] = endpoint_agents_final_df

    # ---- C) Enterprise Tests (Tab: Enterprise Test) ----
    if enterprise_tests_records:
        enterprise_tests_final_df = pd.concat(enterprise_tests_records,
                                              ignore_index=True)
    else:
        enterprise_tests_final_df = pd.DataFrame(
            columns=[
                "accountGroupName", "aid", "testID", "testName", "createdBy",
                "createdDate", "modifiedBy", "modifiedDate", "type", "alertsEnabled",
                "enabled", "direction", "targetAgentID"
            ]
        )
    sheets["Enterprise Test"] = enterprise_tests_final_df

    # ---- D) Scheduled Test Endpoint Agent ----
    if scheduled_tests_records:
        scheduled_tests_final_df = pd.concat(scheduled_tests_records, ignore_index=True)
    else:
        scheduled_tests_final_df = pd.DataFrame(
            columns=[
                "accountGroupName", "aid", "testID", "testName", "server",
                "createdDate", "type", "isEnabled"
            ]
        )
    sheets["Scheduled Test Endpoint Agent"] = scheduled_tests_final_df

    # ---- E) Scheduled Test Assigned Agents ----
    if scheduled_test_assigned_agents_records:
        assigned_agents_final_df = pd.concat(scheduled_test_assigned_agents_records,
                                             ignore_index=True)
    else:
        assigned_agents_final_df = pd.DataFrame(
            columns=["accountGroupName", "aid", "testID", "serverIP", "agentID"]
        )
    sheets["Scheduled Test Assigned Agents"] = assigned_agents_final_df

    # ---- F) Labels ----
    if labels_records:
        labels_final_df = pd.concat(labels_records, ignore_index=True)
    else:
        labels_final_df = pd.DataFrame(
            columns=["accountGroupName", "aid", "id", "name", "color", "matchType"]
        )
    sheets["Labels"] = labels_final_df

    # 4) Write all sheets to a single Excel file
    timestamp_int = int(time.time())
    output_file = f"thousandeyes_data-{timestamp_int}.xlsx"
    logger.info(f"Writing all results to Excel file '{output_file}'...")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    logger.success(
        f"Data successfully written to '{output_file}'. Now applying color fills and number formats...")

    # 5) First: apply color fills in "Labels"
    _apply_color_fills(output_file, sheet_name="Labels", color_col="color")

    # 6) Second: auto-format numeric columns in all sheets
    _auto_format_numbers(output_file)

    logger.success("All post-processing complete. Script finished.")


if __name__ == "__main__":
    main()
