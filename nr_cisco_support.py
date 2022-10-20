#!/usr/bin/env python3
"""
The main function will gather device serial numbers over different input options (argument list, Excel or
dynamically with Nornir) as well as the hostname. With the serial numbers the Cisco support APIs will be
called and the received information will be printed to stdout and optional processed into an Excel report.
Optionally a IBM TSS Maintenance Report can be added with an argument to compare and analyze the IBM TSS
information against the received data from the Cisco support APIs. Also these additional data will be
processed into an Excel report and saved to the local disk.
"""

import os
import sys
import json
from datetime import datetime, timedelta
from colorama import Fore, Style, init
import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_col_to_name
from nornir import InitNornir
from nornir_scrapli.tasks import send_command
from nornir_maze.cisco_support.utils import init_args
from nornir_maze.cisco_support.api_calls import (
    cisco_support_check_authentication,
    get_sni_owner_coverage_by_serial_number,
    print_sni_owner_coverage_by_serial_number,
    get_sni_coverage_summary_by_serial_numbers,
    print_sni_coverage_summary_by_serial_numbers,
    get_eox_by_serial_numbers,
    print_eox_by_serial_numbers,
    verify_cisco_support_api_data,
)
from nornir_maze.utils import (
    print_script_banner,
    print_task_title,
    print_task_name,
    task_host,
    task_info,
    task_error,
    nr_filter_args,
    nr_transform_default_creds_from_env,
    nr_transform_inv_from_env,
    iterate_all,
    get_pandas_column_width,
    construct_filename_with_current_date,
)

init(autoreset=True, strip=False)


__author__ = "Willi Kubny"
__maintainer__ = "Willi Kubny"
__license__ = "MIT"
__email__ = "willi.kubny@kyndryl.com"
__status__ = "Production"


#### Excel Report Constants ##################################################################################

# Change the constants below to adapt the Excel report creation to your needs

# Specify all settings for the title row formatting
TITLE_ROW_HEIGHT = 60
TITLE_FONT_NAME = "Arial"
TITLE_FONT_SIZE = 20
TITLE_FONT_COLOR = "#FFFFFF"
TITLE_BACKGROUND_COLOR = "#FF452C"
# Specify all settings for the title logo (logo placement is in merged cell A1-A3)
TITLE_LOGO = "reports/src/title_logo.png"
TITLE_LOGO_X_SCALE = 1.0
TITLE_LOGO_Y_SCALE = 1.2
TITLE_LOGO_X_OFFSET = 80
TITLE_LOGO_Y_OFFSET = 18
# Specify the title text (title text starts from cell A4)
TITLE_TEXT = "Cisco Maintenance Report"
TITLE_TEXT_WITH_TSS = "Cisco Maintenance Report incl. IBM TSS Analysis"
# Specify the Excel table formatting style
EXCEL_TABLE_STYLE = "Table Style Medium 8"
# Specify the default table text settings
TABLE_FONT_NAME = "Arial"
TABLE_FONT_SIZE = 12
# Specify the grace period in days where a date should be flaged orange before expire and is flaged red
DATE_GRACE_PERIOD = 90
# Specify the list of dict keys and their order for the pandas dataframe -> Key order == excel colums order
# When a key is removed, the column is removed for the Excel report
# fmt: off
EXCEL_COLUMN_ORDER = [
    "host", "sr_no", "sr_no_owner", "is_covered", "coverage_end_date", "coverage_action_needed",
    "api_action_needed", "contract_site_customer_name", "contract_site_address1", "contract_site_city",
    "contract_site_state_province", "contract_site_country", "covered_product_line_end_date",
    "service_contract_number", "service_line_descr", "warranty_end_date", "warranty_type",
    "warranty_type_description", "item_description", "item_type", "orderable_pid", "ErrorDescription",
    "ErrorDataType", "ErrorDataValue", "EOXExternalAnnouncementDate", "EndOfSaleDate",
    "EndOfSWMaintenanceReleases", "EndOfRoutineFailureAnalysisDate", "EndOfServiceContractRenewal",
    "LastDateOfSupport", "EndOfSvcAttachDate", "UpdatedTimeStamp", "MigrationInformation",
    "MigrationProductId", "MigrationProductName", "MigrationStrategy", "MigrationProductInfoURL",
]
EXCEL_COLUMN_ORDER_WITH_TSS = [
    "host", "sr_no", "sr_no_owner", "is_covered", "coverage_end_date", "coverage_action_needed",
    "api_action_needed", "tss_serial", "tss_status", "contract_site_customer_name",
    "contract_site_address1", "contract_site_city", "contract_site_state_province", "contract_site_country",
    "covered_product_line_end_date", "service_contract_number", "tss_contract", "tss_service_level",
    "service_line_descr", "warranty_end_date", "warranty_type", "warranty_type_description",
    "item_description", "item_type", "orderable_pid", "ErrorDescription", "ErrorDataType", "ErrorDataValue",
    "EOXExternalAnnouncementDate", "EndOfSaleDate", "EndOfSWMaintenanceReleases",
    "EndOfRoutineFailureAnalysisDate", "EndOfServiceContractRenewal", "LastDateOfSupport",
    "EndOfSvcAttachDate", "UpdatedTimeStamp", "MigrationInformation", "MigrationProductId",
    "MigrationProductName", "MigrationStrategy", "MigrationProductInfoURL",
]
# Specify all columns with a date for conditional formatting
DATE_COLUMN_LIST = [
    "coverage_end_date", "EOXExternalAnnouncementDate", "EndOfSaleDate", "EndOfSWMaintenanceReleases",
    "EndOfRoutineFailureAnalysisDate", "EndOfServiceContractRenewal", "LastDateOfSupport",
    "EndOfSvcAttachDate",
]
# fmt: on
# Get the current date in the format YYYY-mm-dd
DATE_TODADY = datetime.today().date()


def init_nornir(args):
    """
    This function supports the readability and is used within the main() function. The Nornir inventory will
    be initialized, the default username and password will be transformed and loaded from environment
    variables. The same transformation to load the environment variables is done for the mandatory Cisco
    support API credentials and also for all other inventory keys which start with _env. The function returns
    a filtered Nornir object or quits with an error message in case of issues during the function.
    """
    # pylint: disable=invalid-name

    print_task_title("Initialize Nornir")

    # Initialize Nornir Object with a config file
    nr = InitNornir(config_file="inventory/nr_config.yaml")

    # Transform the Nornir default username and password from environment variables
    nr_transform_default_creds_from_env(nr_obj=nr, verbose=args.verbose)

    # Transform the Nornir inventory and load all env variables staring with "_env" in default.yaml
    nr_transform_inv_from_env(
        iterable=nr.inventory.defaults.data,
        verbose=args.verbose,
        mandatory={
            "cisco_support_api_creds": {
                "env_client_key": "CISCO_SUPPORT_API_KEY",
                "env_client_secret": "CISCO_SUPPORT_API_SECRET",
            },
        },
    )

    # Filter the Nornir inventory based on the provided arguments from init_args
    nr_obj = nr_filter_args(nr_obj=nr, args=args)

    return nr_obj


def prepare_nornir_data(nr_obj, args):
    """
    This function use Nornir to gather and prepare all serial numbers and returns the serials dictionary and
    the report_file string for the destination file of the report.
    """
    task_text = "NORNIR prepare serial numbers"
    print_task_name(text=task_text)

    # Create a dict from the comma separated serials argument with the serial as key and the hostname
    # as value. The hostname is none as there is no possibility to know which host the serial is
    serials = {}

    # Run the Nornir Scrapli task send_command to get show version
    task_result = nr_obj.run(task=send_command, command="show version")

    # from nornir_utils.plugins.functions import print_result
    # print_result(task_result)

    for host in task_result:
        # Snag the specific host results out of the nornir `AggregateResult` object
        host_result = task_result[host][0]
        # Iterate over the whole Scrapli genie_parse_output() and list all key-value pairs in a tuple
        for item in iterate_all(host_result.scrapli_response.genie_parse_output(), returned="key-value"):
            if "system_sn" in item[0]:
                nr_no = item[1]
                # Create a dict for the serial number as key as a key-value pair for the hostname
                serial = {nr_no.upper(): {}}
                serial[nr_no.upper()]["host"] = host

                print(task_host(host=f"HOST: {host} / SN: {nr_no}", changed=False))
                print(task_info(text=task_text, changed=False))
                print(f"'Add {nr_no} to serials dict' -> NornirResult <Success: True>")
                if args.verbose:
                    print("\n" + json.dumps(serial, indent=4))

                # Update the outside loop serials dict with the inside loop serial dict
                serials.update(serial)

    # Get the report_file string from the Nornir inventory for later destination file constructing
    report_file = nr_obj.inventory.defaults.data["cisco_maintenance_report"]["file"]

    # Prepare the Cisco support API key and the secret in a tuple
    api_client_creds = (
        nr_obj.inventory.defaults.data["cisco_support_api_creds"]["env_client_key"],
        nr_obj.inventory.defaults.data["cisco_support_api_creds"]["env_client_secret"],
    )

    # return the serials dict and the report_file string variable and the api_client_creds tuple
    return serials, report_file, api_client_creds


def prepare_static_data(args):
    """
    This function prepare all static serial numbers which can be applied with the --serials ArgParse argument
    or within an Excel document. It returns the serials dictionary and and the report_file string for the
    destination file of the report.
    """
    # pylint: disable=invalid-name

    task_text = "ARGPARSE verify static provided data"
    print_task_name(text=task_text)

    # Create a dict from the comma separated serials argument with the serial as key and the hostname
    # as value. The hostname is none as there is no possibility to know which host the serial is
    serials = {}

    # If the --serials argument is set, verify that the tag has hosts assigned to
    if hasattr(args, "serials"):
        # Create the report_file string for later destination file constructing
        report_file = "reports/cisco_maintenance_report_YYYY-mm-dd.xlsx"

        print(task_info(text=task_text, changed=False))
        print(f"'{task_text}' -> ArgparseResult <Success: True>")

        # Add all serials from args.serials to the serials dict
        for sr_no in args.serials.split(","):
            serials[sr_no.upper()] = {}
            serials[sr_no.upper()]["host"] = None

        print(task_info(text="PYTHON prepare static provided serial numbers", changed=False))
        print("'PYTHON prepare static provided serial numbers' -> ArgparseResult <Success: True>")
        if args.verbose:
            print("\n" + json.dumps(serials, indent=4))

    # If the --excel argument is set, verify that the tag has hosts assigned to
    elif hasattr(args, "excel"):
        # Verify that the excel file exists
        if not os.path.exists(args.excel):
            # If the excel don't exist -> exit the script properly
            print(task_error(text=task_text, changed=False))
            print(f"'{task_text}' -> ArgparseResult <Success: False>")
            print("\n\U0001f4a5 ALERT: FILE NOT FOUND! \U0001f4a5\n")
            print(
                f"{Style.BRIGHT}{Fore.RED}-> Excel file {args.excel} not found\n"
                "-> Verify the file path and the --excel argument\n\n"
            )
            sys.exit(1)

        # Create the report_file string for later destination file constructing
        report_file = args.excel

        print(task_info(text=task_text, changed=False))
        print(f"'{task_text}' -> ArgparseResult <Success: True>")

        # Read the excel file into a pandas dataframe -> Row 0 is the title row
        df = pd.read_excel(rf"{args.excel}", skiprows=[0])

        # Make all serial numbers written in uppercase letters
        df.sr_no = df.sr_no.str.upper()

        # The first fillna will replace all of (None, NAT, np.nan, etc) with Numpy's NaN, then replace
        # Numpy's NaN with python's None
        df = df.fillna(np.nan).replace([np.nan], [None])

        # Add all serials and hostnames from pandas dataframe to the serials dict
        for sr_no, host in zip(df.sr_no, df.host):
            serials[sr_no] = {}
            serials[sr_no]["host"] = host

        # Print the static provided serial numbers
        print(task_info(text="PANDAS prepare static provided Excel", changed=False))
        print("'PANDAS prepare static provided Excel' -> ArgparseResult <Success: True>")
        if args.verbose:
            print("\n" + json.dumps(serials, indent=4))

    else:
        print(task_error(text=task_text, changed=False))
        print(f"'{task_text}' -> ArgparseResult <Success: False>")
        print("\n\n\U0001f4a5 ALERT: NOT SUPPORTET ARGPARSE ARGUMENT FOR FURTHER PROCESSING! \U0001f4a5")
        print(f"\n{Style.BRIGHT}{Fore.RED}-> Analyse the python function for missing Argparse processing\n\n")
        sys.exit(1)

    # Prepare the Cisco support API key and the secret in a tuple
    api_client_creds = (args.api_key, args.api_secret)

    # return the serials dict and the report_file string variable and the api_client_creds tuple
    return serials, report_file, api_client_creds


def prepare_report_data_host(serials_dict):
    """
    This function takes the serials_dict which has been filled with data by various functions and creates a
    host dict with the key "host" and a list of all hostnames as the value. The key will be the pandas
    dataframe column name and the value which is a list will be the colums cell content. The host dict will
    be returned.
    """
    # Define dict key for the hostnames
    host = {}
    host["host"] = []
    # Add all hostnames to the list
    for item in serials_dict.values():
        host["host"].append(item["host"])

    return host


def prepare_report_data_owner_coverage_by_serial_number(serials_dict):
    """
    This function takes the serials_dict which has been filled with data by various functions and creates a
    owner_coverage_status with key-value pairs. The key will be the pandas dataframe column name and the
    value which is a list will be the colums cell content. The host dict will be returned.
    """
    # pylint: disable=consider-using-dict-items

    # Define dict keys for SNIgetOwnerCoverageStatusBySerialNumbers
    owner_coverage_status = {}
    owner_coverage_status["sr_no_owner"] = []
    owner_coverage_status["coverage_end_date"] = []
    # Append the SNIgetOwnerCoverageStatusBySerialNumbers values for each defined dict key
    for header in owner_coverage_status:
        for sr_no in serials_dict.values():
            success = False
            for key, value in sr_no["SNIgetOwnerCoverageStatusBySerialNumbers"].items():
                if header == key:
                    if key in owner_coverage_status:
                        owner_coverage_status[key].append(value)
                        success = True
        # If nothing was appended to the owner_coverage_status dict, append an empty string
        if not success:
            owner_coverage_status[header].append("")

    return owner_coverage_status


def prepare_report_data_coverage_summary_by_serial_numbers(serials_dict):
    """
    This function takes the serials_dict which has been filled with data by various functions and creates a
    coverage_summary with key-value pairs. The key will be the pandas dataframe column name and the value
    which is a list will be the colums cell content. The host dict will be returned.
    """
    # pylint: disable=consider-using-dict-items

    # Define dict keys for SNIgetCoverageSummaryBySerialNumbers
    coverage_summary = {}
    coverage_summary["sr_no"] = []
    coverage_summary["is_covered"] = []
    coverage_summary["contract_site_customer_name"] = []
    coverage_summary["contract_site_address1"] = []
    coverage_summary["contract_site_city"] = []
    coverage_summary["contract_site_state_province"] = []
    coverage_summary["contract_site_country"] = []
    coverage_summary["covered_product_line_end_date"] = []
    coverage_summary["service_contract_number"] = []
    coverage_summary["service_line_descr"] = []
    coverage_summary["warranty_end_date"] = []
    coverage_summary["warranty_type"] = []
    coverage_summary["warranty_type_description"] = []
    coverage_summary["item_description"] = []
    coverage_summary["item_type"] = []
    coverage_summary["orderable_pid"] = []
    # Append the SNIgetCoverageSummaryBySerialNumbers values for each defined dict key
    for header in coverage_summary:
        for sr_no in serials_dict.values():
            success = False
            # Append all general coverage details
            for key, value in sr_no["SNIgetCoverageSummaryBySerialNumbers"].items():
                if header == key:
                    if key in coverage_summary:
                        coverage_summary[key].append(value)
                        success = True
            # Append all the orderable pid details
            for key, value in sr_no["SNIgetCoverageSummaryBySerialNumbers"]["orderable_pid_list"][0].items():
                if header == key:
                    if key in coverage_summary:
                        coverage_summary[key].append(value)
                        success = True
            # If nothing was appended to the coverage_summary dict, append an empty string
            if not success:
                coverage_summary[header].append("")

    return coverage_summary


def prepare_report_data_eox_by_serial_numbers(serials_dict):
    """
    This function takes the serials_dict which has been filled with data by various functions and creates a
    end_of_life with key-value pairs. The key will be the pandas dataframe column name and the value which is
    a list will be the colums cell content. The host dict will be returned.
    """
    # pylint: disable=too-many-nested-blocks,consider-using-dict-items

    # Define dict keys for EOXgetBySerialNumbers
    end_of_life = {}
    end_of_life["EOXExternalAnnouncementDate"] = []
    end_of_life["EndOfSaleDate"] = []
    end_of_life["EndOfSWMaintenanceReleases"] = []
    end_of_life["EndOfSecurityVulSupportDate"] = []
    end_of_life["EndOfRoutineFailureAnalysisDate"] = []
    end_of_life["EndOfServiceContractRenewal"] = []
    end_of_life["LastDateOfSupport"] = []
    end_of_life["EndOfSvcAttachDate"] = []
    end_of_life["UpdatedTimeStamp"] = []
    end_of_life["MigrationInformation"] = []
    end_of_life["MigrationProductId"] = []
    end_of_life["MigrationProductName"] = []
    end_of_life["MigrationStrategy"] = []
    end_of_life["MigrationProductInfoURL"] = []
    end_of_life["ErrorDescription"] = []
    end_of_life["ErrorDataType"] = []
    end_of_life["ErrorDataValue"] = []

    # Append the EOXgetBySerialNumbers values for each defined dict key
    for header in end_of_life:
        for sr_no in serials_dict.values():
            success = False

            # Append all end of life dates
            for key, value in sr_no["EOXgetBySerialNumbers"].items():
                if header == key:
                    if isinstance(value, dict):
                        if "value" in value:
                            end_of_life[key].append(value["value"])
                            success = True
            # Append all migration details
            for key, value in sr_no["EOXgetBySerialNumbers"]["EOXMigrationDetails"].items():
                if header == key:
                    if key in end_of_life:
                        end_of_life[key].append(value)
                        success = True
            # If EOXError exists append the error reason, else append an empty string
            if "EOXError" in sr_no["EOXgetBySerialNumbers"]:
                for key, value in sr_no["EOXgetBySerialNumbers"]["EOXError"].items():
                    if header == key:
                        if key in end_of_life:
                            end_of_life[key].append(value)
                            success = True

            # If nothing was appended to the end_of_life dict, append an empty string
            if not success:
                end_of_life[header].append("")

    return end_of_life


def prepare_report_data_act_needed(serials_dict):
    """
    This function takes the serials_dict an argument and creates a dictionary named act_needed will be
    returned which contains the key value pairs "coverage_action_needed" and "api_action_needed" to create a
    Pandas dataframe later.
    """
    # Define the coverage_action_needed dict and its key value pairs to return as the end of the function
    act_needed = {}
    act_needed["coverage_action_needed"] = []
    act_needed["api_action_needed"] = []

    for records in serials_dict.values():
        # Verify if the user has the correct access rights to access the serial API data
        if "YES" in records["SNIgetOwnerCoverageStatusBySerialNumbers"]["sr_no_owner"]:
            act_needed["api_action_needed"].append(
                "No action needed (API user is associated with contract and device)"
            )
        else:
            act_needed["api_action_needed"].append(
                "Action needed (No associattion between api user, contract and device)"
            )

        # Verify if the serial is covered by Cisco add the coverage_action_needed variable to tss_info
        if "YES" in records["SNIgetCoverageSummaryBySerialNumbers"]["is_covered"]:
            act_needed["coverage_action_needed"].append(
                "No action needed (Device is covered by a maintenance contract)"
            )
        else:
            act_needed["coverage_action_needed"].append(
                "Action needed (Device is not covered by a maintenance contract)"
            )

    return act_needed


def prepare_report_data_tss(serials_dict, file):
    """
    This function takes the serials_dict and a source file which is the IBM TSS report as arguments. The only
    mandatory column is the "Serials" which will be normalized to tss_serial. All other columns can be
    specified with their order and the prefix "tss_" in the EXCEL_COLUMN_ORDER_WITH_TSS constant. A dictionary
    named tss_info will be returned which contains the key value pairs "coverage_action_needed",
    "api_action_needed" and all TSS data to create a Pandas dataframe later.
    """
    # pylint: disable=invalid-name,too-many-branches

    # Define the tss_info dict and its key value pairs to return as the end of the function
    tss_info = {}
    tss_info["coverage_action_needed"] = []
    tss_info["api_action_needed"] = []
    for column in EXCEL_COLUMN_ORDER_WITH_TSS:
        if column.startswith("tss_"):
            tss_info[column] = []

    # Read the excel file into a pandas dataframe -> Row 0 is the title row
    df = pd.read_excel(rf"{file}")

    # Make some data normalization of the TSS report file column headers
    # Make column written in lowercase letters
    df.columns = df.columns.str.lower()
    # Replace column name whitespace with underscore
    df.columns = df.columns.str.replace(" ", "_")
    # Add a prefix to the column name to identify the TSS report columns
    df = df.add_prefix("tss_")

    # Make all serial numbers written in uppercase letters
    df.tss_serial = df.tss_serial.str.upper()

    # The first fillna will replace all of (None, NAT, np.nan, etc) with Numpy's NaN, then replace
    # Numpy's NaN with python's None
    df = df.fillna(np.nan).replace([np.nan], [None])

    # Delete all rows which have not the value "Cisco" in the OEM column
    df = df[df.tss_oem == "Cisco"]

    # Create a list with all IBM TSS serial numbers
    tss_serial_list = df["tss_serial"].tolist()

    # Look for inventory serials which are covered by IBM TSS and add them to the serial_dict
    # It's important to match IBM TSS serials to inventory serials first for the correct order
    for sr_no, records in serials_dict.items():
        records["tss_info"] = {}

        # Covered by IBM TSS if inventory serial number is in all IBM TSS serial numbers
        if sr_no in tss_serial_list:
            # Verify if the user has the correct access rights to access the serial API data
            if "YES" in records["SNIgetOwnerCoverageStatusBySerialNumbers"]["sr_no_owner"]:
                tss_info["api_action_needed"].append(
                    "No action needed (API user is associated with contract and device)"
                )
            else:
                tss_info["api_action_needed"].append(
                    "Action needed (No associattion between api user, contract and device)"
                )

            # Verify if the inventory serial is covered by IBM TSS and is also covered by Cisco
            # Verify if the serial is covered by Cisco add the coverage_action_needed variable to tss_info
            if "YES" in records["SNIgetCoverageSummaryBySerialNumbers"]["is_covered"]:
                tss_info["coverage_action_needed"].append("No action needed (Covered by IBM TSS and Cisco)")
            else:
                tss_info["coverage_action_needed"].append(
                    "Action needed (Covered by IBM TSS, but Cisco coverage missing)"
                )

            # Get the index of the list item and assign the element from the TSS dataframe by its index
            index = tss_serial_list.index(sr_no)
            # Add the data from the TSS dataframe to tss_info
            for column, value in tss_info.items():
                if column.startswith("tss_"):
                    value.append(df[column].values[index])

        # Inventory serial number is not in all IBM TSS serial numbers
        else:
            # Verify if the user has the correct access rights to access the serial API data
            if "YES" in records["SNIgetOwnerCoverageStatusBySerialNumbers"]["sr_no_owner"]:
                tss_info["api_action_needed"].append(
                    "No action needed (API user is associated with contract and device)"
                )
            else:
                tss_info["api_action_needed"].append(
                    "Action needed (No associattion between api user, contract and device)"
                )

            # Verify if the inventory serial is covered by Cisco
            # Add the coverage_action_needed variable to tss_info
            if "YES" in records["SNIgetCoverageSummaryBySerialNumbers"]["is_covered"]:
                tss_info["coverage_action_needed"].append("No action needed (Covered by Cisco SmartNet)")
            else:
                tss_info["coverage_action_needed"].append(
                    "Action needed (Cisco SmartNet or IBM TSS coverage missing)"
                )

            # Add the empty strings for all additional IBM TSS serials to tss_info
            for column, value in tss_info.items():
                if column.startswith("tss_"):
                    value.append("")

    # After the inventory serials have been processed
    # Add IBM TSS serials to tss_info which are not part of the inventory serials
    for tss_serial in tss_serial_list:
        if tss_serial not in serials_dict.keys():
            # Add the coverage_action_needed variable to tss_info
            tss_info["coverage_action_needed"].append(
                "Action needed (Remove serial from IBM TSS inventory as device is decommissioned)"
            )
            tss_info["api_action_needed"].append("No Cisco API data as serial is not part of column sr_no")

            # Get the index of the list item and assign the element from the TSS dataframe by its index
            index = tss_serial_list.index(tss_serial)
            # Add the data from the TSS dataframe to tss_info
            for column, value in tss_info.items():
                if column.startswith("tss_"):
                    value.append(df[column].values[index])

    return tss_info


def create_pandas_dataframe_for_report(serials_dict, tss_report=False, verbose=False):
    """
    Prepare the report data and create a pandas dataframe. The pandas dataframe will be returned
    """
    # pylint: disable=invalid-name

    print_task_name(text="PYTHON prepare report data")

    # Create an empty dict and append the previous dicts to create later the pandas dataframe
    report_data = {}

    if tss_report:
        # Prepare the needed data for the report from the serials dict. The serials dict contains all data
        # that the Cisco support API sent. These functions return a dictionary with the needed data only
        host = prepare_report_data_host(serials_dict=serials_dict)
        owner_coverage_status = prepare_report_data_owner_coverage_by_serial_number(serials_dict=serials_dict)
        coverage_summary = prepare_report_data_coverage_summary_by_serial_numbers(serials_dict=serials_dict)
        end_of_life = prepare_report_data_eox_by_serial_numbers(serials_dict=serials_dict)

        # Analyze the IBM TSS report file and create the tss_info dict
        tss_info = prepare_report_data_tss(serials_dict=serials_dict, file=tss_report)

        # The tss_info dict may have more list elements as TSS serials have been found which are not inside
        # the customer inventory -> Add the differente to all other lists as empty strings
        for _ in range(len(tss_info["tss_serial"]) - len(host["host"])):
            for column in host.values():
                column.append("")
            for column in owner_coverage_status.values():
                column.append("")
            for column in coverage_summary.values():
                column.append("")
            for column in end_of_life.values():
                column.append("")

        # Update the report_data dict with all prepared data dicts
        report_data.update(**host, **owner_coverage_status, **coverage_summary, **end_of_life, **tss_info)
    else:
        # Prepare the needed data for the report from the serials dict. The serials dict contains all data
        # that the Cisco support API sent. These functions return a dictionary with the needed data only
        host = prepare_report_data_host(serials_dict=serials_dict)
        owner_coverage_status = prepare_report_data_owner_coverage_by_serial_number(serials_dict=serials_dict)
        coverage_summary = prepare_report_data_coverage_summary_by_serial_numbers(serials_dict=serials_dict)
        end_of_life = prepare_report_data_eox_by_serial_numbers(serials_dict=serials_dict)
        act_needed = prepare_report_data_act_needed(serials_dict=serials_dict)

        # Update the report_data dict with all prepared data dicts
        report_data.update(**host, **owner_coverage_status, **coverage_summary, **end_of_life, **act_needed)

    print(task_info(text="PYTHON prepare report data dict", changed=False))
    print("'PYTHON prepare report data dict' -> PythonResult <Success: True>")
    if verbose:
        print("\n" + json.dumps(report_data, indent=4))

    # Reorder the data dict according to the key_order list -> This needs Python >= 3.6
    if tss_report:
        report_data = {key: report_data[key] for key in EXCEL_COLUMN_ORDER_WITH_TSS}
    else:
        report_data = {key: report_data[key] for key in EXCEL_COLUMN_ORDER}

    print(task_info(text="PYTHON order report data dict", changed=False))
    print("'PYTHON order report data dict' -> PythonResult <Success: True>")
    if verbose:
        print("\n" + json.dumps(report_data, indent=4))

    # Create a Pandas dataframe for the data dict
    df = pd.DataFrame(report_data)

    # Format each column in the list to a pandas date type for later conditional formatting
    for column in DATE_COLUMN_LIST:
        df[column] = pd.to_datetime(df[column], format="%Y-%m-%d")

    print(task_info(text="PYTHON create pandas dataframe from dict", changed=False))
    print("'PANDAS create dataframe' -> PandasResult <Success: True>")
    if verbose:
        print(df)

    return df


def generate_cisco_maintenance_report(report_file, df, tss_report=False):
    """
    Generate the Cisco Maintenance report Excel file specified by the report_file with the pandas dataframe.
    The function returns None, but saves the Excel file to the local disk.
    """
    # pylint: disable=invalid-name,too-many-locals

    #### Create the xlsx writer, workbook and worksheet objects ##############################################

    # Create a Pandas excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(  # pylint: disable=abstract-class-instantiated
        report_file, engine="xlsxwriter", date_format="yyyy-mm-dd", datetime_format="yyyy-mm-dd"
    )

    # Write the dataframe data to XlsxWriter. Turn off the default header and index and skip one row to allow
    # us to insert a user defined header.
    df.to_excel(writer, sheet_name="Cisco_Maintenance_Report", startrow=2, header=False, index=False)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets["Cisco_Maintenance_Report"]

    # Setting for the whole worksheet
    worksheet.set_zoom(110)
    worksheet.freeze_panes(2, 3)

    # Get the dimensions of the dataframe.
    (max_row, max_col) = df.shape
    # Max_row + 2 because the first two rows are used for title and header
    max_row = max_row + 2
    # Max_com -1 otherwise would be one column to much
    max_col = max_col - 1

    print(task_info(text="PYTHON create pandas writer object using XlsxWriter engine", changed=False))
    print("'PYTHON create pandas writer object using XlsxWriter engine' -> PythonResult <Success: True>")

    #### Create the top title row ############################################################################

    # Set the top row height
    worksheet.set_row(0, TITLE_ROW_HEIGHT)
    # Create a format to use for the merged top row
    title_format = workbook.add_format(
        {
            "font_name": TITLE_FONT_NAME,
            "font_size": TITLE_FONT_SIZE,
            "font_color": TITLE_FONT_COLOR,
            "align": "left",
            "valign": "vcenter",
            "bold": 1,
            "bottom": 1,
            "bg_color": TITLE_BACKGROUND_COLOR,
        }
    )
    # Merge the first three cells in the top row to insert a logo
    worksheet.merge_range(0, 0, 0, 2, None, title_format)
    # Insert a logo to the top row
    worksheet.insert_image(
        "A1",
        TITLE_LOGO,
        {
            "x_scale": TITLE_LOGO_X_SCALE,
            "y_scale": TITLE_LOGO_Y_SCALE,
            "x_offset": TITLE_LOGO_X_OFFSET,
            "y_offset": TITLE_LOGO_Y_OFFSET,
        },
    )
    # Merge from the cell 4 to the max_col and write a title
    if tss_report:
        title_text = f"{TITLE_TEXT_WITH_TSS} (generated by {os.path.basename(__file__)})"
    else:
        title_text = f"{TITLE_TEXT} (generated by {os.path.basename(__file__)})"
    worksheet.merge_range(0, 3, 0, max_col, title_text, title_format)

    print(task_info(text="PYTHON create XlsxWriter title row", changed=False))
    print("'PYTHON create XlsxWriter title row' -> PythonResult <Success: True>")

    ### Create a Excel table structure and add the Pandas dataframe

    # Create a list of column headers, to use in add_table().
    column_settings = [{"header": column} for column in df.columns]

    # Add the Excel table structure. Pandas will add the data.
    # fmt: off
    worksheet.add_table(1, 0, max_row - 1, max_col,
        {
            "columns": column_settings,
            "style": EXCEL_TABLE_STYLE,
        },
    )
    # fmt: on

    table_format = workbook.add_format(
        {"font_name": TABLE_FONT_NAME, "font_size": TABLE_FONT_SIZE, "align": "left", "valign": "vcenter"}
    )
    # Auto-adjust each column width -> +5 on the width makes space for the filter icon
    for index, width in enumerate(get_pandas_column_width(df)):
        worksheet.set_column(index, index - 1, width + 5, table_format)

    print(task_info(text="PYTHON create XlsxWriter table and add pandas dataframe", changed=False))
    print("'PYTHON create XlsxWriter table and add pandas dataframe' -> PythonResult <Success: True>")

    #### Create conditional formating ########################################################################

    # Create a red background format for the conditional formatting
    red_format = workbook.add_format({"bg_color": "#C0504D"})
    # Create a orange background format for the conditional formatting
    orange_format = workbook.add_format({"bg_color": "#F79646"})
    # Create a green background format for the conditional formatting
    green_format = workbook.add_format({"bg_color": "#9BBB59"})

    # Create a conditional formatting for each column in the list.
    column_list = ["sr_no_owner", "is_covered", "coverage_action_needed", "api_action_needed"]
    for column in column_list:
        # Get the column letter by the column name
        target_col = xl_col_to_name(df.columns.get_loc(column))
        # -> Excel requires the value for type cell to be double quoted
        worksheet.conditional_format(
            f"{target_col}3:{target_col}{max_row}",
            {"type": "cell", "criteria": "equal to", "value": '"NO"', "format": red_format},
        )
        worksheet.conditional_format(
            f"{target_col}3:{target_col}{max_row}",
            {"type": "cell", "criteria": "equal to", "value": '"YES"', "format": green_format},
        )
        worksheet.conditional_format(
            f"{target_col}3:{target_col}{max_row}",
            {"type": "text", "criteria": "containing", "value": "No action needed", "format": green_format},
        )
        worksheet.conditional_format(
            f"{target_col}3:{target_col}{max_row}",
            {"type": "text", "criteria": "containing", "value": "Action needed", "format": red_format},
        )

    # Create a conditional formatting for each column with a date. Get the column letter by the column name
    for column in DATE_COLUMN_LIST:
        target_col = xl_col_to_name(df.columns.get_loc(column))
        worksheet.conditional_format(
            f"{target_col}3:{target_col}{max_row}",
            {
                "type": "date",
                "criteria": "between",
                "minimum": DATE_TODADY + timedelta(days=DATE_GRACE_PERIOD),
                "maximum": datetime.strptime("2999-01-01", "%Y-%m-%d"),
                "format": green_format,
            },
        )
        worksheet.conditional_format(
            f"{target_col}3:{target_col}{max_row}",
            {
                "type": "date",
                "criteria": "between",
                "minimum": DATE_TODADY,
                "maximum": DATE_TODADY + timedelta(days=DATE_GRACE_PERIOD),
                "format": orange_format,
            },
        )
        worksheet.conditional_format(
            f"{target_col}3:{target_col}{max_row}",
            {
                "type": "date",
                "criteria": "between",
                "minimum": datetime.strptime("1999-01-01", "%Y-%m-%d"),
                "maximum": DATE_TODADY - timedelta(days=1),
                "format": red_format,
            },
        )

    print(task_info(text="PYTHON create XlsxWriter conditional formating", changed=False))
    print("'PYTHON create XlsxWriter conditional formating' -> PythonResult <Success: True>")

    #### Save the Excel report file to disk ##################################################################

    print_task_name(text="PYTHON generate report Excel file")

    # Close the Pandas Excel writer and output the Excel file.
    writer.close()

    print(task_info(text="PYTHON generate report Excel file", changed=False))
    print("'PYTHON generate report Excel file' -> PythonResult <Success: True>")
    print(f"\n-> Saved information about {df.shape[0]} serials to {report_file}")


def main():
    """
    Main function is executed when the file is directly executed.
    """
    # pylint: disable=invalid-name

    #### Initialize Script and Nornir ########################################################################

    print_script_banner(
        title="Cisco Maint-Check",
        text="Cisco maintenance checker with Nornir and the Cisco support APIs",
    )

    print_task_title("Initialize ArgParse")

    # Initialize the script arguments with ArgParse to define the further script execution
    use_nornir, args = init_args(argparse_prog_name=os.path.basename(__file__))

    if use_nornir:
        # Initialize, transform and filter the Nornir inventory are return the filtered Nornir object
        nr_obj = init_nornir(args=args)
        # Prepare the serials dict and the report_file string for later processing
        serials, report_file, api_client_creds = prepare_nornir_data(nr_obj=nr_obj, args=args)
    else:
        # Prepare the serials dict and the report_file string for later processing
        serials, report_file, api_client_creds = prepare_static_data(args=args)

    #### Get Cisco Support-API Data ##########################################################################

    print_task_title("Check Cisco support API OAuth2 client credentials grant flow")

    # Check the API authentication with the client key and secret to get an access token
    # The script will exit with an error message in case the authentication fails
    if not cisco_support_check_authentication(api_client_creds=api_client_creds, verbose=args.verbose):
        sys.exit(1)

    print_task_title("Gather Cisco support API data for serial numbers")

    # Cisco Support API Call SNIgetOwnerCoverageStatusBySerialNumbers and update the serials dictionary
    serials = get_sni_owner_coverage_by_serial_number(
        serial_dict=serials,
        api_client_creds=api_client_creds,
    )
    # Print the results of get_sni_owner_coverage_by_serial_number()
    print_sni_owner_coverage_by_serial_number(serial_dict=serials, verbose=args.verbose)

    # Cisco Support API Call SNIgetCoverageSummaryBySerialNumbers and update the serials dictionary
    serials = get_sni_coverage_summary_by_serial_numbers(
        serial_dict=serials,
        api_client_creds=api_client_creds,
    )
    # Print the results of get_sni_coverage_summary_by_serial_numbers()
    print_sni_coverage_summary_by_serial_numbers(serial_dict=serials, verbose=args.verbose)

    # Cisco Support API Call EOXgetBySerialNumbers and update the serials dictionary
    serials = get_eox_by_serial_numbers(
        serial_dict=serials,
        api_client_creds=api_client_creds,
    )
    # Print the results of get_eox_by_serial_numbers()
    print_eox_by_serial_numbers(serial_dict=serials, verbose=args.verbose)

    # Verify that the serials dictionary contains no wrong serial numbers
    # The script will exit with an error message in case of invalid serial numbers
    if not verify_cisco_support_api_data(serials_dict=serials, verbose=args.verbose):
        sys.exit(1)

    #### Prepate the Excel report data #######################################################################

    # Exit the script if the --args.report argument is not set
    if not args.report:
        print("\n\u2728 Good news! Script successfully finished! \u2728\n\n")
        sys.exit(0)

    print_task_title("Prepare Cisco maintenance report")

    # Prepare the report data and create a pandas dataframe
    df = create_pandas_dataframe_for_report(serials_dict=serials, tss_report=args.tss, verbose=args.verbose)

    #### Generate Cisco maintenance report Excel #############################################################

    print_task_title("Generate Cisco maintenance report")
    print_task_name(text="PYTHON process report data")

    # Construct the new destination path and filename from the report_file string variable
    report_file = construct_filename_with_current_date(filename=report_file)
    print(task_info(text="PYTHON construct destination file", changed=False))
    print("'PYTHON construct destination file' -> PythonResult <Success: True>")
    print(f"\n-> Constructed {report_file}\n")

    # Generate the Cisco Maintenance report Excel file specified by the report_file with the pandas dataframe
    generate_cisco_maintenance_report(report_file=report_file, df=df, tss_report=args.tss)

    print("\n\u2728 Good news! Script successfully finished! \u2728")

    print("\n")


if __name__ == "__main__":
    main()
