#!/usr/bin/env python3
"""
Docstring needs to be changed
"""

import os
import sys
import json
from pyfiglet import figlet_format
import pandas as pd
import numpy as np
from cisco_support import SNI, EoX
from cisco_support.utils import getToken as cisco_support_get_token
from nornir import InitNornir
from nornir_scrapli.tasks import send_command
from nornir_maze.cisco_support.utils import init_args
from nornir_maze.utils import (
    print_task_title,
    print_task_name,
    task_host,
    task_info,
    task_error,
    nr_filter_args,
    nr_transform_default_creds_from_env,
    nr_transform_inv_from_env,
    iterate_all,
)


__author__ = "Willi Kubny"
__maintainer__ = "Willi Kubny"
__license__ = "MIT"
__email__ = "willi.kubny@kyndryl.com"
__status__ = "Production"


def cisco_support_check_authentication(api_client_creds, verbose=False):
    """
    This function checks to Cisco support API authentication by generating an bearer access token. In case
    of an invalid API client key or secret a error message is printed and the script exits.
    """
    task_name = "CISCO-API check OAuth2 client credentials grant flow"
    print_task_name(text=task_name)

    # Create two variables for the Cisco support API key and secret
    client_id, client_secret = api_client_creds

    try:
        # Try to generate an barer access token
        token = cisco_support_get_token(client_id, client_secret, verify=None, proxies=None)

        print(task_info(text=task_name, changed="False"))
        print("'Bearer access token generation' -> CISCOAPIResult <Success: True>")
        if verbose:
            print(f"\n-> Bearer token: {token}\n")

    except KeyError:
        ansi_red_bold = "\033[1m\u001b[31m"
        ansi_reset = "\033[0m"
        print(task_error(text=task_name, changed="False"))
        print("'Bearer access token generation' -> CISCOAPIResult <Success: False>")
        print("\n\U0001f4a5 ALERT: INVALID API CREDENTIALS PROVIDED! \U0001f4a5")
        print(f"{ansi_red_bold}-> Verify the API client key and secret{ansi_reset}\n")
        sys.exit(1)


def sni_get_owner_coverage_by_serial_number(serial_dict, api_client_creds, verbose=False):
    """
    This function takes the serial_dict which contains all serial numbers and the Cisco support API creds to
    run get the owner coverage by serial number with the cisco-support library. The printout is in Nornir
    style, but there is no use of Nornir. The result of each serial will be added with a new key to the dict.
    The function returns the updated serials dict. The format of the serials_dict need to be as below.
    "<serial>": {
        "host": "<hostname>",
        ...
    },
    """
    print_task_name(text="CISCO-API get owner coverage status by serial number")

    client_key, client_secret = api_client_creds

    sni = SNI(client_key, client_secret)

    owner_coverage_status = sni.getOwnerCoverageStatusBySerialNumbers(serial_dict.keys())

    for item in owner_coverage_status["serial_numbers"]:
        sr_no = item["sr_no"]
        serial_dict[sr_no]["SNIgetOwnerCoverageStatusBySerialNumbers"] = item
        print(task_host(host=f"HOST: {serial_dict[sr_no]['host']} / SN: {sr_no}", changed="False"))

        # Verify if the serial number is associated with the CCO ID
        if "YES" in item["sr_no_owner"]:
            print(task_info(text="Verify provided CCO ID", changed="False"))
            print("'Is associated to the provided CCO ID' -> CISCOAPIResult <Success: True>")
        else:
            print(task_error(text="Verify provided CCO ID", changed="False"))
            print("'Is not associated to the provided CCO ID' -> CISCOAPIResult <Success: False>")

        # Verify if the serial is covered by a service contract
        if "YES" in item["is_covered"]:
            print(task_info(text="Verify service contract", changed="False"))
            print("'Is covered by a service contract' -> CISCOAPIResult <Success: True>")
            # Verify the end date of the service contract coverage
            if item["coverage_end_date"]:
                print(task_info(text="Verify service contract end date", changed="False"))
                print(f"'Coverage end date is {item['coverage_end_date']}' -> CISCOAPIResult <Success: True>")
            else:
                print(task_error(text="Verify service contract end date", changed="False"))
                print("'Coverage end date not available' -> CISCOAPIResult <Success: False>")
        else:
            print(task_error(text="Verify service contract", changed="False"))
            print("'Is not covered by a service contract' -> CISCOAPIResult <Success: False>")

        if verbose:
            print("\n" + json.dumps(item, indent=4))

    return serial_dict


def sni_get_coverage_summary_by_serial_numbers(serial_dict, api_client_creds, verbose=False):
    """
    This function takes the serial_dict which contains all serial numbers and the Cisco support API creds to
    run get the coverage summary by serial number with the cisco-support library. The printout is in Nornir
    style, but there is no use of Nornir. The result of each serial will be added with a new key to the dict.
    The function returns the updated serials dict. The format of the serials_dict need to be as below.
    "<serial>": {
        "host": "<hostname>",
        ...
    },
    """
    task_text = "CISCO-API get coverage summary data by serial number"
    print_task_name(text=task_text)

    client_key, client_secret = api_client_creds

    sni = SNI(client_key, client_secret)

    coverage_summary = sni.getCoverageSummaryBySerialNumbers(serial_dict.keys())

    for item in coverage_summary["serial_numbers"]:
        sr_no = item["sr_no"]
        serial_dict[sr_no]["SNIgetCoverageSummaryBySerialNumbers"] = item
        print(task_host(host=f"HOST: {serial_dict[sr_no]['host']} / SN: {sr_no}", changed="False"))

        if "ErrorResponse" in item:
            error_response = item["ErrorResponse"]["APIError"]
            print(task_error(text=task_text, changed="False"))
            print("'Get SNI data' -> CISCOAPIResult <Success: False>")
            print(f"\n-> {error_response['ErrorDescription']} ({error_response['SuggestedAction']})\n")
        else:
            print(task_info(text=task_text, changed="False"))
            print("'Get SNI data' -> CISCOAPIResult <Success: True>")
            print(f"\n-> Orderable pid: {item['orderable_pid_list'][0]['orderable_pid']}")
            print(f"-> Customer name: {item['contract_site_customer_name']}")
            print(f"-> Customer address: {item['contract_site_address1']}")
            print(f"-> Customer city: {item['contract_site_city']}")
            print(f"-> Customer province: {item['contract_site_state_province']}")
            print(f"-> Customer country: {item['contract_site_country']}")
            print(f"-> Is covered by service contract: {item['is_covered']}")
            print(f"-> Covered product line end date: {item['covered_product_line_end_date']}")
            print(f"-> Service contract number: {item['service_contract_number']}")
            print(f"-> Service contract description: {item['service_line_descr']}")
            print(f"-> Warranty end date: {item['warranty_end_date']}")
            print(f"-> Warranty type: {item['warranty_type']}\n")

        if verbose:
            print(json.dumps(item, indent=4))

    return serial_dict


def eox_by_serial_numbers(serial_dict, api_client_creds, verbose=False):
    """
    This function takes the serial_dict which contains all serial numbers and the Cisco support API creds to
    run get the end of life data by serial number with the cisco-support library. The printout is in Nornir
    style, but there is no use of Nornir. The result of each serial will be added with a new key to the dict.
    The function returns the updated serials dict. The format of the serials_dict need to be as below.
    "<serial>": {
        "host": "<hostname>",
        ...
    },
    """
    task_text = "CISCO-API get EoX data by serial number"
    print_task_name(text=task_text)

    client_key, client_secret = api_client_creds

    eox = EoX(client_key, client_secret)

    end_of_life = eox.getBySerialNumbers(serial_dict.keys())

    for item in end_of_life["EOXRecord"]:
        sr_no = item["EOXInputValue"]
        serial_dict[sr_no]["EOXgetBySerialNumbers"] = item
        print(task_host(host=f"HOST: {serial_dict[sr_no]['host']} / SN: {sr_no}", changed="False"))

        if "EOXError" in item:
            if "No product IDs were found" in item["EOXError"]["ErrorDescription"]:
                print(task_error(text=task_text, changed="False"))
                print("'Get EoX data' -> CISCOAPIResult <Success: False>")
                print(f"\n-> {item['EOXError']['ErrorDescription']} (Serial number does not exist)\n")
            elif "EOX information does not exist" in item["EOXError"]["ErrorDescription"]:
                print(task_info(text=task_text, changed="False"))
                print("'Get EoX data' -> CISCOAPIResult <Success: False>")
                print(f"\n-> {item['EOXError']['ErrorDescription']}\n")
        else:
            print(task_info(text=task_text, changed="False"))
            print(
                f"'Get EoX data (Update timestamp {item['UpdatedTimeStamp']['value']})' "
                + "-> CISCOAPIResult <Success: True>"
            )
            print(f"\n-> EoL product ID: {item['EOLProductID']}")
            print(f"-> Product ID description: {item['ProductIDDescription']}")
            print(f"-> EoL announcement date: {item['EOXExternalAnnouncementDate']['value']}")
            print(f"-> End of sale date: {item['EndOfSaleDate']['value']}")
            print(f"-> End of maintenance release: {item['EndOfSWMaintenanceReleases']['value']}")
            print(f"-> End of vulnerability support: {item['EndOfSecurityVulSupportDate']['value']}")
            print(f"-> Last day of support: {item['LastDateOfSupport']['value']}\n")

        if verbose:
            print(json.dumps(item, indent=4))

    return serial_dict


def verify_cisco_support_api_data(serials_dict, verbose=False):
    """
    This function verifies the serials_dict which has been filled with data by various functions of these
    module like eox_by_serial_numbers, sni_get_coverage_summary_by_serial_numbers, etc. and verifies that
    there are no invalid serial numbers. In case of invalid serial numbers, the script quits with an error
    message.
    """
    failed = False
    task_text = "Verify Cisco support API data"
    print_task_name(text=task_text)

    # Verify that the serials_dict dictionary contains no wrong serial numbers
    for value in iterate_all(iterable=serials_dict, returned="value"):
        if value is not None:
            if "No product IDs were found" in value or "No records found" in value:
                failed = True
                break

    if failed:
        print(task_error(text=task_text, changed="False"))
        print(f"'{task_text}' -> Result <Success: False>")
        print(
            "\n\U0001f4a5 ALERT: INVALID SERIAL NUMBERS PROVIDED! \U0001f4a5\n"
            "\033[1m\u001b[31m"
            "-> Analyse the output for failed tasks to identify the invalid serial numbers\n"
            "-> Run the script with valid serial numbers only again\033[0m\n\n"
        )
        sys.exit(1)

    print(task_info(text=task_text, changed="False"))
    print(f"'{task_text}' -> Result <Success: True>")
    if verbose:
        print("\n" + json.dumps(serials_dict, indent=4))


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
            "cisco_support": {
                "env_client_key": "CISCO_SUPPORT_KEY",
                "env_client_secret": "CISCO_SUPPORT_SECRET",
            },
        },
    )

    # Filter the Nornir inventory based on the provided arguments from init_args
    nr_obj = nr_filter_args(nr_obj=nr, args=args)

    return nr_obj


def prepare_nornir_data(nr_obj, args):
    """
    This function use Nornir to gather and prepare all serial numbers and returns the serials dictionary and
    the Cisco support API credentials from the Nornir inventory.
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

                print(task_host(host=f"HOST: {host} / SN: {nr_no}", changed="False"))
                print(task_info(text=task_text, changed="False"))
                print(f"'Add {nr_no} to serials dict' -> NornirResult <Success: True>")
                if args.verbose:
                    print("\n" + json.dumps(serial, indent=4))

                # Update the outside loop serials dict with the inside loop serial dict
                serials.update(serial)

    # Prepare the Cisco support API key and the secret in a tuple
    api_client_creds = (
        nr_obj.inventory.defaults.data["cisco_support"]["env_client_key"],
        nr_obj.inventory.defaults.data["cisco_support"]["env_client_secret"],
    )

    return serials, api_client_creds


def prepare_static_data(args):
    """
    This function prepare all static serial numbers which can be applied with the --serials ArgParse argument
    or within an Excel document. It returns the serials dictionary and the Cisco support API credentials from
    --api_key and --api_secret ArgParse arguments.
    """
    # pylint: disable=invalid-name

    task_text = "ARGPARSE verify static provided data"
    print_task_name(text=task_text)

    # Create a dict from the comma separated serials argument with the serial as key and the hostname
    # as value. The hostname is none as there is no possibility to know which host the serial is
    serials = {}

    # If the --serials argument is set, verify that the tag has hosts assigned to
    if hasattr(args, "serials"):
        print(task_info(text=task_text, changed="False"))
        print(f"'{task_text}' -> ArgparseResult <Success: True>")

        # Add all serials from args.serials to the serials dict
        for sr_no in args.serials.split(","):
            serials[sr_no.upper()] = {}
            serials[sr_no.upper()]["host"] = None

        print(task_info(text="PYTHON prepare static provided serial numbers", changed="False"))
        print("'PYTHON prepare static provided serial numbers' -> ArgparseResult <Success: True>")
        if args.verbose:
            print("\n" + json.dumps(serials, indent=4))

    # If the --excel argument is set, verify that the tag has hosts assigned to
    elif hasattr(args, "excel"):
        # Verify that the excel file exists
        if not os.path.exists(args.excel):
            # If the excel don't exist -> exit the script properly
            print(task_error(text=task_text, changed="False"))
            print(f"'{task_text}' -> ArgparseResult <Success: False>")
            print()
            print(
                "\n\U0001f4a5 ALERT: FILE NOT FOUND! \U0001f4a5\n"
                f"\033[1m\u001b[31m-> Excel file {args.excel} not found\n"
                "-> Verify the file path and the --excel argument\033[0m\n\n"
            )
            sys.exit()

        print(task_info(text=task_text, changed="False"))
        print(f"'{task_text}' -> ArgparseResult <Success: True>")

        # Read the excel file into a pandas dataframe
        df = pd.read_excel(rf"{args.excel}")

        # Make all serial numbers and hostnames written in uppercase letters
        df.SERIAL_NUMBER = df.SERIAL_NUMBER.str.upper()
        df.HOSTNAME = df.HOSTNAME.str.upper()

        # The first fillna will replace all of (None, NAT, np.nan, etc) with Numpy's NaN, then replace
        # Numpy's NaN with python's None
        df = df.fillna(np.nan).replace([np.nan], [None])

        # Add all serials and hostnames from pandas dataframe to the serials dict
        for sr_no, host in zip(df.SERIAL_NUMBER, df.HOSTNAME):
            serials[sr_no] = {}
            serials[sr_no]["host"] = host

        # Print the static provided serial numbers
        print(task_info(text="PANDAS prepare static provided Excel", changed="False"))
        print("'PANDAS prepare static provided Excel' -> ArgparseResult <Success: True>")
        if args.verbose:
            print("\n" + json.dumps(serials, indent=4))

    else:
        print(task_error(text=task_text, changed="False"))
        print(
            f"'{task_text}' -> ArgparseResult <Success: False>"
            + "\n\n\U0001f4a5 ALERT: NOT SUPPORTET ARGPARSE ARGUMENT FOR FURTHER PROCESSING! \U0001f4a5"
            + "\n\033[1m\u001b[31m-> Analyse the python function for missing Argparse processing\n\n\033[0m"
        )
        sys.exit(1)

    # Prepare the Cisco support API key and the secret in a tuple
    api_client_creds = (args.api_key, args.api_secret)

    return serials, api_client_creds


def main():
    """
    Main function is executed when the file is directly executed.
    """
    #### Initialize Script and Nornir ########################################################################

    # Print a custom script banner
    print(
        f"\n\033[92m{figlet_format('Cisco Maint-Check', width=110)}"
        + "Cisco maintenance checker with Nornir and the Cisco support APIs\033[0m"
    )

    print_task_title("Initialize ArgParse")

    # Initialize the script arguments with ArgParse to define the further script execution
    use_nornir, args = init_args(argparse_prog_name=os.path.basename(__file__))

    if use_nornir:
        # Initialize, transform and filter the Nornir inventory are return the filtered Nornir object
        nr_obj = init_nornir(args=args)
        # Prepare the serials dict and the Cisco support API credentials tuple
        serials, api_client_creds = prepare_nornir_data(nr_obj=nr_obj, args=args)
    else:
        # Prepare the serials dict and the Cisco support API credentials tuple
        serials, api_client_creds = prepare_static_data(args=args)

    #### Get Cisco Support-API Data ##########################################################################

    print_task_title("Check Cisco support API OAuth2 client credentials grant flow")

    # Check the API authentication with the client key and secret to get an access token
    # The script will exit with an error message in case the authentication fails
    cisco_support_check_authentication(api_client_creds=api_client_creds, verbose=args.verbose)

    print_task_title("Gather Cisco support API data for serial numbers")

    # Cisco Support API Call SNIgetOwnerCoverageStatusBySerialNumbers and update the serials dictionary
    serials = sni_get_owner_coverage_by_serial_number(
        serial_dict=serials,
        api_client_creds=api_client_creds,
        verbose=args.verbose,
    )

    # Cisco Support API Call SNIgetCoverageSummaryBySerialNumbers and update the serials dictionary
    serials = sni_get_coverage_summary_by_serial_numbers(
        serial_dict=serials,
        api_client_creds=api_client_creds,
        verbose=args.verbose,
    )

    # Cisco Support API Call EOXgetBySerialNumbers and update the serials dictionary
    serials = eox_by_serial_numbers(
        serial_dict=serials,
        api_client_creds=api_client_creds,
        verbose=args.verbose,
    )

    # Verify that the serials dictionary contains no wrong serial numbers
    # The script will exit with an error message in case of invalid serial numbers
    verify_cisco_support_api_data(serials_dict=serials, verbose=args.verbose)

    #### Create the Excel Report #############################################################################

    print("\n")


if __name__ == "__main__":
    main()
