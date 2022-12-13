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
import argparse
from nornir import InitNornir
from nornir.core import Nornir
from nornir_maze.cisco_support.utils import init_args, prepare_nornir_data, prepare_static_serials
from nornir_maze.cisco_support.reports import (
    create_pandas_dataframe_for_report,
    generate_cisco_maintenance_report,
)
from nornir_maze.cisco_support.api_calls import (
    cisco_support_check_authentication,
    get_sni_owner_coverage_by_serial_number,
    get_sni_coverage_summary_by_serial_numbers,
    get_eox_by_serial_numbers,
    get_ss_suggested_release_by_pid,
    print_sni_owner_coverage_by_serial_number,
    print_sni_coverage_summary_by_serial_numbers,
    print_eox_by_serial_numbers,
    print_get_ss_suggested_release_by_pid,
)
from nornir_maze.utils import (
    print_script_banner,
    print_task_title,
    nr_filter_args,
    nr_transform_default_creds_from_env,
    nr_transform_inv_from_env,
    exit_info,
    exit_error,
    construct_filename_with_current_date,
    load_yaml_file,
)


__author__ = "Willi Kubny"
__maintainer__ = "Willi Kubny"
__license__ = "MIT"
__email__ = "willi.kubny@kyndryl.com"
__status__ = "Production"


#### Functions ###############################################################################################


def _init_nornir(args: argparse.Namespace) -> Nornir:
    """
    This function supports the readability and is used within the main() function. The Nornir inventory will
    be initialized, the default username and password will be transformed and loaded from environment
    variables. The same transformation to load the environment variables is done for the mandatory Cisco
    support API credentials and also for all other inventory keys which start with _env. The function returns
    a filtered Nornir object or quits with an error message in case of issues during the function.
    """
    # pylint: disable=invalid-name

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


def _load_report_yaml_config(report_cfg, args):
    """ """
    # If the report_config file string is available
    if report_cfg["yaml_config"]:
        # Load the report variables from the YAML config file as python dictionary
        config = load_yaml_file(
            file=report_cfg["yaml_config"], text="PYTHON load report yaml config file", verbose=args.verbose
        )
        # Update the report_cfg dict with the loaded yaml config
        report_cfg.update(**config)

    # Select the correct string order based on the TSS arguments
    if args.nornir:
        df_order = "nornir_column_order_with_tss" if args.tss else "nornir_column_order"
    else:
        df_order = "static_column_order_with_tss" if args.tss else "static_column_order"

    # Set the df_order to False if the key don't exist
    report_cfg["df_order"] = report_cfg[df_order] if df_order in report_cfg else False
    # Select the correct dataframe order for all dates regarding conditional formatting
    # Set the df_date_columns to False if the key don't exist
    report_cfg["df_date_columns"] = (
        report_cfg["grace_period_cols"] if "grace_period_cols" in report_cfg else False
    )

    return report_cfg


def main() -> None:
    """
    Main function is executed when the file is directly executed.
    """
    # pylint: disable=invalid-name

    #### Initialize Script and Nornir ########################################################################

    print_script_banner(
        title="Cisco Maint-Check",
        text="Cisco maintenance checker with Nornir, the Cisco support APIs, Pandas and XlsxWriter",
    )

    print_task_title("Initialize ArgParse")
    # Initialize the script arguments with ArgParse to define the further script execution
    args = init_args(argparse_prog_name=os.path.basename(__file__))

    # Create a dict for configuration specifications
    report_cfg = {}

    if args.nornir:
        print_task_title("Initialize Nornir")
        # Initialize, transform and filter the Nornir inventory are return the filtered Nornir object
        nr_obj = _init_nornir(args=args)
        # Prepare the Cisco support API key and the secret in a tuple
        api_creds = (
            nr_obj.inventory.defaults.data["cisco_support_api_creds"]["env_client_key"],
            nr_obj.inventory.defaults.data["cisco_support_api_creds"]["env_client_secret"],
        )
        # Get the report_config string from the Nornir inventory for later YAML file load
        report_cfg["yaml_config"] = (
            nr_obj.inventory.defaults.data["cisco_maintenance_report"]["yaml_config"]
            if args.report
            else False
        )
        # Get the report_file string from the Nornir inventory for later destination file constructing
        report_cfg["excel_file"] = (
            nr_obj.inventory.defaults.data["cisco_maintenance_report"]["excel_file"] if args.report else False
        )
        # Get the ibm_tss_report file from the Nornir inventory
        report_cfg["ibm_tss_file"] = (
            nr_obj.inventory.defaults.data["cisco_maintenance_report"]["ibm_tss_file"] if args.tss else False
        )
        print_task_title("Prepare Nornir Data")
        # Prepare the serials dict for later processing
        serials = prepare_nornir_data(nr_obj=nr_obj, verbose=args.verbose)

    else:
        print_task_title("Prepare Static Data")
        # Prepare the serials dict for later processing
        serials = prepare_static_serials(args=args)
        # Prepare the Cisco support API key and the secret in a tuple
        api_creds = (args.api_key, args.api_secret)
        # Create the report_config string for later YAML file load
        report_cfg["yaml_config"] = "reports/src/report_config.yaml"
        # Create the report_file string for later destination file constructing
        report_cfg["excel_file"] = (
            args.excel if args.excel else "reports/cisco_maintenance_report_YYYY-mm-dd.xlsx"
        )
        # Set the ibm_tss_report file
        report_cfg["ibm_tss_file"] = args.tss if args.tss else False

    #### Get Cisco Support-API Data ##########################################################################

    print_task_title("Check Cisco support API OAuth2 client credentials grant flow")

    # Check the API authentication with the client key and secret to get an access token
    # The script will exit with an error message in case the authentication fails
    if not cisco_support_check_authentication(api_creds=api_creds, verbose=args.verbose, silent=False):
        exit_error(task_text="NORNIR cisco maintenance status", text="Bad news! The script failed!")

    print_task_title("Gather Cisco support API data for serial numbers")

    # Cisco Support API Call SNIgetOwnerCoverageStatusBySerialNumbers and update the serials dictionary
    serials = get_sni_owner_coverage_by_serial_number(serial_dict=serials, api_creds=api_creds)
    # Print the results of get_sni_owner_coverage_by_serial_number()
    print_sni_owner_coverage_by_serial_number(serial_dict=serials, verbose=args.verbose)

    # Cisco Support API Call SNIgetCoverageSummaryBySerialNumbers and update the serials dictionary
    serials = get_sni_coverage_summary_by_serial_numbers(serial_dict=serials, api_creds=api_creds)
    # Print the results of get_sni_coverage_summary_by_serial_numbers()
    print_sni_coverage_summary_by_serial_numbers(serial_dict=serials, verbose=args.verbose)

    # Cisco Support API Call EOXgetBySerialNumbers and update the serials dictionary
    serials = get_eox_by_serial_numbers(serial_dict=serials, api_creds=api_creds)
    # Print the results of get_eox_by_serial_numbers()
    print_eox_by_serial_numbers(serial_dict=serials, verbose=args.verbose)

    # Cisco Support API Call getSuggestedReleasesByProductIDs and update the serials dictionary
    serials = get_ss_suggested_release_by_pid(serial_dict=serials, api_creds=api_creds)
    # Print the results of get_ss_suggested_release_by_pid()
    print_get_ss_suggested_release_by_pid(serial_dict=serials, verbose=args.verbose)

    #### Prepate the Excel report data #######################################################################

    # Exit the script if the args.report argument is not set
    if not args.report:
        exit_info(
            task_text="NORNIR cisco maintenance status", text="Good news! The Script successfully finished!"
        )

    print_task_title("Prepare Cisco maintenance report")

    # Load the yaml report config file
    report_cfg = _load_report_yaml_config(report_cfg=report_cfg, args=args)

    # Prepare the report data and create a pandas dataframe
    df = create_pandas_dataframe_for_report(
        serials_dict=serials,
        args=args,
        df_order=report_cfg["df_order"],
        df_date_columns=report_cfg["df_date_columns"],
        tss_report=report_cfg["ibm_tss_file"],
    )

    #### Generate Cisco maintenance report Excel #############################################################

    print_task_title("Generate Cisco maintenance report")

    # Construct the new destination path and filename from the report_file string variable
    report_cfg["excel_file"] = construct_filename_with_current_date(
        filename=report_cfg["excel_file"],
        name="PYTHON construct destination file",
        silent=False,
    )

    # Generate the Cisco Maintenance report Excel file specified by the report_file with the pandas dataframe
    generate_cisco_maintenance_report(df=df, report_cfg=report_cfg)

    exit_info(
        task_text="NORNIR cisco maintenance status", text="Good news! The Script successfully finished!"
    )


if __name__ == "__main__":
    main()
