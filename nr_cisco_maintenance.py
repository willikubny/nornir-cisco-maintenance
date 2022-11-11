#!/usr/bin/env python3
"""
The main function will gather device serial numbers over different input options (argument list, Excel or
dynamically with Nornir) as well as the hostname. With the serial numbers the Cisco support APIs will be
called and the received information will be printed to stdout and optional processed into an Excel report.
Optionally a IBM TSS Maintenance Report can be added with an argument to compare and analyze the IBM TSS
information against the received data from the Cisco support APIs. Also these additional data will be
processed into an Excel report and saved to the local disk.
"""

import argparse
import os
from datetime import datetime
from nornir import InitNornir
from nornir.core import Nornir
from nornir_maze.cisco_support.utils import init_args, prepare_nornir_serials, prepare_static_serials
from nornir_maze.cisco_support.reports import (
    create_pandas_dataframe_for_report,
    generate_cisco_maintenance_report,
)
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
    nr_filter_args,
    nr_transform_default_creds_from_env,
    nr_transform_inv_from_env,
    exit_info,
    exit_error,
)


__author__ = "Willi Kubny"
__maintainer__ = "Willi Kubny"
__license__ = "MIT"
__email__ = "willi.kubny@kyndryl.com"
__status__ = "Production"


#### Excel Report Constants ##################################################################################

# Change the dictionary values below to adapt the Excel report generation to your needs

report_vars = {
    # Specify all settings for the title row formatting
    "title_row_height": 60,
    "title_font_name": "Calibri",
    "title_font_size": 20,
    "title_font_color": "#FFFFFF",
    "title_background_color": "#FF452C",
    # Specify all settings for the title logo (logo placement is in merged cell A1-A3)
    "title_logo": "reports/src/title_logo.png",
    "title_logo_x_scale": 1.0,
    "title_logo_y_scale": 1.2,
    "title_logo_x_offset": 80,
    "title_logo_y_offset": 18,
    # Specify the title text (title text starts from cell A4)
    "title_text": "Cisco Maintenance Report",
    "title_text_with_tss": "Cisco Maintenance Report incl. IBM TSS Analysis",
    # Specify the Excel table formatting style
    "excel_table_style": "Table Style Medium 8",
    # Specify the default table text settings
    "table_font_name": "Calibri",
    "table_font_size": 11,
    # Specify the grace period in days where a date should be flaged orange before expire and is flaged red
    "date_grace_period": 90,
    # Specify the order of the dict keys for the pandas dataframe -> Key order == excel colums order
    # When a key is removed, the column is removed for the Excel report
    # fmt: off
    "excel_column_order" : [
        "host", "sr_no", "sr_no_owner", "is_covered", "coverage_end_date", "coverage_action_needed",
        "api_action_needed", "contract_site_customer_name", "contract_site_address1", "contract_site_city",
        "contract_site_state_province", "contract_site_country", "covered_product_line_end_date",
        "service_contract_number", "service_line_descr", "warranty_end_date", "warranty_type",
        "warranty_type_description", "item_description", "item_type", "orderable_pid", "ErrorDescription",
        "ErrorDataType", "ErrorDataValue", "EOXExternalAnnouncementDate", "EndOfSaleDate",
        "EndOfSWMaintenanceReleases", "EndOfRoutineFailureAnalysisDate", "EndOfServiceContractRenewal",
        "LastDateOfSupport", "EndOfSvcAttachDate", "UpdatedTimeStamp", "MigrationInformation",
        "MigrationProductId", "MigrationProductName", "MigrationStrategy", "MigrationProductInfoURL",
    ],
    "excel_column_order_with_tss" : [
        "host", "sr_no", "sr_no_owner", "is_covered", "coverage_end_date", "coverage_action_needed",
        "api_action_needed", "tss_serial", "tss_status", "contract_site_customer_name",
        "contract_site_address1", "contract_site_city", "contract_site_state_province",
        "contract_site_country", "covered_product_line_end_date", "service_contract_number", "tss_contract",
        "tss_service_level", "service_line_descr", "warranty_end_date", "warranty_type",
        "warranty_type_description", "item_description", "item_type", "orderable_pid", "ErrorDescription",
        "ErrorDataType", "ErrorDataValue", "EOXExternalAnnouncementDate", "EndOfSaleDate",
        "EndOfSWMaintenanceReleases", "EndOfRoutineFailureAnalysisDate", "EndOfServiceContractRenewal",
        "LastDateOfSupport", "EndOfSvcAttachDate", "UpdatedTimeStamp", "MigrationInformation",
        "MigrationProductId", "MigrationProductName", "MigrationStrategy", "MigrationProductInfoURL",
    ],
    # Specify all columns with a date for conditional formatting
    "date_column_list" : [
        "coverage_end_date", "EOXExternalAnnouncementDate", "EndOfSaleDate", "EndOfSWMaintenanceReleases",
        "EndOfRoutineFailureAnalysisDate", "EndOfServiceContractRenewal", "LastDateOfSupport",
        "EndOfSvcAttachDate",
    ],
    # fmt: on
    # Get the current date in the format YYYY-mm-dd
    "date_today": datetime.today().date(),
}


#### Functions ###############################################################################################


def init_nornir(args: argparse.Namespace) -> Nornir:
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


def main() -> None:
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

        # Prepare the Cisco support API key and the secret in a tuple
        api_creds = (
            nr_obj.inventory.defaults.data["cisco_support_api_creds"]["env_client_key"],
            nr_obj.inventory.defaults.data["cisco_support_api_creds"]["env_client_secret"],
        )
        # Get the report_file string from the Nornir inventory for later destination file constructing
        if args.report:
            report_file = nr_obj.inventory.defaults.data["cisco_maintenance_report"]["file"]

        # Prepare the serials dict for later processing
        serials = prepare_nornir_serials(nr_obj=nr_obj, verbose=args.verbose)
    else:
        # Prepare the serials dict for later processing
        serials = prepare_static_serials(args=args)

        # Prepare the Cisco support API key and the secret in a tuple
        api_creds = (args.api_key, args.api_secret)

        # Create the report_file string for later destination file constructing
        if args.report:
            report_file = args.excel if args.excel else "reports/cisco_maintenance_report_YYYY-mm-dd.xlsx"

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

    # Verify that the serials dictionary contains no wrong serial numbers
    # The script will exit with an error message in case of invalid serial numbers
    if not verify_cisco_support_api_data(serials_dict=serials, verbose=args.verbose, silent=False):
        exit_error(task_text="NORNIR cisco maintenance status", text="Bad news! The script failed!")

    #### Prepate the Excel report data #######################################################################

    # Exit the script if the args.report argument is not set
    if not args.report:
        exit_info(
            task_text="NORNIR cisco maintenance status", text="Good news! The Script successfully finished!"
        )

    print_task_title("Prepare Cisco maintenance report")

    # Prepare the report data and create a pandas dataframe
    df_order = report_vars["excel_column_order_with_tss"] if args.tss else report_vars["excel_column_order"]
    df = create_pandas_dataframe_for_report(
        serials_dict=serials,
        df_order=df_order,
        df_date_columns=report_vars["date_column_list"],
        tss_report=args.tss,
        verbose=args.verbose,
    )

    #### Generate Cisco maintenance report Excel #############################################################

    print_task_title("Generate Cisco maintenance report")

    # Generate the Cisco Maintenance report Excel file specified by the report_file with the pandas dataframe
    generate_cisco_maintenance_report(
        report_vars=report_vars,
        report_file=report_file,
        df=df,
        tss_report=args.tss,
    )

    exit_info(
        task_text="NORNIR cisco maintenance status", text="Good news! The Script successfully finished!"
    )


if __name__ == "__main__":
    main()
