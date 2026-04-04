from ibapi.common import *
import threading
import time
import pandas as pd
import pathlib
from ib_connector_base_class import IBConnector
import utils
import logging
import write_to_excel_helper


class IbAccountInfoFetcher(IBConnector):

    def __init__(self, config_path: pathlib.Path):
        super().__init__()
        self.account_info_output_file, self.deposits_file, self.account_desc = \
            utils.load_account_config(config_path)
        self.account_balance_info = pd.DataFrame(self.create_blank_account_data_structure_dict())
        self.sub_accounts: list[utils.IbAccountId] = []  # Store sub-account identifiers
        self.account_data_received = threading.Event()  # Event to signal data is received
        self.current_account: utils.IbAccountId = None  # Track the current account being processed

    def managedAccounts(self, accountsList):
        """
        Callback to get the list of sub-accounts under the master account.
        """
        self.sub_accounts = list[utils.IbAccountId](accountsList.split(",")[:-1])
        logging.info(f"Managed accounts: {self.sub_accounts}")
        self.account_data_received.set()  # Signal that the sub-account data has been received

    def contractDetails(self, reqId: int, contractDetails):
        logging.info(f"ContractDetails. ReqId: {reqId}, Contract: {contractDetails}")

    def accountSummary(self, reqId, account: utils.IbAccountId, tag: str, value: str, currency: str):
        """
        Handle account summary data for each account.
        """
        if account not in self.account_balance_info:
            # Initialize account balance info for the account if it doesn't exist yet
            self.account_balance_info[account] = self.create_blank_account_data_structure_dict()

        if tag in utils.IbApiConstants.ACCOUNT_BALANCE_FIELDS:
            # Handle base currency (assuming it's USD)
            if currency == "BASE" or currency == "USD":
                self.account_balance_info[account][tag]["USD"] = round(float(value), 2)
            elif currency == "ILS":
                self.account_balance_info[account][tag]["ILS"] = round(float(value), 2)

        logging.debug(f"Account: {account}, Tag: {tag}, Value: {value}, Currency: {currency}")

    def request_account_summary_from_api(self, account_id):
        """
        Fetch account summary for a specific sub-account.
        """
        self.current_account = account_id
        req_id = self.req_ids["account_summary"]
        self.reqAccountSummary(req_id, "All", "$LEDGER")
        time.sleep(2)
        self.cancelAccountSummary(req_id)

    @staticmethod
    def create_blank_account_data_structure_dict():
        return {field: {"USD": 0.0, "ILS": 0.0} for field in utils.IbApiConstants.ACCOUNT_BALANCE_FIELDS}

    def get_total_deposits(self):
        # get total ILS deposits since inception from the manually updated file
        deposits_df = pd.read_excel(self.deposits_file)
        return deposits_df[deposits_df.Amount > 0]["Amount"].sum()  # filter out withdrawals (have negative values)

    def get_account_info(self, write_to_excel):
        """
        Retrieve account information for all sub-accounts.
        """
        # Wait for managedAccounts to populate sub_accounts
        self.account_data_received.wait(timeout=10)  # Wait until managedAccounts is called

        # Prepare a DataFrame for storing aggregated values (sum of all sub-accounts)
        sum_data = self.create_blank_account_data_structure_dict()
        sum_df = pd.DataFrame(sum_data).T  # Transpose to get correct shape

        # Get the latest USD to ILS exchange rate
        exchange_rate = utils.get_usd_to_ils_exchange_rate()

        # For each sub-account, retrieve and process account information
        self.account_balance_info = {}  # Clear any previous data
        for account in self.sub_accounts:
            logging.info(f"Fetching account information for: {account}")

            # Initialize the account data structure before fetching the account summary
            self.account_balance_info[account] = self.create_blank_account_data_structure_dict()

            # Fetch account summary for the current sub-account
            self.request_account_summary_from_api(account)

            # Update ILS values based on the latest exchange rate
            for tag in utils.IbApiConstants.ACCOUNT_BALANCE_FIELDS:
                # Convert USD to ILS using the fetched exchange rate
                usd_value = self.account_balance_info[account][tag]["USD"]
                self.account_balance_info[account][tag]["ILS"] = round(usd_value * exchange_rate, 2)

            # Aggregate the values to the sum DataFrame
            for tag in utils.IbApiConstants.ACCOUNT_BALANCE_FIELDS:
                sum_df.at[tag, "USD"] += self.account_balance_info[account][tag]["USD"]
                sum_df.at[tag, "ILS"] += self.account_balance_info[account][tag]["ILS"]

        # Optionally write to Excel
        if write_to_excel:
            excel_helper = write_to_excel_helper.ExcelHelper(self.account_info_output_file, self.deposits_file, self.account_desc)
            excel_helper.write_account_info_to_excel(sum_df, self.account_balance_info, exchange_rate, self.get_total_deposits())

        return sum_df, self.account_balance_info
