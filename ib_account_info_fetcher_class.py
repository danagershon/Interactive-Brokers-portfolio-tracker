from ibapi.common import TickerId
import threading
import time
import pandas as pd
import pathlib
import logging

from ib_connector_base_class import IBConnector
import utils
from utils import IbApiConstants, DepositsFile, AccountConfigJson
import write_to_excel_helper


class IbAccountInfoFetcher(IBConnector):
    """
    Class to fetch IB account info and write to Excel
    """

    def __init__(self, config_path: pathlib.Path):
        super().__init__()
        self.account_info_output_file, self.deposits_file, self.account_desc = \
            AccountConfigJson.load_account_config(config_path)
        # Per-account balance tables: index=tag (IB balance field name), columns=[USD, ILS]
        self.account_balance_info: dict[utils.IbAccountId, pd.DataFrame] = {}
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

    def accountSummary(self, reqId: TickerId, account: utils.IbAccountId, tag: str, value: str, currency: str):
        """
        Handle account summary data for each account.
        """
        if tag not in IbApiConstants.ACCOUNT_BALANCE_FIELD_LIST:
            logging.debug(f"Unrecognized account summary tag: {tag}")
            return

        if account not in self.account_balance_info:
            self.account_balance_info[account] = self._create_blank_account_df()

        # Handle base currency (assuming it's USD)
        currency_key = (
            IbApiConstants.Currency.USD
            if currency in [IbApiConstants.Currency.BASE, IbApiConstants.Currency.USD]
            else IbApiConstants.Currency.ILS
        )
        self.account_balance_info[account].at[tag, currency_key] = round(float(value), 2)

        logging.debug(f"Account: {account}, Tag: {tag}, Value: {value}, Currency: {currency}")

    def request_account_summary_from_api(self, account_id):
        """
        Fetch account summary for a specific sub-account.
        """
        self.current_account = account_id
        req_id = self.req_ids[IbApiConstants.AccountSummaryReq.ACCOUNT_SUMMARY]
        self.reqAccountSummary(req_id, IbApiConstants.AccountSummaryReq.ALL, IbApiConstants.AccountSummaryReq.LEDGER)
        time.sleep(2)
        self.cancelAccountSummary(req_id)

    @staticmethod
    def _create_blank_account_df() -> pd.DataFrame:
        """
        Create a blank account balance DataFrame with the expected columns and index (IB account summary fields).

        Returns:
            pd.DataFrame - a DataFrame with the expected columns and index (IB account summary fields), filled with 0.0

                                      USD    ILS
            NetLiquidationByCurrency  0.0    0.0
            StockMarketValue          0.0    0.0
            TotalCashBalance          0.0    0.0
            NetDividend               0.0    0.0
            UnrealizedPnL             0.0    0.0
        """
        return pd.DataFrame(
            data=0.0,  # fill with 0.0 for all cells
            index=list(IbApiConstants.ACCOUNT_BALANCE_FIELD_LIST),
            columns=[IbApiConstants.Currency.USD, IbApiConstants.Currency.ILS],
        )

    def get_total_deposits(self) -> tuple[float, float]:
        """
        Get total deposits since inception from the manually updated file.

        Returns:
            (total_ils_deposits, total_usd_deposits)
            
        Note: The USD deposits are pre-converted from ILS.
        """
        deposits_df = pd.read_excel(self.deposits_file)
        deposits_only = deposits_df[
            deposits_df[DepositsFile.ExpectedColumns.TYPE] == DepositsFile.OperationTypes.DEPOSIT
        ]

        # Assumes Amount is in ILS for deposits rows.
        total_ils = float(deposits_only[DepositsFile.ExpectedColumns.AMOUNT].sum())
        # The USD column is assumed to be pre-converted from ILS.
        total_usd = float(deposits_only[DepositsFile.ExpectedColumns.USD].sum())

        return round(total_ils, 2), round(total_usd, 2)

    def get_account_info(self, write_to_excel):
        """
        Retrieve account information for all sub-accounts.
        """
        # Wait for managedAccounts to populate sub_accounts
        self.account_data_received.wait(timeout=10)  # Wait until managedAccounts is called

        # Prepare a DataFrame for storing aggregated values (sum of all sub-accounts)
        sum_df = self._create_blank_account_df()

        # Get the latest USD to ILS exchange rate
        exchange_rate = utils.get_usd_to_ils_exchange_rate()

        # For each sub-account, retrieve and process account information
        self.account_balance_info = {}  # Clear any previous data
        for account in self.sub_accounts:
            logging.info(f"Fetching account information for: {account}")

            # Fetch account summary for the current sub-account
            self.request_account_summary_from_api(account)

            # Update the ILS col in the account balance df by converting each matching USD value to ILS using the fetched exchange rate. then round all ILS values
            self.account_balance_info[account][IbApiConstants.Currency.ILS] = (
                self.account_balance_info[account][IbApiConstants.Currency.USD] * exchange_rate
            ).round(2)

            # Aggregate to sum table (perform element-wise addition of the account balance df with the sum df)
            # there should be no missing values - the account df and sum df have the same index and columns
            sum_df = sum_df.add(self.account_balance_info[account])

        # Optionally write to Excel
        if write_to_excel:
            excel_helper = write_to_excel_helper.ExcelHelper()
            total_ils_deposits, total_usd_deposits = self.get_total_deposits()
            excel_helper.write_account_info_to_excel(self.account_info_output_file, self.account_desc, sum_df, 
                                                     self.account_balance_info, exchange_rate,
                                                     total_ils_deposits, total_usd_deposits)

        return sum_df, self.account_balance_info
