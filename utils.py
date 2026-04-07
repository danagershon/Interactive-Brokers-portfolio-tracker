import requests
from bs4 import BeautifulSoup
import logging
import json
import pathlib
from typing import Any, Dict
from http import HTTPStatus


type IbAccountId = str  # IB Account ID is a string like U11111111


class AccountConfigJson:
    """
    Account config JSON is a JSON file that contains the account information and paths to the output file and deposits file.

    Contains only static methods (the class is not meant to be instantiated).
    """

    class ExpectedKeys:
        ACCOUNT_INFO_OUTPUT_FILE = "account_info_output_file"
        DEPOSITS_FILE = "deposits_file"
        ACCOUNT_DESC = "account_desc"

    @staticmethod
    def get_keys() -> list[str]:
        """
        Get the expected keys for the account config JSON file.
        """
        return [AccountConfigJson.ExpectedKeys.ACCOUNT_INFO_OUTPUT_FILE, 
                AccountConfigJson.ExpectedKeys.DEPOSITS_FILE, 
                AccountConfigJson.ExpectedKeys.ACCOUNT_DESC]

    @staticmethod
    def load_json_file(config_path: pathlib.Path) -> Dict[str, Any]:
        """
        Load the account config JSON file.
        """
        if not config_path.is_file():
            raise FileNotFoundError(
                f"Account config JSON not found: {config_path}. Copy account_config.example.json to "
                f"account_config.json and edit it."
            )

        with open(config_path, encoding="utf-8") as f:
            account_config_data = json.load(f)
        
        return account_config_data

    @staticmethod
    def _validate_account_config(account_config_data: Dict[str, Any]) -> None:
        """
        Validate the account config data.
        """
        required_keys = AccountConfigJson.get_keys()
        missing = [required_key for required_key in required_keys if required_key not in account_config_data]
        if missing:
            raise ValueError(f"Account config missing keys: {missing}")
        if not isinstance(account_config_data[AccountConfigJson.ExpectedKeys.ACCOUNT_DESC], dict):
            raise ValueError("account_desc must be a JSON object mapping account ID strings to labels")

    @staticmethod
    def get_account_config_values(account_config_data: Dict[str, Any]):
        """
        Get the account config values from the account config data.

        Returns:
            - account_info_output_file: pathlib.Path
            - deposits_file: pathlib.Path
            - account_desc: dict[str, str]
        """
        account_info_output_file = pathlib.Path(account_config_data[AccountConfigJson.ExpectedKeys.ACCOUNT_INFO_OUTPUT_FILE])
        deposits_file = pathlib.Path(account_config_data[AccountConfigJson.ExpectedKeys.DEPOSITS_FILE])
        account_desc = account_config_data[AccountConfigJson.ExpectedKeys.ACCOUNT_DESC]

        return account_info_output_file, deposits_file, account_desc

    @staticmethod
    def load_account_config(config_path: pathlib.Path) -> Dict[str, Any]:
        """
        Load account_info_output_file, deposits_file, and account_desc from JSON.
        """
        account_config_data = AccountConfigJson.load_json_file(config_path)
        AccountConfigJson._validate_account_config(account_config_data)

        return AccountConfigJson.get_account_config_values(account_config_data)


class IbApiConstants:
    """
    String constants that IB uses in it's response message
    """
    CLIENT_ID = 0  # IB API client ID

    class Ports:
        IB_GW_PORT = 4001
        TWS_PORT = 7496
    
    class ReqIds:
        ACCOUNT_SUMMARY_REQ_ID = 9001

    class AccountSummaryReq:
        ACCOUNT_SUMMARY = "account_summary"
        ALL = "All"
        LEDGER = "$LEDGER"

    class Currency:
        BASE = "BASE"
        USD = "USD"
        ILS = "ILS"

    class AccountBalanceField:
        NET_LIQUIDATION_BY_CURRENCY = "NetLiquidationByCurrency"
        STOCK_MARKET_VALUE = "StockMarketValue"  # total value of all stocks in the accountm in USD
        TOTAL_CASH_BALANCE = "TotalCashBalance"  # total cash in the account, in USD
        NET_DIVIDEND = "NetDividend"  # net dividend in the account, in USD
        UNREALIZED_PNL = "UnrealizedPnL"  # unrealized profit/loss in the account, in USD

    ACCOUNT_BALANCE_FIELD_LIST = [
        AccountBalanceField.NET_LIQUIDATION_BY_CURRENCY,
        AccountBalanceField.STOCK_MARKET_VALUE,
        AccountBalanceField.TOTAL_CASH_BALANCE,
        AccountBalanceField.NET_DIVIDEND,
        AccountBalanceField.UNREALIZED_PNL
    ]


class DepositsFile:
    """
    Deposits file is an Excel file that contains the deposits and withdrawals of the account.
    """
    class ExpectedColumns:
        """
        The expected columns in the deposits file are (from left to right):
            - Date: The date of the deposit or withdrawal
            - Amount: The amount of the deposit or withdrawal
            - Currency: The currency of the deposit or withdrawal
            - Type: The type of the deposit or withdrawal ("Deposit" or "Withdrawal")
            - Exchange Rate: The exchange rate of the deposit or withdrawal

        There can be more columns in the file, but the expected columns must be in this order.
        """
        DATE = "Date"
        AMOUNT = "Amount"
        CURRENCY = "Currency"
        TYPE = "Type"  # one of the OperationTypes
        EXCHANGE_RATE = "Exchange Rate"

    class OperationTypes:
        DEPOSIT = "Deposit"
        WITHDRAWAL = "Withdrawal"


def get_usd_to_ils_exchange_rate():
    """
    fetch real-time USD to ILS exchange rate from an external source (Globes)
    :return: (float) real-time USD to ILS exchange rate
    """
    # Send a GET request to Globes URL
    url = "https://www.globes.co.il/portal/instrument.aspx?InstrumentID=10463"
    response = requests.get(url)

    if response.status_code != HTTPStatus.OK:
        raise Exception(f"Failed to retrieve the Globes webpage to get the USD to ILS exchange rate. Status code: {response.status_code}")

    # the request was successful, so parse the HTML content
    soup = BeautifulSoup(response.content, 'html.parser')

    # Find the div tag with the specific id, containing the exchange rate
    tag_name, tag_id = 'div', 'bgLastDeal'
    rate_div = soup.find(tag_name, id=tag_id)  # if scraping fails check if the tag / id changed

    # Extract and return the exchange rate
    if not rate_div:
        raise Exception("Exchange rate tag not found")
    exchange_rate = round(float(rate_div.text.strip()), 3)
    logging.info(f"real-time USD to ILS exchange rate: {exchange_rate}")

    return exchange_rate
