import requests
from bs4 import BeautifulSoup
import logging
import json
import pathlib
from typing import Any, Dict


type IbAccountId = str  # IB Account ID is a string like U11111111


class IbApiConstants:
    """
    String constants that IB uses in it's response message
    """

    ACCOUNT_BALANCE_FIELDS = [
        "NetLiquidationByCurrency",
        "StockMarketValue",
        "TotalCashBalance",
        "NetDividend",
        "UnrealizedPnL"
    ]

    class Currency:
        BASE = "BASE"
        USD = "USD"
        ILS = "ILS"


def get_usd_to_ils_exchange_rate():
    """
    fetch real-time USD to ILS exchange rate from an external source (Globes)
    :return: (float) real-time USD to ILS exchange rate
    """
    # Send a GET request to Globes URL
    url = "https://www.globes.co.il/portal/instrument.aspx?InstrumentID=10463"
    response = requests.get(url)

    if response.status_code != 200:
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


def load_account_config(config_path: pathlib.Path) -> Dict[str, Any]:
    """
    Load account_info_output_file, deposits_file, and account_desc from JSON.
    """
    if not config_path.is_file():
        raise FileNotFoundError(
            f"Account config JSON not found: {config_path}. Copy account_config.example.json to "
            f"account_config.json and edit it."
        )

    with open(config_path, encoding="utf-8") as f:
        account_config_data = json.load(f)

    required_keys = ("account_info_output_file", "deposits_file", "account_desc")
    missing = [required_key for required_key in required_keys if required_key not in account_config_data]
    if missing:
        raise ValueError(f"Account config missing keys: {missing}")
    if not isinstance(account_config_data["account_desc"], dict):
        raise ValueError("account_desc must be a JSON object mapping account id strings to labels")

    account_info_output_file = pathlib.Path(account_config_data["account_info_output_file"])
    deposits_file = pathlib.Path(account_config_data["deposits_file"])
    account_desc = account_config_data["account_desc"]

    return account_info_output_file, deposits_file, account_desc
