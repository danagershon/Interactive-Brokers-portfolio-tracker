import requests
from bs4 import BeautifulSoup
import logging


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
