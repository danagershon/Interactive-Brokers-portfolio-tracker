from ibapi.client import EClient
from ibapi.wrapper import EWrapper
from ibapi.contract import Contract
from ibapi.execution import ExecutionFilter
from ibapi.common import *
import threading
import time
from datetime import datetime, timedelta
import yfinance as yf
import openpyxl
import pandas as pd
import requests
from bs4 import BeautifulSoup
from itertools import chain
import argparse

EXCEL_FILES_DIR = "ExcelFiles/"


class IBConnector(EWrapper, EClient):
    LOCALHOST = "127.0.0.1"
    CLIENT_ID = 0

    def __init__(self, host=LOCALHOST, connect_to_IB_GW=True, client_id=CLIENT_ID):
        EClient.__init__(self, self)
        self.connection_thread = threading.Thread(target=self.run)
        self.host = host
        self.port = 4001 if connect_to_IB_GW else 7496  # TWS port
        self.client_id = client_id
        self.req_ids = dict(exchange_rate=101, account_summary=9001)
        self.exchange_rate_received = threading.Event()

    def __enter__(self):
        self.connect_app()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect_app()

    def connect_app(self):
        self.connect(self.host, self.port, self.client_id)
        self.connection_thread.start()
        time.sleep(1)  # Give some time for the connection to establish

    def disconnect_app(self):
        self.disconnect()
        self.connection_thread.join()

    def error(self, reqId, errorCode:int, errorString:str, advancedOrderRejectJson = ""):
        print(f"Error: {reqId} {errorCode} {errorString}")


class IBAccountInfo(IBConnector):

    def __init__(self):
        super().__init__()
        self.account_balance_fields = ["NetLiquidationByCurrency",
                                       "StockMarketValue",
                                       "TotalCashBalance",
                                       "NetDividend",
                                       "UnrealizedPnL"]
        account_balance_dict = {field: {"USD": 0.0, "ILS": 0.0} for field in self.account_balance_fields}
        self.account_balance_info = pd.DataFrame(account_balance_dict)
        self.account_balance_file = 'account_info.xlsx'
        self.deposits_file = 'Deposits.xlsx'

    def contractDetails(self, reqId: int, contractDetails):
        print(f"ContractDetails. ReqId: {reqId}, Contract: {contractDetails}")

    def tickPrice(self, reqId: int, tickType, price, attrib):
        bid_tick = 9  # Assuming tickType 9 is the bid price
        if reqId == self.req_ids['exchange_rate'] and tickType == bid_tick:
            self.account_balance_info["IbExchangeRate"] = price
            self.exchange_rate_received.set()

    def accountSummary(self, reqId: int, account: str, tag: str, value: str, currency: str):
        if tag in self.account_balance_fields:
            self.account_balance_info[tag]["USD"] = round(float(value))

    def request_account_summary_from_api(self):
        req_id = self.req_ids["account_summary"]
        self.reqAccountSummary(req_id, "All", "$LEDGER")
        time.sleep(2)  # Wait for the responses to come in
        self.cancelAccountSummary(req_id)  # Cancel the request to stop receiving updates
        return self.account_balance_info

    @staticmethod
    def get_usd_to_ils_exchange_rate():
        """
        fetch real-time USD to ILS exchange rate from an external source (Globes)
        :return: (float) real-time USD to ILS exchange rate
        """
        # Send a GET request to Globes URL
        url = "https://www.globes.co.il/portal/instrument.aspx?InstrumentID=10463"
        response = requests.get(url)

        # Check if the request was successful
        if response.status_code == 200:
            # Parse the HTML content
            soup = BeautifulSoup(response.content, 'html.parser')

            # Find the div tag with the specific id, containing the exchange rate
            tag_name, tag_id = 'div', 'bgLastDeal'
            rate_div = soup.find(tag_name, id=tag_id)  # if scraping fails check if the tag / id changed

            # Extract and return the exchange rate
            if rate_div:
                exchange_rate = round(float(rate_div.text.strip()), 3)
                print(f"real-time USD to ILS exchange rate: {exchange_rate}")
                return exchange_rate
            else:
                raise Exception("Exchange rate tag not found")
        else:
            raise Exception(f"Failed to retrieve the webpage. Status code: {response.status_code}")

    def request_ib_exchange_rate(self):
        """
        fetch IB USD to ILS exchange rate
        note that the retrieved rate is not in real-time thus cannot be depended on
        :return:
        """
        contract = Contract()
        contract.symbol = "USD"
        contract.secType = "CASH"
        contract.currency = "ILS"
        contract.exchange = "IDEALPRO"
        req_id = self.req_ids['exchange_rate']
        self.reqMktData(req_id, contract, "", False, False, [])
        self.exchange_rate_received.wait(timeout=10)
        self.cancelMktData(req_id)  # Cancel the request to stop receiving updates

    def get_total_deposits(self):
        # get total ILS deposits since inception from the manually updated file
        deposits_df = pd.read_excel(EXCEL_FILES_DIR + self.deposits_file)
        return deposits_df[deposits_df.Amount > 0]["Amount"].sum()  # filter out withdrawals (have negative values)

    def get_account_info(self, write_to_excel):
        # fetch account balance info from IB API
        account_balance_df = self.request_account_summary_from_api()

        # Update ILS values based on the latest exchange rate
        exchange_rate = self.get_usd_to_ils_exchange_rate()
        for tag in self.account_balance_fields:
            account_balance_df[tag]["ILS"] = round(account_balance_df[tag]["USD"] * exchange_rate)

        if write_to_excel:
            self.write_account_info_to_excel(account_balance_df, exchange_rate)

        return account_balance_df

    def write_account_info_to_excel(self, account_balance_df, exchange_rate):
        # Define the headers for the Excel file
        headers = [
            ('Date', 1),  # col 1
            ('Exchange Rate', 1),  # col 2
            ('Net Liquidation', 2),  # cols 3, 4
            ('Stock Market Value', 2),  # cols 5, 6
            ('Total Cash Balance', 2),  # cols 7, 8
            ('Net Dividend', 2),  # cols 9, 10
            ('Unrealized USD PnL', 3),  # cols 11, 12, 13
            ('Total ILS Deposits', 1),  # col 14
            ('Unrealized ILS PnL', 3)  # cols 15, 16, 17
        ]

        # Load the Excel workbook or create a new one if it doesn't exist
        try:
            workbook = openpyxl.load_workbook(EXCEL_FILES_DIR + self.account_balance_file)
            sheet = workbook.active
        except FileNotFoundError:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            self.write_headers(sheet, headers)

        usd_currency_style, nis_currency_style, percentage_style = self.define_excel_styles(workbook)

        curr_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        total_usd_base_value = account_balance_df["NetLiquidationByCurrency"]["USD"] - account_balance_df["UnrealizedPnL"]["USD"]
        total_ils_deposits = self.get_total_deposits()
        unrealized_usd_pnl_percent = account_balance_df["UnrealizedPnL"]["USD"] / total_usd_base_value
        unrealized_ils_pnl_from_deposits = account_balance_df["NetLiquidationByCurrency"]["ILS"] - total_ils_deposits
        unrealized_ils_pnl_from_deposits_usd = unrealized_ils_pnl_from_deposits * (1 / exchange_rate)
        unrealized_ils_pnl_from_deposits_percent = unrealized_ils_pnl_from_deposits / total_ils_deposits

        # Create the account information excel row
        account_balance_fields_flat = chain.from_iterable(zip(*account_balance_df.values.tolist()))
        row_data = ([curr_datetime, exchange_rate] +
                    list(account_balance_fields_flat) + [unrealized_usd_pnl_percent] +
                    [total_ils_deposits, unrealized_ils_pnl_from_deposits_usd, unrealized_ils_pnl_from_deposits,
                     unrealized_ils_pnl_from_deposits_percent])

        # Write the row the Excel file
        next_row = sheet.max_row + 1
        nis_cols = [4, 6, 8, 10, 12, 14, 16]
        usd_cols = [3, 5, 7, 9, 11, 15]
        percent_cols = [13, 17]
        for col_num, value in enumerate(row_data, start=1):
            cell = sheet.cell(row=next_row, column=col_num, value=value)
            # Apply currency styles
            if col_num in nis_cols:
                cell.style = nis_currency_style
            elif col_num in usd_cols:
                cell.style = usd_currency_style
            elif col_num in percent_cols:
                cell.style = percentage_style

        # Save the workbook
        workbook.save(EXCEL_FILES_DIR + self.account_balance_file)

    @staticmethod
    def write_headers(sheet, headers):
        # Write the headers with merged cells for sub-columns
        col = 1
        for header, span in headers:
            if span > 1:
                sheet.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + span - 1)
                sheet.cell(row=1, column=col).value = header
                sheet.cell(row=1, column=col).alignment = openpyxl.styles.Alignment(horizontal='center')
                for i, subcol_name in zip(range(span), ['USD', 'NIS', '%']):
                    sheet.cell(row=2, column=col + i).value = subcol_name
            else:
                sheet.cell(row=1, column=col).value = header
                sheet.cell(row=1, column=col).alignment = openpyxl.styles.Alignment(horizontal='center')
                sheet.cell(row=2, column=col).value = ''
            col += span

    @staticmethod
    def define_excel_styles(workbook):
        # Define currency and percent styles if they don't already exist
        styles = {"usd_currency_style": "$#,##0", "nis_currency_style": "₪#,##0", "percentage_style": "0.00%"}

        for style, style_format in styles.items():
            if style not in workbook.named_styles:
                named_style = openpyxl.styles.NamedStyle(name=style, number_format=style_format)
                workbook.add_named_style(named_style)

        return list(styles.keys())


# Example usage
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process some arguments.')
    parser.add_argument('--write_to_excel', action='store_true',
                        help='Whether to write the account info to an Excel file')
    args = parser.parse_args()

    with IBAccountInfo() as account:
        account_info = account.get_account_info(write_to_excel=args.write_to_excel)
        print(account_info)
