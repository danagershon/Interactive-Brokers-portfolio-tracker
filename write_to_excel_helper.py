import pandas as pd
import openpyxl
from openpyxl.styles import Font
from datetime import datetime
import logging
import pathlib

import utils
from utils import IbApiConstants


class ExcelHelper:

    # Define the headers for the Excel file
    HEADERS = [
        ('Date', 1),  # col 1
        ('Exchange Rate', 1),  # col 2
        ('Type', 1),  # col 3
        ('Net Liquidation', 2),  # cols 4, 5
        ('Stock Market Value', 2),  # cols 6, 7
        ('Total Cash Balance', 2),  # cols 8, 9
        ('Net Dividend', 2),  # cols 10, 11
        ('IB Unrealized USD PnL', 3),  # cols 12, 13, 14
        ('Total ILS Deposits', 1),  # col 15 (only filled in the sum row)
        ('Unrealized ILS PnL', 3),  # cols 16, 17, 18
    ]

    class Colors:
        RED = "FF0000"
        GREEN = "00B050"

    class StyleName:
        USD_CURRENCY_STYLE = "usd_currency_style"
        ILS_CURRENCY_STYLE = "ils_currency_style"
        PERCENTAGE_STYLE = "percentage_style"

    STYLES = {
        StyleName.USD_CURRENCY_STYLE: "$#,##0",
        StyleName.ILS_CURRENCY_STYLE: "₪#,##0",
        StyleName.PERCENTAGE_STYLE: "0.00%"
    }

    SUM_OF_ACCOUNTS_ROW_TYPE = "Sum of Accounts"

    def __init__(self):
        self.workbook = None
        self.sheet = None

    def write_account_info_to_excel(self, account_info_output_file: pathlib.Path, account_desc: dict[str, str], sum_df: pd.DataFrame, 
                                    account_info: dict[str, pd.DataFrame], exchange_rate: float, total_ils_deposits: float):
        """
        Write account info into an Excel file. 
        Rows will be written for: sum of accounts and individual account details.

        Args:
            account_info_output_file: pathlib.Path - the path to the Excel file to write the account info to
            account_desc: dict[str, str] - the description of the account (mapping account ID to description)
            sum_df: pd.DataFrame - the DataFrame for the sum of accounts
            account_info: dict[str, pd.DataFrame] - the DataFrame for the individual accounts
            exchange_rate: float - the current USD to ILS exchange rate
            total_ils_deposits: float - the total ILS deposits of the master account
        """
        self._load_or_create_workbook(account_info_output_file)
        self._define_excel_styles()
        curr_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Write the sum of all accounts (first row)
        self._write_row_to_excel(row_type=self.SUM_OF_ACCOUNTS_ROW_TYPE, df=sum_df, date=curr_datetime, 
                                 exchange_rate=exchange_rate, total_ils_deposits=total_ils_deposits)

        # Write individual account data (second row and onwards)
        for account_id, account_data_df in account_info.items():
            account_name = account_id + " " + account_desc[account_id]
            # No need to write date and exchange rate for individual accounts, since they appear once in the sum row
            self._write_row_to_excel(row_type=account_name, df=account_data_df)

        # Save the workbook
        self.workbook.save(account_info_output_file)

    def _load_or_create_workbook(self, account_info_output_file: pathlib.Path):
        """Load the Excel workbook or create a new one if it doesn't exist"""
        try:
            logging.debug(f"Loading workbook from: {account_info_output_file}")
            self.workbook = openpyxl.load_workbook(account_info_output_file)
            self.sheet = self.workbook.active

            # Check if there is already data in the file (beyond headers)
            if self.sheet.max_row > 2:
                # Explicitly append a blank row by writing an empty string to each column
                self.sheet.append(["" for _ in range(self.sheet.max_column)])

        except FileNotFoundError:
            self.workbook = openpyxl.Workbook()
            self.sheet = self.workbook.active
            utils.ExcelHelper._write_headers(self.sheet, self.HEADERS)

    def _write_row_to_excel(self, row_type: str, df: pd.DataFrame, date: str = None, exchange_rate: float = None, 
                           total_ils_deposits: float = None):
        """
        Write a single row of account data to the Excel file without currency symbols in values.
        Apply currency formatting using Excel styles. Only the first row (Sum of Accounts) should display 'Total ILS Deposits'.

        Args:
            row_type: str - the type of the row (Sum of Accounts or Specific Account)
            df: pd.DataFrame - the DataFrame for the row (contains the account data)
            date: str (optional) - the current date
            exchange_rate: float (optional) - the current USD to ILS exchange rate
            total_ils_deposits: float (optional) - the total deposits of the master account in ILS
        """
        # Ensure df is converted to a DataFrame if it is a dictionary
        if isinstance(df, dict):
            df = pd.DataFrame(df).T  # Transpose to align the dictionary properly as a DataFrame

        # Get the next available row in the sheet
        next_row = self.sheet.max_row + 1

        # Collect data for the row (Date, Exchange Rate, and Row Type (Sum of Accounts or Specific Account))
        row_data = [date, exchange_rate, row_type]

        # Flatten the USD and ILS values from the DataFrame and add to row_data
        for tag in df.index:
            # add the USD value first, then the ILS value
            for currency in [IbApiConstants.Currency.USD, IbApiConstants.Currency.ILS]:
                row_data.append(df.loc[tag, currency])

        # Perform Unrealized PnL calculations
        ib_usd_net_liquidation = df.loc[IbApiConstants.AccountBalanceField.NET_LIQUIDATION_BY_CURRENCY, IbApiConstants.Currency.USD]
        ib_usd_unrealized_pnl = df.loc[IbApiConstants.AccountBalanceField.UNREALIZED_PNL, IbApiConstants.Currency.USD]
        total_usd_base_value = ib_usd_net_liquidation - ib_usd_unrealized_pnl
        unrealized_usd_pnl_percent = ib_usd_unrealized_pnl / total_usd_base_value
        # Add Unrealized USD PnL % calculated from IB data
        row_data.append(unrealized_usd_pnl_percent)

        ils_unrealized_pnl_values = []
        if total_ils_deposits is not None:
            # only the sum row will display ILS/USD PnL
            currency = IbApiConstants.Currency.ILS
            unrealized_ils_pnl_from_deposits = df.loc[IbApiConstants.AccountBalanceField.NET_LIQUIDATION_BY_CURRENCY, currency] - total_ils_deposits
            unrealized_ils_pnl_from_deposits_percent = unrealized_ils_pnl_from_deposits / total_ils_deposits
            ils_to_usd_exchange_rate = 1 / exchange_rate
            unrealized_usd_pnl_from_deposits = unrealized_ils_pnl_from_deposits * ils_to_usd_exchange_rate
            # Add the Total ILS Deposits (only for the sum row, leave blank for individual accounts)
            row_data.append(total_ils_deposits)
            ils_unrealized_pnl_values.extend([unrealized_usd_pnl_from_deposits, 
                                              unrealized_ils_pnl_from_deposits, 
                                              unrealized_ils_pnl_from_deposits_percent])
        else:
            row_data.append("")  # Leave total ILS deposits blank for individual accounts
            ils_unrealized_pnl_values.extend([""] * 3)

        # Add Unrealized ILS PnL fields
        row_data.extend(ils_unrealized_pnl_values)

        # Write the row data into the Excel sheet
        for col_num, value in enumerate(row_data, start=1):
            cell = self.sheet.cell(row=next_row, column=col_num, value=value)
            if col_num <= 3:
                continue  # skip Date, Exchange Rate and Type cols

            subcol_type = self.sheet.cell(row=2, column=col_num).value
            if "%" in subcol_type:
                cell.style = self.StyleName.PERCENTAGE_STYLE  # Apply percentage style to "Unrealized PnL %" columns
            elif IbApiConstants.Currency.USD in subcol_type:
                cell.style = self.StyleName.USD_CURRENCY_STYLE  # Apply USD style for USD columns
            elif IbApiConstants.Currency.ILS in subcol_type:
                cell.style = self.StyleName.ILS_CURRENCY_STYLE  # Apply ILS style for ILS columns

            # Apply red or green font color for Unrealized PnL columns (USD and ILS)
            col_header = self.sheet.cell(row=1, column=col_num).value
            if not col_header:  # we are at ILS or % subcol of PnL
                col_header = self.sheet.cell(row=1, column=col_num - 1).value
            if not col_header:  # we are at % subcol of PnL
                col_header = self.sheet.cell(row=1, column=col_num - 2).value
            if not col_header:
                raise ValueError(f"No col_header found for column {col_num}")
            if "Unrealized USD PnL" in col_header or "Unrealized ILS PnL" in col_header:
                if isinstance(value, (float, int)) and value != "":
                    # Red color for negative values, green for positive values
                    cell.font = Font(color=self.Colors.RED) if value < 0 else Font(color=self.Colors.GREEN)

    def _write_headers(self, headers: list[tuple[str, int]]):
        """
        Write the headers with merged cells for sub-columns
        """
        col = 1
        for header, span in headers:
            if span > 1:
                self.sheet.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + span - 1)
                self.sheet.cell(row=1, column=col).value = header
                self.sheet.cell(row=1, column=col).alignment = openpyxl.styles.Alignment(horizontal='center')
                for i, subcol_name in zip(range(span), [IbApiConstants.Currency.USD, IbApiConstants.Currency.ILS, '%']):
                    self.sheet.cell(row=2, column=col + i).value = subcol_name
            else:
                self.sheet.cell(row=1, column=col).value = header
                self.sheet.cell(row=1, column=col).alignment = openpyxl.styles.Alignment(horizontal='center')
                self.sheet.cell(row=2, column=col).value = ''
            col += span

    def _define_excel_styles(self):
        """
        Define currency and percent styles if they don't already exist
        """
        for style, style_format in self.STYLES.items():
            if style not in self.workbook.named_styles:
                named_style = openpyxl.styles.NamedStyle(name=style, number_format=style_format)
                self.workbook.add_named_style(named_style)
