import pandas as pd
import openpyxl
from openpyxl.styles import Font
from datetime import datetime
import logging
import pathlib
import utils


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
        ('Unrealized USD PnL', 3)  # cols 19, 20, 21
    ]

    def __init__(self, excel_files_dir, account_balance_file, deposits_file, account_desc):
        self.excel_files_dir = excel_files_dir
        self.account_balance_file = account_balance_file
        self.deposits_file = deposits_file
        self.account_desc = account_desc
        self.path = pathlib.Path(self.excel_files_dir) / self.account_balance_file
        self.workbook = self.sheet = None

    def load_or_create_workbook(self):
        """Load the Excel workbook or create a new one if it doesn't exist"""
        try:
            logging.debug(f"Loading workbook from: {self.path}")
            self.workbook = openpyxl.load_workbook(self.path)
            self.sheet = self.workbook.active

            # Check if there is already data in the file (beyond headers)
            if self.sheet.max_row > 2:
                # Explicitly append a blank row by writing an empty string to each column
                self.sheet.append(["" for _ in range(self.sheet.max_column)])

        except FileNotFoundError:
            self.workbook = openpyxl.Workbook()
            self.sheet = self.workbook.active
            utils.ExcelHelper.write_headers(self.sheet, self.HEADERS)

    def write_account_info_to_excel(self, sum_df, account_info, exchange_rate, total_ils_deposits):
        """
        Write account info into an Excel file. Three rows will be written: sum of accounts and individual account details.
        """
        self.load_or_create_workbook()

        # Define the styles in correct order
        usd_currency_style, nis_currency_style, percentage_style = self.define_excel_styles()

        # Get current date and time
        curr_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Write the sum of all accounts (first row)
        self.write_row_to_excel(curr_datetime, exchange_rate, "Sum of Accounts", sum_df, usd_currency_style,
                                nis_currency_style, percentage_style, total_ils_deposits=total_ils_deposits)

        # Write individual account data (second and third rows)
        for account, df in account_info.items():
            account_name = account + " " + self.account_desc[account]
            self.write_row_to_excel(None, None, account_name, df, usd_currency_style,
                                    nis_currency_style, percentage_style)

        # Save the workbook
        self.workbook.save(self.path)

    def write_row_to_excel(self, date, exchange_rate, row_type, df, usd_style, nis_style, percentage_style,
                           total_ils_deposits=None, total_usd_deposits=None):
        """
        Write a single row of account data to the Excel file without currency symbols in values.
        Apply currency formatting using Excel styles. Only the first row (Sum of Accounts) should display 'Total ILS Deposits'.
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
            row_data.append(df.loc[tag, "USD"])  # Add USD value first
            row_data.append(df.loc[tag, "ILS"])  # Then add ILS value

        # Perform Unrealized PnL calculations for each account
        total_usd_base_value = df.loc["NetLiquidationByCurrency", "USD"] - df.loc["UnrealizedPnL", "USD"]
        unrealized_usd_pnl_percent = df.loc["UnrealizedPnL", "USD"] / total_usd_base_value

        # Add Unrealized PnL and ILS PnL fields to the row data

        # Add Unrealized USD PnL %
        row_data.append(unrealized_usd_pnl_percent)

        for total_deposits in [total_ils_deposits, total_usd_deposits]:
            if total_deposits is not None:
                # only the sum row will display ILS/USD PnL
                currency = "ILS" if total_deposits is total_ils_deposits else "USD"
                unrealized_pnl_from_deposits = df.loc["NetLiquidationByCurrency", currency] - total_deposits
                unrealized_pnl_from_deposits_percent = unrealized_pnl_from_deposits / total_deposits
                unrealized_pnl_from_deposits_in_other_currency = unrealized_pnl_from_deposits * \
                                                                 ((1 / exchange_rate) if currency == "ILS" else exchange_rate)
            else:
                unrealized_pnl_from_deposits = ""
                unrealized_pnl_from_deposits_percent = ""
                unrealized_pnl_from_deposits_in_other_currency = ""

            # Add the Total ILS Deposits only for the sum row
            if total_deposits is not None:
                row_data.append(total_deposits)
            else:
                row_data.append("")  # Leave it blank for individual accounts
    
            # Add Unrealized ILS/USD PnL fields
            row_data.append(unrealized_pnl_from_deposits)
            row_data.append(unrealized_pnl_from_deposits_percent)
            row_data.append(unrealized_pnl_from_deposits_in_other_currency)

        # Write the row data into the Excel sheet
        for col_num, value in enumerate(row_data, start=1):
            cell = self.sheet.cell(row=next_row, column=col_num, value=value)
            if col_num <= 3:
                continue  # skip Date, Exchange Rate and Type cols

            subcol_type = self.sheet.cell(2, col_num).value
            subcol_type = subcol_type or "NIS"  # assume col is NIS by default

            if "%" in subcol_type:
                cell.style = percentage_style  # Apply percentage style to "Unrealized PnL %" columns
            elif "USD" in subcol_type:
                cell.style = usd_style  # Apply USD style for USD columns
            elif "NIS" in subcol_type:
                cell.style = nis_style  # Apply NIS style for ILS columns

            # Apply red or green font color for Unrealized PnL columns (USD and ILS)
            col_header = self.sheet.cell(1, col_num).value
            if not col_header:
                col_header = self.sheet.cell(1, col_num - 1).value  # we are at NIS subcol of PnL
            if not col_header:
                col_header = self.sheet.cell(1, col_num - 2).value  # we are at % subcol of PnL
            if col_header and ("Unrealized USD PnL" in col_header or "Unrealized ILS PnL" in col_header):
                if isinstance(value, (float, int)) and value != "":
                    if value < 0:
                        cell.font = Font(color="FF0000")  # Red color for negative values
                    else:
                        cell.font = Font(color="00B050")  # Green color for positive values

    def write_headers(self, headers):
        # Write the headers with merged cells for sub-columns
        col = 1
        for header, span in headers:
            if span > 1:
                self.sheet.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + span - 1)
                self.sheet.cell(row=1, column=col).value = header
                self.sheet.cell(row=1, column=col).alignment = openpyxl.styles.Alignment(horizontal='center')
                for i, subcol_name in zip(range(span), ['USD', 'NIS', '%']):
                    self.sheet.cell(row=2, column=col + i).value = subcol_name
            else:
                self.sheet.cell(row=1, column=col).value = header
                self.sheet.cell(row=1, column=col).alignment = openpyxl.styles.Alignment(horizontal='center')
                self.sheet.cell(row=2, column=col).value = ''
            col += span

    def define_excel_styles(self):
        # Define currency and percent styles if they don't already exist
        styles = {"usd_currency_style": "$#,##0", "nis_currency_style": "₪#,##0", "percentage_style": "0.00%"}

        for style, style_format in styles.items():
            if style not in self.workbook.named_styles:
                named_style = openpyxl.styles.NamedStyle(name=style, number_format=style_format)
                self.workbook.add_named_style(named_style)

        return ["usd_currency_style", "nis_currency_style", "percentage_style"]