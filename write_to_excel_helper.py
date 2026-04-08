import pandas as pd
import openpyxl
from openpyxl.styles import Font
from datetime import datetime
import logging
import pathlib
from typing import Optional, Any

from account_info_excel_schema_class import AccountInfoExcelSchema


class ExcelHelper:

    class Colors:
        RED = "FF0000"
        GREEN = "00B050"

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
        self._write_row_to_excel(
            row_type=self.SUM_OF_ACCOUNTS_ROW_TYPE, df=sum_df,
            # only the sum row will have a date, exchange rate, and total ILS deposits
            date=curr_datetime, exchange_rate=exchange_rate, total_ils_deposits=total_ils_deposits
        )

        # Write individual account data (second row and onwards)
        for account_id, account_data_df in account_info.items():
            account_name = account_id + " " + account_desc[account_id]
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
            self._write_headers()

    def _write_row_to_excel(self, row_type: str, df: pd.DataFrame, date: str=None, exchange_rate: float=None, 
                            total_ils_deposits: float=None):
        """
        Write a single row of account data to the Excel file.

        Args:
            row_type: str - the type of the row (Sum of Accounts or Specific Account)
            df: pd.DataFrame - the DataFrame for the row (contains the account balance data)
            date: str - the current datetime
            exchange_rate: float - the USD to ILS exchange rate
            total_ils_deposits: float - the total ILS deposits of the master account
        """
        # Get the next available row in the sheet
        next_row = self.sheet.max_row + 1

        row_data = []
        inputs = {
            "df": df,
            "row_type": row_type,
            "date": date,
            "exchange_rate": exchange_rate,
            "total_ils_deposits": total_ils_deposits,
        }
        for main_header in AccountInfoExcelSchema.MAIN_HEADER_ROW:
            row_value = AccountInfoExcelSchema.get_row_values(main_header=main_header, inputs=inputs)
            row_data.extend(row_value)

        # Write the row data into the Excel sheet
        flat_cols = AccountInfoExcelSchema.get_flat_columns()
        for col_num, (value, (main_header, subheader)) in enumerate(zip(row_data, flat_cols), start=1):
            cell = self.sheet.cell(row=next_row, column=col_num, value=value)
            self._apply_cell_style(cell=cell, value=value, main_header=main_header, subheader=subheader)

    def _write_headers(self):
        """
        Write a two-row header, with merged cells for repeated groups.
        """
        group_start_col = 1
        curr_group: Optional[str] = None
        flat_cols = AccountInfoExcelSchema.get_flat_columns()

        def flush_group(end_col: int):
            nonlocal group_start_col, curr_group
            if curr_group is None:
                return
            if end_col > group_start_col:
                self.sheet.merge_cells(
                    start_row=1,
                    start_column=group_start_col,
                    end_row=1,
                    end_column=end_col,
                )
            self.sheet.cell(row=1, column=group_start_col).value = curr_group
            self.sheet.cell(row=1, column=group_start_col).alignment = openpyxl.styles.Alignment(horizontal="center")
            curr_group = None

        for idx, (header_group, subheader) in enumerate(flat_cols, start=1):
            if header_group != curr_group:
                flush_group(idx - 1)
                curr_group = header_group
                group_start_col = idx
            self.sheet.cell(row=2, column=idx).value = subheader

        flush_group(len(flat_cols))

    def _define_excel_styles(self):
        """
        Define currency and percent styles if they don't already exist
        """
        for style, style_format in AccountInfoExcelSchema.STYLES.items():
            if style not in self.workbook.named_styles:
                named_style = openpyxl.styles.NamedStyle(name=style, number_format=style_format)
                self.workbook.add_named_style(named_style)

    def _apply_cell_style(self, cell, value: Any, main_header: str, subheader: str):
        style_name = AccountInfoExcelSchema.SUBHEADER_TO_STYLE_NAME.get(subheader)
        if style_name:
            cell.style = style_name

        if AccountInfoExcelSchema.should_color_pnl(main_header, value):
            cell.font = Font(color=self.Colors.RED) if value < 0 else Font(color=self.Colors.GREEN)
