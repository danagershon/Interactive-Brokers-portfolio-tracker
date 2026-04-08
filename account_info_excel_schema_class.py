import pandas as pd
from typing import Any, Callable

from utils import IbApiConstants


class AccountInfoExcelSchema:

    class MainHeader:
        """
        The main headers of the account info Excel file (first row).
        """
        DATE = "Date"
        EXCHANGE_RATE = "Exchange Rate"
        TYPE = "Type"
        NET_LIQUIDATION = "Net Liquidation"
        STOCK_MARKET_VALUE = "Stock Market Value"
        TOTAL_CASH_BALANCE = "Total Cash Balance"
        NET_DIVIDEND = "Net Dividend"
        IB_UNREALIZED_USD_PNL = "IB Unrealized USD PnL"
        TOTAL_ILS_DEPOSITS = "Total ILS Deposits"
        UNREALIZED_ILS_PNL = "Unrealized ILS PnL"

    class SubHeader:
        """
        Optional subheaders for a main header (second row). 
        Relevant for monetary values main headers like Net Liquidation, Stock Market Value, Total Cash Balance, etc.
        """
        USD = IbApiConstants.Currency.USD
        ILS = IbApiConstants.Currency.ILS
        PCT = "%"

    class StyleName:
        USD_CURRENCY_STYLE = "usd_currency_style"
        ILS_CURRENCY_STYLE = "ils_currency_style"
        PERCENTAGE_STYLE = "percentage_style"

    STYLES = {
        StyleName.USD_CURRENCY_STYLE: "$#,##0",
        StyleName.ILS_CURRENCY_STYLE: "₪#,##0",
        StyleName.PERCENTAGE_STYLE: "0.00%"
    }

    SUBHEADER_TO_STYLE_NAME = {
        SubHeader.USD: StyleName.USD_CURRENCY_STYLE,
        SubHeader.ILS: StyleName.ILS_CURRENCY_STYLE,
        SubHeader.PCT: StyleName.PERCENTAGE_STYLE,
    }

    # Define the main header row (first row)
    MAIN_HEADER_ROW = [
        MainHeader.DATE,
        MainHeader.EXCHANGE_RATE,
        MainHeader.TYPE,
        MainHeader.NET_LIQUIDATION,
        MainHeader.STOCK_MARKET_VALUE,
        MainHeader.TOTAL_CASH_BALANCE,
        MainHeader.NET_DIVIDEND,
        MainHeader.IB_UNREALIZED_USD_PNL,
        MainHeader.TOTAL_ILS_DEPOSITS,
        MainHeader.UNREALIZED_ILS_PNL,
    ]

    TOTAL_VALUE_MAIN_HEADERS = [
        MainHeader.NET_LIQUIDATION,
        MainHeader.STOCK_MARKET_VALUE,
        MainHeader.TOTAL_CASH_BALANCE,
        MainHeader.NET_DIVIDEND,
    ]

    TOTAL_VALUE_SUBHEADERS = [
        SubHeader.USD,
        SubHeader.ILS,
    ]

    TOTAL_VALUE_MAIN_HEADER_TO_ACCOUNT_BALANCE_FIELD = {
        MainHeader.NET_LIQUIDATION: IbApiConstants.AccountBalanceField.NET_LIQUIDATION_BY_CURRENCY,
        MainHeader.STOCK_MARKET_VALUE: IbApiConstants.AccountBalanceField.STOCK_MARKET_VALUE,
        MainHeader.TOTAL_CASH_BALANCE: IbApiConstants.AccountBalanceField.TOTAL_CASH_BALANCE,
        MainHeader.NET_DIVIDEND: IbApiConstants.AccountBalanceField.NET_DIVIDEND,
    }

    UNREALIZED_PNL_MAIN_HEADERS = [
        MainHeader.IB_UNREALIZED_USD_PNL,
        MainHeader.UNREALIZED_ILS_PNL,
    ]

    UNREALIZED_PNL_MAIN_HEADER_TO_SOURCE_CURRENCY = {
        MainHeader.IB_UNREALIZED_USD_PNL: IbApiConstants.Currency.USD,
        MainHeader.UNREALIZED_ILS_PNL: IbApiConstants.Currency.ILS,
    }

    UNREALIZED_PNL_SUBHEADERS = [
        SubHeader.USD,
        SubHeader.ILS,
        SubHeader.PCT,
    ]

    # Define the mapping of main headers to their subheaders (second row)
    MAIN_HEADER_TO_SUBHEADERS = {
        MainHeader.DATE: None,
        MainHeader.EXCHANGE_RATE: None,
        MainHeader.TYPE: None,
        MainHeader.NET_LIQUIDATION: TOTAL_VALUE_SUBHEADERS,
        MainHeader.STOCK_MARKET_VALUE: TOTAL_VALUE_SUBHEADERS,
        MainHeader.TOTAL_CASH_BALANCE: TOTAL_VALUE_SUBHEADERS,
        MainHeader.NET_DIVIDEND: TOTAL_VALUE_SUBHEADERS,
        MainHeader.IB_UNREALIZED_USD_PNL: UNREALIZED_PNL_SUBHEADERS,
        MainHeader.TOTAL_ILS_DEPOSITS: [SubHeader.ILS],
        MainHeader.UNREALIZED_ILS_PNL: UNREALIZED_PNL_SUBHEADERS,
    }

    @staticmethod
    def get_flat_columns() -> list[tuple[str, str]]:
        """
        Flatten the schema to the physical Excel column order.

        Returns:
            list[(main_header, subheader)] where subheader is "" for single-column headers.
        """
        cols: list[tuple[str, str]] = []
        for main_header in AccountInfoExcelSchema.MAIN_HEADER_ROW:
            sub_headers = AccountInfoExcelSchema.MAIN_HEADER_TO_SUBHEADERS[main_header]
            if not sub_headers:
                cols.append((main_header, ""))
            else:
                for sub_header in sub_headers:
                    cols.append((main_header, sub_header))
        return cols

    @staticmethod
    def should_color_pnl(main_header: "AccountInfoExcelSchema.MainHeader", value: Any) -> bool:
        return main_header in AccountInfoExcelSchema.UNREALIZED_PNL_MAIN_HEADERS and isinstance(value, (float, int))

    # Signature for per-header value getters.
    # The getter returns the flattened values matching the header's subheaders.
    ValueGetter = Callable[[dict[str, Any]], list[Any]]

    @staticmethod
    def _get_total_value_row(main_header: str, df: pd.DataFrame) -> list[Any]:
        account_balance_field = AccountInfoExcelSchema.TOTAL_VALUE_MAIN_HEADER_TO_ACCOUNT_BALANCE_FIELD[main_header]
        return [df.at[account_balance_field, currency] for currency in AccountInfoExcelSchema.TOTAL_VALUE_SUBHEADERS]

    @staticmethod
    def _get_ib_unrealized_pnl_row(df: pd.DataFrame) -> list[Any]:
        """
        IB-reported unrealized PnL in USD
        """
        net_liq_usd = df.at[IbApiConstants.AccountBalanceField.NET_LIQUIDATION_BY_CURRENCY, IbApiConstants.Currency.USD]
        pnl_usd = df.at[IbApiConstants.AccountBalanceField.UNREALIZED_PNL, IbApiConstants.Currency.USD]
        pnl_ils = df.at[IbApiConstants.AccountBalanceField.UNREALIZED_PNL, IbApiConstants.Currency.ILS]
        base_investment = net_liq_usd - pnl_usd  # cost basis approximation
        pct = pnl_usd / base_investment
        return [pnl_usd, pnl_ils, pct]

    @staticmethod
    def _get_unrealized_pnl_from_deposits_row(df: pd.DataFrame, exchange_rate: float, total_ils_deposits: float) -> list[Any]:
        """
        Unrealized ILS PnL derived from deposits.
        """
        if total_ils_deposits is None:
            return [""] * len(AccountInfoExcelSchema.UNREALIZED_PNL_SUBHEADERS)

        net_liq_ils = df.at[IbApiConstants.AccountBalanceField.NET_LIQUIDATION_BY_CURRENCY, IbApiConstants.Currency.ILS]
        pnl_ils = net_liq_ils - total_ils_deposits
        pnl_usd = round(pnl_ils * (1 / exchange_rate), 2)
        pct = pnl_ils / total_ils_deposits
        return [pnl_usd, pnl_ils, pct]

    @staticmethod
    def _get_date(inputs: dict[str, Any]) -> list[Any]:
        return [inputs.get("date", "")]

    @staticmethod
    def _get_exchange_rate(inputs: dict[str, Any]) -> list[Any]:
        return [inputs.get("exchange_rate", "")]

    @staticmethod
    def _get_row_type(inputs: dict[str, Any]) -> list[Any]:
        return [inputs.get("row_type")]

    @staticmethod
    def _get_total_ils_deposits(inputs: dict[str, Any]) -> list[Any]:
        return [inputs.get("total_ils_deposits", "")]

    @staticmethod
    def _get_total_value_for_header(main_header: str) -> "AccountInfoExcelSchema.ValueGetter":
        return lambda inputs: AccountInfoExcelSchema._get_total_value_row(main_header=main_header, df=inputs["df"])

    @staticmethod
    def _get_ib_unrealized_pnl(inputs: dict[str, Any]) -> list[Any]:
        return AccountInfoExcelSchema._get_ib_unrealized_pnl_row(df=inputs["df"])

    @staticmethod
    def _get_unrealized_pnl_from_deposits(inputs: dict[str, Any]) -> list[Any]:
        return AccountInfoExcelSchema._get_unrealized_pnl_from_deposits_row(
            df=inputs["df"], 
            exchange_rate=inputs["exchange_rate"], 
            total_ils_deposits=inputs["total_ils_deposits"]
        )

    MAIN_HEADER_TO_VALUE_GETTER: dict[str, ValueGetter] = {
        MainHeader.DATE: _get_date.__func__,
        MainHeader.EXCHANGE_RATE: _get_exchange_rate.__func__,
        MainHeader.TYPE: _get_row_type.__func__,
        MainHeader.TOTAL_ILS_DEPOSITS: _get_total_ils_deposits.__func__,

        MainHeader.NET_LIQUIDATION: _get_total_value_for_header(MainHeader.NET_LIQUIDATION),
        MainHeader.STOCK_MARKET_VALUE: _get_total_value_for_header(MainHeader.STOCK_MARKET_VALUE),
        MainHeader.TOTAL_CASH_BALANCE: _get_total_value_for_header(MainHeader.TOTAL_CASH_BALANCE),
        MainHeader.NET_DIVIDEND: _get_total_value_for_header(MainHeader.NET_DIVIDEND),

        MainHeader.IB_UNREALIZED_USD_PNL: _get_ib_unrealized_pnl.__func__,
        MainHeader.UNREALIZED_ILS_PNL: _get_unrealized_pnl_from_deposits.__func__,
    }

    @staticmethod
    def get_row_values(main_header: "AccountInfoExcelSchema.MainHeader", inputs: dict[str, Any]) -> list[Any]:
        """
        Return the *flattened* value(s) to write for a main header, aligned to its subheaders.
        """
        if main_header not in AccountInfoExcelSchema.MAIN_HEADER_ROW:
            raise ValueError(f"Invalid main header: {main_header}")
        
        value_getter = AccountInfoExcelSchema.MAIN_HEADER_TO_VALUE_GETTER.get(main_header)
        if not value_getter:
            raise ValueError(f"Unhandled main header: {main_header}")

        return value_getter(inputs)
