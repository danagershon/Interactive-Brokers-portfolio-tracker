#!/usr/bin/env python3

import logging
import argparse
from ib_account_info_fetcher import IbAccountInfoFetcher


def set_logging_settings():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    # ibapi logs every request/response at INFO; keep only our app messages at INFO
    logging.getLogger("ibapi").setLevel(logging.WARNING)


def parse_arguments():
    parser = argparse.ArgumentParser(description='Process some arguments.')
    parser.add_argument('--write_to_excel', action='store_true',
                        help='Whether to write the account info to an Excel file')
    return parser.parse_args()


if __name__ == "__main__":
    set_logging_settings()
    args = parse_arguments()

    with IbAccountInfoFetcher() as ib_account_info_fetcher:
        sum_info, account_info = ib_account_info_fetcher.get_account_info(write_to_excel=args.write_to_excel)
        print("Sum of all accounts:\n", sum_info)
        print("Individual account information:\n", account_info)
