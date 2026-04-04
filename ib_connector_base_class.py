from ibapi.client import EClient
from ibapi.wrapper import EWrapper
import threading
import time


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
