from ibapi.client import EClient
from ibapi.wrapper import EWrapper
import threading
import time
import logging
from utils import IbApiConstants


class IBConnector(EWrapper, EClient):
    LOCALHOST = "127.0.0.1"

    def __init__(self, host=LOCALHOST, connect_to_IB_GW=True, client_id=IbApiConstants.CLIENT_ID):
        EClient.__init__(self, self)
        self.connection_thread = threading.Thread(target=self.run)
        self.host = host
        self.port = IbApiConstants.Ports.IB_GW_PORT if connect_to_IB_GW else IbApiConstants.Ports.TWS_PORT
        self.client_id = client_id
        self.req_ids = dict(account_summary=IbApiConstants.ReqIds.ACCOUNT_SUMMARY_REQ_ID)

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
        logging.error(f"Error: {reqId} {errorCode} {errorString}")
