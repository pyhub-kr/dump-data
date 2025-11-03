# 아래의 SSL 정책 낮춤을 지정하지 않으면,
# 이 오류는 서버 측 SSL 설정이 약한 Diffie-Hellman 키(‘dh key too small’) 를 사용하고 있어서,
# Python의 기본 SSL 보안 정책이 이를 거부하여 SSL Error가 발생합니다.

import datetime
import ssl
import requests
from urllib3.poolmanager import PoolManager
from requests.adapters import HTTPAdapter
import pandas as pd


class SSLAdapter(HTTPAdapter):
    def __init__(self, ssl_context: ssl.SSLContext, **kwargs):
        self._ssl_context = ssl_context
        super().__init__(**kwargs)

    def init_poolmanager(self, connections, maxsize, block=False, **pool_kwargs):
        self.poolmanager = PoolManager(num_pools=connections, maxsize=maxsize, block=block, ssl_context=self._ssl_context)

context = ssl.create_default_context()
context.set_ciphers('DEFAULT:@SECLEVEL=1')

csv_url = "https://datafile.seoul.go.kr/bigfile/iot/sheet/csv/download.do"
headers = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/26.0.1 Safari/605.1.15",
}
data = {
    "srvType": "S",
    "infId": "OA-15486",
    "serviceKind": 1,
    "pageNo": 1,
    "gridTotalCnt": 5000,
    "ssUserId": "SAMPLE_VIEW",
    "strWhere": "",
    "strOrderby": "DATE DESC",
    "filterCol": "필터선택",
    "txtFilter": "",
}

session = requests.Session()
session.mount("https://", SSLAdapter(context))
res = session.post(csv_url, data=data, headers=headers, timeout=30)
print(res.status_code)

csv_string = res.content.decode("cp949", errors="ignore")


filename = f"{datetime.datetime.now().strftime("%Y%m%d-%H%M%S")}-utf8.csv"

with open(filename, "wt", encoding="utf8") as f_out:
    f_out.write(csv_string)

df = pd.read_csv(filename).fillna("")
print(df.shape)
df.head()

print(f"writen {filename}")

