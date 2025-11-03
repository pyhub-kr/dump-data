# 아래의 SSL 정책 낮춤을 지정하지 않으면,
# 이 오류는 서버 측 SSL 설정이 약한 Diffie-Hellman 키(‘dh key too small’) 를 사용하고 있어서,
# Python의 기본 SSL 보안 정책이 이를 거부하여 SSL Error가 발생합니다.

import csv
import datetime
import io
import json
import ssl
import requests
from urllib3.poolmanager import PoolManager
from requests.adapters import HTTPAdapter


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


filename = f"{datetime.datetime.now().strftime("%Y%m%d-%H%M%S")}-utf8.jsonl"

# CSV 문자열을 파싱하여 JSONL로 변환
csv_reader = csv.DictReader(io.StringIO(csv_string))
records = list(csv_reader)

# 제외할 컬럼 목록
exclude_columns = {"대표이미지", "신청일", "시민/기관", "문화포털상세URL"}

# JSONL 파일로 저장 (불필요한 컬럼 제외)
with open(filename, "wt", encoding="utf8") as f_out:
    for record in records:
        # 제외할 컬럼 삭제
        filtered_record = {k: v for k, v in record.items() if k not in exclude_columns}
        json.dump(filtered_record, f_out, ensure_ascii=False)
        f_out.write("\n")

# JSONL 파일 읽기 및 검증
with open(filename, "rt", encoding="utf8") as f_in:
    lines = f_in.readlines()
    print(f"Total records: {len(lines)}")

    # 첫 5개 레코드 출력
    for i, line in enumerate(lines[:5]):
        record = json.loads(line)
        if i == 0:
            print(f"Fields: {list(record.keys())}")
        print(f"Record {i+1}: {record}")

print(f"Written {filename}")

