import datetime
import json
import time
from itertools import count

import requests


BASE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36",
}


def main():
    # refs: https://paikdabang.com/store/
    url = "https://theborndb.theborn.co.kr/wp-json/api/get_store/"
    params = {
        "state": 9,
        "category": 275,
        "paged": 1,
        "depth1": "",
        "depth2": "",
        "search_string": "",
    }

    shop_list = []

    for page in count(start=1):
        print(".", flush=True, end="")
        params["paged"] = page
        res_obj = requests.get(url, headers=BASE_HEADERS, params=params).json()
        # res_obj["max_count"]
        if not res_obj["results"]:
            break
        shop_list.extend(res_obj["results"])
        time.sleep(0.5)
    print()

    filename = datetime.datetime.now().strftime("paikdabang-%Y%m%d.json")
    with open(filename, "wt", encoding="utf8") as f:
        json_string = json.dumps(shop_list, ensure_ascii=False, indent=4)
        f.write(json_string)

    print(f"created {filename}")


if __name__ == "__main__":
    main()
