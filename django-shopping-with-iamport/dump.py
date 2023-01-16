import os
import json
from pprint import pprint
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from tqdm import tqdm


def get_props_list(query):
    url = "https://search.shopping.naver.com/search/all"
    params = {
        "query": query,
        "frm": "NVSHATC",
    }

    res = requests.get(url, params=params)
    html = res.text
    soup = BeautifulSoup(html, 'html.parser')

    json_string = soup.select_one('#__NEXT_DATA__').text
    item_list = json.loads(json_string)["props"]["pageProps"]["initialState"]["products"]["list"]

    props_list = []
    for row in item_list:
        item = row['item']

        category_name = ""
        for key in ('category4Name', 'category3Name', 'category2Name', 'category1Name'):
            if key in item:
                category_name = item[key]
                break

        desc = ""
        for key in ("smryReview",):
            if key in item:
                desc = (desc + "\n" + item[key]).strip()

        photo_url = item["imageUrl"]

        props_list.append({
            "category_name": category_name,
            "name": item["productName"],  # productTitle
            "price": int(item["price"]),
            "priceUnit": item["priceUnit"],
            "photo_url": photo_url,
            "desc": desc,
        })
    
    return props_list


def main():
    query_list = ["재킷", "점퍼", "아우터", "슬랙스", "니트", "원피스", "스커트", "후드"]

    print("Crawling meta ...")

    props_list = []
    for query in tqdm(query_list):
        props_list.extend(get_props_list(query))

    print("Crawling images ...")

    for idx, props in tqdm(enumerate(props_list)):
        photo_url = props["photo_url"]
        photo_data = requests.get(photo_url).content
        path = Path(f"images/{idx}.jpg")
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("wb") as f:
            f.write(photo_data)

        del props["photo_url"]
        props["photo_path"] = str(path)

    print('Write to file ...')
    with open("product-list.json", "wt") as f:
        json_string = json.dumps(props_list, indent=4, ensure_ascii=False)
        f.write(json_string)


if __name__ == "__main__":
    main()

