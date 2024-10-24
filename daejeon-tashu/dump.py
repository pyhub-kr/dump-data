import requests
import json
import re
import sys
import datetime

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3',
}

url = 'https://www.tashu.or.kr/main.do'
response = requests.get(url, headers=headers)
response.raise_for_status()

# Use regular expression to find the script tag containing the JSON data
pattern = re.compile(r"var station_json = JSON\.parse\('([^']+)'\);")
match = pattern.search(response.text)

if match:
    json_text = match.group(1)
    station_data = json.loads(json_text)

    filename = datetime.datetime.now().strftime("daejeon-tashu-%Y%m%d.json")
    with open(filename, "w") as f:
        json.dump(station_data, f, indent=4, ensure_ascii=False)
    print(f"Saved to {filename}")
else:
    print("JSON data not found", file=sys.stderr)
