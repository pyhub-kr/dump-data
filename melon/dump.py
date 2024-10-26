import argparse
import datetime
import re
import time
from typing import Optional, List, Dict

import requests
import pandas as pd
from bs4 import BeautifulSoup


BASE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36",
}


def get_number_from_string(s: str) -> str:
    matched = re.search(r"\d+", s)
    if matched:
        return matched.group(0)
    return None


def print_today_playlist() -> List[Dict[str, str]]:
    search_url = "https://www.melon.com/dj/today/djtoday_list.htm"

    html = requests.get(search_url, headers=BASE_HEADERS).text
    soup = BeautifulSoup(html, "html.parser")

    for tag in soup.select(".none"):
        tag.extract()

    for wrap_tag in soup.select(".page_header, .rolling"):
        if "page_header" in wrap_tag.get("class"):
            print("#", wrap_tag.text.strip())
        elif "rolling" in wrap_tag.get("class"):
            for detail_tag in wrap_tag.select(".entry a[href*=goDjPlaylistDetail]"):
                playlist_id = re.findall(r"\d+", detail_tag["href"])[-1]
                print("    -", playlist_id, detail_tag.text.strip())


def extract_song_list(page_url: str, filename_fmt: str):
    html = requests.get(page_url, headers=BASE_HEADERS).text

    # HTML 응답 문자열로부터, 필요한 태그 정보를 추출하기 위해, BeautifulSoup4 객체를 생성합니다.
    soup = BeautifulSoup(html, "html.parser")

    # BeautifulSoup4 객체를 통해 노래 정보를 추출해냅니다.
    song_list = []

    for idx, song_tag in enumerate(
        soup.select("#tb_list tbody tr, #pageList tbody tr"), 1
    ):
        play_song_tag = song_tag.select_one("a[href*=playSong]")
        곡명 = play_song_tag.text
        __, 곡일련번호 = re.findall(r"\d+", play_song_tag["href"])
        곡일련번호 = int(곡일련번호)

        artist_tag = song_tag.select_one("a[href*=goArtistDetail]")
        artist_name = artist_tag.text
        artist_uid = int(get_number_from_string(artist_tag["href"]))

        album_tag = song_tag.select_one("a[href*=goAlbumDetail]")
        album_uid = int(get_number_from_string(album_tag["href"]))
        album_name = album_tag["title"]
        순위 = song_tag.select_one(".rank").text

        song_detail_url = f"https://www.melon.com/song/detail.htm?songId={곡일련번호}"
        song_headers = dict(BASE_HEADERS, Referer=page_url)
        song_html = requests.get(song_detail_url, headers=song_headers).text
        song_soup = BeautifulSoup(song_html, "html.parser")
        print(idx, 곡명, artist_name, song_detail_url)
        try:
            커버이미지_주소 = song_soup.select_one(".section_info img")["src"].split(
                "?", 1
            )[0]
        except TypeError:
            커버이미지_주소 = None

        keys = [tag.text.strip() for tag in song_soup.select(".section_info .meta dt")]
        values = [
            tag.text.strip() for tag in song_soup.select(".section_info .meta dd")
        ]
        meta_dict = dict(zip(keys, values))

        lyric_tag = song_soup.select_one(".lyric")
        if lyric_tag:
            inner_html = song_soup.select_one(".lyric").encode_contents().decode("utf8")
            inner_html = re.sub(r"<!--.*?-->", "", inner_html).strip()
            가사 = re.sub(r"<br\s*/?>", "\n", inner_html).strip()
        else:
            가사 = ""

        song = {
            "곡일련번호": 곡일련번호,
            "순위": 순위,
            "곡명": 곡명,
            "artist_uid": artist_uid,
            "artist_name": artist_name,
            "album_uid": album_uid,
            "album_name": album_name,
            # '커버이미지_썸네일_주소': 커버이미지_썸네일_주소,
            "커버이미지_주소": 커버이미지_주소,
            "가사": 가사,
            "장르": [s.strip() for s in meta_dict.get("장르", "").split(",") if s.strip()],
            "발매일": meta_dict.get("발매일", "").replace(".", "-") or None,
        }
        # print(song)

        song_list.append(song)

        time.sleep(0.05)

    # 추출해낸 곡 정보를 Pandas의 DataFrame화 시킵니다.
    song_df = pd.DataFrame(
        song_list,
        columns=[
            "순위",
            "곡일련번호",
            "album_uid",
            "album_name",
            "곡명",
            "artist_uid",
            "artist_name",
            "커버이미지_주소",
            "가사",
            "장르",
            "발매일",
        ],
    ).set_index("곡일련번호")

    # song_df의 인덱스가 노래 id 목록입니다.
    song_id_list = song_df.index

    # 노래별 "좋아요" 정보는 별도로 요청해야합니다. 노래 id 목록을 인자로 넘겨서 좋아요 정보를 획득합니다.
    url = "http://www.melon.com/commonlike/getSongLike.json"
    params = {
        "contsIds": song_id_list,
    }
    result = requests.get(url, headers=BASE_HEADERS, params=params).json()
    like_dict = {int(song["CONTSID"]): song["SUMMCNT"] for song in result["contsLike"]}

    # 좋아요 정보를 song_df에 새로운 필드로 추가합니다.
    song_df["좋아요"] = pd.Series(like_dict)

    # song_df의 상위 5개 Row만 조회합니다.
    # print(song_df.head())

    filename = datetime.datetime.now().strftime(filename_fmt)

    song_df.reset_index(drop=False).to_json(
        filename, orient="records", force_ascii=False, indent=4
    )
    print(f"created {filename}")


if __name__ == "__main__":
    sample_playlist_id = "462121500"

    parser = argparse.ArgumentParser(description="Extract song information.")
    parser.add_argument(
        "--playlist-id",
        type=str,
        help="Playlist ID to extract songs from.",
    )
    parser.add_argument(
        "--sample", action="store_true", help="Indicate if sample data should be used."
    )
    parser.add_argument(
        "--print-today-playlist",
        action="store_true",
        help="Show list of playlists.",
    )
    args = parser.parse_args()

    is_print_today_playlist: bool = args.print_today_playlist

    if is_print_today_playlist:
        print_today_playlist()
    else:
        playlist_id: Optional[str] = args.playlist_id
        use_sample: bool = args.sample

        if use_sample:
            print(
                f"샘플 데이터. 플레이리스트 {sample_playlist_id} 페이지를 추출합니다."
            )
            playlist_id = sample_playlist_id

        if playlist_id:
            page_url = f"https://www.melon.com/mymusic/dj/mymusicdjplaylistview_inform.htm?plylstSeq={playlist_id}"
            print(f"플레이리스트 #{playlist_id} 페이지를 추출합니다.")
            filename_fmt = f"melon-playlist-{playlist_id}-%Y%m%d.json"
        else:
            # 멜론 차트 페이지의 HTML 응답 문자열을 획득합니다.
            page_url = "http://www.melon.com/chart/index.htm"
            print("멜론 TOP100 차트 페이지를 추출합니다.")
            filename_fmt = "melon-%Y%m%d.json"

        print(page_url)
        extract_song_list(page_url, filename_fmt)
