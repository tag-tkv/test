# -*- coding: utf-8 -*-

import hashlib
import io
import re
import time
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests
import streamlit as st


SHEET_NAME = "reviews"
REQUEST_TIMEOUT = 20
MAX_RETRIES = 3
RETRY_BACKOFF_SECONDS = 2

COUNTRY_CANDIDATES = [
    "us", "ru", "gb", "de", "fr", "it", "es", "nl", "tr", "pl", "ua",
    "kz", "ca", "au", "br", "mx", "jp", "kr", "in"
]


class AppStoreScraperError(Exception):
    pass


def safe_request_json(url: str, params: Optional[dict] = None) -> dict:
    last_error = None
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
        ),
        "Accept": "application/json",
    }

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = requests.get(
                url,
                params=params,
                timeout=REQUEST_TIMEOUT,
                headers=headers,
            )

            if response.status_code >= 500:
                raise AppStoreScraperError(f"HTTP {response.status_code} from server")

            if response.status_code != 200:
                raise AppStoreScraperError(f"HTTP {response.status_code}")

            if not response.text.strip():
                raise AppStoreScraperError("Empty response")

            return response.json()

        except (requests.RequestException, ValueError, AppStoreScraperError) as e:
            last_error = e
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_BACKOFF_SECONDS * attempt)
            else:
                break

    raise AppStoreScraperError(f"Network/API error: {last_error}")


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip())


def make_review_uid(date: str, author: str, text: str) -> str:
    raw = f"{date}|{author}|{text}".encode("utf-8")
    return hashlib.sha1(raw).hexdigest()


def extract_app_id_from_input(app_input: str) -> Optional[int]:
    app_input = app_input.strip()

    if re.fullmatch(r"\d{5,20}", app_input):
        return int(app_input)

    match = re.search(r"/id(\d{5,20})", app_input)
    if match:
        return int(match.group(1))

    match = re.search(r"\bid=(\d{5,20})\b", app_input)
    if match:
        return int(match.group(1))

    return None


def lookup_app_by_id(app_id: int, country: str = "us") -> Tuple[int, str]:
    url = "https://itunes.apple.com/lookup"
    params = {
        "id": app_id,
        "country": country,
        "entity": "software",
    }
    data = safe_request_json(url, params=params)

    results = data.get("results", [])
    if not results:
        raise AppStoreScraperError(f"Приложение с Apple ID {app_id} не найдено.")

    app = results[0]
    found_id = app.get("trackId")
    app_name = app.get("trackName") or "Unknown app"

    if not found_id:
        raise AppStoreScraperError("Apple ID lookup вернул некорректные данные.")

    return int(found_id), app_name


def score_search_result(query: str, item: dict) -> int:
    query_norm = normalize_text(query).lower()
    track_name = normalize_text(str(item.get("trackName", ""))).lower()
    seller_name = normalize_text(str(item.get("sellerName", ""))).lower()
    bundle_id = normalize_text(str(item.get("bundleId", ""))).lower()

    score = 0

    if track_name == query_norm:
        score += 100
    if query_norm in track_name:
        score += 50
    if track_name.startswith(query_norm):
        score += 20
    if query_norm in seller_name:
        score += 5
    if query_norm in bundle_id:
        score += 5

    user_rating_count = item.get("userRatingCount", 0) or 0
    score += min(int(user_rating_count), 1000) // 20

    return score


def search_app_by_name(app_name_query: str, country: str = "us") -> Tuple[int, str]:
    url = "https://itunes.apple.com/search"
    params = {
        "term": app_name_query,
        "country": country,
        "entity": "software",
        "limit": 10,
    }
    data = safe_request_json(url, params=params)

    results = data.get("results", [])
    if not results:
        raise AppStoreScraperError(f'Приложение "{app_name_query}" не найдено.')

    ranked = sorted(results, key=lambda x: score_search_result(app_name_query, x), reverse=True)
    best = ranked[0]

    app_id = best.get("trackId")
    app_name = best.get("trackName") or "Unknown app"

    if not app_id:
        raise AppStoreScraperError("Поиск вернул некорректные данные приложения.")

    return int(app_id), app_name


def resolve_app(app_input: str, country: str = "us") -> Tuple[int, str]:
    app_input = app_input.strip()
    if not app_input:
        raise AppStoreScraperError("Укажи Apple ID, ссылку App Store или название приложения.")

    app_id = extract_app_id_from_input(app_input)
    if app_id is not None:
        return lookup_app_by_id(app_id, country=country)

    return search_app_by_name(app_input, country=country)


def extract_rating(entry: dict) -> Optional[int]:
    rating_obj = entry.get("im:rating", {})
    if isinstance(rating_obj, dict):
        label = rating_obj.get("label")
        if label is not None:
            try:
                return int(label)
            except ValueError:
                return None
    return None


def extract_author(entry: dict) -> str:
    author = entry.get("author", {})
    if isinstance(author, dict):
        name = author.get("name", {})
        if isinstance(name, dict):
            return normalize_text(str(name.get("label", "")))
    return ""


def extract_text(entry: dict) -> str:
    content = entry.get("content", {})
    if isinstance(content, dict):
        return normalize_text(str(content.get("label", "")))
    return ""
    

def extract_date(entry: dict) -> str:
    updated = entry.get("updated", {})
    if isinstance(updated, dict):
        return normalize_text(str(updated.get("label", "")))
    return ""


def parse_reviews_from_feed(data: dict) -> List[Dict[str, str]]:
    feed = data.get("feed", {})
    entries = feed.get("entry", [])

    if not entries:
        return []

    if isinstance(entries, dict):
        entries = [entries]

    reviews = []

    for entry in entries:
        rating = extract_rating(entry)
        if rating is None:
            continue

        date_value = extract_date(entry)
        author = extract_author(entry)
        text = extract_text(entry)

        if not (date_value and author and text):
            continue

        reviews.append(
            {
                "date": date_value,
                "author": author,
                "rating": rating,
                "text": text,
            }
        )

    return reviews


def build_review_urls(app_id: int, country: str, page: int) -> List[str]:
    return [
        f"https://itunes.apple.com/{country}/rss/customerreviews/page={page}/id={app_id}/sortby=mostrecent/json",
        f"https://itunes.apple.com/rss/customerreviews/page={page}/id={app_id}/sortby=mostrecent/json?cc={country}",
        f"https://itunes.apple.com/{country}/rss/customerreviews/id={app_id}/sortby=mostrecent/json?page={page}",
        f"https://itunes.apple.com/rss/customerreviews/id={app_id}/sortby=mostrecent/json?cc={country}&page={page}",
    ]


def fetch_reviews_for_country(app_id: int, limit: int, country: str) -> List[Dict[str, str]]:
    collected: List[Dict[str, str]] = []
    seen_uids = set()
    page = 1
    consecutive_empty_pages = 0

    while len(collected) < limit and page <= 20:
        urls = build_review_urls(app_id, country, page)
        page_reviews = []

        for url in urls:
            try:
                data = safe_request_json(url)
                page_reviews = parse_reviews_from_feed(data)
                if page_reviews:
                    break
            except Exception:
                continue

        if not page_reviews:
            consecutive_empty_pages += 1
            if consecutive_empty_pages >= 2:
                break
            page += 1
            continue

        consecutive_empty_pages = 0
        new_on_page = 0

        for review in page_reviews:
            uid = make_review_uid(review["date"], review["author"], review["text"])
            if uid in seen_uids:
                continue
            seen_uids.add(uid)
            collected.append(review)
            new_on_page += 1

            if len(collected) >= limit:
                break

        if new_on_page == 0:
            break

        page += 1

    return collected[:limit]


def fetch_reviews_with_country_fallback(app_id: int, limit: int) -> Tuple[List[Dict[str, str]], Optional[str]]:
    best_reviews = []
    best_country = None

    for country in COUNTRY_CANDIDATES:
        try:
            reviews = fetch_reviews_for_country(app_id, limit, country)
            if reviews:
                return reviews, country

            if len(reviews) > len(best_reviews):
                best_reviews = reviews
                best_country = country
        except Exception:
            continue

    return best_reviews, best_country


def prepare_dataframe(reviews: List[Dict[str, str]]) -> pd.DataFrame:
    if not reviews:
        return pd.DataFrame(columns=["date", "author", "rating", "text"])

    df = pd.DataFrame(reviews, columns=["date", "author", "rating", "text"])
    df["date"] = df["date"].astype(str).map(normalize_text)
    df["author"] = df["author"].astype(str).map(normalize_text)
    df["text"] = df["text"].astype(str).map(normalize_text)
    df["rating"] = pd.to_numeric(df["rating"], errors="coerce").fillna(0).astype(int)
    df["_uid"] = df.apply(lambda row: make_review_uid(row["date"], row["author"], row["text"]), axis=1)
    df = df.drop_duplicates(subset="_uid", keep="first").copy()

    df["_sort_date"] = pd.to_datetime(df["date"], errors="coerce", utc=True)
    df = df.sort_values(by="_sort_date", ascending=False, na_position="last")
    df = df.drop(columns=["_uid", "_sort_date"])

    return df.reset_index(drop=True)


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
    output.seek(0)
    return output.getvalue()


st.set_page_config(page_title="App Store Reviews Scraper", layout="wide")

st.title("Сбор отзывов из Apple App Store")
st.caption("Streamlit-версия без Colab-логики")

with st.form("scraper_form"):
    app_input = st.text_input(
        "Apple ID, ссылка App Store или название приложения",
        placeholder="Например: 6479202680 или https://apps.apple.com/... или ChatGPT"
    )
    limit = st.number_input(
        "Количество последних отзывов",
        min_value=1,
        max_value=5000,
        value=100,
        step=1
    )
    submitted = st.form_submit_button("Собрать отзывы")

if submitted:
    try:
        with st.spinner("Ищу приложение и загружаю отзывы..."):
            app_id, app_name = resolve_app(app_input, country="us")
            reviews, found_country = fetch_reviews_with_country_fallback(app_id, int(limit))

        if not reviews:
            st.error(
                "Отзывы не получены. Возможные причины: у приложения нет публичных отзывов, "
                "RSS-лента недоступна для этого приложения или отзывы отсутствуют в проверенных storefront."
            )
        else:
            df = prepare_dataframe(reviews)
            excel_bytes = dataframe_to_excel_bytes(df)

            st.success("Готово")
            st.write(f"**Приложение:** {app_name}")
            st.write(f"**Apple ID:** {app_id}")
            st.write(f"**Storefront:** {found_country or 'не определён'}")
            st.write(f"**Загружено отзывов:** {len(df)}")

            st.dataframe(df, use_container_width=True)

            st.download_button(
                label="Скачать Excel",
                data=excel_bytes,
                file_name="appstore_reviews.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except AppStoreScraperError as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"Непредвиденная ошибка: {e}")
