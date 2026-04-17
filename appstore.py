# Google Colab / Python 3
# Apple App Store reviews -> Excel
# Append only new reviews, no duplicates
# Save to local Colab, auto-download, or Google Drive

import sys
import subprocess
import importlib
import hashlib
import re
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

REQUIRED_PACKAGES = ["requests", "pandas", "openpyxl"]

for pkg in REQUIRED_PACKAGES:
    try:
        importlib.import_module(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", pkg])

import requests
import pandas as pd

from google.colab import files
from google.colab import drive


DEFAULT_FILENAME = "appstore_reviews.xlsx"
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


def log(message: str) -> None:
    print(message)


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
        raise AppStoreScraperError(f"Application with Apple ID {app_id} not found.")

    app = results[0]
    found_id = app.get("trackId")
    app_name = app.get("trackName") or "Unknown app"

    if not found_id:
        raise AppStoreScraperError("Apple ID lookup returned invalid data.")

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
        raise AppStoreScraperError(f'Application "{app_name_query}" not found.')

    ranked = sorted(results, key=lambda x: score_search_result(app_name_query, x), reverse=True)
    best = ranked[0]

    app_id = best.get("trackId")
    app_name = best.get("trackName") or "Unknown app"

    if not app_id:
        raise AppStoreScraperError("Search returned invalid app data.")

    return int(app_id), app_name


def resolve_app(app_input: str, country: str = "us") -> Tuple[int, str]:
    app_input = app_input.strip()
    if not app_input:
        raise AppStoreScraperError("Application identifier is empty.")

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
                log(f"Reviews found in storefront: {country}")
                return reviews, country

            if len(reviews) > len(best_reviews):
                best_reviews = reviews
                best_country = country
        except Exception:
            continue

    return best_reviews, best_country


def load_existing_reviews(file_path: str) -> pd.DataFrame:
    path = Path(file_path)
    if not path.exists():
        return pd.DataFrame(columns=["date", "author", "rating", "text", "_uid"])

    try:
        df = pd.read_excel(file_path, sheet_name=SHEET_NAME, engine="openpyxl")
    except Exception as e:
        raise AppStoreScraperError(f"Failed to read existing Excel file: {e}")

    expected_cols = ["date", "author", "rating", "text"]
    for col in expected_cols:
        if col not in df.columns:
            df[col] = ""

    df = df[expected_cols].copy()
    df["date"] = df["date"].astype(str).map(normalize_text)
    df["author"] = df["author"].astype(str).map(normalize_text)
    df["text"] = df["text"].astype(str).map(normalize_text)
    df["rating"] = pd.to_numeric(df["rating"], errors="coerce").fillna(0).astype(int)
    df["_uid"] = df.apply(lambda row: make_review_uid(row["date"], row["author"], row["text"]), axis=1)
    return df


def append_only_new_reviews(existing_df: pd.DataFrame, new_reviews: List[Dict[str, str]]) -> Tuple[pd.DataFrame, int]:
    if not new_reviews:
        result = existing_df.copy()
        if "_uid" not in result.columns:
            result["_uid"] = result.apply(lambda row: make_review_uid(row["date"], row["author"], row["text"]), axis=1)
        return result, 0

    new_df = pd.DataFrame(new_reviews, columns=["date", "author", "rating", "text"])
    new_df["date"] = new_df["date"].astype(str).map(normalize_text)
    new_df["author"] = new_df["author"].astype(str).map(normalize_text)
    new_df["text"] = new_df["text"].astype(str).map(normalize_text)
    new_df["rating"] = pd.to_numeric(new_df["rating"], errors="coerce").fillna(0).astype(int)
    new_df["_uid"] = new_df.apply(lambda row: make_review_uid(row["date"], row["author"], row["text"]), axis=1)

    existing_uids = set(existing_df["_uid"].tolist()) if not existing_df.empty else set()
    only_new_df = new_df[~new_df["_uid"].isin(existing_uids)].copy()

    if existing_df.empty:
        combined = only_new_df.copy()
    else:
        combined = pd.concat([existing_df, only_new_df], ignore_index=True)

    combined = combined.drop_duplicates(subset="_uid", keep="first").copy()
    combined["_sort_date"] = pd.to_datetime(combined["date"], errors="coerce", utc=True)
    combined = combined.sort_values(by="_sort_date", ascending=False, na_position="last").drop(columns=["_sort_date"])

    return combined, len(only_new_df)


def save_reviews_to_excel(df: pd.DataFrame, file_path: str) -> None:
    output_df = df[["date", "author", "rating", "text"]].copy()

    try:
        with pd.ExcelWriter(file_path, engine="openpyxl", mode="w") as writer:
            output_df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
    except Exception as e:
        raise AppStoreScraperError(f"Failed to write Excel file: {e}")


def ask_save_mode() -> Tuple[str, str]:
    print("\nChoose where to save the Excel file:")
    print("1 - Save in Colab and immediately download to this computer")
    print("2 - Save to Google Drive folder")
    print("3 - Save only in /content")

    choice = input("Enter 1, 2, or 3: ").strip()

    if choice not in {"1", "2", "3"}:
        raise AppStoreScraperError("Invalid save option. Enter 1, 2, or 3.")

    filename = input(f"Enter file name [{DEFAULT_FILENAME}]: ").strip()
    if not filename:
        filename = DEFAULT_FILENAME
    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"

    return choice, filename


def get_output_path(save_mode: str, filename: str) -> str:
    if save_mode == "1":
        return str(Path("/content") / filename)

    if save_mode == "3":
        return str(Path("/content") / filename)

    if save_mode == "2":
        log("Mounting Google Drive...")
        drive.mount("/content/drive")

        folder = input(
            "Enter Google Drive folder path relative to MyDrive "
            "(example: Reviews/AppStore): "
        ).strip()

        base_dir = Path("/content/drive/MyDrive")
        target_dir = base_dir / folder if folder else base_dir
        target_dir.mkdir(parents=True, exist_ok=True)

        return str(target_dir / filename)

    raise AppStoreScraperError("Unknown save mode.")


def get_user_inputs() -> Tuple[str, int, str]:
    app_input = input("Enter Apple App ID, App Store link, or app name: ").strip()
    limit_input = input("Enter the number of latest reviews to load: ").strip()

    if not app_input:
        raise AppStoreScraperError("Application identifier is required.")

    if not limit_input.isdigit():
        raise AppStoreScraperError("The number of reviews must be a positive integer.")

    limit = int(limit_input)
    if limit <= 0:
        raise AppStoreScraperError("The number of reviews must be greater than 0.")

    save_mode, filename = ask_save_mode()
    output_path = get_output_path(save_mode, filename)

    return app_input, limit, output_path, save_mode


def main():
    try:
        app_input, limit, output_file, save_mode = get_user_inputs()

        app_id, app_name = resolve_app(app_input, country="us")
        log(f"Application found: {app_name} (Apple ID: {app_id})")

        reviews, found_country = fetch_reviews_with_country_fallback(app_id=app_id, limit=limit)

        if not reviews:
            raise AppStoreScraperError(
                "No reviews received for this application. "
                "Possible reasons: no public reviews in tested storefronts, "
                "RSS feed unavailable for this app, or the app has no public reviews."
            )

        if found_country:
            log(f"Storefront used for reviews: {found_country}")

        log(f"Reviews loaded from App Store: {len(reviews)}")

        existing_df = load_existing_reviews(output_file)
        combined_df, added_count = append_only_new_reviews(existing_df, reviews)

        save_reviews_to_excel(combined_df, output_file)

        log(f"New reviews added: {added_count}")
        log(f"Total reviews in file: {len(combined_df)}")
        log(f"Saved file: {Path(output_file).resolve()}")

        if save_mode == "1":
            log("Starting file download to your computer...")
            files.download(output_file)

    except AppStoreScraperError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")


main()