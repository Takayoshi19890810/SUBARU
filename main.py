import os
import json
import time
import re
import random
import requests
import unicodedata
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import gspread

KEYWORDS = ["SUBARU", "subaru", "スバル"]
SPREADSHEET_ID = "1OZNrRF_3CvbVTo-cdq2rxKBxuBNY4ZkAxiHosiMacg4"

def format_datetime(dt_obj):
    return dt_obj.strftime("%Y/%m/%d %H:%M")

def parse_relative_time(pub_label: str, base_time: datetime) -> str:
    pub_label = pub_label.strip().lower()
    try:
        if "分前" in pub_label or "minute" in pub_label:
            m = re.search(r"(\d+)", pub_label)
            if m:
                dt = base_time - timedelta(minutes=int(m.group(1)))
                return format_datetime(dt)
        elif "時間前" in pub_label or "hour" in pub_label:
            h = re.search(r"(\d+)", pub_label)
            if h:
                dt = base_time - timedelta(hours=int(h.group(1)))
                return format_datetime(dt)
        elif "日前" in pub_label or "day" in pub_label:
            d = re.search(r"(\d+)", pub_label)
            if d:
                dt = base_time - timedelta(days=int(d.group(1)))
                return format_datetime(dt)
        elif re.match(r'\d+月\d+日', pub_label):
            dt = datetime.strptime(f"{base_time.year}年{pub_label}", "%Y年%m月%d日")
            return format_datetime(dt)
        elif re.match(r'\d{4}/\d{1,2}/\d{1,2}', pub_label):
            dt = datetime.strptime(pub_label, "%Y/%m/%d")
            return format_datetime(dt)
        elif re.match(r'\d{1,2}:\d{2}', pub_label):
            t = datetime.strptime(pub_label, "%H:%M").time()
            dt = datetime.combine(base_time.date(), t)
            if dt > base_time:
                dt -= timedelta(days=1)
            return format_datetime(dt)
    except:
        pass
    return "取得不可"

def get_last_modified_datetime(url):
    try:
        response = requests.head(url, timeout=5)
        if 'Last-Modified' in response.headers:
            dt = parsedate_to_datetime(response.headers['Last-Modified'])
            jst = dt.astimezone(tz=timedelta(hours=9))
            return format_datetime(jst)
    except:
        pass
    return "取得不可"

def is_japanese(text):
    return any(
        'CJK UNIFIED' in unicodedata.name(ch, '') or
        'HIRAGANA' in unicodedata.name(ch, '') or
        'KATAKANA' in unicodedata.name(ch, '') for ch in text
    )

def get_google_news_with_selenium(keyword: str):
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    url = f"https://news.google.com/search?q={keyword}&hl=ja&gl=JP&ceid=JP:ja"
    driver.get(url)
    time.sleep(5)
    for _ in range(3):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    data = []
    for article in soup.find_all("article"):
        try:
            a_tag = article.select_one("a.JtKRv")
            time_tag = article.select_one("time.hvbAAd")
            source_tag = article.select_one("div.vr1PYe")
            title = a_tag.text.strip()
            href = a_tag.get("href")
            url = "https://news.google.com" + href[1:] if href.startswith("./") else href
            dt = datetime.strptime(time_tag.get("datetime"), "%Y-%m-%dT%H:%M:%SZ") + timedelta(hours=9)
            pub_date = format_datetime(dt)
            source = source_tag.text.strip() if source_tag else "N/A"
            data.append({"タイトル": title, "URL": url, "投稿日": pub_date, "引用元": source})
        except:
            continue
    print(f"✅ Googleニュース件数: {len(data)} 件（{keyword}）")
    return data

def get_yahoo_news_with_selenium(keyword: str):
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    search_url = f"https://news.yahoo.co.jp/search?p={keyword}&ei=utf-8&categories=domestic,world,business,it,science,life,local"
    driver.get(search_url)
    time.sleep(5)
    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    data = []
    for article in soup.find_all("li", class_=re.compile("sc-1u4589e-0")):
        try:
            title_tag = article.find("div", class_=re.compile("sc-3ls169-0"))
            link_tag = article.find("a", href=True)
            time_tag = article.find("time")
            source_tag = article.find("div", class_="sc-n3vj8g-0 yoLqH")

            title = title_tag.text.strip() if title_tag else ""
            url = link_tag["href"] if link_tag else ""
            date_str = time_tag.text.strip() if time_tag else ""
            formatted_date = date_str
            if date_str:
                try:
                    dt = datetime.strptime(re.sub(r'\([月火水木金土日]\)', '', date_str), "%Y/%m/%d %H:%M")
                    formatted_date = format_datetime(dt)
                except:
                    pass

            source_text = source_tag.text.strip() if source_tag else "Yahoo"
            if title and url:
                data.append({"タイトル": title, "URL": url, "投稿日": formatted_date, "引用元": source_text})
        except:
            continue
    print(f"✅ Yahooニュース件数: {len(data)} 件（{keyword}）")
    return data

def get_msn_news_with_selenium(keyword: str):
    now = datetime.utcnow() + timedelta(hours=9)
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    url = f"https://www.bing.com/news/search?q={keyword}&qft=sortbydate%3d'1'&form=YFNR"
    driver.get(url)
    time.sleep(5)
    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    data = []
    for card in soup.select("div.news-card"):
        try:
            title = card.get("data-title", "").strip()
            url = card.get("data-url", "").strip()
            source = card.get("data-author", "").strip()

            pub_tag = card.find("span", attrs={"aria-label": True})
            pub_label = pub_tag["aria-label"].strip() if pub_tag and pub_tag.has_attr("aria-label") else ""
            pub_date = parse_relative_time(pub_label, now)

            if pub_date == "取得不可" and url:
                pub_date = get_last_modified_datetime(url)

            if title and url and is_japanese(title):  # ✅ 日本語のみ
                data.append({"タイトル": title, "URL": url, "投稿日": pub_date, "引用元": source or "MSN"})
        except:
            continue
    print(f"✅ MSNニュース件数: {len(data)} 件（{keyword}）")
    return data

def write_to_spreadsheet(articles, spreadsheet_id, worksheet_name):
    credentials_json_str = os.environ.get('GCP_SERVICE_ACCOUNT_KEY')
    credentials = json.loads(credentials_json_str) if credentials_json_str else json.load(open('credentials.json'))
    gc = gspread.service_account_from_dict(credentials)

    for attempt in range(5):
        try:
            sh = gc.open_by_key(spreadsheet_id)
            try:
                worksheet = sh.worksheet(worksheet_name)
            except gspread.exceptions.WorksheetNotFound:
                worksheet = sh.add_worksheet(title=worksheet_name, rows="1", cols="4")
                worksheet.append_row(['タイトル', 'URL', '投稿日', '引用元'])

            existing_urls = set(row[1] for row in worksheet.get_all_values()[1:] if len(row) > 1)
            new_data = [[a['タイトル'], a['URL'], a['投稿日'], a['引用元']] for a in articles if a['URL'] not in existing_urls]
            if new_data:
                worksheet.append_rows(new_data, value_input_option='USER_ENTERED')
                print(f"✅ {len(new_data)}件をスプレッドシートに追記しました。")
            else:
                print("⚠️ 新しい記事はありません。")
            return
        except gspread.exceptions.APIError as e:
            print(f"⚠️ Google API Error ({attempt+1}/5): {e}")
            time.sleep(5 + random.random() * 5)
    raise RuntimeError("❌ スプレッドシート書き込み失敗")

if __name__ == "__main__":
    for kw in KEYWORDS:
        print(f"\n=== キーワード: {kw} ===")
        for source, fetch_func in {
            "Google": get_google_news_with_selenium,
            "Yahoo": get_yahoo_news_with_selenium,
            "MSN": get_msn_news_with_selenium
        }.items():
            print(f"\n--- {source} News ---")
            articles = fetch_func(kw)
            if articles:
                write_to_spreadsheet(articles, SPREADSHEET_ID, source)
