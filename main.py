# -*- coding: utf-8 -*-
import os
import json
import time
import re
import random
import argparse
import requests
from urllib.parse import urlparse
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import gspread

# =========================
# 既定（引数/環境変数で上書き可）
# =========================
DEFAULT_KEYWORD = "ホンダ"
DEFAULT_SPREADSHEET_ID = "1AwwMGKMHfduwPkrtsik40lkO1z1T8IU_yd41ku-yPi8"

# =========================
# 共通ユーティリティ
# =========================
def format_datetime(dt_obj: datetime) -> str:
    return dt_obj.strftime("%Y/%m/%d %H:%M")

TIME_RE = re.compile(r"(\d+)\s*(分|時間|日)\s*前")
TIME_ONLY_RE = re.compile(r"^\s*(\d+)\s*(分|時間|日)\s*(前|)?\s*$")

def parse_relative_time(pub_label: str, base_time: datetime) -> str:
    if not pub_label:
        return "取得不可"
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
            dt = datetime.strptime(pub_label, "%m月%d日").replace(year=base_time.year)
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

def get_last_modified_datetime(url: str) -> str:
    try:
        res = requests.head(url, timeout=5, allow_redirects=True)
        if 'Last-Modified' in res.headers:
            dt = parsedate_to_datetime(res.headers['Last-Modified'])
            jst = dt + timedelta(hours=9)
            return format_datetime(jst)
    except:
        pass
    return "取得不可"

def make_driver() -> webdriver.Chrome:
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--window-size=1280,2000")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

def resolve_final_url(url: str) -> str:
    """Google News 等の中間URLを最終URLに解決（成功時のみ置換）"""
    try:
        parsed = urlparse(url)
        if "news.google.com" in parsed.netloc:
            res = requests.get(url, timeout=6, allow_redirects=True)
            if res.ok:
                return res.url
    except:
        pass
    return url

def publisher_from_url(url: str) -> str:
    """URLのドメインから媒体名を推定（フォールバック用）"""
    try:
        netloc = urlparse(url).netloc.lower()
        if not netloc:
            return ""
        if netloc.endswith("msn.com"):
            return "MSN"
        if "news.yahoo.co.jp" in netloc:
            return "Yahoo"
        host = netloc.split(":")[0]
        if host.startswith("www."):
            host = host[4:]
        NAME_MAP = {
            "response.jp": "レスポンス（Response.jp）",
            "newsweekjapan.jp": "ニューズウィーク日本版",
            "bloomberg.co.jp": "ブルームバーグ",
            "motor-fan.jp": "Motor-Fan",
            "young-machine.com": "ヤングマシン",
            "as-web.jp": "autosport web",
            "webcg.net": "WebCG",
            "bestcarweb.jp": "ベストカーWeb",
        }
        if host in NAME_MAP:
            return NAME_MAP[host]
        parts = host.split(".")
        base = parts[-2] if len(parts) >= 2 else host
        base = base.replace("-", " ").replace("_", " ")
        return base.capitalize()
    except:
        return ""

def clean_source_text(raw: str) -> str:
    """
    'Merkmal（メルクマール） 1 時間' → 'Merkmal（メルクマール）'
    'MSN による配信 1 分' → ''（= 後段でURLから推定）
    単独の '4 日' '7 時間' → ''（= 後段でURLから推定）
    """
    if not raw:
        return ""
    t = raw.strip()
    t = re.sub(r"\bon\s+MSN\b", "", t, flags=re.IGNORECASE)     # "on MSN"
    t = re.sub(r"MSN\s*による配信", "", t)                        # "MSN による配信"
    t = re.sub(r"(提供|配信)\s*[:：]?", "", t)                   # "提供:","配信:"
    t = t.replace("・", " ").replace("•", " ").replace("·", " ")
    t = re.sub(r"\s*\d+\s*(分|時間|日)\s*(前|)?\s*$", "", t).strip()  # 末尾の時間表現
    t = TIME_RE.sub("", t).strip()                                    # 残存の時間表現
    t = re.sub(r"\s*\(\s*\)\s*$", "", t).strip()                      # 空括弧
    t = re.sub(r"\s{2,}", " ", t).strip()
    t = t.strip("・|•|·|-–—:：").strip()
    if TIME_ONLY_RE.match(t):
        return ""
    return t

def is_timeish(text: str) -> bool:
    """時刻っぽいだけの文字列か判定"""
    if not text:
        return False
    if TIME_ONLY_RE.match(text.strip()):
        return True
    if TIME_RE.search(text):
        return True
    return False

# =========================
# 各サイトのスクレイパ
# =========================
def get_google_news_with_selenium(keyword: str) -> list[dict]:
    driver = make_driver()
    url = f"https://news.google.com/search?q={keyword}&hl=ja&gl=JP&ceid=JP:ja"
    driver.get(url)
    time.sleep(5)
    for _ in range(3):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1.2)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    data: list[dict] = []
    for article in soup.find_all("article"):
        try:
            a_tag = article.select_one("a.JtKRv")
            time_tag = article.select_one("time.hvbAAd")
            if not a_tag or not time_tag:
                continue
            title = a_tag.get_text(strip=True)
            href = a_tag.get("href") or ""
            url = "https://news.google.com" + href[1:] if href.startswith("./") else href
            final_url = resolve_final_url(url)
            guessed_source = publisher_from_url(final_url)

            source = ""
            for sel in ["div.vr1PYe", "div.UOVeFe", "a.wEwyrc"]:
                el = article.select_one(sel)
                if el:
                    source = el.get_text(strip=True)
                    break
            if not source:
                source = guessed_source or "Google"

            dt = datetime.strptime(time_tag.get("datetime"), "%Y-%m-%dT%H:%M:%SZ") + timedelta(hours=9)
            pub_date = format_datetime(dt)

            data.append({"タイトル": title, "URL": final_url, "投稿日": pub_date, "引用元": source})
        except:
            continue
    print(f"✅ Googleニュース件数: {len(data)} 件")
    return data

def get_yahoo_news_with_selenium(keyword: str) -> list[dict]:
    driver = make_driver()
    search_url = (
        f"https://news.yahoo.co.jp/search?p={keyword}"
        f"&ei=utf-8&categories=domestic,world,business,it,science,life,local"
    )
    driver.get(search_url)
    time.sleep(5)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    data: list[dict] = []
    items = soup.find_all("li", class_=re.compile("sc-1u4589e-0"))
    for li in items:
        try:
            title_tag = li.find("div", class_=re.compile("sc-3ls169-0"))
            link_tag = li.find("a", href=True)
            time_tag = li.find("time")

            title = title_tag.get_text(strip=True) if title_tag else ""
            url = link_tag["href"] if link_tag else ""
            date_str = time_tag.get_text(strip=True) if time_tag else ""

            pub_date = "取得不可"
            if date_str:
                ds = re.sub(r'\([月火水木金土日]\)', '', date_str).strip()
                try:
                    pub_date = format_datetime(datetime.strptime(ds, "%Y/%m/%d %H:%M"))
                except:
                    pub_date = ds

            source = ""
            for sel in [
                "div.sc-n3vj8g-0.yoLqH div.sc-110wjhy-8.bsEjY span",
                "div.sc-n3vj8g-0.yoLqH",
                "span",
                "div"
            ]:
                el = li.select_one(sel)
                if el:
                    txt = el.get_text(" ", strip=True)
                    txt = re.sub(r"\d{4}/\d{1,2}/\d{1,2} \d{2}:\d{2}", "", txt)
                    txt = re.sub(r"\([^)]+\)", "", txt)
                    txt = txt.strip()
                    if txt and not txt.isdigit() and any(ch.isalpha() or '\u3040' <= ch <= '\u9FFF' for ch in txt):
                        source = txt
                        break
            if not source:
                source = publisher_from_url(url) or "Yahoo"

            data.append({"タイトル": title, "URL": url, "投稿日": pub_date, "引用元": source})
        except:
            continue
    print(f"✅ Yahoo!ニュース件数: {len(data)} 件")
    return data

def get_msn_news_with_selenium(keyword: str) -> list[dict]:
    """
    MSN(Bingニュース) 強化版：
    - Cookie同意対応
    - a.title / a[data-title] 両対応
    - 周辺テキストから媒体名と相対時刻を分離抽出
    - 'on MSN' や '◯時間前' を引用元から除去
    - 取れない日時は Last-Modified で補完
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException

    now = datetime.utcnow() + timedelta(hours=9)
    driver = make_driver()
    url = (
        f"https://www.bing.com/news/search?q={keyword}"
        "&qft=sortbydate%3D%271%27"
        "&setlang=ja&cc=JP&FORM=HDRSC6"
    )
    driver.get(url)

    # Cookie同意
    try:
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "bnp_btn_accept"))
        ).click()
    except TimeoutException:
        pass

    # 記事読み込み待ち
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "a.title, a[data-title]"))
        )
    except TimeoutException:
        time.sleep(2)

    # Lazy Load対策スクロール
    for _ in range(4):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1.0)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    data: list[dict] = []
    anchors = soup.select("a.title, a[data-title]")
    for a in anchors:
        try:
            title = (a.get("data-title") or a.get_text(strip=True) or "").strip()
            href = a.get("href") or ""
            if not (title and href):
                continue

            parent = a.find_parent(["div", "li"]) or a.parent

            # 1) sourceブロックの文字列を取得
            raw_source = ""
            if parent:
                s_el = parent.select_one("div.source, span.source")
                if s_el:
                    raw_source = s_el.get_text(" ", strip=True)

            # 2) 相対時刻 or ISO datetime を拾う（投稿日用）
            label = ""
            if parent:
                for el in parent.select("[aria-label]"):
                    lab = el.get("aria-label", "").strip()
                    if TIME_RE.search(lab):
                        label = lab
                        break
            if not label and parent:
                for el in parent.select("time"):
                    t = (el.get_text(strip=True) or "").strip()
                    if TIME_RE.search(t):
                        label = t
                        break
                    if el.get("datetime"):
                        label = el.get("datetime")
                        break

            pub_date = "取得不可"
            if label:
                if "T" in label and ":" in label and label.endswith("Z"):
                    try:
                        dt = datetime.strptime(label, "%Y-%m-%dT%H:%M:%SZ") + timedelta(hours=9)
                        pub_date = format_datetime(dt)
                    except:
                        pass
                else:
                    pub_date = parse_relative_time(label, now)
            if pub_date == "取得不可":
                pub_date = get_last_modified_datetime(href)

            # 3) 引用元クリーニング → 周辺候補 → URL推定 の三段構え
            source = clean_source_text(raw_source)
            if (not source) and parent:
                for sel in ["cite", "span.provider", "div.provider", "span.source", "div.source a"]:
                    el = parent.select_one(sel)
                    if el:
                        cand = clean_source_text(el.get_text(" ", strip=True))
                        if cand and not is_timeish(cand):
                            source = cand
                            break
            if (not source) or is_timeish(source):
                source = publisher_from_url(href) or "MSN"

            data.append({"タイトル": title, "URL": href, "投稿日": pub_date, "引用元": source})
        except:
            continue

    print(f"✅ MSNニュース件数: {len(data)} 件")
    return data

# =========================
# スプレッドシート書き込み
# =========================
def write_to_spreadsheet(articles: list[dict], spreadsheet_id: str, worksheet_name: str):
    """
    既存URLと重複しないものだけ追記。シートが無ければ作成。
    認証: GCP_SERVICE_ACCOUNT_KEY（環境） or credentials.json（ローカル）
    """
    credentials_json_str = os.environ.get('GCP_SERVICE_ACCOUNT_KEY')
    credentials = json.loads(credentials_json_str) if credentials_json_str else json.load(open('credentials.json'))
    gc = gspread.service_account_from_dict(credentials)

    for attempt in range(5):
        try:
            sh = gc.open_by_key(spreadsheet_id)
            try:
                ws = sh.worksheet(worksheet_name)
            except gspread.exceptions.WorksheetNotFound:
                ws = sh.add_worksheet(title=worksheet_name, rows="1", cols="4")
                ws.append_row(['タイトル', 'URL', '投稿日', '引用元'])

            existing = ws.get_all_values()
            existing_urls = set(row[1] for row in existing[1:] if len(row) > 1)

            new_rows = [[a['タイトル'], a['URL'], a['投稿日'], a['引用元']]
                        for a in articles if a['URL'] not in existing_urls]

            if new_rows:
                ws.append_rows(new_rows, value_input_option='USER_ENTERED')
                print(f"✅ {len(new_rows)}件をスプレッドシート「{worksheet_name}」に追記しました。")
            else:
                print(f"⚠️ 追記すべき新しいデータはありません。（{worksheet_name}）")
            return
        except gspread.exceptions.APIError as e:
            print(f"⚠️ Google API Error (attempt {attempt + 1}/5): {e}")
            time.sleep(5 + random.random() * 5)

    raise RuntimeError("❌ Googleスプレッドシートへの書き込みに失敗しました（5回試行しても成功せず）")

# =========================
# 設定の解決（引数/環境/既定）
# =========================
def resolve_config() -> tuple[str, str]:
    parser = argparse.ArgumentParser()
    parser.add_argument("--keyword", type=str, default=None, help="検索キーワード（例: ホンダ）")
    parser.add_argument("--sheet", type=str, default=None, help="スプレッドシートID")
    args = parser.parse_args()

    keyword = args.keyword or os.getenv("NEWS_KEYWORD") or DEFAULT_KEYWORD
    spreadsheet_id = args.sheet or os.getenv("SPREADSHEET_ID") or DEFAULT_SPREADSHEET_ID
    print(f"🔎 キーワード: {keyword}")
    print(f"📄 SPREADSHEET_ID: {spreadsheet_id}")
    return keyword, spreadsheet_id

# =========================
# エントリポイント
# =========================
if __name__ == "__main__":
    keyword, spreadsheet_id = resolve_config()

    print("\n--- Google News ---")
    google_news_articles = get_google_news_with_selenium(keyword)
    if google_news_articles:
        write_to_spreadsheet(google_news_articles, spreadsheet_id, "Google")

    print("\n--- Yahoo! News ---")
    yahoo_news_articles = get_yahoo_news_with_selenium(keyword)
    if yahoo_news_articles:
        write_to_spreadsheet(yahoo_news_articles, spreadsheet_id, "Yahoo")

    print("\n--- MSN News ---")
    msn_news_articles = get_msn_news_with_selenium(keyword)
    if msn_news_articles:
        write_to_spreadsheet(msn_news_articles, spreadsheet_id, "MSN")
