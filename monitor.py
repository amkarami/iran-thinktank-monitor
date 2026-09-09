# -*- coding: utf-8 -*-
"""
ایجنت روزانه رصد اندیشکده‌ها برای اخبار/گزارش‌های مرتبط با ایران.

نسخه ۲: به‌جای فید اخبار گوگل‌نیوز (که فقط محتوای «تازه‌ی خبری» را نشان
می‌دهد)، از جست‌وجوی عمومی DuckDuckGo با عبارت site:دامنه Iran استفاده
می‌کند. این کار دقیقاً معادل تایپ کردن «Iran» در جست‌وجوی خودِ سایت است، اما
برای همه‌ی اندیشکده‌ها به‌صورت خودکار.

چون این روش زمان انتشار را محدود نمی‌کند، یک فایل seen_links.json در کنار
اسکریپت نگه می‌داریم که لینک‌های قبلاً دیده‌شده را ثبت می‌کند تا هر بار فقط
موارد "جدید" (که تا امروز دیده نشده بودند) ایمیل شوند.

مراحل:
  1. برای هر اندیشکده، جست‌وجوی DuckDuckGo با site:domain Iran انجام می‌شود.
  2. نتایجی که قبلاً در seen_links.json ثبت نشده‌اند "جدید" محسوب می‌شوند.
  3. اگر حداقل یک نتیجه‌ی جدید پیدا شد، یک PDF ساخته می‌شود.
  4. PDF از طریق جیمیل ایمیل می‌شود.
  5. seen_links.json به‌روزرسانی می‌شود (وورک‌فلو گیت‌هاب آن را کامیت می‌کند).

اجرا: python monitor.py
متغیرهای محیطی لازم:
  GMAIL_USER, GMAIL_APP_PASSWORD, RECIPIENT_EMAIL
"""

import os
import sys
import json
import time
import random
import re
import subprocess
import zipfile
import smtplib
import ssl
import urllib.parse
import urllib.request
from email.message import EmailMessage
from datetime import datetime, timezone

from bs4 import BeautifulSoup
from thinktanks import THINKTANKS

# ----------------------------------------------------------------------
# تنظیمات
# ----------------------------------------------------------------------
MAX_RESULTS_PER_TANK = 15    # حداکثر نتیجه‌ی جست‌وجو برای هر اندیشکده
REQUEST_TIMEOUT = 20         # ثانیه
MIN_SLEEP = 4.0              # حداقل فاصله‌ی بین درخواست‌ها (ثانیه)
MAX_SLEEP = 8.0              # حداکثر فاصله‌ی بین درخواست‌ها (تصادفی، برای شبیه‌سازی رفتار انسانی)
COOLDOWN_AFTER_BLOCK = 45.0  # اگر نشانه‌ی بلاک‌شدن دیده شود، این‌قدر صبر می‌کند

GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
RECIPIENT_EMAIL = os.environ.get("RECIPIENT_EMAIL", "am.karami@gmail.com")
BRAVE_API_KEY = os.environ.get("BRAVE_API_KEY", "")

OUTPUT_PDF = "iran_thinktank_digest.pdf"
SEEN_FILE = "seen_links.json"
ARTICLES_DIR = "articles_pdf"

# --- تنظیمات تبدیل مقاله به PDF ---
ARTICLE_PDF_TIMEOUT = 45      # حداکثر ثانیه برای تبدیل هر مقاله
WKHTMLTOPDF_BIN = "wkhtmltopdf"
MAX_ZIP_SIZE_MB = 20           # حداکثر حجم هر فایل زیپ ایمیلی (زیر محدودیت جیمیل)
MAX_ZIP_SIZE_BYTES = MAX_ZIP_SIZE_MB * 1024 * 1024

HEADERS_BROWSER_POOL = [
    {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )
    },
    {
        "User-Agent": (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 "
            "(KHTML, like Gecko) Version/17.4 Safari/605.1.15"
        )
    },
    {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
        )
    },
]

# ----------------------------------------------------------------------
# مرحله ۱: جست‌وجوی Brave Search API برای هر دامنه
# ----------------------------------------------------------------------

def fetch_brave_results(domain: str, keyword: str = "Iran", max_results: int = MAX_RESULTS_PER_TANK):
    """جست‌وجوی site:domain keyword را از طریق Brave Search API انجام می‌دهد و
    لیستی از دیکشنری‌های {title, link, snippet} برمی‌گرداند."""
    query = f"site:{domain} {keyword}"
    url = "https://api.search.brave.com/res/v1/web/search"
    params = urllib.parse.urlencode({"q": query, "count": max_results})
    full_url = f"{url}?{params}"

    headers = {
        "Accept": "application/json",
        "X-Subscription-Token": BRAVE_API_KEY,
    }
    req = urllib.request.Request(full_url, headers=headers)

    try:
        with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        print(f"  [!] خطا در جست‌وجوی {domain} با Brave API: {e}", file=sys.stderr)
        return []

    items = []
    for r in data.get("web", {}).get("results", [])[:max_results]:
        title = r.get("title", "").strip()
        link = r.get("url", "").strip()
        snippet = r.get("description", "").strip()
        if title and link:
            items.append({"title": title, "link": link, "snippet": snippet})

    return items


# ----------------------------------------------------------------------
# مرحله ۱-ب: روش پیش‌فرض و رایگان (بدون نیاز به هیچ کلید یا ثبت‌نام)
# با DuckDuckGo، با مکانیزم تشخیص بلاک و کول‌داون خودکار.
# ----------------------------------------------------------------------

def looks_blocked(html: str) -> bool:
    """تشخیص می‌دهد که آیا DuckDuckGo به‌جای نتایج واقعی، صفحه‌ی هشدار/بلاک
    برگردانده یا نه (مثلاً هنگام تشخیص ترافیک ربات‌مانند)."""
    lowered = html.lower()
    block_markers = [
        "unusual traffic", "automated", "captcha", "are you a robot",
        "blocked", "rate limit", "too many requests",
    ]
    return any(marker in lowered for marker in block_markers)


def fetch_duckduckgo_results(domain: str, keyword: str = "Iran", max_results: int = MAX_RESULTS_PER_TANK):
    query = f"site:{domain} {keyword}"
    url = "https://html.duckduckgo.com/html/"
    params = urllib.parse.urlencode({"q": query})
    full_url = f"{url}?{params}"

    headers = random.choice(HEADERS_BROWSER_POOL)
    req = urllib.request.Request(full_url, headers=headers)

    try:
        with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT) as resp:
            html = resp.read().decode("utf-8", errors="ignore")
    except Exception as e:
        print(f"  [!] خطا در جست‌وجوی {domain}: {e}", file=sys.stderr)
        return [], False

    if looks_blocked(html):
        print(f"  [!] نشانه‌ی بلاک‌شدن موقت در پاسخ برای {domain} دیده شد.", file=sys.stderr)
        return [], True

    soup = BeautifulSoup(html, "html.parser")
    items = []
    for result in soup.select(".result"):
        a_tag = result.select_one(".result__a")
        if not a_tag or not a_tag.get("href"):
            continue
        real_link = extract_real_url(a_tag["href"])
        if not real_link:
            continue
        title = a_tag.get_text(strip=True)
        snippet_tag = result.select_one(".result__snippet")
        snippet = snippet_tag.get_text(strip=True) if snippet_tag else ""
        if title and real_link:
            items.append({"title": title, "link": real_link, "snippet": snippet})
        if len(items) >= max_results:
            break
    return items, False


def extract_real_url(href: str) -> str:
    if href.startswith("//"):
        href = "https:" + href
    parsed = urllib.parse.urlparse(href)
    if "duckduckgo.com" in parsed.netloc and parsed.path.startswith("/l/"):
        qs = urllib.parse.parse_qs(parsed.query)
        if "uddg" in qs:
            return urllib.parse.unquote(qs["uddg"][0])
        return ""
    return href


def fetch_google_news_rss(domain: str, keyword: str = "Iran", max_results: int = MAX_RESULTS_PER_TANK):
    """فید RSS اخبار گوگل را برای site:domain keyword می‌خواند. این روش هرگز
    توسط گوگل مسدود نمی‌شود (برخلاف DuckDuckGo روی IPهای گیت‌هاب)، اما فقط
    محتوای «تازه‌ی خبری» اخیر را نشان می‌دهد، نه گزارش‌های تحلیلی قدیمی‌تر."""
    query = f"{keyword} site:{domain}"
    encoded = urllib.parse.quote(query)
    url = f"https://news.google.com/rss/search?q={encoded}&hl=en-US&gl=US&ceid=US:en"

    req = urllib.request.Request(url, headers=random.choice(HEADERS_BROWSER_POOL))
    try:
        with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT) as resp:
            data = resp.read()
    except Exception as e:
        print(f"  [!] خطا در گوگل‌نیوز برای {domain}: {e}", file=sys.stderr)
        return []

    try:
        import xml.etree.ElementTree as ET
        root = ET.fromstring(data)
    except Exception as e:
        print(f"  [!] خطا در پارس XML گوگل‌نیوز برای {domain}: {e}", file=sys.stderr)
        return []

    items = []
    for item in root.findall(".//item")[:max_results]:
        title = (item.findtext("title") or "").strip()
        link = (item.findtext("link") or "").strip()
        if title and link:
            items.append({"title": title, "link": link, "snippet": ""})
    return items


def fetch_search_results(domain: str, keyword: str = "Iran", max_results: int = MAX_RESULTS_PER_TANK):
    """تابع اصلی جست‌وجو: دو منبع را ترکیب می‌کند:
      ۱) گوگل‌نیوز RSS (همیشه قابل‌اعتماد، فقط اخبار تازه)
      ۲) Brave API (اگر کلید تنظیم شده) یا DuckDuckGo (تلاش برای پوشش عمیق‌تر،
         ممکن است گاهی توسط IP گیت‌هاب بسته شود)
    نتایج ترکیب و بدون تکرار (بر اساس لینک) برگردانده می‌شوند."""
    combined = {}
    was_blocked = False

    # منبع ۱: گوگل‌نیوز (تضمینی)
    for it in fetch_google_news_rss(domain, keyword, max_results):
        combined[it["link"]] = it

    # منبع ۲: Brave یا DuckDuckGo (بهترین تلاش، برای پوشش عمیق‌تر)
    if BRAVE_API_KEY:
        extra_items = fetch_brave_results(domain, keyword, max_results)
    else:
        extra_items, was_blocked = fetch_duckduckgo_results(domain, keyword, max_results)

    for it in extra_items:
        combined.setdefault(it["link"], it)

    return list(combined.values())[:max_results * 2], was_blocked


# ----------------------------------------------------------------------
# مرحله ۲: مدیریت حافظه‌ی لینک‌های قبلاً دیده‌شده
# ----------------------------------------------------------------------

def load_seen_links() -> set:
    if os.path.exists(SEEN_FILE):
        try:
            with open(SEEN_FILE, "r", encoding="utf-8") as f:
                return set(json.load(f))
        except Exception as e:
            print(f"[!] خطا در خواندن {SEEN_FILE}: {e}", file=sys.stderr)
    return set()


def save_seen_links(links: set):
    with open(SEEN_FILE, "w", encoding="utf-8") as f:
        json.dump(sorted(links), f, ensure_ascii=False, indent=2)


# ----------------------------------------------------------------------
# مرحله ۳: جمع‌آوری نتایج جدید برای همه‌ی اندیشکده‌ها
# ----------------------------------------------------------------------

def collect_all_news(old_seen: set):
    """خروجی: (results, all_links_this_run)
    results: dict[country][tank_name] -> list[items جدید]
    all_links_this_run: مجموعه‌ی همه‌ی لینک‌هایی که امروز دیده شدند (جدید و قدیمی)
    """
    results = {}
    all_links_this_run = set()
    consecutive_blocks = 0

    if not BRAVE_API_KEY:
        print("[i] از DuckDuckGo (رایگان، بدون نیاز به کلید) استفاده می‌شود.", file=sys.stderr)

    for country, tanks in THINKTANKS.items():
        results[country] = {}
        for fa_name, en_name, domain in tanks:
            print(f"در حال جست‌وجو: {country} | {fa_name} ({domain}) ...")
            items, was_blocked = fetch_search_results(domain)

            if was_blocked:
                consecutive_blocks += 1
                print(f"  [!] کول‌داون {COOLDOWN_AFTER_BLOCK} ثانیه‌ای برای دور زدن بلاک موقت ...")
                time.sleep(COOLDOWN_AFTER_BLOCK)
                # یک تلاش دوم بعد از کول‌داون
                items, was_blocked_retry = fetch_search_results(domain)
                if was_blocked_retry:
                    print(f"  [!] بعد از کول‌داون هم بلاک بود؛ از {domain} صرف‌نظر شد این‌بار.")
            else:
                consecutive_blocks = 0

            new_items = []
            for it in items:
                all_links_this_run.add(it["link"])
                if it["link"] not in old_seen:
                    new_items.append(it)

            if new_items:
                results[country][f"{fa_name} ({en_name})"] = new_items

            time.sleep(random.uniform(MIN_SLEEP, MAX_SLEEP))

    return results, all_links_this_run


# ----------------------------------------------------------------------
# مرحله ۵: تبدیل خودِ صفحه‌ی وب مقاله به PDF
# ----------------------------------------------------------------------

def slugify_filename(index: int, domain: str) -> str:
    """یک نام فایل امن و یکتا برای هر مقاله می‌سازد، مثل 001_brookings_edu.pdf"""
    safe_domain = re.sub(r"[^a-zA-Z0-9]+", "_", domain).strip("_")[:40]
    return f"{index:03d}_{safe_domain}.pdf"


def convert_url_to_pdf(url: str, output_path: str) -> bool:
    """صفحه‌ی وب مقاله را با wkhtmltopdf به یک فایل PDF تبدیل می‌کند.
    خروجی: True اگر موفق بود (فایل واقعاً ساخته و غیر خالی بود)، وگرنه False."""
    cmd = [
        WKHTMLTOPDF_BIN,
        "--quiet",
        "--load-error-handling", "ignore",
        "--load-media-error-handling", "ignore",
        "--javascript-delay", "1200",
        "--no-stop-slow-scripts",
        "--print-media-type",
        "--disable-smart-shrinking",
        url, output_path,
    ]
    try:
        subprocess.run(
            cmd, timeout=ARTICLE_PDF_TIMEOUT,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
    except subprocess.TimeoutExpired:
        print(f"    [!] تایم‌اوت در تبدیل {url}", file=sys.stderr)
        return False
    except Exception as e:
        print(f"    [!] خطا در تبدیل {url}: {e}", file=sys.stderr)
        return False

    # حتی اگر wkhtmltopdf کد خطا برگرداند، گاهی فایل PDF جزئی ساخته می‌شود؛
    # فقط وجود و حداقل حجم منطقی را چک می‌کنیم.
    if os.path.exists(output_path) and os.path.getsize(output_path) > 2048:
        return True
    return False


def convert_all_articles(flat_items: list) -> list:
    """برای هر آیتم {title, link, country, tank}، صفحه را به PDF تبدیل می‌کند.
    خروجی: لیستی از dict های موفق با کلید اضافه‌ی 'pdf_path'."""
    os.makedirs(ARTICLES_DIR, exist_ok=True)
    successful = []
    total = len(flat_items)

    for i, item in enumerate(flat_items, start=1):
        domain = urllib.parse.urlparse(item["link"]).netloc
        filename = slugify_filename(i, domain)
        output_path = os.path.join(ARTICLES_DIR, filename)

        print(f"[{i}/{total}] تبدیل به PDF: {item['title'][:60]} ({domain}) ...")
        ok = convert_url_to_pdf(item["link"], output_path)

        if ok:
            item["pdf_path"] = output_path
            item["pdf_filename"] = filename
            successful.append(item)
        else:
            print(f"    [!] تبدیل ناموفق بود، این مقاله رد شد: {item['link']}")

    return successful


# ----------------------------------------------------------------------
# مرحله ۶: بسته‌بندی PDFها در فایل(های) زیپ (با رعایت سقف حجم ایمیل)
# ----------------------------------------------------------------------

def build_zip_chunks(successful_items: list) -> list:
    """PDFهای موفق را در یک یا چند فایل زیپ (هرکدام زیر MAX_ZIP_SIZE_MB)
    بسته‌بندی می‌کند. خروجی: لیستی از مسیر فایل‌های زیپ ساخته‌شده."""
    chunks = []
    current_zip_items = []
    current_size = 0
    chunk_index = 1

    def flush_chunk():
        nonlocal chunk_index, current_zip_items, current_size
        if not current_zip_items:
            return
        zip_path = f"iran_articles_part{chunk_index}.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for it in current_zip_items:
                zf.write(it["pdf_path"], arcname=it["pdf_filename"])
        chunks.append({"zip_path": zip_path, "items": list(current_zip_items)})
        chunk_index += 1
        current_zip_items = []
        current_size = 0

    for item in successful_items:
        size = os.path.getsize(item["pdf_path"])
        if current_size + size > MAX_ZIP_SIZE_BYTES and current_zip_items:
            flush_chunk()
        current_zip_items.append(item)
        current_size += size

    flush_chunk()
    return chunks


# ----------------------------------------------------------------------
# مرحله ۷: ارسال ایمیل (یک ایمیل به‌ازای هر بخش زیپ)
# ----------------------------------------------------------------------

def send_email_with_zip(zip_path: str, items: list, part_num: int, total_parts: int, failed_count: int):
    if not GMAIL_USER or not GMAIL_APP_PASSWORD:
        print("[!] GMAIL_USER یا GMAIL_APP_PASSWORD تنظیم نشده است. ایمیل ارسال نشد.", file=sys.stderr)
        sys.exit(1)

    today_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    part_label = f" (بخش {part_num} از {total_parts})" if total_parts > 1 else ""

    msg = EmailMessage()
    msg["Subject"] = f"مقالات تازه اندیشکده‌ها درباره ایران — {today_str}{part_label} ({len(items)} مقاله)"
    msg["From"] = GMAIL_USER
    msg["To"] = RECIPIENT_EMAIL

    body_lines = [
        "سلام،", "",
        f"{len(items)} مقاله‌ی جدید درباره ایران (به‌صورت فایل PDF کامل) در این بخش پیوست شده است.",
    ]
    if part_num == 1 and failed_count:
        body_lines.append(f"({failed_count} مقاله‌ی دیگر به‌دلیل محدودیت سایت مبدأ قابل تبدیل نبود و رد شد.)")
    body_lines += ["", "فهرست این بخش:"]
    for it in items:
        body_lines.append(f"- {it['title']} — {it['link']}")
    body_lines += ["", "این ایمیل به‌صورت خودکار تولید شده است."]

    msg.set_content("\n".join(body_lines))

    with open(zip_path, "rb") as f:
        zip_data = f.read()
    msg.add_attachment(
        zip_data, maintype="application", subtype="zip",
        filename=f"iran_articles_{today_str}_part{part_num}.zip"
    )

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        server.send_message(msg)
    print(f"[✓] ایمیل بخش {part_num}/{total_parts} با موفقیت ارسال شد ({len(items)} مقاله).")


# ----------------------------------------------------------------------
# اجرای اصلی
# ----------------------------------------------------------------------

def main():
    print("=== شروع رصد روزانه اندیشکده‌ها (تبدیل مقالات به PDF) ===")

    old_seen = load_seen_links()
    print(f"تعداد لینک‌های قبلاً دیده‌شده: {len(old_seen)}")

    results, all_links_this_run = collect_all_news(old_seen)

    # تبدیل دیکشنری results به یک لیست تخت از آیتم‌ها برای پردازش راحت‌تر
    flat_items = []
    for country, tanks in results.items():
        for tank_name, items in tanks.items():
            for it in items:
                flat_items.append({
                    "title": it["title"],
                    "link": it["link"],
                    "country": country,
                    "tank": tank_name,
                })

    total_new = len(flat_items)
    print(f"جمعاً {total_new} گزارش/خبر جدید (که قبلاً دیده نشده بود) یافت شد.")

    if total_new > 0:
        print(f"شروع تبدیل {total_new} مقاله به PDF (ممکن است طول بکشد) ...")
        successful = convert_all_articles(flat_items)
        failed_count = total_new - len(successful)
        print(f"[✓] {len(successful)} مقاله با موفقیت به PDF تبدیل شد ({failed_count} مورد ناموفق).")

        if successful:
            chunks = build_zip_chunks(successful)
            total_parts = len(chunks)
            print(f"[✓] {total_parts} فایل زیپ ساخته شد.")
            for idx, chunk in enumerate(chunks, start=1):
                send_email_with_zip(
                    chunk["zip_path"], chunk["items"], idx, total_parts, failed_count
                )
        else:
            print("[!] هیچ مقاله‌ای با موفقیت تبدیل نشد؛ ایمیلی ارسال نمی‌شود.")
    else:
        print("هیچ مورد جدیدی یافت نشد؛ ایمیلی ارسال نمی‌شود.")

    # به‌روزرسانی حافظه با همه‌ی لینک‌های دیده‌شده تا امروز
    updated_seen = old_seen | all_links_this_run
    save_seen_links(updated_seen)
    print(f"[✓] فایل {SEEN_FILE} به‌روزرسانی شد ({len(updated_seen)} لینک ثبت‌شده).")


if __name__ == "__main__":
    main()
