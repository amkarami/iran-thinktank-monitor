# -*- coding: utf-8 -*-
"""
ایجنت روزانه رصد اندیشکده‌ها برای اخبار/گزارش‌های مرتبط با ایران.

مراحل:
  1. برای هر اندیشکده، فید RSS گوگل‌نیوز با فیلتر site: و کلمه‌کلیدی Iran را می‌خواند.
  2. نتایج ۴۸ ساعت اخیر را نگه می‌دارد (برای جلوگیری از جا ماندن به‌خاطر اختلاف منطقه زمانی/تأخیر).
  3. اگر حداقل یک نتیجه پیدا شد، یک PDF می‌سازد.
  4. PDF را از طریق جیمیل (با App Password) به آدرس مقصد ایمیل می‌کند.

اجرا: python monitor.py
متغیرهای محیطی لازم (از GitHub Secrets تزریق می‌شوند):
  GMAIL_USER          آدرس جیمیل فرستنده
  GMAIL_APP_PASSWORD  رمز App Password (نه رمز عبور معمولی جیمیل)
  RECIPIENT_EMAIL     آدرس گیرنده (پیش‌فرض: siadenichi@gmail.com)
"""

import os
import sys
import time
import smtplib
import ssl
import urllib.parse
import urllib.request
from email.message import EmailMessage
from datetime import datetime, timedelta, timezone
import xml.etree.ElementTree as ET

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import arabic_reshaper
from bidi.algorithm import get_display

from thinktanks import THINKTANKS

# ----------------------------------------------------------------------
# فونت فارسی/عربی (برای اینکه حروف در PDF درست بهم بچسبند و راست‌به‌چپ باشند)
# ----------------------------------------------------------------------
FONT_NAME = "Vazirmatn"
FONT_FILE = "Vazirmatn-Regular.ttf"
FONT_URL = "https://github.com/rastikerdar/vazirmatn/raw/master/fonts/ttf/Vazirmatn-Regular.ttf"


def ensure_persian_font():
    """اگر فونت فارسی محلی موجود نبود، آن را دانلود و در ReportLab ثبت می‌کند."""
    if not os.path.exists(FONT_FILE):
        print("در حال دانلود فونت فارسی Vazirmatn ...")
        try:
            urllib.request.urlretrieve(FONT_URL, FONT_FILE)
        except Exception as e:
            print(f"[!] دانلود فونت فارسی ناموفق بود: {e}", file=sys.stderr)
            return False
    try:
        pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_FILE))
        return True
    except Exception as e:
        print(f"[!] ثبت فونت فارسی ناموفق بود: {e}", file=sys.stderr)
        return False


def fa(text: str) -> str:
    """متن فارسی/عربی را برای نمایش صحیح در PDF (اتصال حروف + راست‌به‌چپ) آماده می‌کند."""
    reshaped = arabic_reshaper.reshape(text)
    return get_display(reshaped)

# ----------------------------------------------------------------------
# تنظیمات
# ----------------------------------------------------------------------
LOOKBACK_HOURS = 48          # فقط اخبار این‌قدر ساعت اخیر را نگه دار
MAX_ITEMS_PER_TANK = 6       # حداکثر خبر برای هر اندیشکده در PDF
REQUEST_TIMEOUT = 15         # ثانیه
SLEEP_BETWEEN_REQUESTS = 1.0 # برای احترام به سرویس گوگل‌نیوز

GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
RECIPIENT_EMAIL = os.environ.get("RECIPIENT_EMAIL", "siadenichi@gmail.com")

OUTPUT_PDF = "iran_thinktank_digest.pdf"

# ----------------------------------------------------------------------
# مرحله ۱: دریافت اخبار از گوگل‌نیوز RSS برای هر اندیشکده
# ----------------------------------------------------------------------

def fetch_google_news_rss(domain: str, keyword: str = "Iran"):
    """فید RSS گوگل‌نیوز را برای site:domain keyword می‌خواند و لیست آیتم‌ها را برمی‌گرداند."""
    query = f'{keyword} site:{domain}'
    encoded = urllib.parse.quote(query)
    url = f"https://news.google.com/rss/search?q={encoded}&hl=en-US&gl=US&ceid=US:en"

    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT) as resp:
            data = resp.read()
    except Exception as e:
        print(f"  [!] خطا در خواندن {domain}: {e}", file=sys.stderr)
        return []

    try:
        root = ET.fromstring(data)
    except ET.ParseError as e:
        print(f"  [!] خطا در پارس XML برای {domain}: {e}", file=sys.stderr)
        return []

    items = []
    for item in root.findall(".//item"):
        title = (item.findtext("title") or "").strip()
        link = (item.findtext("link") or "").strip()
        pub_date_raw = (item.findtext("pubDate") or "").strip()
        source_el = item.find("source")
        source_name = source_el.text.strip() if source_el is not None and source_el.text else domain

        pub_dt = None
        for fmt in ("%a, %d %b %Y %H:%M:%S %Z", "%a, %d %b %Y %H:%M:%S %z"):
            try:
                pub_dt = datetime.strptime(pub_date_raw, fmt)
                if pub_dt.tzinfo is None:
                    pub_dt = pub_dt.replace(tzinfo=timezone.utc)
                break
            except ValueError:
                continue

        if title and link:
            items.append({
                "title": title,
                "link": link,
                "pub_dt": pub_dt,
                "source": source_name,
            })
    return items


def is_recent(item, hours=LOOKBACK_HOURS):
    if item["pub_dt"] is None:
        return True  # اگر تاریخ قابل‌تشخیص نبود، برای احتیاط نگه‌اش دار
    cutoff = datetime.now(timezone.utc) - timedelta(hours=hours)
    return item["pub_dt"] >= cutoff


def collect_all_news():
    """برای همه اندیشکده‌ها اخبار را جمع می‌کند. خروجی: dict[country][tank_name] -> list[items]"""
    results = {}
    for country, tanks in THINKTANKS.items():
        results[country] = {}
        for fa_name, en_name, domain in tanks:
            print(f"در حال بررسی: {country} | {fa_name} ({domain}) ...")
            items = fetch_google_news_rss(domain)
            recent_items = [it for it in items if is_recent(it)]
            recent_items.sort(key=lambda x: x["pub_dt"] or datetime.min.replace(tzinfo=timezone.utc), reverse=True)
            recent_items = recent_items[:MAX_ITEMS_PER_TANK]
            if recent_items:
                results[country][f"{fa_name} ({en_name})"] = recent_items
            time.sleep(SLEEP_BETWEEN_REQUESTS)
    return results


# ----------------------------------------------------------------------
# مرحله ۲: ساخت PDF
# ----------------------------------------------------------------------

def build_pdf(results: dict, output_path: str):
    font_ok = ensure_persian_font()
    base_font = FONT_NAME if font_ok else "Helvetica"

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "TitleFa", parent=styles["Title"], fontName=base_font, fontSize=18,
        spaceAfter=6, alignment=1
    )
    date_style = ParagraphStyle(
        "DateFa", parent=styles["Normal"], fontName=base_font, fontSize=10,
        alignment=1, textColor="#555555"
    )
    country_style = ParagraphStyle(
        "CountryFa", parent=styles["Heading1"], fontName=base_font, fontSize=14,
        spaceBefore=16, spaceAfter=8, textColor="#1a3c6e", alignment=TA_RIGHT
    )
    tank_style = ParagraphStyle(
        "TankFa", parent=styles["Heading2"], fontName=base_font, fontSize=11.5,
        spaceBefore=10, spaceAfter=4, textColor="#333333", alignment=TA_RIGHT
    )
    item_style = ParagraphStyle(
        "ItemFa", parent=styles["Normal"], fontName=base_font, fontSize=9.5,
        spaceAfter=3, leading=13, alignment=TA_RIGHT
    )
    empty_style = ParagraphStyle(
        "EmptyFa", parent=styles["Normal"], fontName=base_font, fontSize=10,
        textColor="#888888", spaceAfter=10, alignment=TA_RIGHT
    )

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm,
        leftMargin=1.8 * cm, rightMargin=1.8 * cm,
    )

    flow = []
    today_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    flow.append(Paragraph(fa("گزارش روزانه اندیشکده‌ها درباره ایران"), title_style))
    flow.append(Paragraph(f"Daily Iran Think-Tank Digest &mdash; {today_str} (UTC)", date_style))
    flow.append(Spacer(1, 0.6 * cm))
    flow.append(HRFlowable(width="100%", color="#cccccc"))

    total_items = 0
    for country, tanks in results.items():
        if not tanks:
            continue
        flow.append(Paragraph(fa(country), country_style))
        for tank_name, items in tanks.items():
            flow.append(Paragraph(fa(tank_name), tank_style))
            for it in items:
                total_items += 1
                date_txt = it["pub_dt"].strftime("%Y-%m-%d %H:%M UTC") if it["pub_dt"] else ""
                # عنوان و منبع خبر معمولاً انگلیسی است، پس نیازی به reshape ندارد
                link_html = (
                    f'&#8226; <a href="{it["link"]}" color="blue">{it["title"]}</a>'
                    f'<br/><font size="8" color="#777777">{it["source"]} &mdash; {date_txt}</font>'
                )
                flow.append(Paragraph(link_html, item_style))
            flow.append(Spacer(1, 0.15 * cm))

    if total_items == 0:
        flow.append(Spacer(1, 1 * cm))
        flow.append(Paragraph(fa("امروز هیچ خبر یا گزارش مرتبط با ایران یافت نشد."), empty_style))

    doc.build(flow)
    return total_items


# ----------------------------------------------------------------------
# مرحله ۳: ارسال ایمیل با پیوست PDF
# ----------------------------------------------------------------------

def send_email_with_attachment(pdf_path: str, total_items: int):
    if not GMAIL_USER or not GMAIL_APP_PASSWORD:
        print("[!] GMAIL_USER یا GMAIL_APP_PASSWORD تنظیم نشده است. ایمیل ارسال نشد.", file=sys.stderr)
        sys.exit(1)

    msg = EmailMessage()
    today_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    msg["Subject"] = f"گزارش روزانه اندیشکده‌ها درباره ایران — {today_str} ({total_items} خبر)"
    msg["From"] = GMAIL_USER
    msg["To"] = RECIPIENT_EMAIL
    msg.set_content(
        f"سلام،\n\nگزارش روزانه رصد اندیشکده‌های آمریکا، بریتانیا، آلمان، چین، "
        f"عربستان، امارات و قطر درباره ایران پیوست شده است.\n"
        f"تعداد کل خبرها/گزارش‌های یافت‌شده: {total_items}\n\n"
        f"این ایمیل به‌صورت خودکار تولید شده است."
    )

    with open(pdf_path, "rb") as f:
        pdf_data = f.read()
    msg.add_attachment(
        pdf_data, maintype="application", subtype="pdf",
        filename=f"iran_thinktank_digest_{today_str}.pdf"
    )

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        server.send_message(msg)
    print(f"[✓] ایمیل با موفقیت به {RECIPIENT_EMAIL} ارسال شد.")


# ----------------------------------------------------------------------
# اجرای اصلی
# ----------------------------------------------------------------------

def main():
    print("=== شروع رصد روزانه اندیشکده‌ها ===")
    results = collect_all_news()

    total_items = sum(len(items) for tanks in results.values() for items in tanks.values())
    print(f"جمعاً {total_items} خبر/گزارش مرتبط با ایران یافت شد.")

    if total_items == 0:
        print("هیچ خبری یافت نشد؛ طبق درخواست، ایمیلی ارسال نمی‌شود.")
        return

    build_pdf(results, OUTPUT_PDF)
    print(f"[✓] فایل PDF ساخته شد: {OUTPUT_PDF}")

    send_email_with_attachment(OUTPUT_PDF, total_items)


if __name__ == "__main__":
    main()
