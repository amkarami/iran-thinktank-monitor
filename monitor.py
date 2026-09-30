"""Iran think-tank daily monitor. Set REGION=all to scan and send daily reports."""
import os
import sys
import json
import time
import smtplib
import logging
from datetime import datetime, timedelta, timezone
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
RECIPIENT_EMAIL = "siadenichi@gmail.com"
SENDER_EMAIL = os.environ.get("GMAIL_USER", "")
SENDER_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
REGION = os.environ.get("REGION", "all")
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; IranThinkTankMonitor/1.0)", "Accept": "text/html,application/xhtml+xml"}
IRAN_KEYWORDS = ["iran", "iranian", "tehran", "jcpoa", "irgc", "iaea", "khamenei", "rouhani", "raisi", "pezeshkian", "nuclear deal", "sanctions iran", "persian gulf", "revolutionary guard", "sepah", "quds force", "hezbollah iran", "hamas iran", "axis of resistance", "isfahan", "natanz", "fordow", "arak reactor", "strait of hormuz"]
REGION_FA = {"Americas":"آمریکا", "Europe":"اروپا", "Asia_Pacific":"آسیا و اقیانوسیه", "Middle_East":"خاورمیانه"}
# Each tuple is (think tank, search URL, country/area).
_RAW = {
"Americas": [
("Brookings","https://www.brookings.edu/search/?s=iran","USA"),("CFR","https://www.cfr.org/search?term=iran","USA"),("RAND","https://www.rand.org/search.html#q=iran","USA"),("Carnegie Endowment","https://carnegieendowment.org/search?q=iran","USA"),("CSIS","https://www.csis.org/search/iran","USA"),("Wilson Center","https://www.wilsoncenter.org/search?term=iran","USA"),("Atlantic Council","https://www.atlanticcouncil.org/?s=iran","USA"),("Heritage Foundation","https://www.heritage.org/search#q=iran","USA"),("AEI","https://www.aei.org/search/?q=iran","USA"),("Hudson Institute","https://www.hudson.org/search#q=iran","USA"),("Cato Institute","https://www.cato.org/search?q=iran","USA"),("Stimson Center","https://www.stimson.org/?s=iran","USA"),("USIP","https://www.usip.org/search?term=iran","USA"),("Washington Institute","https://www.washingtoninstitute.org/search?q=iran","USA"),("FDD","https://www.fdd.org/?s=iran","USA"),("Middle East Institute","https://www.mei.edu/search?q=iran","USA"),("CAP","https://www.americanprogress.org/?s=iran","USA"),("New America","https://www.newamerica.org/search/?q=iran","USA"),("Quincy Institute","https://quincyinst.org/?s=iran","USA"),("National Interest","https://nationalinterest.org/search/iran","USA"),("FPRI","https://www.fpri.org/?s=iran","USA"),("Woodrow Wilson Center","https://www.wilsoncenter.org/search?term=iran","USA"),("Belfer Center Harvard","https://www.belfercenter.org/search?term=iran","USA"),("Hoover Institution","https://www.hoover.org/search#q=iran","USA"),("IISS Americas","https://www.iiss.org/search#q=iran","USA"),("Carnegie Global","https://carnegieendowment.org/search?q=iran","USA"),("Third Way","https://www.thirdway.org/search?q=iran","USA"),("Defense Priorities","https://defensepriorities.org/?s=iran","USA"),("Niskanen Center","https://www.niskanencenter.org/?s=iran","USA"),("R Street Institute","https://www.rstreet.org/?s=iran","USA"),("Arms Control Association","https://www.armscontrol.org/search?q=iran","USA"),("Nuclear Threat Initiative","https://www.nti.org/search/?q=iran","USA"),("Canadian Global Affairs Institute","https://www.cgai.ca/search?q=iran","Canada"),("IRPP Canada","https://irpp.org/?s=iran","Canada")],
"Europe": [
("Chatham House","https://www.chathamhouse.org/search?search_api_fulltext=iran","UK"),("IISS","https://www.iiss.org/search#q=iran","UK"),("ECFR","https://ecfr.eu/?s=iran","Germany/EU"),("SWP Berlin","https://www.swp-berlin.org/en/search?q=iran","Germany"),("SIPRI","https://www.sipri.org/search?search_api_fulltext=iran","Sweden"),("IFRI","https://www.ifri.org/en/search?search_api_fulltext=iran","France"),("IAI Italy","https://www.iai.it/en/search?q=iran","Italy"),("Clingendael","https://www.clingendael.org/search?q=iran","Netherlands"),("EGMONT Institute","https://www.egmontinstitute.be/?s=iran","Belgium"),("FRIDE Spain","https://www.fride.org/search?q=iran","Spain"),("Real Instituto Elcano","https://www.realinstitutoelcano.org/en/search/?q=iran","Spain"),("NUPI Norway","https://www.nupi.no/en/search?q=iran","Norway"),("FIIA Finland","https://www.fiia.fi/en/search?q=iran","Finland"),("DIIS Denmark","https://www.diis.dk/en/search?q=iran","Denmark"),("PISM Poland","https://www.pism.pl/en/search?q=iran","Poland"),("IEP Paris","https://www.iep-paris.fr/en/search?q=iran","France"),("CIDOB Barcelona","https://www.cidob.org/en/search?q=iran","Spain"),("Bruegel","https://www.bruegel.org/search?q=iran","Belgium"),("European Policy Centre","https://www.epc.eu/en/search?q=iran","Belgium"),("Friends of Europe","https://www.friendsofeurope.org/?s=iran","Belgium"),("CEIP Europe","https://carnegieeurope.eu/search?q=iran","Belgium"),("ISPI Italy","https://www.ispionline.it/en/search?q=iran","Italy"),("CSS ETH Zurich","https://css.ethz.ch/en/search.html?q=iran","Switzerland"),("Geneva Centre for Security Policy","https://www.gcsp.ch/search?q=iran","Switzerland"),("IRSEM France","https://www.irsem.fr/en/search?q=iran","France"),("RAND Europe","https://www.rand.org/randeurope/research.html","UK"),("Henry Jackson Society","https://henryjacksonsociety.org/?s=iran","UK"),("Policy Exchange UK","https://policyexchange.org.uk/search/?q=iran","UK"),("RUSI","https://rusi.org/search?search=iran","UK"),("Wilton Park","https://www.wiltonpark.org.uk/search/?q=iran","UK"),("Koerber Stiftung","https://koerber-stiftung.de/en/search/?q=iran","Germany"),("WIIW Vienna","https://wiiw.ac.at/search.html?q=iran","Austria")],
"Asia_Pacific": [
("ORF India","https://www.orfonline.org/search/?q=iran","India"),("IDSA India","https://idsa.in/search?q=iran","India"),("CPR India","https://www.cprindia.org/search?q=iran","India"),("IISS Asia","https://www.iiss.org/search#q=iran","Singapore"),("ISEAS Singapore","https://www.iseas.edu.sg/search/?q=iran","Singapore"),("RSIS Singapore","https://www.rsis.edu.sg/search/?q=iran","Singapore"),("Lowy Institute","https://www.lowyinstitute.org/search?q=iran","Australia"),("ASPI Australia","https://www.aspi.org.au/search?q=iran","Australia"),("AIIA Australia","https://www.aiia.asn.au/?s=iran","Australia"),("NIDS Japan","https://www.nids.mod.go.jp/english/research/index.html","Japan"),("JIIA Japan","https://www.jiia.or.jp/en/search?q=iran","Japan"),("EAI South Korea","https://www.eai.or.kr/main/eng/search/search.asp?q=iran","South Korea"),("ASAN Institute","https://en.asaninst.org/search/?q=iran","South Korea"),("Stimson Asia","https://www.stimson.org/?s=iran","USA/Asia"),("SIPRI Asia","https://www.sipri.org/search?search_api_fulltext=iran","Sweden/Asia"),("IISS Bahrain","https://www.iiss.org/search#q=iran","Bahrain"),("Pacific Forum CSIS","https://pacforum.org/search?q=iran","USA/Asia"),("NIST Thailand","https://www.nist.or.th/search?q=iran","Thailand"),("CARI Malaysia","https://cariasean.org/?s=iran","Malaysia"),("CISS China","https://www.ciss.com.cn/en/search?q=iran","China"),("CSIS Indonesia","https://csis.or.id/search?q=iran","Indonesia")],
"Middle_East": [
("INSS Israel","https://www.inss.org.il/search/?q=iran","Israel"),("BESA Center","https://besacenter.org/?s=iran","Israel"),("ITIC Israel","https://www.terrorism-info.org.il/en/search/?q=iran","Israel"),("MEI Washington","https://www.mei.edu/search?q=iran","USA/ME"),("Arab Reform Initiative","https://www.arab-reform.net/search?q=iran","France/ME"),("Issam Fares Institute AUB","https://www.aub.edu.lb/ifi/search?q=iran","Lebanon"),("KFCRIS Saudi Arabia","https://kfcris.com/en/search?q=iran","Saudi Arabia"),("Gulf International Forum","https://gulfif.org/?s=iran","USA/Gulf"),("Emirates Policy Center","https://epc.ae/en/search?q=iran","UAE"),("Al-Ahram Center Egypt","https://acpss.ahram.org.eg/search?q=iran","Egypt"),("Ahram Online Strategic","https://english.ahram.org.eg/search.aspx?q=iran","Egypt"),("Turkish Policy Quarterly","http://turkishpolicy.com/search?q=iran","Turkey"),("TESEV Turkey","https://www.tesev.org.tr/en/search?q=iran","Turkey"),("ORSAM Turkey","https://orsam.org.tr/en/search?q=iran","Turkey"),("Jordan Institute of Diplomacy","https://www.jid.org.jo/en/search?q=iran","Jordan"),("Rasanah IIIS","https://rasanah-iiis.org/en/?s=iran","Saudi Arabia")]
}
THINK_TANKS = {r: [{"name":n,"url":u,"country":c} for n,u,c in rows] for r,rows in _RAW.items()}


def _clean(value):
    return " ".join(str(value or "").replace("\xa0", " ").split())


def _parse_date(value):
    text = _clean(value).strip(" ,|")
    if not text:
        return None
    now = datetime.now(timezone.utc)
    low = text.lower()
    if low in ("today", "just now"):
        return now
    if low == "yesterday":
        return now - timedelta(days=1)
    import re
    m = re.search(r"(\d+)\s+(minute|minutes|hour|hours|day|days)\s+ago", low)
    if m:
        amount = int(m.group(1)); unit = m.group(2)
        return now - (timedelta(minutes=amount) if unit.startswith("minute") else timedelta(hours=amount) if unit.startswith("hour") else timedelta(days=amount))
    candidate = text[:-1] + "+00:00" if text.endswith("Z") else text
    try:
        dt = datetime.fromisoformat(candidate)
        return dt.replace(tzinfo=timezone.utc) if dt.tzinfo is None else dt.astimezone(timezone.utc)
    except ValueError:
        pass
    formats = ["%Y-%m-%d %H:%M:%S","%Y-%m-%d %H:%M","%Y-%m-%d","%Y/%m/%d","%Y.%m.%d","%B %d, %Y","%b %d, %Y","%d %B %Y","%d %b %Y","%d/%m/%Y","%m/%d/%Y","%d-%m-%Y","%m-%d-%Y","%d.%m.%Y"]
    for fmt in formats:
        try:
            return datetime.strptime(text,fmt).replace(tzinfo=timezone.utc)
        except ValueError:
            continue
    return None


def is_recent(date_str, hours=48):
    if date_str is None or not str(date_str).strip():
        return True
    parsed = _parse_date(date_str)
    if parsed is None:
        logging.info("Skipping unparseable date: %s", _clean(date_str)[:80])
        return True
    return parsed >= datetime.now(timezone.utc) - timedelta(hours=hours)


def contains_iran_keyword(text):
    value = str(text or "").lower()
    return any(keyword in value for keyword in IRAN_KEYWORDS)


def translate_to_persian(text, context="title") -> str:
    """Translate text to Persian through the configured LiteLLM-compatible API."""
    original = str(text or "").strip()
    if not original:
        return ""
    import re
    endpoint = "https://llm.aiprc.ir/v1/chat/completions"
    api_key = "secret-e4a1bc30"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    chunks = [original[i:i + 4500] for i in range(0, len(original), 4500)]
    translated_chunks = []
    for chunk in chunks:
        payload = {
            "model": "gemma4-12b",
            "messages": [
                {"role": "system", "content": "Translate the provided text into natural, accurate Persian. Return only the translation, preserving names and meaning."},
                {"role": "user", "content": f"Context: {context}\n\nText:\n{chunk}"},
            ],
            "temperature": 0.2,
        }
        result = None
        for attempt, delay in enumerate((0, 2, 4)):
            if delay:
                time.sleep(delay)
            try:
                response = requests.post(endpoint, headers=headers, json=payload, timeout=30)
                response.raise_for_status()
                result = response.json()["choices"][0]["message"]["content"]
                if isinstance(result, list):
                    result = "".join(part.get("text", "") if isinstance(part, dict) else str(part) for part in result)
                result = str(result or "").strip()
                if not result:
                    raise ValueError("Empty translation response")
                break
            except Exception as exc:
                logging.warning("Translation attempt %d/3 failed (%s): %s", attempt + 1, type(exc).__name__, exc)
        if not result:
            return original
        translated_chunks.append(result)
    translated = " ".join(translated_chunks).strip()
    if not re.search(r"[\u0600-\u06FF]", translated):
        logging.warning("Translation sanity check failed; keeping original text")
        return original
    return translated


def fetch_page(url, timeout=15):
    try:
        response = requests.get(url, headers=HEADERS, timeout=timeout)
        response.raise_for_status()
        try:
            return BeautifulSoup(response.text, "lxml")
        except Exception:
            return BeautifulSoup(response.text, "html.parser")
    except requests.RequestException as exc:
        logging.warning("Could not fetch %s: %s", url, exc)
        return None


def _summary_text(value):
    """Normalize a summary, removing any escaped HTML and common UI noise."""
    if not value:
        return ""
    text = BeautifulSoup(str(value), "html.parser").get_text(" ", strip=True)
    text = _clean(text)
    import re
    text = re.sub(r"\b(?:read more|read article|share|tweet|email|print)\b", " ", text, flags=re.IGNORECASE)
    return _clean(text)[:200]


def _valid_summary(value):
    text = _summary_text(value)
    junk = {"search", "home", "about", "contact", "menu", "news", "read more", "loading", "login", "subscribe", "skip to content"}
    return text if len(text) >= 40 and text.casefold().strip(" .,!?:;-") not in junk else ""


def _article_summary(node, soup):
    # Prefer article-local content and metadata; only then inspect page-level material.
    for container in (node, soup):
        if container is None:
            continue
        for attrs in ({"name": "description"}, {"property": "og:description"}, {"name": "twitter:description"}):
            meta = container.find("meta", attrs=attrs)
            if meta:
                value = _valid_summary(meta.get("content", ""))
                if value:
                    return value[:200]
    for container in (node, soup):
        if container is None:
            continue
        parts = []
        for fragment in container.stripped_strings:
            fragment = _valid_summary(fragment)
            if fragment:
                parts.append(fragment)
            if sum(len(part) + 1 for part in parts) >= 200:
                break
        candidate = _valid_summary(" ".join(parts))
        if candidate:
            return candidate[:200]
    return ""


def _extract_date(node, soup):
    # Prioritized machine-readable dates, then common visible date elements.
    time_node = node.find("time") if node else None
    if time_node and time_node.get("datetime"):
        return _clean(time_node.get("datetime"))[:120]
    for script in (soup.find_all("script", attrs={"type": "application/ld+json"}) if soup else []):
        try:
            data = json.loads(script.string or script.get_text())
        except (TypeError, ValueError):
            continue
        stack = data if isinstance(data, list) else [data]
        while stack:
            item = stack.pop(0)
            if isinstance(item, dict):
                date_value = item.get("datePublished")
                if date_value:
                    return _clean(date_value)[:120]
                stack.extend(value for value in item.values() if isinstance(value, (dict, list)))
            elif isinstance(item, list):
                stack.extend(item)
    meta = soup.find("meta", attrs={"property": "article:published_time"}) if soup else None
    if meta and meta.get("content"):
        return _clean(meta.get("content"))[:120]
    selectors = ["time[datetime]", ".date", ".published", ".entry-date", ".post-date", ".article-date", ".publication-date", ".timestamp", "[class*=date]", "[class*=published]"]
    for container in (node, soup):
        if container is None:
            continue
        for selector in selectors:
            candidate = container.select_one(selector)
            if candidate:
                value = candidate.get("datetime") or candidate.get_text(" ", strip=True)
                if value:
                    return _clean(value)[:120]
    return "—"


def extract_articles(soup, base_url):
    if soup is None:
        return []
    selectors = ["article", ".search-result", ".result-item", ".post", ".entry", ".card", ".item"]
    nodes = []
    for selector in selectors:
        nodes.extend(soup.select(selector))
    if not nodes:
        nodes = soup.find_all("a", href=True)
    articles, seen = [], set()
    for node in nodes:
        if node.name == "a":
            title_node = node
            href = node.get("href", "")
            date_text = ""
        else:
            title_node = next((node.find(tag) for tag in ("h1","h2","h3","h4") if node.find(tag)), None)
            if title_node is None:
                title_node = node.select_one(".title, .entry-title, .post-title, [class*='title']")
            anchor = node.find("a", href=True)
            href = anchor.get("href", "") if anchor else ""
            date_text = _extract_date(node, soup)
            if title_node is None and anchor:
                title_node = anchor
        title = _clean(title_node.get_text(" ", strip=True)) if title_node else ""
        link = urljoin(base_url.split("#",1)[0], href).split("#",1)[0]
        if not link.startswith(("http://","https://")) or link in seen or len(title) < 15 or " " not in title:
            continue
        if link.lower().endswith((".jpg",".jpeg",".png",".gif",".svg",".css",".js",".zip")):
            continue
        summary = _article_summary(node, soup)
        if not contains_iran_keyword(title + " " + summary):
            continue
        seen.add(link)
        articles.append({"title":title[:300],"url":link,"date":_clean(date_text)[:120] or "—","summary":summary})
        if len(articles) >= 8:
            break
    return articles


def translate_articles(articles):
    """Translate article titles and summaries, retaining originals on failure."""
    for index, article in enumerate(articles or []):
        article["title_fa"] = translate_to_persian(article.get("title", ""), context="title") or article.get("title", "")
        article["summary_fa"] = translate_to_persian(article.get("summary", ""), context="summary") or article.get("summary", "")
        if index < len(articles) - 1:
            time.sleep(0.5)
    return articles


def _save_json(region_key, date_str, articles):
    path = "results_%s_%s.json" % (region_key,date_str)
    with open(path,"w",encoding="utf-8") as handle:
        json.dump({"region":region_key,"date":date_str,"count":len(articles),"articles":articles},handle,ensure_ascii=False,indent=2)
    logging.info("Saved %s (%d articles)",path,len(articles))
    return path


def scan_region(region_key):
    date_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    tanks = THINK_TANKS.get(region_key)
    if tanks is None:
        logging.error("Unknown region: %s",region_key)
        _save_json(region_key,date_str,[])
        return []
    results, seen = [], set()
    for index,tank in enumerate(tanks):
        if index:
            time.sleep(0.5)
        try:
            soup = fetch_page(tank["url"])
            for article in extract_articles(soup,tank["url"]):
                if article["url"] in seen or not is_recent(article.get("date"),48):
                    continue
                seen.add(article["url"])
                article.update({"think_tank":tank["name"],"country":tank["country"],"region":region_key})
                results.append(article)
        except Exception:
            logging.exception("Error scanning %s",tank["name"])
    translate_articles(results)
    _save_json(region_key,date_str,results)
    return results


def _dedupe(data):
    seen, output = set(), []
    for item in data or []:
        url = (item.get("url") or "").strip()
        key = url.lower() if url else (item.get("title"),item.get("think_tank"))
        if key not in seen:
            seen.add(key); output.append(item)
    return output


def build_excel(data, filepath):
    wb = Workbook(); ws = wb.active; ws.title = "گزارش"; ws.sheet_view.rightToLeft = True
    labels = ["ردیف","عنوان گزارش","قاره کشور","تاریخ","آدرس وب‌سایت","چکیده فارسی"]
    navy = PatternFill("solid",fgColor="1F3864"); alternate = PatternFill("solid",fgColor="DCE6F1")
    white_bold = Font(name="Tahoma",bold=True,color="FFFFFF",size=12)
    normal = Font(name="Tahoma",size=11); linkfont = Font(name="Tahoma",size=11,color="0563C1",underline="single")
    border = Border(*( [Side(style="thin",color="BFBFBF")] * 4 ))
    for col,label in enumerate(labels,1):
        c=ws.cell(1,col,label); c.fill=navy; c.font=white_bold; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=border
    ws.row_dimensions[1].height=30
    for i,item in enumerate(data,1):
        values=[i,item.get("title_fa") or item.get("title") or "بدون عنوان", "%s / %s"%(REGION_FA.get(item.get("region",""),item.get("region","")),item.get("country","")), item.get("date") or "—",item.get("url","") or "",item.get("summary_fa") or item.get("summary") or ""]
        for col,value in enumerate(values,1):
            if col != 1 and isinstance(value,str) and value[:1] in ("=","+","-","@"):
                value="'"+value
            c=ws.cell(i+1,col,value); c.font=linkfont if col==5 else normal; c.border=border
            c.alignment=Alignment(horizontal="center" if col in (1,3,4) else "right",vertical="center",wrap_text=True)
            if i%2==0: c.fill=alternate
            if col==5 and isinstance(value,str) and value.startswith("http") and len(value)<=255: c.hyperlink=value
        ws.row_dimensions[i+1].height=55
    for col,width in enumerate([10,42,28,22,46,62],1): ws.column_dimensions[get_column_letter(col)].width=width
    ws.freeze_panes="A2"; ws.auto_filter.ref="A1:F%d"%max(1,len(data)+1)
    wb.save(filepath); return filepath


def _rtl(paragraph, alignment=WD_ALIGN_PARAGRAPH.RIGHT):
    paragraph.alignment=alignment
    ppr=paragraph._p.get_or_add_pPr(); bidi=OxmlElement("w:bidi"); bidi.set(qn("w:val"),"1"); ppr.append(bidi)

def _run_style(run,size=12,bold=False,color=None,underline=False):
    run.font.name="Tahoma"; run.font.size=Pt(size); run.bold=bold; run.underline=underline
    if color: run.font.color.rgb=color
    rpr=run._element.get_or_add_rPr(); fonts=OxmlElement("w:rFonts")
    for key in ("ascii","hAnsi","cs"): fonts.set(qn("w:"+key),"Tahoma")
    rpr.append(fonts); rtl=OxmlElement("w:rtl"); rtl.set(qn("w:val"),"1"); rpr.append(rtl)

def build_word(data, filepath):
    doc=Document(); section=doc.sections[0]; section.top_margin=Cm(1.8); section.bottom_margin=Cm(1.8)
    title=doc.add_heading("گزارش روزانه اندیشکده‌ها",0); _rtl(title,WD_ALIGN_PARAGRAPH.CENTER)
    for r in title.runs: _run_style(r,22,True,RGBColor(31,56,100))
    date_str=(data[0].get("_report_date") if data else None) or datetime.now(timezone.utc).strftime("%Y-%m-%d")
    p=doc.add_paragraph(); _rtl(p,WD_ALIGN_PARAGRAPH.CENTER); _run_style(p.add_run("تاریخ گزارش: "+date_str),12)
    for item in data:
        h=doc.add_heading(item.get("title_fa") or item.get("title") or "بدون عنوان",2); _rtl(h)
        for r in h.runs: _run_style(r,16,True,RGBColor(31,56,100))
        p=doc.add_paragraph(); _rtl(p); _run_style(p.add_run("%s | %s | %s"%(item.get("think_tank",""),item.get("country",""),item.get("date") or "—")),11)
        p=doc.add_paragraph(); _rtl(p); _run_style(p.add_run(item.get("url","") or ""),11,False,RGBColor(5,99,193),True)
        p=doc.add_paragraph(); _rtl(p); _run_style(p.add_run(item.get("summary_fa") or item.get("summary") or ""),12)
        p=doc.add_paragraph(); _rtl(p); _run_style(p.add_run("─"*60),10,False,RGBColor(136,136,136))
    doc.core_properties.title="گزارش روزانه اندیشکده‌ها"; doc.save(filepath); return filepath


def build_no_result_pdf(filepath):
    pdf=canvas.Canvas(filepath,pagesize=A4); width,height=A4; pdf.setTitle("No reports found today")
    pdf.setFont("Helvetica-Bold",18); pdf.drawCentredString(width/2,height/2+30,"No reports found today")
    message="امروز گزارشی یافت نشد"
    font_path=next((p for p in ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf","/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf") if os.path.exists(p)),None)
    if font_path:
        try:
            import arabic_reshaper
            from bidi.algorithm import get_display
            pdfmetrics.registerFont(TTFont("Persian",font_path)); pdf.setFont("Persian",16)
            pdf.drawCentredString(width/2,height/2-8,get_display(arabic_reshaper.reshape(message)))
        except Exception:
            pdf.setFont("Helvetica",12); pdf.drawCentredString(width/2,height/2-8,"Persian notice: see email body")
    else:
        pdf.setFont("Helvetica",12); pdf.drawCentredString(width/2,height/2-8,"Persian notice: see email body")
    pdf.showPage(); pdf.save(); return filepath


def send_email(attachments, has_results, date_str):
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        raise RuntimeError("GMAIL_USER and GMAIL_APP_PASSWORD must be configured")
    if has_results:
        subject="گزارش روزانه اندیشکده‌ها — "+date_str
        body="با سلام،\n\nگزارش روزانه اندیشکده‌ها درباره ایران برای تاریخ %s آماده است. فایل اکسل و فایل Word پیوست شده‌اند.\n\nبا احترام"%date_str
    else:
        subject="گزارش روزانه اندیشکده‌ها — گزارشی یافت نشد — "+date_str
        body="با سلام،\n\nامروز گزارشی یافت نشد.\nفایل PDF پیوست شده است.\n\nبا احترام"
    msg=MIMEMultipart(); msg["From"]=SENDER_EMAIL; msg["To"]=RECIPIENT_EMAIL; msg["Subject"]=Header(subject,"utf-8"); msg.attach(MIMEText(body,"plain","utf-8"))
    for path in attachments:
        with open(path,"rb") as handle: payload=handle.read()
        part=MIMEBase("application","octet-stream"); part.set_payload(payload); encoders.encode_base64(part)
        part.add_header("Content-Disposition","attachment",filename=os.path.basename(path)); msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com",465,timeout=60) as server:
        server.login(SENDER_EMAIL,SENDER_PASSWORD); server.sendmail(SENDER_EMAIL,[RECIPIENT_EMAIL],msg.as_string())
    logging.info("Email sent to %s",RECIPIENT_EMAIL)


def _build_and_send(data, date_str):
    rows=_dedupe(data)
    for item in rows: item["_report_date"]=date_str
    if rows:
        excel="iran_thinktanks_%s.xlsx"%date_str; word="iran_thinktanks_%s.docx"%date_str
        build_excel(rows,excel); build_word(rows,word); send_email([excel,word],True,date_str); return [excel,word]
    pdf="no_report_%s.pdf"%date_str; build_no_result_pdf(pdf); send_email([pdf],False,date_str); return [pdf]


def main():
    date_str=datetime.now(timezone.utc).strftime("%Y-%m-%d")
    requested=(REGION or "all").strip()
    regions=list(THINK_TANKS) if requested=="all" else [requested]
    if any(region not in THINK_TANKS for region in regions):
        logging.error("Unknown REGION: %s",requested); return 2
    results=[]
    for region in regions: results.extend(scan_region(region))
    if requested=="all": _build_and_send(results,date_str)
    return 0

if __name__ == "__main__":
    sys.exit(main())
