import concurrent.futures
import datetime
import io
import os
import random
import re
import ssl
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# 1. 套件載入與環境設定
# ---------------------------------------------------------------------------
try:
    from bs4 import BeautifulSoup
except ImportError:
    st.error("❌ 系統缺少 'bs4' 套件！請在終端機執行：pip install beautifulsoup4")
    st.stop()

try:
    import openpyxl
except ImportError:
    st.error("❌ 系統缺少 'openpyxl' 套件！請在終端機執行：pip install openpyxl")
    st.stop()

# 全域 SSL context，避免爬取特定網站時因憑證問題報錯崩潰
ssl_context = ssl.create_default_context()
ssl_context.check_hostname = False
ssl_context.verify_mode = ssl.CERT_NONE

# 頁面配置
st.set_page_config(
    page_title="彰化家扶輿情自動檢索與報表生成系統",
    page_icon="📰",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 初始化 Session State
if "search_history" not in st.session_state:
    st.session_state["search_history"] = []

# CSS 樣式與動畫
st.markdown(
    """
<style>
    .main { background-color: #f8f9fa; }
    .main-header { font-size: 2.2rem; font-weight: 700; color: #1F2937; margin-bottom: 0.2rem; }
    .sub-header { color: #6B7280; font-size: 1.0rem; margin-bottom: 0.8rem; }
    .warning-bar {
        background-color: #FEF2F2;
        border-left: 5px solid #EF4444;
        padding: 0.75rem 1rem;
        border-radius: 6px;
        margin-bottom: 1.5rem;
    }
    .warning-text { color: #DC2626; font-weight: 700; font-size: 0.95rem; margin: 0; }
    .search-card {
        background-color: #ffffff;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        border: 1px solid #E5E7EB;
        margin-bottom: 1.5rem;
    }
    [data-testid="stSidebar"] { background-color: #f1f5f9; border-right: 1px solid #E2E8F0; }
    div.stButton > button:first-child {
        background: linear-gradient(90deg, #2563EB, #1D4ED8);
        color: white; border: none; padding: 0.6rem 1.2rem;
        font-weight: 600; border-radius: 8px; transition: all 0.3s ease;
    }
    div.stButton > button:first-child:hover {
        background: linear-gradient(90deg, #1D4ED8, #1E40AF);
        transform: translateY(-1px);
    }

    /* 🎈 5秒上升氣球動畫 CSS */
    .balloon-container {
        position: fixed;
        bottom: -100px;
        left: 0;
        width: 100vw;
        height: 100vh;
        pointer-events: none;
        z-index: 99999;
        overflow: hidden;
    }
    .balloon {
        position: absolute;
        bottom: -100px;
        width: 40px;
        height: 55px;
        background-color: #FF5722;
        border-radius: 50% 50% 50% 50% / 40% 40% 60% 60%;
        animation: floatUp 5s ease-in-out forwards;
    }
    .balloon::after {
        content: "";
        position: absolute;
        bottom: -12px;
        left: 18px;
        width: 2px;
        height: 15px;
        background-color: #888;
    }
    @keyframes floatUp {
        0% { transform: translateY(0) rotate(0deg); opacity: 1; }
        100% { transform: translateY(-120vh) rotate(15deg); opacity: 0; }
    }
</style>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# 2. 標題與警示區塊
# ---------------------------------------------------------------------------
st.markdown('<div class="main-header">📰 彰化家扶中心輿情自動檢索與報表生成系統</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">支援多重轉載媒體檢索，可辨識記者姓名，並且不進行去重</div>', unsafe_allow_html=True)

st.markdown(
    """
<div class="warning-bar">
    <p class="warning-text">※本系統為個人自主開發，旨在優化查詢媒體露出流程與準確度，請勿用於非法行為😈</p>
    <p class="warning-text">※已強化全網電子報巡查與撰文記者姓名識別，並加入雙重檢核功能，過濾非相關新聞🌏</p>
    <p class="warning-text">※檢索資料庫（database.csv）為「彰化家扶」常見出報媒體清單，開發者將不定期更新👀</p>
    <p class="warning-text">※開發者保有此系統所有權，敬請尊重開發者之權利。若有不法，將依中華民國相關法規追究⚠️</p>
</div>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# 3. 核心工具：擷取 Domain
# ---------------------------------------------------------------------------

def extract_domain(url_or_html):
    if not url_or_html or not isinstance(url_or_html, str):
        return ""
    
    text = url_or_html.strip()
    if "<" in text and ">" in text:
        try:
            soup = BeautifulSoup(text, "html.parser")
            a_tag = soup.find("a")
            if a_tag and a_tag.get("href"):
                text = a_tag.get("href")
            else:
                text = soup.get_text()
        except Exception:
            pass

    domain = re.sub(r"^https?://", "", text.strip(), flags=re.IGNORECASE)
    domain = domain.split('/')[0].split('?')[0].split('#')[0]
    domain = re.sub(r"^(www\.|news\.)", "", domain, flags=re.IGNORECASE)
    return domain.lower()

# ---------------------------------------------------------------------------
# 4. 側邊欄與 database.csv 完整讀取
# ---------------------------------------------------------------------------
st.sidebar.title("⚙️ 系統功能導覽")
sidebar_option = st.sidebar.radio(
    "請選擇功能模組：",
    ["🔍 檢索系統", "💡 系統簡介", "📌 系統須知", "🔐 系統管理員"],
    index=0,
    label_visibility="collapsed",
)

st.sidebar.markdown("---")
db_file_path = "database.csv"
media_type_map = {}
csv_media_list = []  # 保存 CSV 裡面讀到的每一筆 (媒體名稱, 網域)

if os.path.exists(db_file_path):
    try:
        db_df = pd.read_csv(db_file_path, encoding="utf-8").dropna(how="all")
        st.sidebar.success("✅ database.csv 已連線")
        
        if len(db_df) > 0:
            st.sidebar.markdown("⏳ **正在解析媒體清單資料庫...**")
            domain_progress_bar = st.sidebar.progress(0)
            domain_status_text = st.sidebar.empty()
            
            total_rows = len(db_df)
            media_col = db_df.columns[0]
            type_col = db_df.columns[1] if len(db_df.columns) >= 2 else db_df.columns[0]

            for idx, row in db_df.iterrows():
                m_name = str(row[media_col]).strip()
                m_type = str(row[type_col]).strip() if len(db_df.columns) >= 2 else "非三大報全國性"
                
                if m_name and m_name.lower() != "nan":
                    media_type_map[m_name] = m_type

                    possible_url_or_html = ""
                    if len(row) >= 3 and pd.notna(row.iloc[2]):
                        possible_url_or_html = str(row.iloc[2])
                    elif "." in m_name:
                        possible_url_or_html = m_name

                    clean_dom = extract_domain(possible_url_or_html)
                    if clean_dom:
                        media_type_map[clean_dom] = m_type
                    
                    # 無論有沒有網域，都強制記錄下來
                    csv_media_list.append((m_name, clean_dom))

                pct = int((idx + 1) / total_rows * 100)
                domain_progress_bar.progress(pct)
                domain_status_text.markdown(f"📑已完成 **{pct}％**")

            domain_status_text.markdown(f"📑已載入 CSV 內 **{len(csv_media_list)}** 家媒體目標")
    except Exception as e:
        st.sidebar.error(f"❌ 讀取 database.csv 失敗: {e}")

# ---------------------------------------------------------------------------
# 5. 核心：重導向爬蟲 / 記者探針 V3 / 雙軌強迫檢索
# ---------------------------------------------------------------------------

def fetch_article_data(url):
    """取得新聞網頁內文與 Meta Tag 記者姓名"""
    if not url or not isinstance(url, str) or not url.startswith(("http://", "https://")):
        return "", ""

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7",
    }
    
    try:
        req = urllib.request.Request(url, headers=headers)
        opener = urllib.request.build_opener(
            urllib.request.HTTPSHandler(context=ssl_context),
            urllib.request.HTTPRedirectHandler()
        )
        with opener.open(req, timeout=7) as response:
            html_bytes = response.read()
            charset = response.headers.get_param("charset") or "utf-8"
            try:
                html = html_bytes.decode(charset, errors="ignore")
            except Exception:
                html = html_bytes.decode("utf-8", errors="ignore")

            soup = BeautifulSoup(html, "html.parser")

            meta_reporter = ""
            meta_candidates = [
                soup.find("meta", attrs={"name": re.compile(r"author|reporter|bnews:author|dnews:author", re.I)}),
                soup.find("meta", attrs={"property": re.compile(r"author|article:author|og:article:author", re.I)}),
                soup.find("meta", attrs={"itemprop": "author"}),
            ]
            for meta in meta_candidates:
                if meta and meta.get("content"):
                    meta_reporter = meta.get("content").strip()
                    break

            for script in soup(["script", "style", "noscript", "header", "footer", "nav", "aside", "iframe"]):
                script.extract()

            text = soup.get_text(separator=" ")
            clean_text = re.sub(r"\s+", " ", text).strip()
            return clean_text[:3500], meta_reporter
    except Exception:
        return "", ""


def reporter_detector_sensor_v3(article_text, meta_reporter=""):
    """記者探針 V3：辨識台灣媒體常見署名"""
    if meta_reporter:
        clean_meta = re.sub(r"(記者|特派員|專題小組|編輯|責任編輯|\s+)", "", meta_reporter).strip()
        if 2 <= len(clean_meta) <= 4 and not re.search(r"(新聞|中心|即時|綜合|報導|頻道|社|網)", clean_meta):
            return clean_meta

    if not article_text or not isinstance(article_text, str):
        return "編輯部"

    clean_text = re.sub(r"\s+", " ", article_text).strip()

    patterns = [
        r"[（\(〔\[【]\s*(?:特派|實習|攝影|駐地)?記者\s*([\u4e00-\u9fa5]{2,4})\s*[\/／\s\-_]*[\u4e00-\u9fa5]*\s*(?:報導|攝影|特稿|專訪)?\s*[）\)〕\]】]",
        r"(?:特派|實習|攝影|駐地)?記者\s*([\u4e00-\u9fa5]{2,4})\s*[\/／]\s*[\u4e00-\u9fa5]{2,6}\s*報導",
        r"(?:文|圖|攝影|撰文|責任編輯|稿源)\s*[:：\/／]\s*([\u4e00-\u9fa5]{2,4})",
        r"[（\(〔\[]\s*([\u4e00-\u9fa5]{2,4})\s*[\/／]\s*(?:彰化|地方|即時|綜合|生活|社會|報導)\s*[:：]?\s*(?:報導)?\s*[）\)〕\]]",
        r"([\u4e00-\u9fa5]{2,4})\s*[\/／]\s*(?:彰化|台中|台北|高雄|地方|即時|綜合|專題|生活)+\s*報導",
        r"【\s*記者\s*([\u4e00-\u9fa5]{2,4})\s*[\/／]",
        r"(?:特派|攝影)?記者\s*([\u4e00-\u9fa5]{2,4})\s*(?:報導|專訪|隨筆)",
    ]

    exclude_words = {
        "彰化", "台中", "台北", "地方", "即時", "綜合", "專題", "社會", "生活",
        "新聞", "家扶", "中心", "本報", "特別", "責任", "編輯", "焦點", "總會", "報導", 
        "公益", "台灣", "全國", "網路", "媒體", "業務", "客服", "版權", "所有"
    }

    prefix_text = clean_text[:600]

    for pattern in patterns:
        matches = re.finditer(pattern, prefix_text)
        for match in matches:
            reporter_name = match.group(1).strip()
            if 2 <= len(reporter_name) <= 4 and reporter_name not in exclude_words:
                return reporter_name

    return "編輯部"


def parse_media_from_url_or_title(title, url, source_elem_text=None, default_media=""):
    """解析媒體名稱，優先使用預設傳入的媒體名稱"""
    if source_elem_text and str(source_elem_text).strip():
        return str(source_elem_text).strip()

    title = str(title) if title else ""
    url = str(url) if url else ""

    domain_map = {
        "owlting.com": "奧丁丁新聞", "886.news": "警政時報", "taichung.news": "台中時報",
        "nantoutimes.com": "南投時報", "pingtungtimes.com.tw": "屏東時報", "taipeipost.org": "台北郵報",
        "marketersgo.com": "行銷人", "gothe.tw": "走遊", "tdn.today": "善思新聞網",
        "ltvnews.net": "在地人新聞", "firenews.com.tw": "火報", "tc.news": "台中新聞網",
        "tn.news": "台灣新聞網", "peopo.org": "PeoPo公民新聞", "cdns.com.tw": "中華日報",
        "ksnews.com.tw": "更生日報", "taiwanhot.net": "台灣好新聞", "ettoday.net": "ETtoday新聞雲",
        "ltn.com.tw": "自由時報", "udn.com": "聯合報", "chinatimes.com": "中國時報",
        "cna.com.tw": "中央社", "pchome.com.tw": "PChome新聞", "yam.com": "蕃新聞",
        "yahoo.com": "Yahoo奇摩新聞", "thehubnews.com.tw": "新頭條", "mch-mediagroup.com": "互傳媒",
        "watchmedia01.com": "觀傳媒", "changhua01.com": "彰化人彰化事", "commerce.media": "商傳媒",
        "today.line.me": "LINE TODAY", "storm.mg": "風傳媒", "setn.com": "三立新聞網",
        "tvbs.com.tw": "TVBS新聞網", "ebc.net.tw": "東森新聞", "ctee.com.tw": "工商時報",
        "nowNews.com": "NOWnews今日新聞", "mirrormedia.mg": "鏡週刊"
    }

    clean_dom = extract_domain(url)
    if clean_dom in domain_map:
        return domain_map[clean_dom]

    for domain_key, name in domain_map.items():
        if domain_key in clean_dom:
            return name

    try:
        match = re.search(r"[\-\|｜\_]\s*([^\-\|｜\_]+)$", title)
        if match:
            possible_media = match.group(1).strip()
            if len(possible_media) <= 12:
                return possible_media
    except Exception:
        pass

    return default_media if default_media else "地方網路新聞"


def parse_pub_year(pub_date_str):
    if not pub_date_str:
        return None
    try:
        dt = parsedate_to_datetime(pub_date_str)
        return dt.year
    except Exception:
        match = re.search(r"\b(20\d{2})\b", pub_date_str)
        if match:
            return int(match.group(1))
    return None


def fetch_google_news_rss(query, target_year=None, default_media=""):
    results = []
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    }
    encoded_query = urllib.parse.quote(query)
    rss_url = f"https://news.google.com/rss/search?q={encoded_query}&hl=zh-TW&gl=TW&ceid=TW:zh-Hant"

    try:
        req = urllib.request.Request(rss_url, headers=headers)
        with urllib.request.urlopen(req, timeout=8, context=ssl_context) as response:
            xml_data = response.read()

        root = ET.fromstring(xml_data)
        for item in root.findall(".//item"):
            try:
                title = item.find("title").text if item.find("title") is not None else ""
                link = item.find("link").text if item.find("link") is not None else ""
                pub_date = item.find("pubDate").text if item.find("pubDate") is not None else ""
                source_elem = item.find("source")
                source_text = source_elem.text if source_elem is not None else ""

                if target_year:
                    item_year = parse_pub_year(pub_date)
                    if item_year and item_year != target_year:
                        continue

                if title and link:
                    media_name = parse_media_from_url_or_title(title, link, source_text, default_media)
                    results.append(
                        {
                            "title": title,
                            "url": link,
                            "media_name": media_name,
                            "date": pub_date,
                        }
                    )
            except Exception:
                continue
    except Exception:
        pass

    return results


def lookup_media_type(media_name, media_map, url=""):
    if url:
        dom = extract_domain(url)
        if dom in media_map:
            return media_map[dom]

    if not media_name:
        return "非三大報全國性"
    
    m_name = str(media_name).strip()
    if m_name in media_map:
        return media_map[m_name]
    for k, v in media_map.items():
        if k in m_name or m_name in k:
            return v
    return "非三大報全國性"


def clean_title_local(title):
    if not title:
        return ""
    try:
        cleaned = re.sub(r"\s*[\-\|｜\_]\s*.*$", "", str(title))
        return cleaned.strip()
    except Exception:
        return str(title)


def trigger_5s_balloon_animation():
    balloon_colors = ["#FF5722", "#2196F3", "#4CAF50", "#FFEB3B", "#9C27B0", "#E91E63"]
    balloons_html = '<div class="balloon-container">'
    for i in range(25):
        left = random.randint(5, 95)
        delay = random.uniform(0, 1.5)
        color = random.choice(balloon_colors)
        balloons_html += f'<div class="balloon" style="left:{left}vw; background-color:{color}; animation-delay:{delay}s;"></div>'
    balloons_html += '</div>'
    
    st.markdown(balloons_html, unsafe_allow_html=True)
    st.balloons()


def process_single_article(item, office, staff_name, org, keyword, media_map):
    """處理單篇新聞的爬取與驗證，確保新聞內文/標題同時滿足 機構名稱 與 關鍵字"""
    cleaned_title = clean_title_local(item["title"])
    media_name = item["media_name"]
    m_type = lookup_media_type(media_name, media_map, item["url"])

    # 擷取真實網頁與 Meta 標籤
    article_snippet, meta_reporter = fetch_article_data(item["url"])
    combined_text = f"標題：{item['title']}\n內文：{article_snippet}"

    short_org = org.replace("彰化", "").replace("中心", "") if "家扶" in org else org

    # 機構名稱與關鍵字雙重檢核
    has_org = (org in cleaned_title) or (org in article_snippet) or (short_org in cleaned_title) or (short_org in article_snippet)
    has_keyword = (keyword in cleaned_title) or (keyword in article_snippet)

    if not (has_org and has_keyword):
        return None

    reporter_name = reporter_detector_sensor_v3(combined_text, meta_reporter)

    return {
        "服務處": office,
        "查報同工": staff_name,
        "媒體名稱": media_name,
        "媒體類別": m_type,
        "新聞標題": cleaned_title,
        "記者": reporter_name,
        "新聞連結": item["url"],
    }


def run_news_pipeline(
    office, staff_name, org, keyword, year, media_map, csv_media_list
):
    st.session_state["search_history"].append(
        {
            "檢索時間": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "服務處": office,
            "同工姓名": staff_name,
            "機構": org,
            "關鍵字": keyword,
            "目標年份": year,
        }
    )

    raw_results = []
    short_org = org.replace("彰化", "").replace("中心", "") if "家扶" in org else org

    # =================================================----------------------
    # 📌 第一階段：全網廣泛關鍵字掃描
    # =================================================----------------------
    first_stage_queries = [
        f'"{org}" "{keyword}"',             
        f'{org} {keyword}',                 
        f'{short_org} {keyword}',            
        f'彰化家扶 {keyword}',
    ]

    with st.spinner(f'🕷️ [第一階段] 正在依「({org}) ({keyword})」進行全網廣泛檢索 (目標年份：{year})...'):
        for q in first_stage_queries:
            res = fetch_google_news_rss(q, target_year=year)
            raw_results.extend(res)

    # =================================================----------------------
    # 📌 第二階段：【100% 強制巡查】database.csv 內的每一家媒體
    # =================================================----------------------
    st.info("🔎 [第二階段] 正在對 database.csv 內的媒體進行 100% 逐一專屬強迫巡查...")
    
    # 預設重點轉載媒體與 CSV 目標融合
    mandatory_targets = [
        ("奧丁丁新聞", "owlting.com"), ("PChome新聞", "pchome.com.tw"),
        ("Yahoo奇摩新聞", "yahoo.com"), ("蕃新聞", "yam.com"),
        ("LINE TODAY", "today.line.me"), ("警政時報", "886.news"),
        ("台中時報", "taichung.news"), ("南投時報", "nantoutimes.com"),
        ("屏東時報", "pingtungtimes.com.tw"), ("台北郵報", "taipeipost.org"),
        ("新頭條", "thehubnews.com.tw"), ("互傳媒", "mch-mediagroup.com"),
        ("觀傳媒", "watchmedia01.com"), ("彰化人彰化事", "changhua01.com"),
        ("走遊", "gothe.tw"), ("台灣好新聞", "taiwanhot.net"),
        ("在地人新聞", "ltvnews.net"), ("善思新聞網", "tdn.today"),
        ("商傳媒", "commerce.media"), ("火報", "firenews.com.tw")
    ]
    
    # 聯集：確保所有 mandatory 與 CSV 內的媒體都被納入
    all_targets = mandatory_targets + (csv_media_list if csv_media_list else [])
    
    # 依「媒體名稱」去重，確保不重複搜尋
    unique_search_targets = {}
    for m, d in all_targets:
        m_clean = m.strip()
        if m_clean and m_clean not in unique_search_targets:
            unique_search_targets[m_clean] = d

    fallback_progress = st.progress(0)
    fallback_status = st.empty()
    total_targets = len(unique_search_targets)

    for i, (m_name, dom) in enumerate(unique_search_targets.items()):
        pct = int((i + 1) / total_targets * 100)
        fallback_status.markdown(f"🔎 深度對點巡查：《**{m_name}**》 ({i+1}/{total_targets})...")
        fallback_progress.progress(pct)

        site_res = []

        # 法一：如果有網域，優先執行 site: 網域搜尋
        if dom:
            site_query = f'site:{dom} {org} {keyword}'
            site_res = fetch_google_news_rss(site_query, target_year=year, default_media=m_name)

        # 法二：如果無網域，或 site: 無結果，【強制】使用 "媒體名稱" + "機構" + "關鍵字" 進行點對點撈取
        if not site_res:
            name_query = f'"{m_name}" "{org}" "{keyword}"'
            site_res = fetch_google_news_rss(name_query, target_year=year, default_media=m_name)
            
            # 備用放寬法
            if not site_res:
                name_query_soft = f'{m_name} {short_org} {keyword}'
                site_res = fetch_google_news_rss(name_query_soft, target_year=year, default_media=m_name)

        for res in site_res:
            # 強制矯正媒體名稱，防止被誤判成「地方網路新聞」
            res["media_name"] = m_name
            raw_results.append(res)

    fallback_progress.empty()
    fallback_status.empty()

    # 僅針對相同 URL 進行網址去重，完整保留轉載新聞
    unique_raw = []
    seen_urls = set()
    for r in raw_results:
        if r["url"] not in seen_urls:
            seen_urls.add(r["url"])
            unique_raw.append(r)
    raw_results = unique_raw

    if not raw_results:
        st.error("❌ 經兩階段檢索後，仍未抓取到相關報導，請檢查關鍵字或目標年份設定。")
        return []

    # =================================================----------------------
    # 📌 第三階段：多線程併行網頁解析與記者探針辨識
    # =================================================----------------------
    results = []
    progress_text_slot = st.empty()
    progress_bar = st.progress(0)
    total_items = len(raw_results)

    st.info("✈️ [第三階段] 啟用多線程併行解析網頁與記者姓名探針...")

    completed = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        future_to_item = {
            executor.submit(process_single_article, item, office, staff_name, org, keyword, media_map): item 
            for item in raw_results
        }
        for future in concurrent.futures.as_completed(future_to_item):
            completed += 1
            percent = int((completed / total_items) * 100)
            progress_text_slot.markdown(f"✈️ **新聞內文解析與記者姓名探針辨識中：{percent}% ({completed}/{total_items})**")
            progress_bar.progress(percent)

            res = future.result()
            if res:
                results.append(res)

    progress_text_slot.empty()
    progress_bar.empty()
    return results


# ---------------------------------------------------------------------------
# 6. UI 介面與主流程控制
# ---------------------------------------------------------------------------
if sidebar_option == "🔍 檢索系統":
    st.markdown('<div class="search-card">', unsafe_allow_html=True)
    st.subheader("🔍 新聞搜尋條件")

    col1, col2 = st.columns(2)
    with col1:
        office = st.selectbox(
            "🏢 選擇服務處：",
            ["全部", "和美兒童館", "員林服務處", "彰化服務處", "二林服務處", "田中服務處"],
        )
        org = st.text_input(
            "🏛️ 搜尋機構名稱：", value="", placeholder="e.g. 彰化家扶"
        )
        year_input = st.text_input(
            "📅 目標年份：", value="", placeholder="e.g. 2026"
        )

    with col2:
        staff_name = st.text_input("👤 主責同工姓名：", value="", placeholder="e.g. 張小明")
        keyword = st.text_input(
            "🔑 搜尋新聞關鍵字：", value="", placeholder="e.g. 課輔班、相見歡、寒冬送暖"
        )

    search_button = st.button("🚀 開始全網檢索與生成報表", use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    if search_button:
        target_org = org.strip() if org.strip() else "彰化家扶"
        
        try:
            clean_year_str = re.sub(r"\D", "", year_input.strip())
            year = int(clean_year_str) if clean_year_str else datetime.date.today().year
        except ValueError:
            year = datetime.date.today().year

        if not keyword.strip() or not staff_name.strip():
            st.warning("⚠️ 請完整填寫「搜尋新聞關鍵字」與「主責同工姓名」！")
        else:
            final_data = run_news_pipeline(
                office, staff_name.strip(), target_org, keyword.strip(), year, media_type_map, csv_media_list
            )

            if final_data:
                df_result = pd.DataFrame(final_data)
                df_result = df_result.drop_duplicates(subset=["新聞連結"])

                st.success(f"🎉 成功捕捉到 {len(df_result)} 筆相關新聞報導！")
                
                trigger_5s_balloon_animation()

                st.dataframe(df_result, use_container_width=True)

                try:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        df_result.to_excel(
                            writer, index=False, sheet_name="新聞輿情統計"
                        )

                    st.download_button(
                        label="📥 下載輿情統計 Excel 報表",
                        data=output.getvalue(),
                        file_name=f"{target_org}_{keyword}_精準輿情報表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"❌ 產出 Excel 報表時發生錯誤：{e}")
            else:
                st.info("ℹ️ 未能找到符合條件的新聞，或過濾後無相關結果。建議擴大關鍵字範圍試試！")

elif sidebar_option == "💡 系統簡介":
    st.subheader("💡 全網廣泛檢索系統")
    st.markdown(
        """
    **彰化家扶中心輿情自動檢索與報表生成系統**旨在幫助同工快速彙整網路媒體報導。

    1. **全開放式彈性檢索**：移除死板的查詢框架，讓搜尋引擎能涵蓋「彰化家扶」、「家扶中心」等變體露出。
    2. **雙軌強迫檢索**：先執行全網廣泛搜尋，再**100% 強制巡查 `database.csv` 內的每一家媒體**，絕不遺漏。
    3. **網域與媒體名稱雙解析**：即便 CSV 無提供網域，也能透過媒體名稱強迫點對點精準搜尋。
    4. **內文探針與簡稱彈性過濾**：自動爬取新聞內文，偵測機構全稱與簡稱，以防止漏抓新聞。
    5. **記者署名識別功能**：涵蓋標準型、括號型、無記者字樣型、複合角色型（文／圖／攝影）等常見新聞署名格式。
    6. **無 API 依賴防爆機制**：100% Python 運算，防止觸發 Google 反爬蟲機制，並免除 API 配置與額度限制。
    """
    )

elif sidebar_option == "📌 系統須知":
    st.subheader("📌 系統須知與使用規範")
    st.success("※本系統已優化「兩階段精準檢索」、「簡稱彈性比對」與「記者姓名探針 V3」，提升抓取量與精準度📈")
    st.warning(
        """
    1. **使用規範**：本系統僅供彰化家扶內部輿情檢索使用，嚴禁用於商業爬蟲或任何非法用途🚫
    2. **資料準確性**：報表匯出後，請務必人工進行二次核對，確保無遺漏✅
    3. **非網路新聞補充**：紙本報紙、電視新聞、廣播節目、社群新聞等露出請務必人工補充✏️
    4. **報導存檔備查**：電子報或社群新聞請下載成PDF檔、YouTube露出下載成JPG檔，放置於查報資料夾中備查📁
    """
    )

elif sidebar_option == "🔐 系統管理員":
    st.subheader("🔐 系統管理員及開發者後台")
    admin_key = st.text_input("🔑 請輸入管理員金鑰：", type="password")

    if admin_key == "Automation_initiator114077":
        st.success("🔓 驗證成功，歡迎進入管理員後台！")
        col_m1, col_m2 = st.columns(2)
        col_m1.metric("📅 今日日期", str(datetime.date.today()))
        col_m2.metric(
            "🔍 累積檢索次數", f"{len(st.session_state['search_history'])} 筆"
        )

        st.markdown("---")
        if st.session_state["search_history"]:
            history_df = pd.DataFrame(st.session_state["search_history"])
            st.dataframe(history_df, use_container_width=True)

            try:
                history_output = io.BytesIO()
                with pd.ExcelWriter(history_output, engine="openpyxl") as writer:
                    history_df.to_excel(
                        writer, index=False, sheet_name="系統使用統計"
                    )

                st.download_button(
                    label="📥 匯出管理員統計報表 (Excel)",
                    data=history_output.getvalue(),
                    file_name=f"系統使用紀錄_{datetime.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"❌ 產出管理員報表失敗：{e}")
        else:
            st.info("目前尚無搜尋歷史紀錄。")
    elif admin_key:
        st.error("❌ 金鑰錯誤！")
