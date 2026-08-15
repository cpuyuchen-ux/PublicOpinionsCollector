import datetime
import io
import json
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
# 1. 套件載入與環境防呆
# ---------------------------------------------------------------------------
try:
    from bs4 import BeautifulSoup
except ImportError:
    st.error("❌ 系統缺少 'bs4' 套件！請在終端機執行：pip install beautifulsoup4")
    st.stop()

try:
    import openpyxl
except ImportError:
    st.error("❌ 系統缺少 'openpyxl' 套件（匯出 Excel 必備）！請在終端機執行：pip install openpyxl")
    st.stop()

try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None

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
if "api_count_today" not in st.session_state:
    st.session_state["api_count_today"] = 0
if "last_api_date" not in st.session_state:
    st.session_state["last_api_date"] = datetime.date.today()
if "search_history" not in st.session_state:
    st.session_state["search_history"] = []

# 每日 API 計算器重置
if st.session_state["last_api_date"] != datetime.date.today():
    st.session_state["api_count_today"] = 0
    st.session_state["last_api_date"] = datetime.date.today()

# CSS 樣式
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
    div[data-testid="stRadio"] div[role="radiogroup"] > label > div:first-child {
        display: none !important;
    }
    div[data-testid="stRadio"] div[role="radiogroup"] > label {
        padding: 6px 10px;
        border-radius: 8px;
        transition: background-color 0.2s ease;
        margin-bottom: 4px;
        cursor: pointer;
    }
    div[data-testid="stRadio"] div[role="radiogroup"] > label p {
        font-size: 1.1rem !important;
        font-weight: 600 !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# 2. 標題與警示橫幅區塊
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="main-header">📰 彰化家扶中心輿情自動檢索與報表生成系統</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-header">支援全網頁小報無 API 本地深度檢索、雙重關聯過濾與記者精準辨識</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="warning-bar">
    <p class="warning-text">※本系統為個人自主開發，意在優化查報工作流程與準確度，請勿用於非法行為😈</p>
    <p class="warning-text">※已強化全網電子報擷取與記者識別演算法，加入雙重檢核功能，精準過濾無關新聞🌏</p>
    <p class="warning-text">※檢索資料庫（database.csv）為「彰化家扶」常見出報媒體清單，資料庫將不定期更新👀</p>
    <p class="warning-text">※開發者保有此系統所有權，敬請尊重開發者之權利。若有不法，將依中華民國相關法規追究🔧</p>
</div>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# 3. 核心工具：自動把網址/HTML/字串轉成純 Domain
# ---------------------------------------------------------------------------

def extract_domain(url_or_html):
    """
    將傳入的網址、HTML標籤或字串自動清理為純粹的主 Domain 
    (例如：<a href="https://news.ltn.com.tw/"> -> ltn.com.tw)
    """
    if not url_or_html or not isinstance(url_or_html, str):
        return ""
    
    text = url_or_html.strip()
    # 若包含 HTML 標籤，嘗試解析出 href 或純文字
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

    # 去除首尾空白與 http://, https://
    domain = re.sub(r"^https?://", "", text.strip(), flags=re.IGNORECASE)
    # 截斷第一個斜線 / 之後的所有路徑與參數
    domain = domain.split('/')[0].split('?')[0].split('#')[0]
    # 去除前綴如 www. 或 news.
    domain = re.sub(r"^(www\.|news\.)", "", domain, flags=re.IGNORECASE)
    return domain.lower()

# ---------------------------------------------------------------------------
# 4. 側邊欄與資料庫讀取 (讀取 database.csv 與自動將 C 欄 HTML 轉 Domain)
# ---------------------------------------------------------------------------
st.sidebar.title("⚙️ 系統功能導覽")
sidebar_option = st.sidebar.radio(
    "請選擇功能模組：",
    ["🔍 檢索系統", "💡 系統簡介", "📌 系統須知", "🔐 系統管理員"],
    index=0,
    label_visibility="collapsed",
)

st.sidebar.markdown("---")
api_key = st.secrets.get("GEMINI_API_KEY", "")
if not api_key:
    api_key = st.sidebar.text_input(
        "🔑 輸入 Gemini API Key:",
        type="password",
        help="請輸入您的 Gemini API Key",
    )

st.sidebar.markdown("---")
db_file_path = "database.csv"
media_type_map = {}
db_domains = []

if os.path.exists(db_file_path):
    try:
        # A欄: 媒體名稱, B欄: 三大報全國性/非三大報全國性, C欄: HTML
        db_df = pd.read_csv(db_file_path, encoding="utf-8").dropna(how="all")
        st.sidebar.success("✅ database.csv 已連線")
        if len(db_df.columns) >= 2:
            media_col = db_df.columns[0]
            type_col = db_df.columns[1]
            
            for _, row in db_df.iterrows():
                m_name = str(row[media_col]).strip()
                m_type = str(row[type_col]).strip()
                media_type_map[m_name] = m_type
                
                # 自動將 C 欄 HTML 轉成純 Domain
                possible_url_or_html = ""
                if len(row) >= 3 and pd.notna(row.iloc[2]):
                    possible_url_or_html = str(row.iloc[2])
                elif "." in m_name:
                    possible_url_or_html = m_name
                    
                clean_dom = extract_domain(possible_url_or_html)
                if clean_dom:
                    db_domains.append(clean_dom)
                    media_type_map[clean_dom] = m_type

            db_domains = list(set(db_domains))
            if db_domains:
                st.sidebar.info(f"🎯 已鎖定資料庫 {len(db_domains)} 個媒體站內檢索點")
    except Exception as e:
        st.sidebar.error(f"❌ 讀取 database.csv 失敗: {e}")

# ---------------------------------------------------------------------------
# 5. 核心演算法：HTTP Fetch + 雙重過濾 + 精準記者探針 + 目標年份過濾
# ---------------------------------------------------------------------------

def fetch_article_text(url):
    """嘗試取得新聞網頁的前段內文，用於精準抓取記者姓名與過濾雜訊"""
    if not url or not isinstance(url, str) or not url.startswith(("http://", "https://")):
        return ""

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    }
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=4, context=ssl_context) as response:
            charset = response.headers.get_param("charset") or "utf-8"
            try:
                html = response.read().decode(charset, errors="replace")
            except Exception:
                html = response.read().decode("utf-8", errors="ignore")

            soup = BeautifulSoup(html, "html.parser")

            for script in soup(["script", "style", "noscript", "header", "footer", "nav"]):
                script.extract()

            text = soup.get_text(separator=" ")
            clean_text = re.sub(r"\s+", " ", text).strip()
            return clean_text[:1200]
    except Exception:
        return ""


def reporter_detector_sensor(article_text):
    """
    高精度新聞記者內文探針
    支援格式：
    - 記者陳雅芳/彰化報導、記者 林明佑／台中報導
    - (記者陳雅芳/彰化報導)、〔記者林明佑/彰化報導〕
    - 陳雅芳/彰化報導、林明佑／即時報導
    - 文/陳雅芳、圖／林明佑
    """
    if not article_text or not isinstance(article_text, str):
        return "編輯部"

    clean_text = re.sub(r"\s+", " ", article_text).strip()

    patterns = [
        # 1. 包含「記者」+ 姓名 + 斜線 + 地名/類別 + 報導
        r"記者\s*([\u4e00-\u9fa5]{2,4})\s*[\/／]\s*[\u4e00-\u9fa5]{2,6}\s*報導",
        
        # 2. 括號/方括號包覆
        r"[（\(〔\[]\s*記者\s*([\u4e00-\u9fa5]{2,4})\s*[\/／\s]*[\u4e00-\u9fa5]*\s*報導\s*[）\)〕\]]",
        
        # 3. 無「記者」二字直接接地名報導
        r"(?<!新聞)(?<!中心)(?<!家扶)\b([\u4e00-\u9fa5]{2,4})\s*[\/／]\s*(?:彰化|台中|台北|高雄|地方|即時|綜合|專題)+\s*報導",
        
        # 4. 單純「記者 姓名 報導」
        r"記者\s*([\u4e00-\u9fa5]{2,4})\s*報導",
        
        # 5. 文/圖/攝影 署名
        r"(?:文|圖|攝影|責任編輯)\s*[\/／]\s*([\u4e00-\u9fa5]{2,4})"
    ]

    exclude_words = {
        "彰化", "台中", "台北", "地方", "即時", "綜合", "專題", "社會", "生活",
        "新聞", "家扶", "中心", "本報", "特別", "責任", "編輯", "焦點", "總會"
    }

    for pattern in patterns:
        matches = re.finditer(pattern, clean_text)
        for match in matches:
            reporter_name = match.group(1).strip()
            if 2 <= len(reporter_name) <= 4 and reporter_name not in exclude_words:
                return reporter_name

    return "編輯部"


def parse_media_from_url_or_title(title, url, source_elem_text=None):
    """本地辨識媒體名稱 (結合 Domain 判讀)"""
    title = str(title) if title else ""
    url = str(url) if url else ""

    if source_elem_text and str(source_elem_text).strip():
        return str(source_elem_text).strip()

    domain_map = {
        "owlting.com": "奧丁丁新聞",
        "886.news": "警政時報",
        "taichung.news": "台中時報",
        "nantoutimes.com": "南投時報",
        "pingtungtimes.com.tw": "屏東時報",
        "taipeipost.org": "台北郵報",
        "marketersgo.com": "行銷人",
        "gothe.tw": "走遊",
        "tdn.today": "善思新聞網",
        "ltvnews.net": "在地人新聞",
        "firenews.com.tw": "火報",
        "tc.news": "台中新聞網",
        "tn.news": "台灣新聞網",
        "peopo.org": "PeoPo公民新聞",
        "cdns.com.tw": "中華日報",
        "ksnews.com.tw": "更生日報",
        "taiwanhot.net": "台灣好新聞",
        "ettoday.net": "ETtoday新聞雲",
        "ltn.com.tw": "自由時報",
        "udn.com": "聯合報",
        "chinatimes.com": "中國時報",
        "cna.com.tw": "中央社",
        "pchome.com.tw": "PChome新聞",
        "yam.com": "蕃新聞",
        "yahoo.com": "Yahoo奇摩新聞",
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

    return "地方網路新聞"


def parse_pub_year(pub_date_str):
    """解析 RSS 發布日期字串並提取年份"""
    if not pub_date_str:
        return None
    try:
        dt = parsedate_to_datetime(pub_date_str)
        return dt.year
    except Exception:
        # 正則表達式抓取 4 位數年份
        match = re.search(r"\b(20\d{2})\b", pub_date_str)
        if match:
            return int(match.group(1))
    return None


def fetch_google_news_rss(org, keyword, site_domains=None, target_year=None):
    """
    ⚡ 高穩定 Google News RSS 檢索引擎
    避開反爬蟲：針對 database 提供之媒體 Domain 進行站內搜尋 (Site-Search)，
    並加入微幅延遲與目標年份 (target_year) 過濾機制。
    """
    results = []
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    }

    # 建立搜尋 Query 列表
    queries = []
    if site_domains and len(site_domains) > 0:
        # 分批 site:domain，每組 5 個，避免 URL 過長引起 Google 429 阻擋
        chunk_size = 5
        for i in range(0, len(site_domains), chunk_size):
            chunk = site_domains[i:i + chunk_size]
            valid_sites = [f"site:{extract_domain(d)}" for d in chunk if extract_domain(d)]
            if valid_sites:
                sites_query = " OR ".join(valid_sites)
                queries.append(f'"{org}" "{keyword}" ({sites_query})')
    else:
        queries.append(f'"{org}" "{keyword}"')

    for q in queries:
        encoded_query = urllib.parse.quote(q)
        rss_url = f"https://news.google.com/rss/search?q={encoded_query}&hl=zh-TW&gl=TW&ceid=TW:zh-Hant"

        try:
            req = urllib.request.Request(rss_url, headers=headers)
            with urllib.request.urlopen(req, timeout=10, context=ssl_context) as response:
                xml_data = response.read()

            root = ET.fromstring(xml_data)
            for item in root.findall(".//item"):
                try:
                    title = item.find("title").text if item.find("title") is not None else ""
                    link = item.find("link").text if item.find("link") is not None else ""
                    pub_date = item.find("pubDate").text if item.find("pubDate") is not None else ""
                    source_elem = item.find("source")
                    source_text = source_elem.text if source_elem is not None else ""

                    # 🎯 目標年份 Filter：過濾掉非當年度的報導
                    if target_year:
                        item_year = parse_pub_year(pub_date)
                        if item_year and item_year != target_year:
                            continue

                    if title and link:
                        media_name = parse_media_from_url_or_title(title, link, source_text)
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
        except Exception as e:
            st.warning(f"⚠️ 檢索出現微幅異常 (可能受到頻率限制)：{e}")

        # 避開反爬蟲微幅休眠
        if len(queries) > 1:
            time.sleep(random.uniform(0.3, 0.8))

    return results


def lookup_media_type(media_name, media_map, url=""):
    """對照媒體類別 (支援名稱與 Domain 對照)"""
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
    """標題清理 (去除網站後綴)"""
    if not title:
        return ""
    try:
        cleaned = re.sub(r"\s*[\-\|｜\_]\s*.*$", "", str(title))
        return cleaned.strip()
    except Exception:
        return str(title)


def run_news_pipeline(
    office, staff_name, org, keyword, year, media_map, db_domains, GEMINI_API_KEY
):
    # 記錄檢索歷史
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

    # 1. 啟動第一階段：全網精準 RSS 檢索 (帶入目標年份過濾)
    with st.spinner(f"🕷️ [第一階段] 正在搜羅全網新聞報導『"{org}" "{keyword}"』 (目標年份：{year})..."):
        raw_results = fetch_google_news_rss(org, keyword, target_year=year)

    # 2. 自動二次探針：針對 database.csv 提供的媒體 Domain 逐一/分批進行站內搜尋 (Site-Search)
    if db_domains:
        st.info("ℹ️ 啟動 database.csv 媒體網域站內二次檢索 (Site-Search)...")
        with st.spinner("🔎 [第二階段] 針對指定媒體 Domain 進行站內精準檢索與年份過濾..."):
            site_results = fetch_google_news_rss(org, keyword, site_domains=db_domains, target_year=year)
            raw_results.extend(site_results)

    # 去重初步結果
    unique_raw = []
    seen_urls = set()
    for r in raw_results:
        if r["url"] not in seen_urls:
            seen_urls.add(r["url"])
            unique_raw.append(r)
    raw_results = unique_raw

    if not raw_results:
        st.error("❌ 經全網與媒體站內二次檢索後，仍未抓取到相關報導，請更換關鍵字。")
        return []

    # 3. 初始化 Gemini Client
    client = None
    if genai and GEMINI_API_KEY:
        try:
            client = genai.Client(api_key=GEMINI_API_KEY)
        except Exception as e:
            st.sidebar.warning(f"⚠️ Gemini 初始化失敗：{e}")

    results = []
    progress_text_slot = st.empty()
    progress_bar = st.progress(0)
    total_items = len(raw_results)

    for i, item in enumerate(raw_results):
        percent = int((i + 1) / total_items * 100)
        progress_text_slot.markdown(f"✈️ **新聞深度解析與精準記者探針辨識中：{percent}%**")
        progress_bar.progress(percent)

        cleaned_title = clean_title_local(item["title"])
        media_name = item["media_name"]
        m_type = lookup_media_type(media_name, media_map, item["url"])

        # 抓取內文前段
        article_snippet = fetch_article_text(item["url"])
        combined_text = f"標題：{item['title']}\n內文開頭：{article_snippet}"

        # 本地硬過濾 (嚴格雙重比對)
        if (org not in cleaned_title and org not in article_snippet) and \
           (keyword not in cleaned_title and keyword not in article_snippet):
            continue

        # 執行精準記者探針
        reporter_name = reporter_detector_sensor(combined_text)

        # Gemini AI 語意檢核 (若 API 可用)
        is_relevant = True
        if client:
            try:
                st.session_state["api_count_today"] += 1
                prompt = f"""
                你是一個新聞輿情分析助手。請分析以下新聞內容：
                {combined_text}

                請執行以下任務：
                1. 判斷這篇新聞是否與「{org}」以及「{keyword}」高度相關？ (填寫 true 或 false)
                2. 清理新聞標題，移除媒體名稱、頻道或來源後綴。
                3. 辨識記者/撰稿人姓名 (若無則填 '{reporter_name}')。

                傳回 JSON 格式：
                {{"is_relevant": true, "title": "純標題", "reporter": "記者姓名"}}
                """
                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json"
                    ),
                )
                
                raw_json = response.text.strip()
                if raw_json.startswith("```json"):
                    raw_json = raw_json.split("```json")[1].split("```")[0].strip()
                elif raw_json.startswith("```"):
                    raw_json = raw_json.split("```")[1].split("```")[0].strip()

                parsed = json.loads(raw_json)
                is_relevant = parsed.get("is_relevant", True)
                cleaned_title = parsed.get("title", cleaned_title)
                reporter_name = parsed.get("reporter", reporter_name)
            except Exception:
                pass

        if not is_relevant:
            continue

        results.append(
            {
                "服務處": office,
                "查報同工": staff_name,
                "媒體名稱": media_name,
                "媒體類別": m_type,
                "新聞標題": cleaned_title,
                "記者": reporter_name,
                "新聞連結": item["url"],
            }
        )

    progress_text_slot.empty()
    progress_bar.empty()
    return results


# ---------------------------------------------------------------------------
# 6. UI 與主流程控制
# ---------------------------------------------------------------------------
if sidebar_option == "🔍 檢索系統":
    st.markdown('<div class="search-card">', unsafe_allow_html=True)
    st.subheader("🔍 新聞輿情搜尋條件")

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
            "📅 目標年份：", value="", placeholder=f"e.g. {datetime.date.today().year}"
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
                office, staff_name.strip(), target_org, keyword.strip(), year, media_type_map, db_domains, api_key
            )

            if final_data:
                df_result = pd.DataFrame(final_data)
                df_result = df_result.drop_duplicates(subset=["新聞連結"])

                st.success(f"🎉 成功捕捉到 {len(df_result)} 筆高品質新聞！")
                st.balloons()
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
    st.subheader("💡 全網小報檢索系統特點")
    st.markdown(
        """
    **彰化家扶中心輿情自動檢索與報表生成系統**旨在幫助同工快速彙整網路媒體報導。

    * **即時檢索**：自動爬取 Google 最新相關新聞與網頁報導。
    * **網域自動轉換 (Domain Extractor)**：不論資料庫 C 欄輸入何種 HTML 格式或網址，自動轉為標準主網域執行 `site:` 精準定位。
    * **精準記者探針 (Reporter Sensor)**：內建高精準度新聞內文正規表示式探針，自動識別專利署名格式（如：`記者陳雅芳/彰化報導`）。
    * **雙重過濾與年份篩選**：自動採集網頁內文與標題，透過嚴格比對與目標年份 Filter 排除非當年度或無關新聞雜訊。
    * **一鍵報表**：自動產出包含服務處、主責查詢同工、媒體分類與超連結的標準化 Excel 檔案。
    """
    )

elif sidebar_option == "📌 系統須知":
    st.subheader("📌 系統須知與使用規範")
    st.success("※本版本已升級「 Domain 自動轉譯」與「精準記者內文探針」，大幅提高記者命中率與檢索品質📈")
    st.warning(
        """
    1. **遵守使用規範**：本系統僅供彰化家扶內部輿情檢索使用，嚴禁用於商業爬蟲或任何非法用途！
    2. **雙保險機制**：系統優先採用 Gemini Flash 模型與內文爬取；若網頁被防爬蟲阻擋或 API 額滿，會自動降級至本地精準 Sensor 演算法！
    3. **資料準確性**：報表匯出後，請人工進行二次核對，確保無遺漏。
    4. **非網路新聞補充**：紙本報紙、電視新聞等露出請務必人工補充。
    """
    )

elif sidebar_option == "🔐 系統管理員":
    st.subheader("🔐 系統管理員後台")
    admin_key = st.text_input("🔑 請輸入管理員金鑰：", type="password")

    if admin_key == "Automation_initiator114077":
        st.success("🔓 驗證成功，歡迎進入管理員後台！")
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("📅 今日日期", str(st.session_state["last_api_date"]))
        col_m2.metric(
            "📡 今日 API 請求次數", f"{st.session_state['api_count_today']} 次"
        )
        col_m3.metric(
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
