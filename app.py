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

# CSS 樣式 (包含 5 秒氣球動畫)
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
    <p class="warning-text">※本系統為個人自主開發，旨在優化查詢媒體露出流程與準確度，請勿用於非法行為😈</p>
    <p class="warning-text">※已強化全網電子報擷查與撰文記者識別，並加入雙重檢核功能，精準過濾非相關新聞🌏</p>
    <p class="warning-text">※檢索資料庫（database.csv）為「彰化家扶」常見出報媒體清單，開發者將不定期更新👀</p>
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
# 4. 側邊欄與資料庫讀取 (含進度條顯示：📑已完成XX％)
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
db_media_list = []  # 保存 (媒體名稱, Domain) 對照

if os.path.exists(db_file_path):
    try:
        db_df = pd.read_csv(db_file_path, encoding="utf-8").dropna(how="all")
        st.sidebar.success("✅ database.csv 已連線")
        
        if len(db_df) > 0:
            st.sidebar.markdown("⏳ **正在將媒體清單網址/HTML 轉化為 Domain...**")
            domain_progress_bar = st.sidebar.progress(0)
            domain_status_text = st.sidebar.empty()
            
            total_rows = len(db_df)
            media_col = db_df.columns[0]
            type_col = db_df.columns[1] if len(db_df.columns) >= 2 else db_df.columns[0]

            for idx, row in db_df.iterrows():
                m_name = str(row[media_col]).strip()
                m_type = str(row[type_col]).strip() if len(db_df.columns) >= 2 else "非三大報全國性"
                media_type_map[m_name] = m_type

                possible_url_or_html = ""
                if len(row) >= 3 and pd.notna(row.iloc[2]):
                    possible_url_or_html = str(row.iloc[2])
                elif "." in m_name:
                    possible_url_or_html = m_name

                clean_dom = extract_domain(possible_url_or_html)
                if clean_dom:
                    db_domains.append(clean_dom)
                    media_type_map[clean_dom] = m_type
                    db_media_list.append((m_name, clean_dom))

                pct = int((idx + 1) / total_rows * 100)
                domain_progress_bar.progress(pct)
                domain_status_text.markdown(f"📑已完成 **{pct}％**")

            db_domains = list(set(db_domains))
            domain_status_text.markdown(f"📑已完成 **100％** (已載入 {len(db_domains)} 個站點)")
    except Exception as e:
        st.sidebar.error(f"❌ 讀取 database.csv 失敗: {e}")

# ---------------------------------------------------------------------------
# 5. 核心演算法：網頁擷取 + 全能記者探針 + 雙重篩選機制
# ---------------------------------------------------------------------------

def fetch_article_text(url):
    """嘗試取得新聞網頁的內文，用於精準抓取記者姓名與過濾雜訊"""
    if not url or not isinstance(url, str) or not url.startswith(("http://", "https://")):
        return ""

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    }
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=5, context=ssl_context) as response:
            charset = response.headers.get_param("charset") or "utf-8"
            try:
                html = response.read().decode(charset, errors="replace")
            except Exception:
                html = response.read().decode("utf-8", errors="ignore")

            soup = BeautifulSoup(html, "html.parser")

            for script in soup(["script", "style", "noscript", "header", "footer", "nav", "aside"]):
                script.extract()

            text = soup.get_text(separator=" ")
            clean_text = re.sub(r"\s+", " ", text).strip()
            return clean_text[:2000] # 讀取前 2000 字
    except Exception:
        return ""


def reporter_detector_sensor(article_text):
    """
    高精度新聞記者內文探針
    支援常見 5 大署名型態：
    1. 標準型：記者陳雅芳/彰化報導、記者 林明佑／台中報導
    2. 括號型：（記者張小明／彰化報導）、〔記者李四／彰化報導〕
    3. 無「記者」字樣型：陳雅芳／彰化報導、林明佑/彰化報導
    4. 複合角色型：文／陳雅芳、圖／林明佑、攝影記者 張小明
    """
    if not article_text or not isinstance(article_text, str):
        return "編輯部"

    clean_text = re.sub(r"\s+", " ", article_text).strip()

    patterns = [
        # 1. 帶「記者/攝影記者」+ 姓名 + 地名/類別 + 報導
        r"(?:攝影)?記者\s*([\u4e00-\u9fa5]{2,4})\s*[\/／\s]\s*[\u4e00-\u9fa5]{2,6}\s*報導",
        
        # 2. 括號/方括號包覆 (記者張小明／彰化報導) / 〔記者李四／彰化報導〕
        r"[（\(〔\[]\s*(?:攝影)?記者\s*([\u4e00-\u9fa5]{2,4})\s*[\/／\s]*[\u4e00-\u9fa5]*\s*報導\s*[）\)〕\]]",
        
        # 3. 複合角色型：文／陳雅芳、圖／林明佑、文/圖 陳雅芳
        r"(?:文|圖|攝影|撰文|責任編輯)\s*[\/／]\s*([\u4e00-\u9fa5]{2,4})",
        
        # 4. 無「記者」二字直接接地名報導 (例如：陳雅芳／彰化報導)
        r"(?<!新聞)(?<!中心)(?<!家扶)\b([\u4e00-\u9fa5]{2,4})\s*[\/／]\s*(?:彰化|台中|台北|高雄|地方|即時|綜合|專題|生活)+\s*報導",
        
        # 5. 單純「記者 姓名 報導」
        r"(?:攝影)?記者\s*([\u4e00-\u9fa5]{2,4})\s*報導",
    ]

    exclude_words = {
        "彰化", "台中", "台北", "地方", "即時", "綜合", "專題", "社會", "生活",
        "新聞", "家扶", "中心", "本報", "特別", "責任", "編輯", "焦點", "總會", "報導"
    }

    for pattern in patterns:
        matches = re.finditer(pattern, clean_text)
        for match in matches:
            reporter_name = match.group(1).strip()
            if 2 <= len(reporter_name) <= 4 and reporter_name not in exclude_words:
                return reporter_name

    return "編輯部"


def parse_media_from_url_or_title(title, url, source_elem_text=None):
    """本地辨識媒體名稱 (結合 Domain 與標題對照)"""
    title = str(title) if title else ""
    url = str(url) if url else ""

    if source_elem_text and str(source_elem_text).strip():
        return str(source_elem_text).strip()

    domain_map = {
        "owlting.com": "奧丁丁新聞", "886.news": "警政時報", "taichung.news": "台中時報",
        "nantoutimes.com": "南投時報", "pingtungtimes.com.tw": "屏東時報", "taipeipost.org": "台北郵報",
        "marketersgo.com": "行銷人", "gothe.tw": "走遊", "tdn.today": "善思新聞網",
        "ltvnews.net": "在地人新聞", "firenews.com.tw": "火報", "tc.news": "台中新聞網",
        "tn.news": "台灣新聞網", "peopo.org": "PeoPo公民新聞", "cdns.com.tw": "中華日報",
        "ksnews.com.tw": "更生日報", "taiwanhot.net": "台灣好新聞", "ettoday.net": "ETtoday新聞雲",
        "ltn.com.tw": "自由時報", "udn.com": "聯合報", "chinatimes.com": "中國時報",
        "cna.com.tw": "中央社", "pchome.com.tw": "PChome新聞", "yam.com": "蕃新聞",
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
        match = re.search(r"\b(20\d{2})\b", pub_date_str)
        if match:
            return int(match.group(1))
    return None


def fetch_google_news_rss(query, target_year=None):
    """標準 Google News RSS 檢索"""
    results = []
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
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
    except Exception:
        pass

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


def trigger_5s_balloon_animation():
    """發射 5 秒升空氣球動畫"""
    balloon_colors = ["#FF5722", "#2196F3", "#4CAF50", "#FFEB3B", "#9C27B0", "#E91E63"]
    balloons_html = '<div class="balloon-container">'
    for i in range(25):
        left = random.randint(5, 95)
        delay = random.uniform(0, 1.5)
        color = random.choice(balloon_colors)
        balloons_html += f'<div class="balloon" style="left:{left}vw; background-color:{color}; animation-delay:{delay}s;"></div>'
    balloons_html += '</div>'
    
    st.markdown(balloons_html, unsafe_allow_html=True)
    st.balloons()  # Streamlit 原生氣球補強


def run_news_pipeline(
    office, staff_name, org, keyword, year, media_map, db_domains, db_media_list, GEMINI_API_KEY
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

    raw_results = []

    # -----------------------------------------------------------------------
    # 第一階段：全網 Google 搜尋 「"機構名稱" "新聞關鍵字"」
    # -----------------------------------------------------------------------
    with st.spinner(f'🕷️ [階段一] 正在執行 Google 全網報導檢索：『"{org}" "{keyword}"』 (目標年份：{year})...'):
        primary_query = f'"{org}" "{keyword}"'
        raw_results = fetch_google_news_rss(primary_query, target_year=year)

    # -----------------------------------------------------------------------
    # 第二階段：備援邏輯 (若全網查無結果，自動對媒體清單進行站內搜尋 site:domain)
    # -----------------------------------------------------------------------
    if not raw_results and (db_domains or db_media_list):
        st.warning("⚠️ 全網搜尋未發現直接報導，自動啟用「媒體站內精準二次檢索 (Fallback Logic)」...")
        
        fallback_progress = st.progress(0)
        fallback_status = st.empty()
        
        # 以 db_media_list 或 db_domains 逐一/分批進行站內搜尋
        search_targets = db_media_list if db_media_list else [(f"站點{i}", d) for i, d in enumerate(db_domains)]
        total_targets = len(search_targets)

        for i, (m_name, dom) in enumerate(search_targets):
            pct = int((i + 1) / total_targets * 100)
            # 修正處：將內部雙引號改為單引號或用斜線轉義
            fallback_status.markdown(f"🔎 正在站內搜索《{m_name}》 ({dom})：`site:{dom} '{org}' '{keyword}'`...")
            fallback_progress.progress(pct)

            site_query = f'site:{dom} "{org}" "{keyword}"'
            site_res = fetch_google_news_rss(site_query, target_year=year)
            
            for res in site_res:
                res["media_name"] = m_name if m_name else res["media_name"]
                raw_results.append(res)

            time.sleep(random.uniform(0.1, 0.3))

        fallback_progress.empty()
        fallback_status.empty()

    # 網址去重
    unique_raw = []
    seen_urls = set()
    for r in raw_results:
        if r["url"] not in seen_urls:
            seen_urls.add(r["url"])
            unique_raw.append(r)
    raw_results = unique_raw

    if not raw_results:
        st.error("❌ 經全網與媒體站內二次檢索後，仍未抓取到相關報導，請檢查關鍵字或年份設定。")
        return []

    # -----------------------------------------------------------------------
    # 第三階段：新聞內文抓取 + 雙重過濾 + 記者探針
    # -----------------------------------------------------------------------
    client = None
    if genai and GEMINI_API_KEY:
        try:
            client = genai.Client(api_key=GEMINI_API_KEY)
        except Exception:
            pass

    results = []
    progress_text_slot = st.empty()
    progress_bar = st.progress(0)
    total_items = len(raw_results)

    for i, item in enumerate(raw_results):
        percent = int((i + 1) / total_items * 100)
        progress_text_slot.markdown(f"✈️ **新聞深度解析、雙重過濾與記者探針辨識中：{percent}%**")
        progress_bar.progress(percent)

        cleaned_title = clean_title_local(item["title"])
        media_name = item["media_name"]
        m_type = lookup_media_type(media_name, media_map, item["url"])

        # 1. 抓取新聞內文前段 (抓取內文解決自動搜尋抓不到/標題無關鍵字的問題)
        article_snippet = fetch_article_text(item["url"])
        combined_text = f"標題：{item['title']}\n內文：{article_snippet}"

        # 2. 嚴格雙重過濾 (標題或內文必須「同時包含」機構名稱與關鍵字，提升抓取精準度)
        has_org = (org in cleaned_title) or (org in article_snippet)
        has_keyword = (keyword in cleaned_title) or (keyword in article_snippet)

        if not (has_org and has_keyword):
            continue

        # 3. 執行全能記者探針 (支援 5 大常見署名型態)
        reporter_name = reporter_detector_sensor(combined_text)

        # 4. Gemini AI 輔助校準 (若 API 金鑰可用)
        is_relevant = True
        if client:
            try:
                st.session_state["api_count_today"] += 1
                prompt = f"""
                你是一個專業新聞輿情分析助手。請分析以下新聞內容：
                {combined_text[:1500]}

                請執行以下任務：
                1. 判斷這篇新聞是否確實報導了「{org}」以及與「{keyword}」相關事件？ (填寫 true 或 false)
                2. 清理新聞標題，去除媒體名稱或頻道後綴。
                3. 精準辨識撰稿記者/作者姓名 (若內文無法辨識則填 '{reporter_name}')。

                請傳回 JSON 格式：
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
            "🏛️ 搜尋機構名稱：", value="彰化家扶", placeholder="e.g. 彰化家扶"
        )
        year_input = st.text_input(
            "📅 目標年份：", value=str(datetime.date.today().year), placeholder=f"e.g. {datetime.date.today().year}"
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
                office, staff_name.strip(), target_org, keyword.strip(), year, media_type_map, db_domains, db_media_list, api_key
            )

            if final_data:
                df_result = pd.DataFrame(final_data)
                df_result = df_result.drop_duplicates(subset=["新聞連結"])

                st.success(f"🎉 成功捕捉到 {len(df_result)} 筆高品質相關新聞！")
                
                # 🎈 觸發 5 秒升空氣球動畫
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
    st.subheader("💡 全網小報檢索系統特點")
    st.markdown(
        """
    **彰化家扶中心輿情自動檢索與報表生成系統**旨在幫助同工快速彙整網路媒體報導。

    * **雙軌備援檢索**：先執行全網搜尋，若無結果則自動進入 `database.csv` 媒體清單進行 `site:` 站內搜尋。
    * **網域自動轉換與進度顯示**：自動解析 HTML 或網址轉為純 Domain，並即時顯示 `📑已完成XX％` 進度條。
    * **內文探針與雙重過濾**：自動爬取新聞內文前段，檢測標題與內文是否同時符合「機構名稱」與「關鍵字」，擺脫無關新聞。
    * **全能記者署名識別**：涵蓋標準型、括號型、無記者字樣型、複合角色型（文／圖／攝影）等 5 大新聞署名格式。
    * **互動視覺體驗**：完成擷取後自動發射 5 秒升空氣球動畫，增加系統活潑度與互動性。
    """
    )

elif sidebar_option == "📌 系統須知":
    st.subheader("📌 系統須知與使用規範")
    st.success("※本版本已升級「站內備援檢索」、「內文雙重檢測」與「全能記者探針」，精準度全面提升📈")
    st.warning(
        """
    1. **使用規範**：本系統僅供彰化家扶內部輿情檢索使用，嚴禁用於商業爬蟲或任何非法用途！
    2. **雙保險機制**：系統優先採用 Gemini Flash 模型與內文爬取；若網頁被 Google 防爬蟲機制阻擋或 API 額滿，會自動轉化為 Python 防爆演算。
    3. **資料準確性**：報表匯出後，請務必人工進行二次核對，確保無遺漏。
    4. **非網路新聞補充**：紙本報紙、電視新聞、廣播節目、社群新聞等露出請務必人工補充。
    5. **報導存檔備查**：電子報或社群新聞請下載成PDF檔、YouTube露出下載成JPG檔，放置於查報資料夾中備查。
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
