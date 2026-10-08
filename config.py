"""
Central configuration for the Trip.com unmanned/automated vs. conventional hotel review study.

Everything that a researcher might want to change (paths, stay conditions, politeness delays,
keyword dictionaries, matching weights, CSS selectors) lives here so the methodology is
documented in one place.
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

# ------------------------------------------------------------------ paths
ROOT = Path(__file__).resolve().parent
INPUT_DIR = ROOT / "inputs"
OUTPUT_DIR = ROOT / "outputs"
STAGE_DIR = OUTPUT_DIR / "stages"
WORK_DIR = ROOT / "work"
DEBUG_DIR = WORK_DIR / "debug"
META_CACHE_DIR = WORK_DIR / "hotel_meta"          # one JSON per Trip.com hotel (resume)
REVIEW_DIR = WORK_DIR / "reviews"                 # one JSONL per hotel_id (resume)
RAW_RESPONSE_DIR = WORK_DIR / "raw_responses"     # raw captured JSON (reproducibility)
DATASET_DIR = ROOT / "dataset"
IMAGE_DIR = DATASET_DIR / "images"

# Source workbooks (read-only — never overwritten)
UNMANNED_SOURCE_NAME = "호텔_표본_재구성_PhysicalAI100_매칭계획.xlsx"
CONVENTIONAL_SOURCE_NAME = "유인호텔_리스트.xlsx"
MANUAL_CONTROL_CANDIDATES_NAME = "manual_control_candidates.xlsx"   # optional, see README

# Pipeline artefacts
FINAL_UNMANNED = OUTPUT_DIR / "01_final_unmanned_automated_100.xlsx"
FINAL_CONVENTIONAL = OUTPUT_DIR / "02_final_conventional_100.xlsx"
FINAL_PAIRS = OUTPUT_DIR / "03_final_hotel_pairs_100.xlsx"
DRAFT_CONVENTIONAL = OUTPUT_DIR / "02_conventional_matched_DRAFT.xlsx"
DRAFT_PAIRS = OUTPUT_DIR / "03_hotel_pairs_DRAFT.xlsx"
REVIEWS_CSV = OUTPUT_DIR / "04_tripcom_reviews_raw.csv"
CRAWL_LOG = OUTPUT_DIR / "05_crawling_log.xlsx"
SAMPLE_SUMMARY = OUTPUT_DIR / "06_sample_summary.xlsx"

STAGE1_CANDIDATES = STAGE_DIR / "S1_unmanned_candidates_validated.xlsx"
STAGE1_CONTROLS = STAGE_DIR / "S1_control_pool_validated.xlsx"
STAGE2_MATCHING_LOG = STAGE_DIR / "S2_matching_log.xlsx"
PAIRS_CURRENT_CSV = WORK_DIR / "pairs_current.csv"          # what the crawler reads
CRAWL_STATE = WORK_DIR / "crawl_state.json"
IMAGE_MANIFEST = WORK_DIR / "image_manifest.csv"
RUN_CONFIG = WORK_DIR / "run_config.json"                    # frozen stay dates etc.

# ------------------------------------------------------------------ sample design
N_PER_GROUP = 100
MAX_REVIEWS_PER_HOTEL = 100
MAX_IMAGES_PER_REVIEW = 10

GROUP_UNMANNED = "UNMANNED_AUTOMATED"
GROUP_CONVENTIONAL = "CONVENTIONAL"
TYPE_PA = "PHYSICAL_AI"
TYPE_AO = "AUTOMATION_ONLY"
TYPE_CONV = "CONVENTIONAL"
TYPE_UNCERTAIN = "UNCERTAIN"

# Stay conditions used for every price lookup. The dates are frozen into work/run_config.json
# on the first run so that a resumed run keeps using identical conditions.
DEFAULT_CHECKIN_OFFSET_DAYS = 28        # ~4 weeks ahead, avoids same-day price noise
STAY_NIGHTS = 1
ADULTS = 2
ROOMS = 1
CHILDREN = 0

# ------------------------------------------------------------------ politeness / robustness
HEADLESS = True
LOCALE = "en-US"
TIMEZONE = "Asia/Seoul"
NAV_TIMEOUT_MS = 60_000
ACTION_DELAY = (2.5, 4.5)               # seconds between in-page actions (page clicks)
HOTEL_DELAY = (5.0, 10.0)               # seconds between hotels
BATCH_SIZE = 10                         # hotels per batch
BATCH_COOLDOWN = (25.0, 45.0)         # pause after each batch
MAX_RETRIES = 2                         # per hotel, across runs
CONSECUTIVE_BLOCK_LIMIT = 3             # this many blocked hotels in a row -> long pause
BLOCK_COOLDOWN = (600.0, 900.0)         # long pause (no bypass, just wait)
MAX_BLOCK_COOLDOWNS = 2                 # after this many long pauses, stop and let user resume later
MAX_REVIEW_PAGES = 40                   # hard cap on pagination clicks per hotel
IMAGE_WORKERS = 8                       # thread pool for CDN downloads (originals average ~1 MB)
IMAGE_DELAY = (0.1, 0.4)
IMAGE_TIMEOUT = 30
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
)
SAVE_DEBUG_SNAPSHOTS = True             # HTML + captured-URL list for first hotels and failures
DEBUG_FIRST_N_HOTELS = 3

# Research-ethics option: reviewer names are personal data. The spec asks to store them,
# so default is False; set True to store a salted SHA-256 pseudonym instead.
PSEUDONYMIZE_REVIEWER_NAME = False
PSEUDONYM_SALT = "change-me-before-running"

# ------------------------------------------------------------------ Trip.com specifics
TRIP_BASE = "https://www.trip.com"
# Review endpoint observed in 2026 (restapi/soa2/34308/getHotelCommentInfo). Direct replay of this
# endpoint is blocked (HTTP 430), so the crawler only *listens* to responses the page itself makes.
REVIEW_RESPONSE_PATTERNS = ["getHotelCommentInfo", "comment", "Comment", "review", "Review"]
HOTEL_LIST_RESPONSE_PATTERNS = ["fetchHotelList", "hotelList", "HotelList"]
DETAIL_URL_PRICE_PARAMS = "?checkIn={checkin}&checkOut={checkout}&adult={adults}&children={children}&crn={rooms}&curr=USD"
LIST_URL_TEMPLATE = (
    TRIP_BASE + "/hotels/list?city={city_id}&checkin={checkin_slash}&checkout={checkout_slash}"
    "&adult={adults}&children={children}&crn={rooms}&searchWord={keyword}&curr=USD"
)

# Ctrip/Trip.com city IDs (used only as a fallback when the cityId cannot be read from the
# unmanned hotel's own detail page).
CITY_IDS = {
    "Beijing": 1, "Shanghai": 2, "Tianjin": 3, "Chongqing": 4, "Harbin": 5, "Dalian": 6,
    "Qingdao": 7, "Nanjing": 12, "Suzhou": 14, "Hangzhou": 17, "Chengdu": 28, "Shenzhen": 30,
    "Guangzhou": 32, "Kunming": 34, "Jinan": 144, "Changsha": 206, "Hefei": 278, "Wuhan": 477,
    "Zhengzhou": 559, "Tokyo": 228,
}

# Text signals that the page is a CAPTCHA / login wall. We never try to solve or bypass these.
BLOCK_TEXT_SIGNALS = [
    "verify you are human", "security verification", "slide to complete", "drag the slider",
    "please complete the verification", "captcha", "jigsaw", "unusual traffic",
    "请完成安全验证", "拖动滑块", "安全验证", "访问受限",
]
LOGIN_TEXT_SIGNALS = ["sign in to continue", "log in to view", "please sign in", "请登录后"]
BLOCK_HTTP_STATUSES = {403, 429, 430}

# UI selectors tried in order (text/role based so they survive CSS-class churn).
SORT_OPEN_SELECTORS = [
    # 2026-10 review page: custom dropdown = read-only <input value="Most relevant"> next to "Sort by:"
    "input[readonly][value='Most relevant']",
    "input[readonly][value='Most Relevant']",
    "text=/^\\s*(Sort|Sort by|Recommended|Most relevant|排序|并び替え|정렬)\\b/i",
    "[class*='sort'] >> nth=0",
]
SORT_NEWEST_SELECTORS = [
    "text=/^(Most recent|Newest|Latest|Newest first|最新点评|最新|新しい順|최신순)/i",
]
NEXT_PAGE_SELECTORS = [
    "button[aria-label*='next' i]:not([disabled])",
    "a[aria-label*='next' i]",
    "li[class*='next']:not([class*='disabled'])",
    "[class*='pagination'] [class*='next']:not([class*='disabled'])",
    "text=/^(Next|Next page|下一页|次へ|다음)$/i",
]
LOAD_MORE_SELECTORS = [
    "text=/^(Show more reviews|Load more|More reviews|See more reviews|查看更多点评|もっと見る|더보기)$/i",
]

# Review photos are delivered as CDN-relative paths (e.g. "/0230...E5.jpg"); the page itself renders
# them from this base (observed 2026-10: https://ak-d.tripcdn.com/images/<name>_R_214_214_R5.webp).
REVIEW_IMAGE_BASE = "https://ak-d.tripcdn.com/images"
# Trip.com facility ids (shareData.hotel.ctripFacilityIds on the public review page).
# 364 = "Service robots" (observed live 2026-10-06 in hotelHotFacilityTags).
ROBOT_FACILITY_IDS = [364]
AUTOMATION_FACILITY_IDS: list[int] = []     # filled from work/facility_id_names.json once observed
# Anonymous visitors get at most 5 pages x 15 reviews on the public review page.
SEO_LIST_URL = TRIP_BASE + "/hotels/{city_url}-hotels-list-{city_id}/"
# The review list's load-more control is a blue button whose label is exactly this text; every
# review card ALSO has a small "Show More" text link (expands long text), which must not be clicked.
LOAD_MORE_BUTTON_TEXTS = ["Show More", "Show more", "Show more reviews", "Load more"]

# ------------------------------------------------------------------ classification keywords
ROBOT_KEYWORDS = [
    "service robot", "delivery robot", "robot", "robotic", "humanoid", "机器人", "送物机器人",
    "ロボット", "로봇",
]
# Strong evidence of front-desk automation (core touch point).
AUTOMATION_STRONG_KEYWORDS = [
    "self check-in", "self-check-in", "self-service check-in", "self service check-in",
    "self check-out", "self-check-out", "automated check-in", "automatic check-in",
    "check-in kiosk", "kiosk", "hologram", "unmanned", "自助入住", "自助办理入住", "自助机",
    "无人", "自動チェックイン", "セルフチェックイン", "ホログラム", "무인", "셀프 체크인", "키오스크",
]
# Weak signals: recorded but NOT sufficient for AUTOMATION_ONLY ("Express check-in/check-out" is
# a standard staffed-service facility label on Trip.com).
AUTOMATION_WEAK_KEYWORDS = [
    "express check-in", "contactless check-in", "mobile check-in", "digital key", "smart lock",
]
STAFFED_KEYWORDS = [
    "24-hour front desk", "front desk hours", "front desk", "concierge", "multilingual staff",
    "前台", "フロント", "프런트",
]

# ------------------------------------------------------------------ matching
# Weights implement the spec's priority order lexicographically:
#   same city  >  same brand  >  same star  >  star ±1  >  similar price  >  similar review volume.
# (e.g. same city alone (200) always beats a nearby city with every other bonus (90+50+40+12+6);
#  exact star (40) beats star±1 plus a perfect price (25+12).)
MATCH_WEIGHTS = {
    "same_city": 200.0,
    "nearby_city": 90.0,        # same metro area / prefecture (NEARBY_CITIES)
    "same_country": 10.0,       # last resort, flagged SAME_COUNTRY_ONLY
    "same_brand": 50.0,
    "same_group": 25.0,
    "star_exact": 40.0,
    "star_pm1": 25.0,
    "star_far": -30.0,          # |diff| > 1
    "star_unknown": 5.0,
    "price_max": 12.0,          # scaled by log-ratio similarity (0 at 2x difference)
    "reviews_max": 6.0,         # scaled by log-ratio similarity (0 at 10x difference)
    "area_token": 5.0,          # shared neighbourhood token (tie-breaker)
    "grade_a": 3.0,             # staffed front desk explicitly confirmed in source list
}
MIN_CONTROLS_PER_UNMANNED_IN_CITY = 2   # (legacy) discovery target margin
CONTROL_MARGIN = 0.34                   # discovery verifies n + ceil(0.34 n) (min +1) controls per city
MIN_CONTROL_REVIEWS = 10                # a control needs Trip.com reviews to sample (seed pool)
# Property types that are not comparable hotels (apartment / homestay / hostel) or whose name itself
# signals a tech-operated concept (smart / unmanned / cinema-room budget concepts): never a control.
NON_CONTROL_NAME_RE = (r"apartment|homestay|hostel|youth|capsule|guest ?house|b&b|villa|residence|smart|intelligen|"
                       r"unmanned|self[- ]service|cinema|movie|e-?sports|gaming|公寓|民宿|青年旅|旅舍|智|无人|無人|影院|觀影|观影|电竞|電競")
MIN_DISCOVERED_CONTROL_REVIEWS = 50     # newly discovered controls: enough reviews for a comparable sample

# Same metro area / prefecture fallback (never crosses a country border).
NEARBY_CITIES = {
    ("Japan", "Urayasu"): ["Tokyo"],
    ("Japan", "Komatsu"): ["Kanazawa"],
    ("Japan", "Izumisano"): ["Osaka"],
    ("Japan", "Gamagori"): ["Nagoya"],
}

# Brand patterns (regex, brand, parent group). First match wins.
BRAND_PATTERNS = [
    (r"henn[\s\-]?na", "Henn na Hotel", "H.I.S. Hotel Holdings"),
    (r"watermark", "Watermark Hotel", "H.I.S. Hotel Holdings"),
    (r"\bji hotel\b|全季", "JI Hotel", "H World (Huazhu)"),
    (r"crystal orange|orange hotel|桔子", "Orange Hotel", "H World (Huazhu)"),
    (r"hanting|汉庭", "Hanting", "H World (Huazhu)"),
    (r"ni hao hotel|你好酒店", "Ni Hao Hotel", "H World (Huazhu)"),
    (r"starway|星程", "Starway", "H World (Huazhu)"),
    (r"manxin|漫心", "Manxin", "H World (Huazhu)"),
    (r"citigo", "CitiGO", "H World (Huazhu)"),
    (r"atour|亚朵", "Atour Hotel", "Atour Group"),
    (r"zhotel", "ZHotel", "Atour Group"),
    (r"metropolo|metropolis hotel|锦江都城", "Metropolo", "Jin Jiang"),
    (r"jinjiang inn|锦江之星", "Jinjiang Inn", "Jin Jiang"),
    (r"vienna", "Vienna Hotel", "Jin Jiang"),
    (r"7 ?days", "7 Days Inn", "Jin Jiang"),
    (r"home ?inn|如家", "Home Inn", "BTG Homeinns"),
    (r"yunji", "Yunji", "Yunji"),
    (r"even hotel", "EVEN Hotel", "IHG"),
    (r"holiday inn", "Holiday Inn", "IHG"),
    (r"crowne plaza", "Crowne Plaza", "IHG"),
    (r"intercontinental", "InterContinental", "IHG"),
    (r"hotel indigo", "Hotel Indigo", "IHG"),
    (r"renaissance", "Renaissance", "Marriott"),
    (r"courtyard", "Courtyard", "Marriott"),
    (r"four points", "Four Points", "Marriott"),
    (r"mercure", "Mercure", "Accor"),
    (r"\bibis\b", "ibis", "Accor"),
    (r"novotel", "Novotel", "Accor"),
    (r"citadines", "Citadines", "Ascott"),
    (r"r&b hotel", "R&B Hotel", "Washington Hotel Corp."),
    (r"washington hotel", "Washington Hotel", "Washington Hotel Corp."),
    (r"mitsui garden", "Mitsui Garden Hotel", "Mitsui Fudosan"),
    (r"keio presso", "Keio Presso Inn", "Keio"),
    (r"koko hotel", "KOKO HOTEL", "KOKO HOTEL"),
    (r"granbell", "Granbell Hotel", "Granbell"),
    (r"tokyu stay", "Tokyu Stay", "Tokyu"),
    (r"gracery", "Hotel Gracery", "Fujita Kanko"),
    (r"jr kyushu", "JR Kyushu Hotel", "JR Kyushu"),
    (r"jal city", "Hotel JAL City", "Okura Nikko"),
    (r"keikyu ex inn", "Keikyu EX Inn", "Keikyu"),
    (r"monday", "hotel MONday", "MONday"),
    (r"tobu hotel", "Tobu Hotel", "Tobu"),
    (r"richmond hotel", "Richmond Hotel", "Richmond"),
    (r"monterey", "Hotel Monterey", "Monterey"),
    (r"daiwa roynet", "Daiwa Roynet Hotel", "Daiwa House"),
    (r"solaria nishitetsu", "Solaria Nishitetsu Hotel", "Nishitetsu"),
    (r"nishitetsu hotel croom", "Nishitetsu Hotel Croom", "Nishitetsu"),
    (r"onefive", "The OneFive", "Fukuoka Realty"),
    (r"chisun", "Chisun", "Solare"),
    (r"onyado nono", "Onyado Nono", "Kyoritsu Maintenance"),
    (r"dormy inn", "Dormy Inn", "Kyoritsu Maintenance"),
    (r"cross hotel|cross life", "Cross Hotel", "Orix"),
    (r"granvia", "Hotel Granvia", "JR West"),
    (r"hankyu respire", "Hotel Hankyu Respire", "Hankyu Hanshin"),
    (r"androoms", "hotel androoms", "Solare"),
    (r"sotetsu grand fresa", "Sotetsu Grand Fresa", "Sotetsu"),
    (r"sotetsu fresa inn", "Sotetsu Fresa Inn", "Sotetsu"),
    (r"quintessa", "Quintessa Hotel", "Quintessa"),
    (r"keihan", "Hotel Keihan", "Keihan"),
    (r"hearton", "Hearton Hotel", "Hearton"),
    (r"oriental hotel", "Oriental Hotel", "Oriental Hotels & Resorts"),
    (r"miyako hotel", "Miyako Hotel", "Kintetsu"),
    (r"glad one", "Hotel Glad One", "M's"),
    (r"agora", "Agora", "Agora Hospitalities"),
    (r"travelodge", "Travelodge", "Travelodge"),
    (r"comfort (hotel|inn)", "Comfort Hotel", "Choice Hotels"),
    (r"smile hotel", "Smile Hotel", "Smile Hotel"),
    (r"royal park", "Royal Park Hotel", "Mitsubishi Estate"),
    (r"mystays", "HOTEL MYSTAYS", "MyStays"),
    (r"\bl7\b|lotte", "L7 / Lotte", "Lotte Hotels"),
    (r"arlo", "Arlo", "Arlo"),
]

# Neighbourhood tokens used as a tie-breaker inside the same city.
AREA_TOKENS = [
    "ginza", "haneda", "asakusa", "shinjuku", "namba", "nipponbashi", "shinsaibashi", "umeda",
    "hachijo", "shichijo", "gojo", "karasuma", "hakata", "tenjin", "nakasu", "korinbo", "fushimi",
    "myeongdong", "midtown", "tenmonkan", "kokubuncho", "akasaka", "hamamatsucho",
]
