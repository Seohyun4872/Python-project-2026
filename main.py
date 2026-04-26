import time
import schedule
import gspread
from datetime import datetime
from google.oauth2.service_account import Credentials

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

SPREADSHEET_ID = "1l8Zo1vkGe2mwrEZ3p5vKjSalT5NwXb9-Kk4q33SXs34"
SHEET_NAME = "ranking_log"
SERVICE_ACCOUNT_FILE = "service_account.json"

TARGET_BRAND_KR = "로긴 앤 로지"
TARGET_BRAND_DATA = "roginnrosie"

BASE_URL = "https://www.musinsa.com/main/musinsa/ranking?gf=A&storeCode=musinsa&sectionId=199&contentsId=&categoryCode=000&ageBand=AGE_BAND_ALL&subPan=product"

TABS = ["전체", "NEW", "급상승"]


def connect_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=scopes
    )

    gc = gspread.authorize(credentials)
    sh = gc.open_by_key(SPREADSHEET_ID)

    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=1000, cols=10)
        ws.append_row([
            "수집시간",
            "순위 카테고리",
            "순위",
            "브랜드명",
            "제품명",
            "할인율",
            "판매가격",
            "상품 URL",
        ])

    return ws


def create_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    return driver


def click_tab(driver, tab_name):
    elements = driver.find_elements(By.XPATH, f"//*[normalize-space(text())='{tab_name}']")

    for el in elements:
        try:
            driver.execute_script("arguments[0].click();", el)
            time.sleep(3)
            return True
        except:
            continue

    print(f"[경고] {tab_name} 탭 클릭 실패")
    return False


def scroll_to_150(driver):
    for _ in range(12):
        driver.execute_script("window.scrollBy(0, 1200);")
        time.sleep(1)


def parse_price(price):
    if not price:
        return ""
    try:
        return f"{int(price):,}원"
    except:
        return price


def crawl_tab(driver, tab_name):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []

    scroll_to_150(driver)

    product_links = driver.find_elements(
        By.CSS_SELECTOR,
        "a[data-item-id][data-brand][href*='/products/']"
    )

    seen = set()

    for item in product_links:
        try:
            url = item.get_attribute("href")
            item_id = item.get_attribute("data-item-id")
            brand_data = item.get_attribute("data-brand")
            price = item.get_attribute("data-price")
            discount_rate = item.get_attribute("data-discount-rate")

            if not item_id or item_id in seen:
                continue

            seen.add(item_id)

            if brand_data != TARGET_BRAND_DATA:
                continue

            rank = item.get_attribute("data-index")

            text = item.text.strip()
            lines = [x.strip() for x in text.split("\n") if x.strip()]

            product_name = ""
            for line in lines:
                if line not in [TARGET_BRAND_KR] and "%" not in line and "원" not in line:
                    if line != rank:
                        product_name = line
                        break

            rows.append([
                now,
                tab_name,
                rank,
                TARGET_BRAND_KR,
                product_name,
                f"{discount_rate}%" if discount_rate else "",
                parse_price(price),
                url,
            ])

        except Exception as e:
            print("상품 파싱 오류:", e)

    return rows


def run_crawling():
    print("크롤링 시작:", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    ws = connect_sheet()
    driver = create_driver()

    all_rows = []

    try:
        driver.get(BASE_URL)
        time.sleep(5)

        for tab in TABS:
            print(f"{tab} 수집 중...")

            click_tab(driver, tab)

            rows = crawl_tab(driver, tab)
            all_rows.extend(rows)

            print(f"{tab}: 로긴 앤 로지 {len(rows)}개 발견")

        if all_rows:
            ws.append_rows(all_rows, value_input_option="USER_ENTERED")
            print(f"구글시트 저장 완료: {len(all_rows)}행")
        else:
            print("이번 수집에서는 로긴 앤 로지 상품이 없습니다.")

    finally:
        driver.quit()

    print("크롤링 종료")


run_crawling()

schedule.every(30).minutes.do(run_crawling)

while True:
    schedule.run_pending()
    time.sleep(1)