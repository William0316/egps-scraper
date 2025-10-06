import re
import requests
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# === Google Sheets 設定 ===
SERVICE_ACCOUNT_JSON = "service_account/kiehls-scraper.json"

SPREADSHEET_ID = "1LjyB8BfL6MrEdATkj_3cuhdJcF_CAxiQbGnGtzhJtCY"

GSCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_JSON, scopes=GSCOPE)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(SPREADSHEET_ID)

BASE_URL = "https://www.egps.com.tw/products.asp"
HEADERS = {"User-Agent": "Mozilla/5.0"}

# ---------------- 爬蟲部分（維持改版骨架，可穩定抓商品） ----------------
def fetch_page(brand: str, page: int) -> BeautifulSoup:
    """抓取單頁 HTML"""
    if page == 1:
        payload = {"type": "search", "t1": brand, "t2": "", "t3": "", "page": page}
        res = requests.post(BASE_URL, data=payload, headers=HEADERS)
    else:
        res = requests.get(f"{BASE_URL}?type=search&t1={brand}&page={page}", headers=HEADERS)
    res.encoding = "big5"
    return BeautifulSoup(res.text, "lxml")


def scrape_brand(brand: str):
    """完整抓取該品牌所有商品"""
    rows = []
    page = 1
    while True:
        soup = fetch_page(brand, page)
        products = soup.select("td[width='160'][valign='top']")
        if not products:
            if page == 1:
                print("⚠️ 沒有符合的商品")
            else:
                print(f"⚠️ 第 {page} 頁沒有商品，停止。")
            break

        print(f"✅ 抓到第 {page} 頁，共 {len(products)} 筆商品")

        for p in products:
            shop = p.select_one("span")
            shop = shop.get_text(strip=True) if shop else ""

            a = p.select_one("a.a_table_list_txt")
            name_full = a.get_text(" ", strip=True) if a else ""
            url = "https://www.egps.com.tw/" + a["href"] if a else ""

            img_tag = p.select_one("img")
            img_url = "https://www.egps.com.tw/" + img_tag["src"] if img_tag else ""
            img_formula = f'=IMAGE("{img_url}", 4, 80, 80)' if img_url else ""

            price = ""
            price_span = p.find("span", class_="shopping_Price")
            if price_span:
                digits = re.sub(r"[^\d]", "", price_span.get_text())
                price = int(digits) if digits else ""

            brand_name = ""
            if name_full:
                parts = name_full.split()
                brand_name = parts[0] if parts else ""

            rows.append([shop, brand_name, "", name_full, price, url, img_formula])

        has_next = soup.select_one(f"a[href*='products.asp?page={page+1}']")
        if not has_next:
            print("✅ 已到最後一頁，結束。")
            break
        page += 1

    return rows


# ---------------- Google Sheets 寫入與比對 ----------------
def write_to_sheet(data):
    today = datetime.today().strftime("%Y-%m-%d")

    # 建立或取得「商品變動追蹤」表
    try:
        tracking_ws = spreadsheet.worksheet("商品變動追蹤")
    except gspread.WorksheetNotFound:
        tracking_ws = spreadsheet.add_worksheet("商品變動追蹤", rows=200, cols=10)
        spreadsheet.reorder_worksheets([tracking_ws] + spreadsheet.worksheets())
        tracking_ws.update("A1:E1", [["日期", "狀態", "商品名稱", "型號", "價格"]])

    # 建立今天的工作表
    try:
        ws_today = spreadsheet.worksheet(today)
        spreadsheet.del_worksheet(ws_today)  # 若存在則重建
    except gspread.WorksheetNotFound:
        pass
    ws_today = spreadsheet.add_worksheet(today, rows=200, cols=10)

    headers = ["店名", "品牌", "型號", "完整名稱", "價格", "連結", "圖片"]
    ws_today.update("A1:G1", [headers], value_input_option="USER_ENTERED")
    ws_today.update(f"A2:G{len(data)+1}", data, value_input_option="USER_ENTERED")

    # 型號欄公式（再轉成純值）
    model_formulas = []
    for i in range(len(data)):
        r = i + 2
        formula = f'=REGEXEXTRACT(D{r}, "\\d{{5,6}}[A-Z]{{0,4}}")'
        model_formulas.append([formula])
    ws_today.update(f"C2:C{len(data)+1}", model_formulas, value_input_option="USER_ENTERED")
    model_values = ws_today.get(f"C2:C{len(data)+1}")
    ws_today.update(f"C2:C{len(data)+1}", model_values, value_input_option="RAW")

    print(f"✅ 已成功寫入工作表 {today}，共 {len(data)} 筆資料")

    # ---------------- 比對新增 / 下架 ----------------
    sheets = spreadsheet.worksheets()
    date_sheets = [s.title for s in sheets if re.match(r"\d{4}-\d{2}-\d{2}", s.title)]
    date_sheets.sort()

    if len(date_sheets) >= 2:
        ws_yesterday = spreadsheet.worksheet(date_sheets[-2])
        yesterday_data = ws_yesterday.get("D2:E")  # D=完整名稱, E=價格
        today_data = ws_today.get("D2:E")

        def extract_model(name):
            match = re.search(r"\d{5,6}[A-Z]{0,4}", name)
            return match.group(0) if match else ""

        yesterday_dict = {
            row[0]: {
                "price": row[1] if len(row) > 1 else "",
                "model": extract_model(row[0])
            }
            for row in yesterday_data if row
        }
        today_dict = {
            row[0]: {
                "price": row[1] if len(row) > 1 else "",
                "model": extract_model(row[0])
            }
            for row in today_data if row
        }

        yesterday_names = set(yesterday_dict.keys())
        today_names = set(today_dict.keys())

        gone = yesterday_names - today_names
        new = today_names - yesterday_names

        rows_to_append = []
        for g in gone:
            rows_to_append.append([today, "下架", g, yesterday_dict[g]["model"], yesterday_dict[g]["price"]])
        for n in new:
            rows_to_append.append([today, "新增", n, today_dict[n]["model"], today_dict[n]["price"]])

        if rows_to_append:
            tracking_ws.append_rows(rows_to_append, value_input_option="USER_ENTERED")
            print(f"⚠️ 今日商品變動 {len(rows_to_append)} 筆，已更新追蹤表")
        else:
            print("✅ 今日商品數量無變動")
    else:
        print("ℹ️ 無前一日資料，略過比對。")


# ---------------- 主程式 ----------------
if __name__ == "__main__":
    brand = "Rolex"
    print(f"開始抓取品牌：{brand}")
    data = scrape_brand(brand)
    if data:
        write_to_sheet(data)
    else:
        print("⚠️ 沒有抓到任何商品。")
