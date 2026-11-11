"""
Steadfast (steadfast.com.bd/pricing) rate scraper
Price API : GET https://steadfast.com.bd/pricing/quote?from=<id>&to=<id>&category=1&weight=<kg>&service=1&delivery=0&cod=0

Install : python -m pip install requests pandas openpyxl
Run     : python steadfast_scraper.py
Output  : steadfast_rates.xlsx  (sheets: Home Delivery, Point Delivery, Express, Long)
"""
import html
import re
import time
import requests
import pandas as pd

SITE = "https://steadfast.com.bd"
HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/json",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": SITE + "/pricing",
}

# ---------------- SETTINGS ----------------
FROM_ID = 1                          # 1 = Dhaka City
WEIGHTS_KG = [0.5, 1, 2, 3, 5, 10]   # website-er weight option (chaile 1.5, 4 ityadi add korte paren)
CATEGORIES = {1: "Regular"}          # Document/Book-o lagle: {1: "Regular", 2: "Document", 3: "Book"}
SERVICE = 1
DELAY = 0.3
OUT_FILE = "steadfast_rates.xlsx"
# ------------------------------------------

# Page source (datalist) theke paoa district list: ID -> naam (fallback, live page parse na hole eta use hobe)
DISTRICTS = {
    1: "Dhaka City", 2: "Chittagong", 3: "Sylhet", 4: "Rangpur", 5: "Rajshahi", 6: "Khulna",
    7: "Barishal", 8: "Noakhali", 9: "Cumilla", 10: "Narshindi", 11: "Mymensingh", 12: "Kishoreganj",
    13: "Narayanganj", 14: "Gazipur", 15: "Dinajpur", 16: "Bogra", 17: "Tangail", 18: "Dhaka Sub-Urban",
    19: "Brahmanbaria", 20: "Chandpur", 21: "Feni", 22: "Natore", 23: "Gaibandha", 24: "Nilphamari",
    26: "Jamalpur", 27: "Kustia", 28: "Faridpur", 29: "Laxmipur", 30: "Chuadanga", 31: "Magura",
    32: "Manikganj", 33: "Sirajganj", 34: "Habiganj", 35: "Munshiganj", 36: "Jashore", 37: "Lalmonirhat",
    38: "Bhola", 39: "Rajbari", 40: "Moulvibazar", 41: "Pabna", 42: "Cox's Bazar", 43: "Jhenaidah",
    44: "Gopalganj", 45: "Patuakhali", 46: "Thakurgaon", 47: "Rangamati", 48: "Narail", 49: "Pirojpur",
    50: "Netrokona", 51: "Meherpur", 52: "Khagrachori", 53: "Shatkhira", 54: "Panchgarh", 55: "Naogaon",
    56: "Kurigram", 57: "Barguna", 58: "Chapainawabganj", 59: "Jhalokati", 60: "Sherpur", 61: "Bandarban",
    62: "Madaripur", 63: "Shariatpur", 64: "Joypurhat", 65: "Sunamganj", 66: "Bagerhat",
}

s = requests.Session()
s.headers.update(HEADERS)


def load_districts():
    """Live page-er <datalist> theke list nei; na pele upore-r fallback list."""
    try:
        t = requests.get(SITE + "/pricing", headers={"User-Agent": "Mozilla/5.0"}, timeout=20).text
        found = {int(i): html.unescape(n) for n, i in re.findall(r'<option value="([^"]*)"[^>]*data-id="(\d+)"', t)}
        if len(found) > 50:
            print(f"{len(found)} district live page theke nilam")
            return found
    except requests.RequestException as e:
        print("live page paoa gelo na:", e)
    print(f"{len(DISTRICTS)} district built-in list theke nilam")
    return DISTRICTS


def quote(to_id, weight, category, retries=3):
    params = {"from": FROM_ID, "to": to_id, "category": category, "weight": weight,
              "service": SERVICE, "delivery": 0, "cod": 0}
    for attempt in range(1, retries + 1):
        try:
            r = s.get(SITE + "/pricing/quote", params=params, timeout=20)
            if r.status_code == 200:
                return r.json()
            print(f"  HTTP {r.status_code} to={to_id} w={weight}")
        except (requests.RequestException, ValueError) as e:
            print(f"  error ({attempt}/{retries}) to={to_id} w={weight}: {e}")
        time.sleep(1.5 * attempt)
    return None


def main():
    districts = load_districts()
    pickup = districts.get(FROM_ID, str(FROM_ID))
    rows = []
    total = len(districts) * len(WEIGHTS_KG) * len(CATEGORIES)
    done = 0
    for cat_id, cat_name in CATEGORIES.items():
        for to_id in sorted(districts):
            for w in WEIGHTS_KG:
                j = quote(to_id, w, cat_id)
                done += 1
                if not j:
                    continue
                opt = j.get("options") or {}
                rows.append({
                    "Category": cat_name,
                    "Pickup": pickup,
                    "To ID": to_id,
                    "Delivery District": districts[to_id],
                    "Lane": j.get("lane"),
                    "Available": j.get("available"),
                    "Weight (kg)": w,
                    "Charge (Home)": opt.get("home", j.get("charge")),
                    "Charge (Point)": opt.get("point"),
                    "Charge (Express)": opt.get("express"),
                })
                time.sleep(DELAY)
            print(f"{done}/{total}  {districts[to_id]}")

    df = pd.DataFrame(rows)
    if df.empty:
        print("Kono data pawa jayni")
        return

    def matrix(col):
        m = df.pivot_table(index=["Category", "Pickup", "To ID", "Delivery District", "Lane"],
                           columns="Weight (kg)", values=col, aggfunc="first").reset_index()
        m.columns = [c if isinstance(c, str) else f"{c}kg" for c in m.columns]
        return m

    with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as xw:
        matrix("Charge (Home)").to_excel(xw, sheet_name="Home Delivery", index=False)
        matrix("Charge (Point)").to_excel(xw, sheet_name="Point Delivery", index=False)
        matrix("Charge (Express)").to_excel(xw, sheet_name="Express", index=False)
        df.to_excel(xw, sheet_name="Long", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                w = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = w + 3

    print(f"Done -> {OUT_FILE}")


if __name__ == "__main__":
    main()
