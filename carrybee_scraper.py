"""
CarryBee shipping-rate scraper
- Cities list : GET /api/v2/country/1/cities
- Price       : GET /api/v2/orders/price-calculation?delivery_type=1&item_type=1&weight=<g>&from=<id>&to=<id>

Install : pip install requests pandas openpyxl
Run     : python carrybee_scraper.py
Output  : carrybee_rates.xlsx  (sheets: Matrix, Long)
"""
import time
import requests
import pandas as pd

BASE = "https://api-ops.carrybee.com/api/v2"
HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Origin": "https://carrybee.com",
    "Referer": "https://carrybee.com/",
}

# ---------------- SETTINGS ----------------
PICKUP_NAMES = ["Dhaka"]          # chaile ["Dhaka", "Chattogram"] ba ALL_PICKUPS = True
ALL_PICKUPS = False               # True dile 64 x 64 x weights = ~24,000 request (onek somoy lagbe)
WEIGHTS_G = [200, 300, 500, 1000, 2000, 5000]   # calculator-er 6-ta weight (gram)
DELIVERY_TYPE = 1
ITEM_TYPE = 1
DELAY = 0.3                       # request-er majhe biroti (second)
OUT_FILE = "carrybee_rates.xlsx"
# ------------------------------------------

session = requests.Session()
session.headers.update(HEADERS)


def get_json(url, params=None, retries=3):
    for attempt in range(1, retries + 1):
        try:
            r = session.get(url, params=params, timeout=20)
            if r.status_code == 200:
                return r.json()
            print(f"  HTTP {r.status_code} for {params}")
        except requests.RequestException as e:
            print(f"  error ({attempt}/{retries}): {e}")
        time.sleep(1.5 * attempt)
    return None


def main():
    data = get_json(f"{BASE}/country/1/cities")
    cities = data["data"]["cities"]
    id_to_name = {c["id"]: c["name"] for c in cities}
    name_to_id = {c["name"].lower(): c["id"] for c in cities}
    print(f"{len(cities)} cities loaded")

    if ALL_PICKUPS:
        pickup_ids = list(id_to_name)
    else:
        pickup_ids = [name_to_id[n.lower()] for n in PICKUP_NAMES]

    rows = []
    total = len(pickup_ids) * len(cities) * len(WEIGHTS_G)
    done = 0
    for frm in pickup_ids:
        for to in id_to_name:
            for w in WEIGHTS_G:
                params = {
                    "delivery_type": DELIVERY_TYPE,
                    "item_type": ITEM_TYPE,
                    "weight": w,
                    "from": frm,
                    "to": to,
                }
                j = get_json(f"{BASE}/orders/price-calculation", params)
                fee = None
                if j and not j.get("error"):
                    fee = j["data"].get("delivery_fee")
                rows.append({
                    "Pickup": id_to_name[frm],
                    "Delivery": id_to_name[to],
                    "Weight (g)": w,
                    "Delivery Fee (BDT)": fee,
                })
                done += 1
                if done % 25 == 0:
                    print(f"{done}/{total}")
                time.sleep(DELAY)

    long_df = pd.DataFrame(rows)
    matrix = long_df.pivot_table(
        index=["Pickup", "Delivery"],
        columns="Weight (g)",
        values="Delivery Fee (BDT)",
        aggfunc="first",
    ).reset_index()
    matrix.columns = [f"{c}g" if isinstance(c, int) else c for c in matrix.columns]

    with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as xw:
        matrix.to_excel(xw, sheet_name="Matrix", index=False)
        long_df.to_excel(xw, sheet_name="Long", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = width + 3

    print(f"Done -> {OUT_FILE}")


if __name__ == "__main__":
    main()