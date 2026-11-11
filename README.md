# courier-rate-scrapers

Python scripts that pull the public delivery-rate calculators of Bangladeshi courier companies and save every route and weight slab to Excel, so rates can be compared side by side.

| Script | Courier | Source | Output |
|---|---|---|---|
| `carrybee_scraper.py` | [CarryBee](https://carrybee.com/calculator) | public price-calculation API used by the calculator page | `carrybee_rates.xlsx` |
| `steadfast_scraper.py` | [Steadfast](https://steadfast.com.bd/pricing) | public `/pricing/quote` endpoint used by the pricing page | `steadfast_rates.xlsx` |

## How it works

Both calculator pages are thin front ends over a JSON API. I found the endpoints by recording the page in Chrome DevTools (a HAR file) and then call them directly with `requests`. No browser automation and no login is needed.

- **CarryBee:** loads the district list, then requests the price for every destination and weight (200 g to 5 kg).
- **Steadfast:** reads the district list from the pricing page, then requests Home, Point and Express charges for every destination and weight (0.5 kg to 10 kg).

## Setup

```
python -m pip install -r requirements.txt
```

## Run

```
python carrybee_scraper.py
python steadfast_scraper.py
```

Each run takes a few minutes (roughly 400 requests per script, with a short delay between them) and prints progress as it goes.

## Output

**CarryBee** (`carrybee_rates.xlsx`)
- `Matrix`: one row per destination, one column per weight.
- `Long`: `Pickup | Delivery | Weight (g) | Delivery Fee (BDT)`, good for filtering and pivot tables.

**Steadfast** (`steadfast_rates.xlsx`)
- `Home Delivery`, `Point Delivery`, `Express`: one row per destination, one column per weight.
- `Long`: every quote with category, lane, availability and all three charges.

## Settings

Edit the constants at the top of each script.

| Setting | Script | Meaning |
|---|---|---|
| `PICKUP_NAMES` / `ALL_PICKUPS` | CarryBee | Pickup city, or every city as pickup (about 24,000 requests) |
| `FROM_ID` | Steadfast | Pickup district ID (`1` = Dhaka City) |
| `WEIGHTS_G` / `WEIGHTS_KG` | both | Weight slabs to query |
| `CATEGORIES` | Steadfast | Add Document (`2`) and Book (`3`) next to Regular (`1`) |
| `DELAY` | both | Pause between requests, in seconds |

## Notes and limits

- These are the **standard public rates**. Merchants on negotiated rates are billed differently.
- District-level rates only. Thana or area-level pricing is not covered.
- Steadfast COD charges are not included.
- The endpoints are undocumented and can change without notice. If a script breaks, capture a fresh HAR from the calculator page and compare the request URL and response shape.
- Keep the delay in place and check each site's terms before heavy use. This project is meant for comparing rates for your own business, not for hammering the servers.
- `.har` and `.xlsx` files are git-ignored. HAR files can contain cookies and tokens, so never commit them.

## License

MIT. See [LICENSE](LICENSE).
