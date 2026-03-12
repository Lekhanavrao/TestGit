"""
Flight Fare Automation Script using Playwright
Automates login, flight search, filter application, and fare data extraction
from app.itilite.com, saving results to an Excel file every hour.
"""

import time
import schedule
import pandas as pd
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ─────────────────────────── Constants ─────────────────────────────────────
LOGIN_URL       = "https://app.itilite.com"
EMAIL           = "prod_dashboard_test@itilite.com"
PASSWORD        = "Dashboard@1234"
ORIGIN          = "LAS"
DESTINATION     = "LAX"
EXCEL_FILE      = "flight_fare_details.xlsx"
TARGET_AIRLINES = ["Delta", "United", "Southwest"]


# ─────────────────────────── Login ─────────────────────────────────────────
def login(page) -> bool:
    """Log in to app.itilite.com with the provided credentials."""
    print("[INFO] Navigating to login page...")
    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    page.wait_for_load_state("networkidle", timeout=30_000)

    try:
        # Accept cookie / consent banner if present
        for selector in ["button:has-text('Accept')", "button:has-text('OK')", "[data-testid='cookie-accept']"]:
            if page.locator(selector).count() > 0:
                page.locator(selector).first.click()
                break

        # Fill email
        email_selectors = [
            "input[type='email']",
            "input[name='email']",
            "input[placeholder*='mail' i]",
            "#email",
        ]
        for sel in email_selectors:
            locator = page.locator(sel)
            if locator.count() > 0:
                locator.first.fill(EMAIL)
                print(f"[INFO] Filled email using selector: {sel}")
                break

        page.wait_for_timeout(500)

        # Click Next / Continue if the form is multi-step
        for btn in ["button:has-text('Next')", "button:has-text('Continue')", "button[type='submit']"]:
            if page.locator(btn).count() > 0:
                page.locator(btn).first.click()
                page.wait_for_timeout(1_000)
                break

        # Fill password
        password_selectors = [
            "input[type='password']",
            "input[name='password']",
            "#password",
        ]
        for sel in password_selectors:
            locator = page.locator(sel)
            if locator.count() > 0:
                locator.first.fill(PASSWORD)
                print(f"[INFO] Filled password using selector: {sel}")
                break

        page.wait_for_timeout(500)

        # Submit login form
        submit_selectors = [
            "button[type='submit']",
            "button:has-text('Sign In')",
            "button:has-text('Login')",
            "button:has-text('Log In')",
        ]
        for sel in submit_selectors:
            if page.locator(sel).count() > 0:
                page.locator(sel).first.click()
                print(f"[INFO] Clicked submit using selector: {sel}")
                break

        # Wait for dashboard / post-login page
        page.wait_for_load_state("networkidle", timeout=40_000)
        print("[SUCCESS] Login completed.")
        return True

    except PlaywrightTimeoutError as exc:
        print(f"[ERROR] Login timed out: {exc}")
        return False
    except Exception as exc:
        print(f"[ERROR] Login failed: {exc}")
        return False


# ─────────────────────────── Search Flight ─────────────────────────────────
def search_flight(page) -> bool:
    """Navigate to flight search and submit a one-way LAS → LAX search."""
    print("[INFO] Starting flight search...")
    try:
        # Navigate to the flights section
        flight_nav_selectors = [
            "a:has-text('Flights')",
            "a:has-text('Flight')",
            "[data-testid='flights-nav']",
            "nav a[href*='flight']",
        ]
        for sel in flight_nav_selectors:
            if page.locator(sel).count() > 0:
                page.locator(sel).first.click()
                page.wait_for_load_state("networkidle", timeout=20_000)
                print(f"[INFO] Navigated to Flights via: {sel}")
                break

        page.wait_for_timeout(1_500)

        # Select One-Way trip type
        one_way_selectors = [
            "label:has-text('One Way')",
            "input[value='one_way']",
            "input[value='oneway']",
            "button:has-text('One Way')",
            "[data-testid='one-way']",
        ]
        for sel in one_way_selectors:
            if page.locator(sel).count() > 0:
                page.locator(sel).first.click()
                print(f"[INFO] Selected One Way using: {sel}")
                page.wait_for_timeout(500)
                break

        # Fill Origin
        origin_selectors = [
            "input[placeholder*='From' i]",
            "input[placeholder*='Origin' i]",
            "[data-testid='origin-input']",
            "input[name='from']",
        ]
        for sel in origin_selectors:
            if page.locator(sel).count() > 0:
                origin_input = page.locator(sel).first
                origin_input.click()
                origin_input.fill("")
                origin_input.type(ORIGIN, delay=80)
                page.wait_for_timeout(1_200)
                # Click first suggestion
                for sug in ["li:has-text('LAS')", "[role='option']:has-text('LAS')", ".autocomplete-item:has-text('LAS')"]:
                    if page.locator(sug).count() > 0:
                        page.locator(sug).first.click()
                        break
                print(f"[INFO] Entered origin {ORIGIN}")
                break

        page.wait_for_timeout(800)

        # Fill Destination
        dest_selectors = [
            "input[placeholder*='To' i]",
            "input[placeholder*='Destination' i]",
            "[data-testid='destination-input']",
            "input[name='to']",
        ]
        for sel in dest_selectors:
            if page.locator(sel).count() > 0:
                dest_input = page.locator(sel).first
                dest_input.click()
                dest_input.fill("")
                dest_input.type(DESTINATION, delay=80)
                page.wait_for_timeout(1_200)
                for sug in ["li:has-text('LAX')", "[role='option']:has-text('LAX')", ".autocomplete-item:has-text('LAX')"]:
                    if page.locator(sug).count() > 0:
                        page.locator(sug).first.click()
                        break
                print(f"[INFO] Entered destination {DESTINATION}")
                break

        page.wait_for_timeout(800)

        # Click Search / Submit
        search_selectors = [
            "button:has-text('Search')",
            "button[type='submit']",
            "[data-testid='search-button']",
            "button:has-text('Find Flights')",
        ]
        for sel in search_selectors:
            if page.locator(sel).count() > 0:
                page.locator(sel).first.click()
                print(f"[INFO] Clicked search using: {sel}")
                break

        page.wait_for_load_state("networkidle", timeout=40_000)
        print("[SUCCESS] Flight search submitted.")
        return True

    except PlaywrightTimeoutError as exc:
        print(f"[ERROR] Flight search timed out: {exc}")
        return False
    except Exception as exc:
        print(f"[ERROR] Flight search failed: {exc}")
        return False


# ─────────────────────────── Apply Filters ─────────────────────────────────
def apply_filters(page) -> bool:
    """Apply Non-Stop, Delta/United/Southwest airline, and Economy cabin filters."""
    print("[INFO] Applying filters...")
    try:
        page.wait_for_load_state("networkidle", timeout=30_000)
        page.wait_for_timeout(2_000)

        # ── Non-Stop filter ──────────────────────────────────────────────────
        nonstop_selectors = [
            "label:has-text('Non Stop')",
            "label:has-text('Nonstop')",
            "input[value='non_stop'] + label",
            "[data-testid='filter-nonstop']",
            "label:has-text('0 Stops')",
        ]
        nonstop_applied = False
        for sel in nonstop_selectors:
            if page.locator(sel).count() > 0:
                page.locator(sel).first.click()
                nonstop_applied = True
                print(f"[INFO] Applied Non-Stop filter via: {sel}")
                page.wait_for_timeout(1_000)
                break
        if not nonstop_applied:
            print("[WARN] Non-Stop filter selector not found; skipping.")

        # ── Airline filters ──────────────────────────────────────────────────
        for airline in TARGET_AIRLINES:
            airline_selectors = [
                f"label:has-text('{airline}')",
                f"input[value='{airline.lower()}'] + label",
                f"[data-testid='filter-airline-{airline.lower()}']",
            ]
            applied = False
            for sel in airline_selectors:
                if page.locator(sel).count() > 0:
                    page.locator(sel).first.click()
                    applied = True
                    print(f"[INFO] Applied {airline} airline filter.")
                    page.wait_for_timeout(800)
                    break
            if not applied:
                print(f"[WARN] Airline filter for '{airline}' not found; skipping.")

        # ── Cabin Class: Economy ─────────────────────────────────────────────
        economy_selectors = [
            "label:has-text('Economy')",
            "input[value='economy'] + label",
            "[data-testid='filter-economy']",
            "button:has-text('Economy')",
        ]
        economy_applied = False
        for sel in economy_selectors:
            if page.locator(sel).count() > 0:
                page.locator(sel).first.click()
                economy_applied = True
                print(f"[INFO] Applied Economy cabin filter via: {sel}")
                page.wait_for_timeout(1_000)
                break
        if not economy_applied:
            print("[WARN] Economy filter selector not found; skipping.")

        page.wait_for_load_state("networkidle", timeout=20_000)
        print("[SUCCESS] Filters applied.")
        return True

    except PlaywrightTimeoutError as exc:
        print(f"[ERROR] Applying filters timed out: {exc}")
        return False
    except Exception as exc:
        print(f"[ERROR] Applying filters failed: {exc}")
        return False


# ─────────────────────────── Capture Fares ─────────────────────────────────
def capture_airline_fares(page) -> list[dict]:
    """
    For each target airline, click 'View Fares', extract all fare card details,
    and return a list of dictionaries.
    """
    all_fares: list[dict] = []

    for airline in TARGET_AIRLINES:
        print(f"\n[INFO] Processing fares for: {airline}")
        try:
            # Locate the airline's result card / row
            airline_row_selectors = [
                f"[data-airline*='{airline}' i]",
                f".flight-card:has-text('{airline}')",
                f".airline-row:has-text('{airline}')",
                f"div:has-text('{airline}'):has(button:has-text('View Fares'))",
            ]

            airline_row = None
            for sel in airline_row_selectors:
                rows = page.locator(sel)
                if rows.count() > 0:
                    airline_row = rows.first
                    print(f"[INFO] Found {airline} row via: {sel}")
                    break

            if airline_row is None:
                print(f"[WARN] Could not locate result row for '{airline}'; skipping.")
                continue

            # Click "View Fares" within that row
            view_fares_btn = airline_row.locator(
                "button:has-text('View Fares'), a:has-text('View Fares'), [data-testid='view-fares']"
            )
            if view_fares_btn.count() == 0:
                # Fallback: look for it anywhere on the page near the airline name
                view_fares_btn = page.locator(
                    f"text={airline} >> .. >> button:has-text('View Fares')"
                )

            if view_fares_btn.count() == 0:
                print(f"[WARN] 'View Fares' button not found for {airline}; skipping.")
                continue

            view_fares_btn.first.click()
            print(f"[INFO] Clicked 'View Fares' for {airline}")
            page.wait_for_timeout(2_000)
            page.wait_for_load_state("networkidle", timeout=20_000)

            # ── Scrape fare cards ────────────────────────────────────────────
            fare_card_selectors = [
                ".fare-card",
                "[data-testid='fare-card']",
                ".brand-fare-card",
                ".fare-option",
                "[class*='fareCard']",
                "[class*='fare-card']",
            ]

            fare_cards = None
            for sel in fare_card_selectors:
                cards = page.locator(sel)
                if cards.count() > 0:
                    fare_cards = cards
                    print(f"[INFO] Found {cards.count()} fare card(s) for {airline} via: {sel}")
                    break

            if fare_cards is None or fare_cards.count() == 0:
                print(f"[WARN] No fare cards found for {airline}; skipping.")
                continue

            card_count = fare_cards.count()
            for i in range(card_count):
                card = fare_cards.nth(i)
                try:
                    # ── Airline Name ──────────────────────────────────────────
                    airline_name = airline  # already known from outer loop

                    # ── Brand Fare Name ───────────────────────────────────────
                    brand_fare_name = ""
                    brand_name_selectors = [
                        ".fare-name", ".brand-name", "[data-testid='fare-name']",
                        "h3", "h4", ".fare-title", "[class*='fareName']",
                    ]
                    for bsel in brand_name_selectors:
                        el = card.locator(bsel)
                        if el.count() > 0:
                            brand_fare_name = el.first.inner_text().strip()
                            break

                    # ── Brand Inclusions ──────────────────────────────────────
                    inclusions: list[str] = []
                    inclusion_selectors = [
                        ".fare-inclusion", ".inclusion-item", "[data-testid='inclusion']",
                        "li", ".amenity", "[class*='inclusion']", "[class*='amenity']",
                    ]
                    for isel in inclusion_selectors:
                        items = card.locator(isel)
                        if items.count() > 0:
                            for j in range(items.count()):
                                text = items.nth(j).inner_text().strip()
                                if text:
                                    inclusions.append(text)
                            break
                    brand_inclusions = " | ".join(inclusions) if inclusions else "N/A"

                    # ── Route ─────────────────────────────────────────────────
                    route = f"{ORIGIN}-{DESTINATION}"
                    route_selectors = [
                        ".route", "[data-testid='route']", ".flight-route",
                        "[class*='route']",
                    ]
                    for rsel in route_selectors:
                        el = card.locator(rsel)
                        if el.count() > 0:
                            extracted = el.first.inner_text().strip()
                            if extracted:
                                route = extracted
                            break

                    # ── Date / Time ───────────────────────────────────────────
                    date_time = ""
                    dt_selectors = [
                        ".departure-time", ".flight-time", "[data-testid='departure-time']",
                        ".time", "[class*='time']", "[class*='date']",
                    ]
                    for dtsel in dt_selectors:
                        el = card.locator(dtsel)
                        if el.count() > 0:
                            date_time = el.first.inner_text().strip()
                            break

                    fare_record = {
                        "Airline":            airline_name,
                        "Brand Fare Name":    brand_fare_name or "N/A",
                        "Brand Inclusions":   brand_inclusions,
                        "Route":              route,
                        "Date/Time":          date_time or "N/A",
                        "Captured At":        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                    all_fares.append(fare_record)
                    print(f"  [INFO] Fare {i+1}: {brand_fare_name or 'N/A'} | {brand_inclusions[:60]}")

                except Exception as exc:
                    print(f"  [WARN] Error extracting fare card {i+1} for {airline}: {exc}")
                    continue

            # Close fare panel / go back if needed
            close_selectors = [
                "button:has-text('Close')", "button[aria-label='Close']",
                ".modal-close", "[data-testid='close-fares']",
            ]
            for csel in close_selectors:
                if page.locator(csel).count() > 0:
                    page.locator(csel).first.click()
                    page.wait_for_timeout(800)
                    break

        except PlaywrightTimeoutError as exc:
            print(f"[ERROR] Timeout while processing {airline}: {exc}")
            continue
        except Exception as exc:
            print(f"[ERROR] Unexpected error for {airline}: {exc}")
            continue

    print(f"\n[INFO] Total fare records captured: {len(all_fares)}")
    return all_fares


# ─────────────────────────── Export to Excel ───────────────────────────────
def export_to_excel(fares: list[dict]) -> None:
    """Append fare records to the Excel file, creating it if necessary."""
    if not fares:
        print("[WARN] No fare data to export.")
        return

    new_df = pd.DataFrame(fares, columns=[
        "Airline", "Brand Fare Name", "Brand Inclusions", "Route", "Date/Time", "Captured At"
    ])

    excel_path = Path(EXCEL_FILE)
    if excel_path.exists():
        existing_df = pd.read_excel(excel_path, sheet_name="Flight Fares")
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        combined_df = new_df

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
        combined_df.to_excel(writer, sheet_name="Flight Fares", index=False)

        # Auto-fit column widths
        ws = writer.sheets["Flight Fares"]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

    print(f"[SUCCESS] Data exported to '{EXCEL_FILE}' ({len(new_df)} new rows, {len(combined_df)} total).")


# ─────────────────────────── Main Automation Job ───────────────────────────
def run_automation() -> None:
    """Single end-to-end automation run."""
    print(f"\n{'='*60}")
    print(f"[START] Automation run at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}")

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, slow_mo=200)
        context = browser.new_context(viewport={"width": 1440, "height": 900})
        page = context.new_page()

        try:
            if not login(page):
                print("[ERROR] Login failed — aborting this run.")
                return

            if not search_flight(page):
                print("[ERROR] Flight search failed — aborting this run.")
                return

            if not apply_filters(page):
                print("[WARN] Filter application had issues — continuing anyway.")

            fares = capture_airline_fares(page)
            export_to_excel(fares)

        except Exception as exc:
            print(f"[CRITICAL] Unhandled exception during automation: {exc}")
        finally:
            context.close()
            browser.close()

    print(f"[END] Automation run completed at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")


# ─────────────────────────── Scheduler ─────────────────────────────────────
def main() -> None:
    print("[INFO] Flight Fare Automation started.")
    print(f"[INFO] Schedule: every 1 hour | Output: {EXCEL_FILE}")
    print("[INFO] Running first iteration immediately...\n")

    run_automation()  # execute once immediately on startup

    schedule.every(1).hours.do(run_automation)
    print("[INFO] Scheduler active — next run in 1 hour. Press Ctrl+C to stop.\n")

    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == "__main__":
    main()
