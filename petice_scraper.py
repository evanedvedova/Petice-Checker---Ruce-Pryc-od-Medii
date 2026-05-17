"""
Petice Scraper — ruceprycodmedii.cz
=====================================
Stáhne podpisy a uloží je do CSV. Průběžně ukládá každou stránku.
Při pádu lze pokračovat z checkpointu.

Závislosti:
    pip install playwright
    playwright install chromium

Výstup:
    petice_data_YYYYMMDD_HHMMSS.csv   ← raw data
    checkpoint_scraper.json            ← průběžný stav (smaže se po dokončení)
    petice_scraper.log
"""

import csv
import json
import time
import unicodedata
from pathlib import Path
from datetime import datetime

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# ============================================================================
# KONFIG
# ============================================================================

BASE_URL        = "https://ruceprycodmedii.cz/podepsali/"
RUN_ID          = datetime.now().strftime("%Y%m%d_%H%M%S")
DEFAULT_CSV     = f"petice_data_{RUN_ID}.csv"
CHECKPOINT_FILE = "checkpoint_scraper.json"
LOG_FILE        = "petice_scraper.log"

PAGE_LOAD_TIMEOUT = 30_000
PAGE_DELAY        = 0.8
HEADLESS          = True

CSV_COLUMNS = ["cislo", "jmeno_prijmeni", "mesto", "povolani", "email"]

# ============================================================================
# LOGGING
# ============================================================================

def log(msg: str):
    ts   = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

# ============================================================================
# CHECKPOINT
# ============================================================================

def save_checkpoint(next_page: int, output_file: str, seen_numbers: set):
    data = {
        "next_page":    next_page,
        "output_file":  output_file,
        "seen_numbers": list(seen_numbers),
        "saved_at":     datetime.now().isoformat(),
    }
    with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)

def load_checkpoint():
    if not Path(CHECKPOINT_FILE).exists():
        return None
    with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def delete_checkpoint():
    p = Path(CHECKPOINT_FILE)
    if p.exists():
        p.unlink()
        log("Checkpoint smazán.")

# ============================================================================
# CSV
# ============================================================================

def init_csv(filepath: str):
    """Vytvoří CSV soubor s hlavičkou (UTF-8-BOM pro Excel)."""
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        writer.writeheader()
    log(f"Nový CSV soubor: {filepath}")

def append_to_csv(filepath: str, rows: list):
    """Přidá řádky do existujícího CSV."""
    with open(filepath, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        for row in rows:
            writer.writerow({col: row.get(col, "") for col in CSV_COLUMNS})

def read_csv(filepath: str) -> tuple[list, set]:
    """Načte existující CSV — vrátí (záznamy, množina čísel)."""
    rows = []
    seen = set()
    with open(filepath, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
            seen.add(int(row["cislo"]))
    return rows, seen

# ============================================================================
# PARSING
# ============================================================================

def parse_rows(page_obj) -> list:
    results = []
    rows = page_obj.locator("tbody tr[data-slot='table-row']").all()
    for row in rows:
        cells = row.locator("td[data-slot='table-cell']").all()
        if len(cells) < 3:
            continue
        cislo    = cells[0].inner_text().strip()
        jmeno    = cells[1].inner_text().strip()
        mesto    = cells[2].inner_text().strip()
        povolani = cells[3].inner_text().strip() if len(cells) > 3 else ""
        if not cislo.isdigit():
            continue
        results.append({
            "cislo":          cislo,
            "jmeno_prijmeni": jmeno,
            "mesto":          mesto,
            "povolani":       povolani,
            "email":          "",
        })
    return results

# ============================================================================
# SCRAPER
# ============================================================================

def scrape(start_signature: int = 1, end_signature: int = None):
    seen_numbers = set()
    output_file  = DEFAULT_CSV
    page_num     = None

    # --- Checkpoint ---
    checkpoint = load_checkpoint()
    if checkpoint:
        answer = input("\nNalezen checkpoint. Pokračovat? [Y/n]: ").strip().lower()
        if answer in ("", "y", "yes"):
            page_num     = checkpoint["next_page"]
            output_file  = checkpoint["output_file"]
            seen_numbers = set(checkpoint["seen_numbers"])
            log(f"RESUME od stránky {page_num}, soubor: {output_file}")
        else:
            log("Začínám znovu — starý checkpoint ignorován.")

    started = time.time()

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=HEADLESS)
        context = browser.new_context(user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        ))
        bpage = context.new_page()

        # --- Zjistit start stránku ---
        if page_num is None:
                log("Měřím velikost stránky...")
                try:
                    bpage.goto(f"{BASE_URL}1?sort=signature_number&dir=asc", timeout=PAGE_LOAD_TIMEOUT)
                    bpage.wait_for_selector("tbody tr[data-slot='table-row']", timeout=PAGE_LOAD_TIMEOUT)
                    first_page_rows = parse_rows(bpage)
                    page_size = len(first_page_rows) or 50
                    nums_on_p1 = [int(r["cislo"]) for r in first_page_rows]
                    sig_span = max(nums_on_p1)
                    log(f"Čísla na stránce 1: {min(nums_on_p1)}–{max(nums_on_p1)}, řádků={page_size}, span={sig_span}")
                except Exception as e:
                    log(f"Měření selhalo: {e}, defaultuji sig_span=25")
                    sig_span = 25
                page_num = max(1, int(start_signature / sig_span) - 10)
                log(f"Startovní stránka: {page_num} (start_sig={start_signature}, sig_span={sig_span})")
                init_csv(output_file)

        # --- Hlavní smyčka ---
        while True:
            url = f"{BASE_URL}{page_num}?sort=signature_number&dir=asc"
            log(f"PAGE {page_num}")

            try:
                bpage.goto(url, timeout=PAGE_LOAD_TIMEOUT)
                bpage.wait_for_selector("tbody tr[data-slot='table-row']", timeout=PAGE_LOAD_TIMEOUT)
            except PlaywrightTimeout:
                log("TIMEOUT — konec stránek nebo problém se sítí")
                break
            except Exception as e:
                log(f"ERROR: {e}")
                break

            sigs = parse_rows(bpage)
            if not sigs:
                log("NO DATA — konec stránek")
                break

            nums = [int(s["cislo"]) for s in sigs]
            log(f"  čísla: {min(nums)}–{max(nums)}")

            stop_after = False
            to_save    = []
            for s in sigs:
                cislo = int(s["cislo"])
                if cislo in seen_numbers:
                    continue
                seen_numbers.add(cislo)
                if cislo < start_signature:
                    continue
                if end_signature and cislo > end_signature:
                    stop_after = True
                    break
                to_save.append(s)

            if to_save:
                append_to_csv(output_file, to_save)

            elapsed = round(time.time() - started)
            log(f"  uloženo={len(to_save)}  elapsed={elapsed}s")

            save_checkpoint(page_num + 1, output_file, seen_numbers)
            log("  checkpoint uložen")

            if stop_after:
                log("Dosažena horní hranice rozsahu.")
                break

            page_num += 1
            time.sleep(PAGE_DELAY)

        browser.close()

    log(f"Scraping hotov. Data: {output_file}")
    delete_checkpoint()
    return output_file

# ============================================================================
# MAIN
# ============================================================================

def get_range() -> tuple:
    print("=" * 60)
    print("  PETICE SCRAPER — ruceprycodmedii.cz")
    print("=" * 60)

    while True:
        raw = input("\nOd čísla podpisu (Enter = od 1): ").strip()
        if not raw:
            start = 1; break
        if raw.isdigit() and int(raw) > 0:
            start = int(raw); break
        print("  ❌ Zadej kladné celé číslo nebo Enter.")

    while True:
        raw = input("Do čísla podpisu  (Enter = až do konce): ").strip()
        if not raw:
            end = None; break
        if raw.isdigit() and int(raw) >= start:
            end = int(raw); break
        print(f"  ❌ Zadej číslo ≥ {start} nebo Enter.")

    print(f"\n  Rozsah: {start} — {end if end else 'konec'}\n")
    return start, end


def main():
    start, end = get_range()
    output_file = scrape(start, end)
    print(f"\n✅ Hotovo. Data uložena do: {output_file}")
    print(f"   Spusť petice_analyzer.py a zadej tento soubor k analýze.")


if __name__ == "__main__":
    main()
