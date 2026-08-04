import json
import os
import sys
import asyncio
from scraper import run_scraper
from excel_export import save_to_excel

def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def load_categories(path):
    sections = []
    current_section = None
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            if line.startswith('#'):
                current_section = {'section': line.lstrip('#').strip(), 'categories': []}
                sections.append(current_section)
            elif current_section:
                current_section['categories'].append(line)
            else:
                # Ako neko nije stavio # Naslov na pocetku
                current_section = {'section': 'Ostalo', 'categories': [line]}
                sections.append(current_section)
    return sections

def print_menu(title, options):
    print(f"\n--- {title} ---")
    for i, opt in enumerate(options, 1):
        print(f" {i}. {opt}")
    while True:
        try:
            choice = int(input(f"Izaberite (1-{len(options)}): "))
            if 1 <= choice <= len(options):
                return choice - 1
            print("Pogrešan unos, pokušajte ponovo.")
        except ValueError:
            print("Molimo unesite broj.")

def main():
    print("="*50)
    print(" GOOGLE MAPS SCRAPER v1.1 - ARCHVIZ & DESIGN")
    print("="*50)

    # Učitaj konfiguraciju
    try:
        config = load_json("config.json")
    except Exception:
        config = {"max_results_per_query": 20, "headless": True, "workers": 3, "output_dir": "rezultati"}

    # Učitaj kategorije
    try:
        sections = load_categories("kategorije.txt")
    except Exception:
        print("[!] Fajl kategorije.txt nije pronađen.")
        sys.exit(1)

    print("\n--- IZBOR LOKACIJE ---")
    odabrana_drzava = input("Unesite ime države (npr. 'Nemačka', 'Srbija', 'USA'): ").strip()
    if not odabrana_drzava:
        print("[!] Morate uneti državu!")
        sys.exit(1)
        
    unos_gradova = input(f"Unesite ime grada u {odabrana_drzava} (ili više gradova odvojenih zarezom, npr. 'Berlin, Hamburg').\nOstavite prazno za pretragu na nivou cele države: ").strip()
    
    lokacije_za_pretragu = []
    if unos_gradova:
        gradovi = [g.strip() for g in unos_gradova.split(",") if g.strip()]
        for g in gradovi:
            lokacije_za_pretragu.append(f"{g}, {odabrana_drzava}")
    else:
        lokacije_za_pretragu.append(odabrana_drzava)

    print("\nPriprema zadataka...")
    tasks_list = []
    for lokacija in lokacije_za_pretragu:
        for sec in sections:
            sekcija_naziv = sec['section']
            for kat in sec['categories']:
                tasks_list.append((lokacija, sekcija_naziv, kat))

    ukupno_zadataka = len(tasks_list)
    print(f"Ukupno jedinstvenih pretraga za izvršavanje: {ukupno_zadataka}")
    print("Pokrećem Playwright u asinhronom režimu...")

    # Pokreni scraping
    results_map = asyncio.run(run_scraper(
        tasks_list,
        headless=config.get("headless", True),
        max_workers=config.get("workers", 3),
        max_per=config.get("max_results_per_query", 20),
        scrape_emails=config.get("scrape_emails", False)
    ))

    # Snimanje u Excel
    print("\nSnimam podatke u Excel fajlove...")
    save_to_excel(results_map, config.get("output_dir", "rezultati"))
    
    print("\n" + "="*50)
    print(" SVE ZAVRŠENO USPEŠNO!")
    print("="*50)

if __name__ == "__main__":
    main()
