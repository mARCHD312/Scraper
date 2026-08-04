import asyncio
import re
from datetime import datetime
from playwright.async_api import async_playwright

async def scrape_one_query(page, query: str, location: str, max_results: int, scrape_emails: bool = False) -> list:
    """Scrape jedne pretrage za lokaciju. location može biti 'Grad' ili 'Država' ili 'Grad, Država'."""
    results = []
    full_query = f"{query}, {location}"
    search_url = f"https://www.google.com/maps/search/{full_query.replace(' ', '+')}"

    try:
        await page.goto(search_url, wait_until="domcontentloaded", timeout=30000)
        await asyncio.sleep(2)

        # Prihvati cookies
        for sel in [
            'button:has-text("Accept all")',
            'button:has-text("Prihvati sve")',
            'form:nth-child(2) button',
        ]:
            try:
                btn = page.locator(sel).first
                if await btn.is_visible(timeout=1000):
                    await btn.click()
                    await asyncio.sleep(1)
                    break
            except Exception:
                pass

        # Čekaj feed
        try:
            await page.wait_for_selector('div[role="feed"]', timeout=10000)
        except Exception:
            print(f"    [!] Feed nije pronađen za: {full_query} (možda nema rezultata)")
            return []

        # Skrolaj
        last_count = 0
        no_change = 0
        while True:
            cards = await page.locator('div[role="feed"] a[href*="maps/place"]').all()
            if len(cards) >= max_results:
                break
            try:
                await page.locator('div[role="feed"]').evaluate("el => el.scrollBy(0, 2000)")
            except Exception:
                pass
            await asyncio.sleep(1.5)
            
            try:
                if await page.locator('div[role="feed"] span:has-text("end of the list")').count():
                    break
            except Exception:
                pass
            
            if len(cards) == last_count:
                no_change += 1
                if no_change >= 3:
                    break
            else:
                no_change = 0
            last_count = len(cards)

        # Ekstrakcija stavki
        items = await page.locator('div[role="feed"] > div').all()
        processed = 0

        for item in items:
            if processed >= max_results:
                break
            try:
                link_el = item.locator('a[href*="maps/place"]').first
                if not await link_el.count():
                    continue
                href = await link_el.get_attribute("href") or ""

                # Pronalazak naziva (brzi fallback)
                name = ""
                for sel in [
                    '[class*="fontHeadlineSmall"]',
                    'div[class*="qBF1Pd"]',
                    'div[class*="NrDZNb"]',
                ]:
                    try:
                        el = item.locator(sel).first
                        if await el.count():
                            txt = (await el.inner_text(timeout=800)).strip()
                            if txt and not txt.lower().startswith("rezultat"):
                                name = txt
                                break
                    except Exception:
                        pass

                if not name:
                    continue

                # Brzi podaci iz teksta kartice
                card_text = ""
                try:
                    card_text = await item.inner_text(timeout=1000)
                except Exception:
                    pass
                lines = [l.strip() for l in card_text.splitlines() if l.strip()]

                ocjena = next((l.replace(",", ".") for l in lines if re.match(r"^[1-5][.,][0-9]$", l)), "")
                broj_rec = ""
                for l in lines:
                    m = re.search(r"\(([\d.,]+)\)", l)
                    if m:
                        broj_rec = m.group(1).replace(".", "").replace(",", "")
                        break

                maps_url = href if href.startswith("http") else f"https://www.google.com{href}"

                # Klikni za više detalja
                extra = {}
                try:
                    await link_el.click(timeout=5000)
                    await asyncio.sleep(2)
                    await page.wait_for_selector('h1', timeout=4000)
                    extra = await extract_detail_page(page, name)
                    await page.go_back(timeout=6000, wait_until="domcontentloaded")
                    await asyncio.sleep(1)
                except Exception:
                    # Oporavak u slucaju greske unazad
                    try:
                        await page.goto(search_url, wait_until="domcontentloaded", timeout=15000)
                        await asyncio.sleep(2)
                    except Exception:
                        pass

                results.append({
                    "redni_broj":     processed + 1,
                    "naziv":          extra.get("naziv") or name,
                    "kategorija":     extra.get("kategorija", ""),
                    "adresa":         extra.get("adresa", ""),
                    "telefon":        extra.get("telefon", ""),
                    "web_stranica":   extra.get("web_stranica", ""),
                    "radno_vrijeme":  extra.get("radno_vrijeme", ""),
                    "status":         extra.get("trenutno_otvoreno", ""),
                    "ocjena":         extra.get("ocjena") or ocjena,
                    "broj_recenzija": extra.get("broj_recenzija") or broj_rec,
                    "plus_code":      extra.get("plus_code", ""),
                    "opis":           extra.get("opis", ""),
                    "google_maps_url": maps_url,
                    "datum_scrape":   datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "pozvano":        "",
                    "zainteresovan":  "",
                })
                processed += 1
                print(f"      [{processed}] {results[-1]['naziv'][:50]}")

            except Exception as e:
                # pass over errors for individual items
                continue

    except Exception as e:
        print(f"    [!] Greška pretrage: {e}")

    if scrape_emails:
        print(f"    [*] Pokrećem email scraping za {len(results)} rezultata...")
        for r in results:
            if r.get("web_stranica"):
                # Pokusaj pretrage
                print(f"      Tražim email na: {r['web_stranica']}")
                r["email"] = await extract_email_from_website(page, r["web_stranica"])
            else:
                r["email"] = ""

    return results

async def extract_email_from_website(page, url: str) -> str:
    if not url.startswith("http"): url = "http://" + url
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=12000)
        await asyncio.sleep(2)
        content = await page.content()
        
        # Pronalazak emailova u tekstu
        emails = set(re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', content))
        # Filtriranje false-positive emailova slika, css fajlova itd.
        valid_emails = [e for e in emails if not any(x in e.lower() for x in ['.png', '.jpg', '.jpeg', '.gif', '.css', '.js', '.webp', 'example'])]
        if valid_emails:
            return valid_emails[0]
            
        # Ako nema, trazi /contact ili /kontakt
        for link in await page.locator("a").all():
            try:
                href = await link.get_attribute("href")
                if href and ("contact" in href.lower() or "kontakt" in href.lower() or "about" in href.lower()):
                    if href.startswith("/"):
                        href = url.rstrip("/") + href
                    elif not href.startswith("http"):
                        continue
                    
                    await page.goto(href, wait_until="domcontentloaded", timeout=10000)
                    content = await page.content()
                    emails = set(re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', content))
                    valid_emails = [e for e in emails if not any(x in e.lower() for x in ['.png', '.jpg', '.jpeg', '.gif', '.css', '.js', '.webp', 'example'])]
                    if valid_emails:
                        return valid_emails[0]
                    break # Proveravamo samo prvi "kontakt" link
            except Exception:
                pass
    except Exception:
        pass
    return ""

async def extract_detail_page(page, fallback_name: str = "") -> dict:
    """Ekstrahuje detalje sa otvorene stranice biznisa."""
    data = {
        "naziv": fallback_name, "kategorija": "", "adresa": "", "telefon": "",
        "web_stranica": "", "radno_vrijeme": "", "trenutno_otvoreno": "",
        "ocjena": "", "broj_recenzija": "", "plus_code": "", "opis": ""
    }

    try:
        # Kategorija
        for sel in ['button[jsaction*="category"]', 'span[jsaction*="category"]']:
            el = page.locator(sel).first
            if await el.count():
                data["kategorija"] = (await el.inner_text(timeout=1000)).strip()
                break

        # Adresa
        for sel in ['[data-item-id="address"]', 'button[data-tooltip="Copy address"]']:
            el = page.locator(sel).first
            if await el.count():
                data["adresa"] = (await el.inner_text(timeout=1000)).strip()
                break

        # Telefon
        for sel in ['[data-item-id^="phone:"]', 'a[href^="tel:"]']:
            el = page.locator(sel).first
            if await el.count():
                val = (await el.get_attribute("href") or await el.inner_text(timeout=1000) or "")
                data["telefon"] = val.replace("tel:", "").strip()
                break

        # Web stranica
        for sel in ['a[data-item-id="authority"]']:
            el = page.locator(sel).first
            if await el.count():
                val = await el.get_attribute("href") or (await el.inner_text(timeout=1000)).strip()
                if val and "google" not in val:
                    data["web_stranica"] = val
                break

        # Radno vrijeme
        el = page.locator('[data-item-id*="oh"], [aria-label*="hours"]').first
        if await el.count():
            val = await el.get_attribute("aria-label") or await el.inner_text(timeout=1000)
            data["radno_vrijeme"] = val.strip().replace("\n", " | ")
            
    except Exception:
        pass

    return data

async def worker(queue, browser, headless, max_per, scrape_emails):
    """Worker koji uzima zadatke (kategorija, lokacija) iz reda i parsira ih koristeći jedan browser."""
    # Svaki worker dobija svoj kontekst da se pretrage ne bi preklapale
    context = await browser.new_context(
        locale="hr",
        viewport={"width": 1366, "height": 900},
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    )
    page = await context.new_page()
    
    results_map = {} # Kljuc je (lokacija, sekcija, kategorija)
    
    while True:
        task = await queue.get()
        if task is None:
            # End of queue
            break
        
        lokacija, sekcija, kategorija = task
        print(f"  -> Pokrećem pretragu: {kategorija} @ {lokacija}")
        res = await scrape_one_query(page, kategorija, lokacija, max_per, scrape_emails)
        results_map[(lokacija, sekcija, kategorija)] = res
        queue.task_done()
        
    await context.close()
    return results_map

async def run_scraper(tasks_list, headless=True, max_workers=3, max_per=20, scrape_emails=False):
    """Glavna funkcija koja orkestrira scraping koristeći asyncio."""
    queue = asyncio.Queue()
    for task in tasks_list:
        queue.put_nowait(task)
        
    # Dodaj 'None' zadatke da ugasiš workere na kraju
    for _ in range(max_workers):
        queue.put_nowait(None)
        
    all_results = []
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        
        workers = [
            asyncio.create_task(worker(queue, browser, headless, max_per, scrape_emails))
            for _ in range(max_workers)
        ]
        
        results_dicts = await asyncio.gather(*workers)
        
        await browser.close()
        
        # Spajanje svih mapa iz workera
        final_map = {}
        for r_map in results_dicts:
            final_map.update(r_map)
            
    return final_map
