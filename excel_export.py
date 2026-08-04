import os
import re

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


KOLONE = [
    "redni_broj", "naziv", "kategorija", "adresa", "telefon", "email",
    "web_stranica", "radno_vrijeme", "status", "ocjena",
    "broj_recenzija", "plus_code", "opis", "google_maps_url",
    "datum_scrape", "pozvano", "zainteresovan",
]

ZAGLAVLJE = {
    "redni_broj":     "Br.",
    "naziv":          "Naziv",
    "kategorija":     "Kategorija",
    "adresa":         "Adresa",
    "telefon":        "Telefon",
    "email":          "Email",
    "web_stranica":   "Web stranica",
    "radno_vrijeme":  "Radno vrijeme",
    "status":         "Status",
    "ocjena":         "Ocjena",
    "broj_recenzija": "Recenzije",
    "plus_code":      "Plus Code",
    "opis":           "Opis",
    "google_maps_url": "Google Maps URL",
    "datum_scrape":   "Datum scrape",
    "pozvano":        "Pozvano (DA/NE)",
    "zainteresovan":  "Zainteresovan (DA/NE)",
}

SIRINE = {
    "redni_broj": 5, "naziv": 35, "kategorija": 20, "adresa": 35,
    "telefon": 16, "email": 25, "web_stranica": 35, "radno_vrijeme": 40, "status": 14,
    "ocjena": 8, "broj_recenzija": 10, "plus_code": 14, "opis": 40,
    "google_maps_url": 45, "datum_scrape": 16,
    "pozvano": 16, "zainteresovan": 20,
}


def safe_sheet_name(name: str) -> str:
    name = re.sub(r"[\\/*?:\[\]]", "", name)
    return name[:31]


def safe_filename(name: str) -> str:
    name = re.sub(r"[^\w\s\-]", "", name, flags=re.UNICODE)
    name = name.strip().replace(" ", "_")
    return name or "rezultati"


def format_sheet(ws):
    header_fill  = PatternFill("solid", fgColor="1F4E79")
    header_font  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    body_font    = Font(name="Arial", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align   = Alignment(horizontal="left",   vertical="center")
    wrap_align   = Alignment(horizontal="left",   vertical="top", wrap_text=True)
    thin         = Side(style="thin", color="CCCCCC")
    border       = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Stil zaglavlja
    for cell in ws[1]:
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = center_align
        cell.border    = border

    # Stil tijela
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font   = body_font
            cell.border = border
            idx = cell.column - 1
            key = KOLONE[idx] if idx < len(KOLONE) else ""
            if key in ("redni_broj", "ocjena", "broj_recenzija", "status", "pozvano", "zainteresovan"):
                cell.alignment = center_align
            elif key in ("radno_vrijeme", "opis", "google_maps_url"):
                cell.alignment = wrap_align
            else:
                cell.alignment = left_align

    # Sirine kolona
    for i, key in enumerate(KOLONE, start=1):
        ws.column_dimensions[get_column_letter(i)].width = SIRINE.get(key, 15)

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 22


def save_to_excel(results_map: dict, output_dir: str, skip_duplicates: bool = True):
    """
    results_map je oblika:
    { (lokacija, sekcija, kategorija): [ lista_rezultata ], ... }
    
    Kreira folder po lokaciji, excel fajl po sekciji, sheet po kategoriji.
    """
    if not EXCEL_AVAILABLE:
        print("[UPOZORENJE] openpyxl nije instaliran. Ne mogu snimiti Excel.")
        return

    os.makedirs(output_dir, exist_ok=True)
    
    # Reorganizacija mape za lakse snimanje:
    # tree[lokacija][sekcija][kategorija] = results
    tree = {}
    for (lokacija, sekcija, kategorija), results in results_map.items():
        if not results:
            continue
            
        if lokacija not in tree: tree[lokacija] = {}
        if sekcija not in tree[lokacija]: tree[lokacija][sekcija] = {}
        tree[lokacija][sekcija][kategorija] = results
        
    for lokacija, sekcije in tree.items():
        lokacija_folder = os.path.join(output_dir, safe_filename(lokacija))
        os.makedirs(lokacija_folder, exist_ok=True)
        
        for sekcija, kategorije in sekcije.items():
            filename = safe_filename(sekcija) + ".xlsx"
            filepath = os.path.join(lokacija_folder, filename)
            
            if os.path.exists(filepath):
                wb = openpyxl.load_workbook(filepath)
            else:
                wb = openpyxl.Workbook()
                default = wb.active
                if default and default.title in ("Sheet", "Sheet1"):
                    del wb[default.title]

            for kategorija, results in kategorije.items():
                sname = safe_sheet_name(kategorija)
                existing_urls = set()
                
                if sname in wb.sheetnames:
                    if not skip_duplicates:
                        del wb[sname]
                        ws = wb.create_sheet(title=sname)
                        ws.append([ZAGLAVLJE.get(k, k) for k in KOLONE])
                    else:
                        ws = wb[sname]
                        url_col_idx = KOLONE.index("google_maps_url") + 1
                        for row_idx in range(2, ws.max_row + 1):
                            val = ws.cell(row=row_idx, column=url_col_idx).value
                            if val:
                                existing_urls.add(val)
                else:
                    ws = wb.create_sheet(title=sname)
                    ws.append([ZAGLAVLJE.get(k, k) for k in KOLONE])

                added_count = 0
                for row in results:
                    url = row.get("google_maps_url", "")
                    if skip_duplicates and url in existing_urls:
                        continue
                    
                    if skip_duplicates:
                        row["redni_broj"] = ws.max_row # max_row je trenutno zadnji red
                        
                    ws.append([row.get(k, "") for k in KOLONE])
                    existing_urls.add(url)
                    added_count += 1
                
                if added_count > 0:
                    format_sheet(ws)

            wb.save(filepath)
            print(f"  [+] Sačuvano: {filepath}")
