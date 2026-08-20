import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import sys
import asyncio
import os
import traceback

from scraper import run_scraper
from excel_export import save_to_excel
from main import load_categories, load_json

class ConsoleRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)
        
        # Obezbeđuje se da se tekst odmah renderuje
        try:
            self.text_widget.update_idletasks()
        except:
            pass

    def flush(self):
        pass

class ScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Maps Scraper - v1.1 GUI")
        self.root.geometry("750x650")
        
        # Main Frame
        main_frame = ttk.Frame(root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.stop_flag = {"stop": False}

        # Top frame for inputs
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=5)

        # Country
        ttk.Label(input_frame, text="Država (npr. Nemačka):", width=25).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.country_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.country_var, width=50).grid(row=0, column=1, sticky=tk.W, pady=5)

        # Cities
        ttk.Label(input_frame, text="Gradovi (zarezom odvojeni):", width=25).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.cities_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.cities_var, width=50).grid(row=1, column=1, sticky=tk.W, pady=5)
        ttk.Label(input_frame, text="(ostavite prazno za celu državu)", font=("Arial", 8)).grid(row=1, column=2, sticky=tk.W, padx=5)

        # Categories file
        ttk.Label(input_frame, text="Fajl sa kategorijama:", width=25).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.cat_file_var = tk.StringVar(value="kategorije.txt")
        file_frame = ttk.Frame(input_frame)
        file_frame.grid(row=2, column=1, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.cat_file_var, width=38).pack(side=tk.LEFT)
        ttk.Button(file_frame, text="Izaberi...", command=self.browse_file).pack(side=tk.LEFT, padx=5)

        # Options
        options_frame = ttk.LabelFrame(main_frame, text="Opcije", padding=10)
        options_frame.pack(fill=tk.X, pady=10)

        self.skip_duplicates_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Ne briši stare podatke, samo dodaj nove (Preporučeno da sačuvate 'Pozvano' kolone)", variable=self.skip_duplicates_var).pack(anchor=tk.W, pady=2)

        try:
            cfg = load_json("config.json")
            default_emails = cfg.get("scrape_emails", True)
        except:
            default_emails = True
            
        self.scrape_emails_var = tk.BooleanVar(value=default_emails)
        ttk.Checkbutton(options_frame, text="Učitaj sajtove klijenata i traži Email adrese (Usporava pretragu)", variable=self.scrape_emails_var).pack(anchor=tk.W, pady=2)

        # Start/Stop buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)

        self.start_btn = ttk.Button(button_frame, text="POKRENI PRETRAGU", command=self.start_scraping)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = ttk.Button(button_frame, text="ZAUSTAVI PRETRAGU", command=self.stop_scraping, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        # Console
        console_frame = ttk.LabelFrame(main_frame, text="Konzola (Uživo)", padding=10)
        console_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # Dodajemo Scrollbar
        scrollbar = ttk.Scrollbar(console_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.console_text = tk.Text(console_frame, bg="black", fg="white", font=("Consolas", 9), yscrollcommand=scrollbar.set)
        self.console_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.console_text.yview)

        # Redirect stdout
        sys.stdout = ConsoleRedirector(self.console_text)
        print("GUI pokrenut. Spreman za rad!")

    def browse_file(self):
        filename = filedialog.askopenfilename(
            initialdir=os.getcwd(),
            title="Izaberi fajl sa kategorijama",
            filetypes=(("Text files", "*.txt"), ("All files", "*.*"))
        )
        if filename:
            self.cat_file_var.set(filename)

    def stop_scraping(self):
        self.stop_flag["stop"] = True
        self.stop_btn.config(state=tk.DISABLED)
        print("\n[!] Komanda za ZAUSTAVLJANJE je poslata!")
        print("[!] Sačekajte malo da skripta završi započete pretrage i snimi podatke u Excel...\n")

    def start_scraping(self):
        country = self.country_var.get().strip()
        if not country:
            messagebox.showerror("Greška", "Morate uneti državu!")
            return

        cat_file = self.cat_file_var.get().strip()
        if not os.path.exists(cat_file):
            messagebox.showerror("Greška", f"Fajl '{cat_file}' ne postoji!")
            return
            
        self.stop_flag["stop"] = False
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.console_text.delete(1.0, tk.END)
        print(f"[{threading.current_thread().name}] Priprema zadataka...")
        
        # Pokreće se u posebnoj niti da ne bi blokiralo prozor
        threading.Thread(target=self.run_process, daemon=True).start()

    def run_process(self):
        try:
            country = self.country_var.get().strip()
            cities_input = self.cities_var.get().strip()
            cat_file = self.cat_file_var.get().strip()
            scrape_emails = self.scrape_emails_var.get()
            skip_dupes = self.skip_duplicates_var.get()

            # Učitavanje
            sections = load_categories(cat_file)
            try:
                config = load_json("config.json")
            except Exception:
                config = {"max_results_per_query": 150, "headless": True, "workers": 3, "output_dir": "rezultati"}

            # Parsiranje lokacija
            lokacije_za_pretragu = []
            if cities_input:
                gradovi = [g.strip() for g in cities_input.split(",") if g.strip()]
                for g in gradovi:
                    lokacije_za_pretragu.append(f"{g}, {country}")
            else:
                lokacije_za_pretragu.append(country)

            # Generisanje zadataka
            tasks_list = []
            for lokacija in lokacije_za_pretragu:
                for sec in sections:
                    sekcija_naziv = sec['section']
                    for kat in sec['categories']:
                        tasks_list.append((lokacija, sekcija_naziv, kat))

            print(f"Ukupno jedinstvenih upita: {len(tasks_list)}")
            print("Pokrećem Playwright...")
            
            # Scrape asinhrono
            results_map = asyncio.run(run_scraper(
                tasks_list,
                headless=config.get("headless", True),
                max_workers=config.get("workers", 3),
                max_per=config.get("max_results_per_query", 150),
                scrape_emails=scrape_emails,
                stop_flag=self.stop_flag
            ))

            print("\nSnimam podatke u Excel...")
            save_to_excel(results_map, config.get("output_dir", "rezultati"), skip_duplicates=skip_dupes)
            
            print("\n" + "="*40)
            print(" ZAVRŠENO USPEŠNO!")
            print("="*40)
            
            # Message box mora da se prikaže glavnoj niti (ali tk radi OK i ovako sa showinfo ponekad na Windowsu)
            # Najbolje je prebaciti nazad u glavnu nit
            self.root.after(0, lambda: messagebox.showinfo("Završeno", "Pretraga je uspešno završena! Podaci su sačuvani u Excel."))
            
        except Exception as e:
            print(f"\n[!] GREŠKA: {e}")
            traceback.print_exc()
            self.root.after(0, lambda: messagebox.showerror("Greška u radu", f"Došlo je do greške:\n{e}"))
        finally:
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))


if __name__ == "__main__":
    root = tk.Tk()
    app = ScraperGUI(root)
    root.mainloop()
