"""
Nmbrs Tijdregistratie Automatisch Invullen
==========================================
Vereisten installeren (eenmalig):
  pip install playwright
  playwright install chromium

Gebruik:
  python nmbrs_uren_invullen.py
"""

import csv
import time
import shutil
import threading
import os
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# ─────────────────────────────────────────
ENV_PAD = Path(__file__).parent / '.env'
try:
    from dotenv import load_dotenv, set_key
    load_dotenv(ENV_PAD)
    DOTENV_BESCHIKBAAR = True
except ImportError:
    DOTENV_BESCHIKBAAR = False

EMAIL      = os.getenv("EMAIL", "")
WACHTWOORD = os.getenv("WACHTWOORD", "")
# ─────────────────────────────────────────

NMBRS_LOGIN_URL = "https://login.nmbrsapp.com/index.html"
NMBRS_START_URL = "https://onlinesalarisportal.nmbrs.nl"

MAANDEN_NL = {
    "januari": 1, "februari": 2, "maart": 3, "april": 4,
    "mei": 5, "juni": 6, "juli": 7, "augustus": 8,
    "september": 9, "oktober": 10, "november": 11, "december": 12
}

def parse_datum(datum_str):
    delen = datum_str.strip().split()
    dag   = int(delen[1])
    maand = MAANDEN_NL[delen[2].lower()]
    jaar  = int(delen[3])
    return datetime(jaar, maand, dag)

def lees_csv(bestand):
    rijen = []
    with open(bestand, newline='', encoding='utf-8-sig') as f:
        # Detecteer automatisch komma of puntkomma
        sample = f.read(1024)
        f.seek(0)
        scheidingsteken = ';' if sample.count(';') > sample.count(',') else ','
        reader = csv.DictReader(f, delimiter=scheidingsteken)
        for row in reader:
            if row['VAN'] and row['TOT']:
                datum = parse_datum(row['DATUM'])
                van_uur, van_min = row['VAN'].split(':')
                tot_uur, tot_min = row['TOT'].split(':')
                rijen.append({
                    'datum':   datum.strftime('%d-%m-%Y'),
                    'van_uur': van_uur.lstrip('0') or '0',
                    'van_min': van_min,
                    'tot_uur': tot_uur.lstrip('0') or '0',
                    'tot_min': tot_min,
                })
    return rijen

def sluit_popup(page, selector, naam, timeout=5000):
    try:
        page.click(selector, timeout=timeout)
        time.sleep(1)
    except Exception:
        pass

def archiveer_csv(csv_pad):
    archief_map = Path(csv_pad).parent / "Archief"
    archief_map.mkdir(exist_ok=True)
    doel = archief_map / Path(csv_pad).name
    shutil.move(csv_pad, doel)
    return doel

def voer_tijdregistraties_in(email, wachtwoord, rijen, log_func, klaar_func, focus_func=None):
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            page = browser.new_page()
            if focus_func:
                focus_func()

            log_func("Inloggen bij Nmbrs...")
            page.goto(NMBRS_LOGIN_URL)
            page.wait_for_load_state('networkidle')

            sluit_popup(page, '#onetrust-accept-btn-handler', 'Cookie popup')

            page.fill('input[type="email"]', email)
            page.click('#LoginButton')
            page.wait_for_load_state('networkidle')
            time.sleep(1)

            sluit_popup(page, '#onetrust-accept-btn-handler', 'Cookie popup')

            page.fill('input[type="password"]', wachtwoord)
            page.click('#LoginButton')
            page.wait_for_load_state('networkidle')
            time.sleep(2)

            # Account kiezen — werkt met 0, 1 of meerdere profielen
            try:
                profielen = page.locator('ul.account-picker-profile')
                aantal = profielen.count()
                if aantal > 0:
                    profielen.first.click()
                    page.wait_for_load_state('networkidle')
                    time.sleep(2)
                    log_func("Account gekozen!")
            except Exception:
                pass

            log_func("Ingelogd!")
            sluit_popup(page, '.walkme-custom-balloon-close-button', 'Walkme popup')

            log_func("Naar Tijdregistratie navigeren...")
            try:
                page.click('#widgetCopilotTabMenu', timeout=5000)
                time.sleep(1)
            except Exception:
                pass

            try:
                page.click('#tabmap_timeregistration-248', timeout=5000)
            except Exception:
                page.locator('span.card--list_text:has-text("Tijdregistratie")').first.click()

            page.wait_for_load_state('networkidle')
            time.sleep(2)
            log_func("Tijdregistratie pagina geladen!")

            # CSRF token onderscheppen
            csrf_token = {'value': ''}
            def handle_request(request):
                if 'TijdregistratieEditHandler' in request.url and request.method == 'POST':
                    token = request.headers.get('__antixsrftoken', '')
                    if token:
                        csrf_token['value'] = token

            page.on('request', handle_request)
            try:
                page.evaluate("""
                    () => {
                        const links = document.querySelectorAll('a, button, [onclick]');
                        for (const el of links) {
                            const text = el.textContent + (el.getAttribute('onclick') || '');
                            if (text.includes('+') || el.classList.contains('plus')) {
                                el.click(); return true;
                            }
                        }
                        return false;
                    }
                """)
                time.sleep(2)
            except Exception:
                pass
            try:
                page.keyboard.press('Escape')
            except Exception:
                pass
            page.remove_listener('request', handle_request)

            # Per dag invullen
            succes = 0
            mislukt = 0
            overgeslagen = 0
            for rij in rijen:
                datum = rij['datum']

                # Stap 1: controleer of identieke registratie al bestaat
                al_aanwezig = page.evaluate(f"""
                    async () => {{
                        let csrfToken = '';
                        if (window.__antixsrftoken) csrfToken = window.__antixsrftoken;
                        if (!csrfToken) {{
                            const meta = document.querySelector('meta[name="__antixsrftoken"]');
                            if (meta) csrfToken = meta.content;
                        }}
                        if (!csrfToken) {{
                            const inputs = document.querySelectorAll('input[type="hidden"]');
                            for (const h of inputs) {{
                                if (h.name.toLowerCase().includes('token') || h.name.toLowerCase().includes('xsrf')) {{
                                    csrfToken = h.value; break;
                                }}
                            }}
                        }}
                        const headers = {{'Content-Type': 'application/x-www-form-urlencoded'}};
                        if (csrfToken) headers['__antixsrftoken'] = csrfToken;
                        try {{
                            const checkData = new URLSearchParams();
                            checkData.append('action', 'get');
                            checkData.append('args', JSON.stringify({{
                                "popupwindow_tijdEdit_datum": "{datum}"
                            }}));
                            const resp = await fetch('/handlers/popups/MedewerkerLogin/TijdregistratieEditHandler.ashx?rnd=' + Math.random(), {{
                                method: 'POST', headers: headers, body: checkData
                            }});
                            const text = await resp.text();
                            const json = JSON.parse(text);
                            if (Array.isArray(json)) {{
                                return json.some(e =>
                                    String(e.starttijdUur  ?? e.van_uur  ?? '') === '{rij['van_uur']}'  &&
                                    String(e.starttijdMinuut ?? e.van_min ?? '') === '{rij['van_min']}' &&
                                    String(e.eindtijdUur   ?? e.tot_uur  ?? '') === '{rij['tot_uur']}'  &&
                                    String(e.eindtijdMinuut  ?? e.tot_min ?? '') === '{rij['tot_min']}'
                                );
                            }}
                        }} catch(e) {{}}
                        return false;
                    }}
                """)

                if al_aanwezig:
                    log_func(f"  ⏭️  {datum} — overgeslagen (identiek al aanwezig)")
                    overgeslagen += 1
                    time.sleep(0.2)
                    continue

                # Stap 2: invoeren
                result = page.evaluate(f"""
                    async () => {{
                        let csrfToken = '';
                        if (window.__antixsrftoken) csrfToken = window.__antixsrftoken;
                        if (!csrfToken) {{
                            const meta = document.querySelector('meta[name="__antixsrftoken"]');
                            if (meta) csrfToken = meta.content;
                        }}
                        if (!csrfToken) {{
                            const inputs = document.querySelectorAll('input[type="hidden"]');
                            for (const h of inputs) {{
                                if (h.name.toLowerCase().includes('token') || h.name.toLowerCase().includes('xsrf')) {{
                                    csrfToken = h.value; break;
                                }}
                            }}
                        }}
                        const data = new URLSearchParams();
                        data.append('action', 'save');
                        data.append('args', JSON.stringify({{
                            "handler_url": "/handlers/popups/MedewerkerLogin/TijdregistratieEditHandler.ashx",
                            "popupwindow_tijdEdit_id": "0",
                            "Hidden1": "/handlers/popups/MedewerkerLogin/TijdregistratieEditHandler.ashx",
                            "popupwindow_tijdEdit_datum": "{datum}",
                            "popupwindow_tijdEdit_starttijdUur": "{rij['van_uur']}",
                            "popupwindow_tijdEdit_starttijdMinuut": "{rij['van_min']}",
                            "popupwindow_tijdEdit_eindtijdUur": "{rij['tot_uur']}",
                            "popupwindow_tijdEdit_eindtijdMinuut": "{rij['tot_min']}",
                            "popupwindow_tijdEdit_project": "0",
                            "popupwindow_tijdEdit_status": "2"
                        }}));
                        const headers = {{'Content-Type': 'application/x-www-form-urlencoded'}};
                        if (csrfToken) headers['__antixsrftoken'] = csrfToken;
                        const resp = await fetch('/handlers/popups/MedewerkerLogin/TijdregistratieEditHandler.ashx?rnd=' + Math.random(), {{
                            method: 'POST', headers: headers, body: data
                        }});
                        const text = await resp.text();
                        return {{ status: resp.status, body: text.substring(0, 200) }};
                    }}
                """)

                body = result.get('body', '')
                status = result.get('status')

                if 'access_denied' in body or 'CSRF' in body:
                    log_func(f"  ❌ {datum} — CSRF geweigerd")
                    mislukt += 1
                elif status == 200:
                    log_func(f"  ✅ {datum} — ingevoerd")
                    succes += 1
                else:
                    log_func(f"  ❌ {datum} — mislukt (status {status})")
                    mislukt += 1

                time.sleep(0.5)

            time.sleep(2)
            browser.close()
            klaar_func(succes, mislukt, overgeslagen)

    except Exception as e:
        log_func(f"\n❌ Fout: {e}")
        klaar_func(0, 0, 0)


# ── GUI ──────────────────────────────────────────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Nmbrs Uren Invullen")
        self.root.geometry("560x700")
        self.root.resizable(False, False)
        self.root.configure(bg="#1a1a2e")
        self.csv_pad = None
        self._build_ui()

    def _label(self, parent, tekst):
        tk.Label(parent, text=tekst, font=("Courier New", 9),
                 bg="#1a1a2e", fg="#a0a0c0").pack(anchor="w")

    def _build_ui(self):
        # Titel
        tk.Label(self.root, text="NMBRS", font=("Courier New", 28, "bold"),
                 bg="#1a1a2e", fg="#e94560").pack(pady=(28, 0))
        tk.Label(self.root, text="Uren Automatisch Invullen",
                 font=("Courier New", 10), bg="#1a1a2e", fg="#a0a0c0").pack(pady=(2, 20))

        frame = tk.Frame(self.root, bg="#1a1a2e")
        frame.pack(padx=30, fill="x")

        # Login velden
        self._label(frame, "E-mailadres")
        self.email_var = tk.StringVar(value=EMAIL)
        tk.Entry(frame, textvariable=self.email_var,
                 font=("Courier New", 10), bg="#16213e", fg="#e0e0e0",
                 insertbackground="white", relief="flat",
                 highlightthickness=1, highlightbackground="#0f3460",
                 highlightcolor="#e94560").pack(fill="x", ipady=6, pady=(4, 12))

        self._label(frame, "Wachtwoord")
        self.pass_var = tk.StringVar(value=WACHTWOORD)
        tk.Entry(frame, textvariable=self.pass_var, show="●",
                 font=("Courier New", 10), bg="#16213e", fg="#e0e0e0",
                 insertbackground="white", relief="flat",
                 highlightthickness=1, highlightbackground="#0f3460",
                 highlightcolor="#e94560").pack(fill="x", ipady=6, pady=(4, 12))

        self.onthoud_var = tk.BooleanVar(value=bool(EMAIL and WACHTWOORD))
        tk.Checkbutton(frame, text="Onthoud inloggegevens",
                       variable=self.onthoud_var,
                       font=("Courier New", 9), bg="#1a1a2e", fg="#a0a0c0",
                       selectcolor="#16213e", activebackground="#1a1a2e",
                       activeforeground="#e0e0e0").pack(anchor="w", pady=(0, 12))

        # CSV selectie
        self._label(frame, "CSV bestand")
        row = tk.Frame(frame, bg="#1a1a2e")
        row.pack(fill="x", pady=(4, 0))

        self.csv_label = tk.Label(row, text="Geen bestand geselecteerd",
                                  font=("Courier New", 9), bg="#16213e", fg="#a0a0c0",
                                  anchor="w", padx=10, pady=8, relief="flat", width=38)
        self.csv_label.pack(side="left", fill="x", expand=True)

        tk.Button(row, text="Bladeren", font=("Courier New", 9, "bold"),
                  bg="#e94560", fg="white", relief="flat", padx=12, pady=8,
                  cursor="hand2", command=self.kies_csv).pack(side="right", padx=(8, 0))

        # Log venster
        tk.Label(self.root, text="Log", font=("Courier New", 9),
                 bg="#1a1a2e", fg="#a0a0c0").pack(anchor="w", padx=30, pady=(20, 4))

        self.log = scrolledtext.ScrolledText(
            self.root, height=14, font=("Courier New", 9),
            bg="#0f3460", fg="#e0e0e0", relief="flat",
            insertbackground="white", state="disabled",
            padx=10, pady=8
        )
        self.log.pack(padx=30, fill="x")

        self.start_btn = tk.Button(
            self.root, text="▶  START VERWERKING",
            font=("Courier New", 11, "bold"),
            bg="#e94560", fg="white", relief="flat",
            padx=20, pady=12, cursor="hand2",
            command=self.start
        )
        self.start_btn.pack(pady=20)

    def _breng_naar_voren(self):
        self.root.lift()
        self.root.focus_force()

    def kies_csv(self):
        pad = filedialog.askopenfilename(
            title="Selecteer CSV bestand",
            filetypes=[("CSV bestanden", "*.csv")]
        )
        if pad:
            self.csv_pad = pad
            self.csv_label.config(text=Path(pad).name, fg="#e0e0e0")

    def log_schrijf(self, tekst):
        self.log.config(state="normal")
        self.log.insert("end", tekst + "\n")
        self.log.see("end")
        self.log.config(state="disabled")
        self.root.update_idletasks()

    def start(self):
        email      = self.email_var.get().strip()
        wachtwoord = self.pass_var.get()

        if self.onthoud_var.get():
            if DOTENV_BESCHIKBAAR:
                set_key(str(ENV_PAD), 'EMAIL', email)
                set_key(str(ENV_PAD), 'WACHTWOORD', wachtwoord)
        else:
            if ENV_PAD.exists():
                ENV_PAD.unlink()

        if not email or not wachtwoord:
            messagebox.showwarning("Inloggegevens", "Vul je e-mailadres en wachtwoord in.")
            return

        if not self.csv_pad:
            messagebox.showwarning("Geen bestand", "Selecteer eerst een CSV bestand.")
            return

        try:
            rijen = lees_csv(self.csv_pad)
        except Exception as e:
            messagebox.showerror("Fout", f"CSV kon niet worden gelezen:\n{e}")
            return

        if not rijen:
            messagebox.showinfo("Leeg", "Geen werkdagen met uren gevonden in de CSV.")
            return

        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

        self.start_btn.config(state="disabled", text="Bezig...")
        self.log_schrijf(f"📂 {len(rijen)} werkdagen gevonden")
        self.log_schrijf("─" * 42)

        def run():
            voer_tijdregistraties_in(email, wachtwoord, rijen, self.log_schrijf, self.klaar,
                                     lambda: self.root.after(1500, self._breng_naar_voren))

        threading.Thread(target=run, daemon=True).start()

    def klaar(self, succes, mislukt, overgeslagen=0):
        self.log_schrijf("─" * 42)
        self.log_schrijf(f"✅ {succes} ingevoerd   ⏭️  {overgeslagen} overgeslagen   ❌ {mislukt} mislukt")

        if succes > 0:
            try:
                doel = archiveer_csv(self.csv_pad)
                self.log_schrijf(f"📁 CSV gearchiveerd naar: {Path(doel).parent.name}/")
            except Exception as e:
                self.log_schrijf(f"⚠️  Archiveren mislukt: {e}")

        self.start_btn.config(state="normal", text="▶  START VERWERKING")
        self.csv_pad = None
        self.csv_label.config(text="Geen bestand geselecteerd", fg="#a0a0c0")


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()