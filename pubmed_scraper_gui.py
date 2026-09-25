"""
PubMed Affiliations Scraper – GUI + Firefox
Versiune îmbunătățită cu detectare robustă a butonului Expand
"""

import threading
import time
import os
from urllib.parse import urlparse, parse_qs

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from selenium import webdriver
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ─────────────────────────────────────────────
# Logică scraping
# ─────────────────────────────────────────────

def extract_ids_from_url(url: str) -> list:
    parsed = urlparse(url.strip())
    params = parse_qs(parsed.query)
    term = params.get("term", [""])[0]
    return [i.strip() for i in term.split(",") if i.strip().isdigit()]


def load_article_ids(filepath: str) -> list:
    seen, ids = set(), []
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            for pmid in extract_ids_from_url(line.strip()):
                if pmid not in seen:
                    seen.add(pmid)
                    ids.append(pmid)
    return ids


def build_firefox_driver(headless: bool = True):
    options = FirefoxOptions()
    if headless:
        options.add_argument("--headless")
    options.set_preference("general.useragent.override",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0")
    options.set_preference("dom.webdriver.enabled", False)
    options.set_preference("useAutomationExtension", False)
    service = FirefoxService(GeckoDriverManager().install())
    return webdriver.Firefox(service=service, options=options)


def click_expand_affiliations(driver, wait):
    """
    Încearcă toate metodele cunoscute pentru a apăsa butonul Expand din secțiunea Affiliations.
    Returnează True dacă a reușit.
    """

    # Metoda 1 – buton cu text exact "+ Expand" sau "Expand"
    try:
        btns = driver.find_elements(By.TAG_NAME, "button")
        for btn in btns:
            txt = btn.text.strip().lower()
            if "expand" in txt or txt == "+":
                driver.execute_script("arguments[0].scrollIntoView(true);", btn)
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(1.2)
                return True
    except Exception:
        pass

    # Metoda 2 – <a> cu text "expand" sau "show"
    try:
        links = driver.find_elements(By.TAG_NAME, "a")
        for a in links:
            txt = a.text.strip().lower()
            if "expand" in txt or ("show" in txt and "affil" in txt.lower()):
                driver.execute_script("arguments[0].click();", a)
                time.sleep(1.2)
                return True
    except Exception:
        pass

    # Metoda 3 – orice element cu aria-label care contine "affiliation"
    try:
        el = driver.find_element(By.XPATH,
            "//*[contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'affiliation')]"
            "[self::button or self::a]")
        driver.execute_script("arguments[0].click();", el)
        time.sleep(1.2)
        return True
    except Exception:
        pass

    # Metoda 4 – selector CSS generice PubMed cunoscute
    css_selectors = [
        "button.affiliations-button",
        "button[class*='affiliation']",
        "a[class*='affiliation']",
        ".authors-affiliations button",
        ".author-list button",
        ".expand-affiliations",
        "[data-ga-action='toggle_affiliations']",
        "button.toggle-affiliations",
        ".show-affiliations",
        "#affiliations button",
        "section.affiliations button",
        # PubMed foloseste uneori un <a> cu href=#
        "a[href='#'][class*='expan']",
        "a[href='#'][class*='toggl']",
    ]
    for sel in css_selectors:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            driver.execute_script("arguments[0].click();", el)
            time.sleep(1.2)
            return True
        except Exception:
            continue

    # Metoda 5 – JavaScript: cauta orice buton care contine "+" in text si e langa "affiliation"
    try:
        result = driver.execute_script("""
            var allBtns = document.querySelectorAll('button, a[href]');
            for (var i=0; i<allBtns.length; i++) {
                var el = allBtns[i];
                var txt = el.textContent.trim().toLowerCase();
                var parent = el.parentElement ? el.parentElement.textContent.toLowerCase() : '';
                if ((txt.includes('expand') || txt === '+' || txt.includes('show')) 
                    && (parent.includes('affil') || el.id.includes('affil') || 
                        (el.className && el.className.includes && el.className.includes('affil')))) {
                    el.click();
                    return 'clicked: ' + el.tagName + ' - ' + el.textContent.trim();
                }
            }
            return null;
        """)
        if result:
            time.sleep(1.2)
            return True
    except Exception:
        pass

    # Metoda 6 – JavaScript brut: expanda TOATE elementele ascunse din sectiunea de autori
    try:
        driver.execute_script("""
            // Forteaza vizibilitatea tuturor elementelor hidden din zona autorilor
            var selectors = [
                '.authors-affiliations', '.affiliations', '#affiliations',
                '[class*="affiliation"]', '[id*="affiliation"]'
            ];
            selectors.forEach(function(sel) {
                var els = document.querySelectorAll(sel);
                els.forEach(function(el) {
                    el.style.display = 'block';
                    el.style.visibility = 'visible';
                    el.style.maxHeight = 'none';
                    el.style.overflow = 'visible';
                    el.setAttribute('aria-hidden', 'false');
                    el.classList.remove('hidden', 'collapsed', 'hide');
                });
            });
        """)
        time.sleep(0.8)
    except Exception:
        pass

    return False


def collect_affiliations(driver) -> str:
    """
    Colectează textul affiliations după expand, folosind multiple strategii.
    """
    aff_text = ""

    # Strategie 1 – <li> din lista de affiliations
    selectors_list = [
        "ul.item-list li",
        "div.affiliations ol li",
        "div.affiliations ul li",
        "ol.affiliation-list li",
        "#affiliations li",
        "section.affiliations li",
        "[class*='affiliation'] li",
        "[id*='affiliation'] li",
    ]
    for sel in selectors_list:
        try:
            items = driver.find_elements(By.CSS_SELECTOR, sel)
            if items:
                lines = [it.text.strip() for it in items if it.text.strip()]
                if lines:
                    aff_text = "\n".join(lines)
                    return aff_text
        except Exception:
            continue

    # Strategie 2 – bloc text complet
    selectors_block = [
        "div.affiliations",
        "section.affiliations",
        "#affiliations",
        "[class*='affiliation-list']",
        "[class*='affiliations']",
        "[id*='affiliations']",
        ".authors-affiliations",
        ".author-affiliations",
    ]
    for sel in selectors_block:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            txt = el.text.strip()
            if txt and len(txt) > 5:
                aff_text = txt
                return aff_text
        except Exception:
            continue

    # Strategie 3 – XPath
    xpaths = [
        "//*[contains(@class,'affiliation')]",
        "//*[contains(@id,'affiliation')]",
        "//section[contains(@class,'affil')]",
        "//div[contains(@class,'affil')]",
    ]
    for xp in xpaths:
        try:
            els = driver.find_elements(By.XPATH, xp)
            for el in els:
                txt = el.text.strip()
                if txt and len(txt) > 10:
                    return txt
        except Exception:
            continue

    # Strategie 4 – JavaScript: cauta in DOM orice nod care contine nr + institutie
    try:
        result = driver.execute_script("""
            var candidates = document.querySelectorAll('li, p, div, span');
            var affiliations = [];
            var pattern = /^\\d+\\s+[A-Z]/;
            for (var i = 0; i < candidates.length; i++) {
                var el = candidates[i];
                var txt = el.textContent.trim();
                // Cauta elemente care par a fi affiliations (incep cu nr + text)
                if (pattern.test(txt) && txt.length > 20 && txt.length < 500
                    && el.children.length < 3) {
                    affiliations.push(txt);
                }
            }
            return affiliations.join('\\n');
        """)
        if result and len(result) > 10:
            return result
    except Exception:
        pass

    return ""


def scrape_article(driver, pmid: str) -> dict:
    url = f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/"
    result = {"url": url, "title": "", "affiliations": ""}
    try:
        driver.get(url)
        wait = WebDriverWait(driver, 15)

        # Asteapta sa se incarce pagina
        try:
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.heading-title")))
        except Exception:
            pass
        time.sleep(1.0)

        # Titlu
        try:
            title_el = driver.find_element(By.CSS_SELECTOR, "h1.heading-title")
            result["title"] = title_el.text.strip()
        except Exception:
            try:
                title_el = driver.find_element(By.CSS_SELECTOR, ".heading-title")
                result["title"] = title_el.text.strip()
            except Exception:
                result["title"] = "N/A"

        # Apasa Expand
        click_expand_affiliations(driver, wait)

        # Colecteaza affiliations
        aff = collect_affiliations(driver)
        result["affiliations"] = aff if aff else "Affiliations negăsite"

    except Exception as e:
        result["title"] = f"EROARE: {e}"
        result["affiliations"] = ""
    return result


# ─────────────────────────────────────────────
# Scriere Excel
# ─────────────────────────────────────────────

AFFILIATION_KEYWORDS = (
    "university", "universit", "institute", "instit", "department", "hospital",
    "center", "centre", "laboratory", "laborator", "school", "college",
    "faculty", "clinic", "division", "unit", "academy", "research",
)


def _looks_like_affiliation(text: str) -> bool:
    """True daca textul contine cuvinte tipice unei afilieri reale
    (universitate, institut, spital etc.). Folosit ca sa filtram
    textul gresit prins uneori din pagina (ex. citarea revistei)."""
    low = text.lower()
    return any(kw in low for kw in AFFILIATION_KEYWORDS)


TARGET_COUNTRIES = ["Lithuania", "Latvia", "Estonia", "Ukraine", "Belarus", "Kazakhstan"]


def _matched_countries(text: str) -> list:
    """Returneaza lista tarilor din TARGET_COUNTRIES gasite in textul afilierii."""
    low = text.lower()
    return [c for c in TARGET_COUNTRIES if c.lower() in low]


def _problem_reason(rec: dict) -> str:
    """Motivul problemei daca afilierea lipseste/e invalida/eroare, altfel string gol."""
    if rec["title"].startswith("EROARE:"):
        return rec["title"]
    aff = rec["affiliations"]
    if not aff or aff == "Affiliations negăsite":
        return "Affiliations negăsite"
    if not _looks_like_affiliation(aff):
        return "Affiliations negăsite"
    return ""


def write_excel(records: list, output_path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "PubMed Affiliations"
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 100

    HEADER_FILL = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
    HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    SEP_FILL    = PatternFill("solid", start_color="2E86C1", end_color="2E86C1")
    URL_FONT    = Font(name="Arial", bold=True, color="1F4E79", size=10)
    LABEL_FONT  = Font(name="Arial", bold=True, color="2E4053", size=10)
    BODY_FONT   = Font(name="Arial", size=10)
    THIN        = Side(style="thin", color="BDBDBD")
    BORDER      = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    CENTER      = Alignment(horizontal="center", vertical="center")
    WRAP        = Alignment(wrap_text=True, vertical="top")

    ws.merge_cells("A1:B1")
    c = ws["A1"]
    c.value = "PubMed – Raport Affiliations"
    c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER
    ws.row_dimensions[1].height = 28

    row = 2
    for idx, rec in enumerate(records, 1):
        ws.merge_cells(f"A{row}:B{row}")
        sc = ws.cell(row=row, column=1, value=f"── Articol #{idx} ──")
        sc.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        sc.fill = SEP_FILL; sc.alignment = CENTER
        ws.row_dimensions[row].height = 18
        row += 1

        for label, value, is_url in [
            ("Pagina:", rec["url"], True),
            ("Titlu:", rec["title"], False),
            ("Affiliations:", rec["affiliations"], False),
        ]:
            lc = ws.cell(row=row, column=1, value=label)
            lc.font = LABEL_FONT; lc.border = BORDER; lc.alignment = WRAP

            vc = ws.cell(row=row, column=2, value=value)
            vc.font = URL_FONT if is_url else BODY_FONT
            vc.border = BORDER; vc.alignment = WRAP
            if is_url:
                vc.hyperlink = value

            lines = value.count("\n") + 1 if value else 1
            ws.row_dimensions[row].height = max(18, lines * 15)
            row += 1

        row += 1

    # ── Sheet separat cu articolele fara afiliere gasita ──
    problems = [(rec, _problem_reason(rec)) for rec in records]
    problems = [(rec, reason) for rec, reason in problems if reason]

    ws2 = wb.create_sheet("Afiliere_Lipsa")
    ws2.column_dimensions["A"].width = 45
    ws2.column_dimensions["B"].width = 60
    ws2.column_dimensions["C"].width = 35

    ws2.merge_cells("A1:C1")
    h = ws2["A1"]
    h.value = f"Articole cu Afiliere Lipsă ({len(problems)})"
    h.font = HEADER_FONT; h.fill = PatternFill("solid", start_color="C0392B", end_color="C0392B")
    h.alignment = CENTER
    ws2.row_dimensions[1].height = 24

    for col, title in enumerate(["Pagina", "Titlu", "Motiv"], start=1):
        hc = ws2.cell(row=2, column=col, value=title)
        hc.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        hc.fill = HEADER_FILL; hc.alignment = CENTER; hc.border = BORDER

    row2 = 3
    for rec, reason in problems:
        uc = ws2.cell(row=row2, column=1, value=rec["url"])
        uc.font = URL_FONT; uc.border = BORDER; uc.alignment = WRAP
        uc.hyperlink = rec["url"]

        tc = ws2.cell(row=row2, column=2, value=rec["title"])
        tc.font = BODY_FONT; tc.border = BORDER; tc.alignment = WRAP

        rc = ws2.cell(row=row2, column=3, value=reason)
        rc.font = Font(name="Arial", color="C0392B", size=10, bold=True)
        rc.border = BORDER; rc.alignment = WRAP

        ws2.row_dimensions[row2].height = 18
        row2 += 1

    if problems:
        ws2.freeze_panes = "A3"

    # ── Sheet separat cu articolele afiliate in tarile tinta ──
    target_hits = [(rec, _matched_countries(rec["affiliations"])) for rec in records]
    target_hits = [(rec, countries) for rec, countries in target_hits if countries]

    ws3 = wb.create_sheet("Tari_Tinta")
    ws3.column_dimensions["A"].width = 45
    ws3.column_dimensions["B"].width = 60
    ws3.column_dimensions["C"].width = 25
    ws3.column_dimensions["D"].width = 60

    ws3.merge_cells("A1:D1")
    h3 = ws3["A1"]
    h3.value = f"Articole cu Afiliere in {', '.join(TARGET_COUNTRIES)} ({len(target_hits)})"
    h3.font = HEADER_FONT; h3.fill = PatternFill("solid", start_color="27AE60", end_color="27AE60")
    h3.alignment = CENTER
    ws3.row_dimensions[1].height = 24

    for col, title in enumerate(["Pagina", "Titlu", "Tara detectata", "Affiliations"], start=1):
        hc = ws3.cell(row=2, column=col, value=title)
        hc.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        hc.fill = HEADER_FILL; hc.alignment = CENTER; hc.border = BORDER

    row3 = 3
    for rec, countries in target_hits:
        uc = ws3.cell(row=row3, column=1, value=rec["url"])
        uc.font = URL_FONT; uc.border = BORDER; uc.alignment = WRAP
        uc.hyperlink = rec["url"]

        tc = ws3.cell(row=row3, column=2, value=rec["title"])
        tc.font = BODY_FONT; tc.border = BORDER; tc.alignment = WRAP

        cc = ws3.cell(row=row3, column=3, value=", ".join(countries))
        cc.font = Font(name="Arial", color="27AE60", size=10, bold=True)
        cc.border = BORDER; cc.alignment = WRAP

        ac = ws3.cell(row=row3, column=4, value=rec["affiliations"])
        ac.font = BODY_FONT; ac.border = BORDER; ac.alignment = WRAP

        lines = rec["affiliations"].count("\n") + 1 if rec["affiliations"] else 1
        ws3.row_dimensions[row3].height = max(18, lines * 15)
        row3 += 1

    if target_hits:
        ws3.freeze_panes = "A3"

    wb.save(output_path)


# ─────────────────────────────────────────────
# Interfață grafică
# ─────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PubMed Affiliations Scraper")
        self.resizable(False, False)
        self.configure(bg="#F0F4F8")
        self._running = False
        self._build_ui()
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"+{x}+{y}")

    def _build_ui(self):
        BG   = "#F0F4F8"
        BLUE = "#1F4E79"

        title_fr = tk.Frame(self, bg=BLUE)
        title_fr.pack(fill="x")
        tk.Label(title_fr, text="🔬 PubMed Affiliations Scraper",
                 bg=BLUE, fg="white", font=("Arial", 15, "bold"), pady=12).pack()

        body = tk.Frame(self, bg=BG, padx=18, pady=14)
        body.pack(fill="both")

        tk.Label(body, text="Fișier URL-uri (.txt):", bg=BG,
                 font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=4)
        self.input_var = tk.StringVar()
        tk.Entry(body, textvariable=self.input_var, width=52,
                 font=("Arial", 10)).grid(row=0, column=1, padx=6)
        tk.Button(body, text="Browse…", command=self._pick_input,
                  bg="#2E86C1", fg="white", font=("Arial", 9, "bold"),
                  relief="flat", padx=8).grid(row=0, column=2)

        tk.Label(body, text="Salvează Excel (.xlsx):", bg=BG,
                 font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", pady=4)
        self.output_var = tk.StringVar(value="pubmed_affiliations.xlsx")
        tk.Entry(body, textvariable=self.output_var, width=52,
                 font=("Arial", 10)).grid(row=1, column=1, padx=6)
        tk.Button(body, text="Browse…", command=self._pick_output,
                  bg="#2E86C1", fg="white", font=("Arial", 9, "bold"),
                  relief="flat", padx=8).grid(row=1, column=2)

        opts = tk.Frame(body, bg=BG)
        opts.grid(row=2, column=0, columnspan=3, sticky="w", pady=8)

        self.headless_var = tk.BooleanVar(value=False)
        tk.Checkbutton(opts, text="Headless (Firefox invizibil)",
                       variable=self.headless_var, bg=BG,
                       font=("Arial", 10)).pack(side="left", padx=(0, 20))

        tk.Label(opts, text="Pauză între pagini (sec):", bg=BG,
                 font=("Arial", 10)).pack(side="left")
        self.delay_var = tk.DoubleVar(value=2.0)
        tk.Spinbox(opts, from_=0.5, to=10.0, increment=0.5,
                   textvariable=self.delay_var, width=5,
                   font=("Arial", 10)).pack(side="left", padx=6)

        btn_fr = tk.Frame(body, bg=BG)
        btn_fr.grid(row=3, column=0, columnspan=3, pady=6)
        self.start_btn = tk.Button(btn_fr, text="▶  START Scraping",
                                   command=self._start,
                                   bg="#1F4E79", fg="white",
                                   font=("Arial", 12, "bold"),
                                   relief="flat", padx=20, pady=8)
        self.start_btn.pack(side="left", padx=8)
        self.stop_btn = tk.Button(btn_fr, text="■  STOP",
                                  command=self._stop,
                                  bg="#C0392B", fg="white",
                                  font=("Arial", 12, "bold"),
                                  relief="flat", padx=20, pady=8,
                                  state="disabled")
        self.stop_btn.pack(side="left", padx=8)

        self.progress_var = tk.IntVar(value=0)
        self.progress = ttk.Progressbar(body, variable=self.progress_var,
                                        maximum=100, length=560)
        self.progress.grid(row=4, column=0, columnspan=3, pady=(6, 2))
        self.status_var = tk.StringVar(value="Gata de pornire.")
        tk.Label(body, textvariable=self.status_var, bg=BG,
                 font=("Arial", 9), fg="#555").grid(row=5, column=0,
                 columnspan=3, sticky="w")

        tk.Label(body, text="Log:", bg=BG,
                 font=("Arial", 10, "bold")).grid(row=6, column=0,
                 sticky="w", pady=(10, 2))
        log_fr = tk.Frame(body)
        log_fr.grid(row=7, column=0, columnspan=3, sticky="nsew")
        self.log = tk.Text(log_fr, height=14, width=78, state="disabled",
                           font=("Consolas", 9), bg="#1A1A2E", fg="#E0E0E0",
                           relief="flat", wrap="word")
        sb = ttk.Scrollbar(log_fr, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def _pick_input(self):
        p = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All", "*.*")])
        if p:
            self.input_var.set(p)

    def _pick_output(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                         filetypes=[("Excel", "*.xlsx")])
        if p:
            self.output_var.set(p)

    def _log(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_status(self, msg: str):
        self.status_var.set(msg)

    def _start(self):
        inp = self.input_var.get().strip()
        out = self.output_var.get().strip()
        if not inp:
            messagebox.showerror("Eroare", "Selectează fișierul .txt cu URL-urile!")
            return
        if not out:
            messagebox.showerror("Eroare", "Specifică fișierul de output .xlsx!")
            return
        if not os.path.isfile(inp):
            messagebox.showerror("Eroare", f"Fișierul nu există:\n{inp}")
            return

        self._running = True
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.progress_var.set(0)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

        t = threading.Thread(target=self._worker, args=(inp, out), daemon=True)
        t.start()

    def _stop(self):
        self._running = False
        self._log("⛔ Stop solicitat de utilizator...")
        self._set_status("Oprire în curs…")

    def _worker(self, inp: str, out: str):
        try:
            self._log(f"📂 Citesc: {inp}")
            ids = load_article_ids(inp)
            total = len(ids)
            if not ids:
                self.after(0, lambda: messagebox.showerror(
                    "Eroare", "Nu s-au găsit ID-uri valide în fișier!"))
                return
            self._log(f"🔢 Total articole: {total}")
            self._log("🦊 Pornesc Firefox…")
            self._set_status("Pornire Firefox…")

            driver = build_firefox_driver(headless=self.headless_var.get())
            records = []
            try:
                for i, pmid in enumerate(ids, 1):
                    if not self._running:
                        break
                    url = f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/"
                    self._log(f"[{i}/{total}] {url}")
                    self._set_status(f"Procesez {i}/{total}: PMID {pmid}")

                    rec = scrape_article(driver, pmid)
                    records.append(rec)

                    preview = rec["title"][:65] + "…" if len(rec["title"]) > 65 else rec["title"]
                    n_aff = rec["affiliations"].count("\n") + 1 if rec["affiliations"] and rec["affiliations"] != "Affiliations negăsite" else 0
                    self._log(f"    📌 {preview}")
                    if n_aff:
                        self._log(f"    🏛  {n_aff} affiliations găsite ✓")
                    else:
                        self._log(f"    ⚠️  Affiliations negăsite")

                    pct = int(i / total * 100)
                    self.after(0, lambda p=pct: self.progress_var.set(p))
                    time.sleep(self.delay_var.get())

            finally:
                driver.quit()
                self._log("🛑 Firefox închis.")

            if records:
                self._log(f"\n💾 Salvez Excel: {out}")
                write_excel(records, out)
                self._log(f"✅ Gata! {len(records)} articole salvate în:\n   {out}")
                self._set_status(f"✅ Finalizat! {len(records)} articole salvate.")
                self.after(0, lambda: messagebox.showinfo(
                    "Finalizat!", f"Raportul a fost salvat!\n\n{out}"))
            else:
                self._log("⚠️ Niciun articol procesat.")
                self._set_status("Niciun articol procesat.")

        except Exception as e:
            self._log(f"❌ EROARE: {e}")
            self._set_status("Eroare!")
            self.after(0, lambda: messagebox.showerror("Eroare", str(e)))
        finally:
            self._running = False
            self.after(0, lambda: self.start_btn.configure(state="normal"))
            self.after(0, lambda: self.stop_btn.configure(state="disabled"))


if __name__ == "__main__":
    app = App()
    app.mainloop()
