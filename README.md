# PubMed Affiliations Scraper

Aplicație desktop (Tkinter + Selenium + Firefox) care citește o listă de PMID-uri PubMed, extrage titlul și afilierile autorilor pentru fiecare articol și generează un raport Excel.

## Cerințe

- Windows cu **Python 3.10+** instalat
- **Mozilla Firefox** instalat
- Acces la internet (pentru descărcarea automată a `geckodriver` și pentru accesarea PubMed)

## Instalare

### 1. Verifică Python

```
python --version
```

### 2. Instalează dependențele Python

Din folderul proiectului, fără drepturi de admin (instalare pe userul local):

```
pip install --user selenium webdriver_manager openpyxl
```

**Dacă ești într-o rețea corporate** cu un index intern de pip (ex. Artifactory/Nexus) care cere autentificare sau dă erori, instalează direct din PyPI public:

```
pip install --user --index-url https://pypi.org/simple --trusted-host pypi.org --trusted-host files.pythonhosted.org selenium webdriver_manager openpyxl
```

### 3. Erori SSL la descărcarea geckodriver

Dacă la prima rulare apare o eroare de tipul `SSLCertVerificationError` / `certificate verify failed` (frecvent în rețele corporate cu inspecție TLS), instalează:

```
pip install --user --index-url https://pypi.org/simple --trusted-host pypi.org --trusted-host files.pythonhosted.org pip_system_certs
```

Acest pachet face Python să folosească magazinul de certificate Windows (unde certificatul corporate e deja de încredere), în loc de bundle-ul intern `certifi`.

### 4. Firefox

Instalează [Mozilla Firefox](https://www.mozilla.org/firefox/) dacă nu e deja prezent pe sistem. `geckodriver` (driverul necesar Selenium) este descărcat automat de `webdriver_manager` la prima rulare — nu trebuie instalat manual.

## Utilizare

1. Pornește aplicația:

   ```
   python pubmed_scraper_gui.py
   ```

2. Pregătește un fișier `.txt` cu URL-uri PubMed, câte unul pe linie, în formatul:

   ```
   https://pubmed.ncbi.nlm.nih.gov/?term=<PMID>
   ```

   (parametrul `term` poate conține și mai multe PMID-uri separate prin virgulă)

3. În interfață:
   - Selectează fișierul `.txt` de input
   - Alege unde se salvează fișierul `.xlsx` de output
   - Opțional: bifează **Headless** ca Firefox să ruleze invizibil, și ajustează pauza dintre pagini
   - Apasă **START Scraping**

## Structura raportului Excel

- **PubMed Affiliations** — toate articolele procesate: pagina, titlul și afilierile găsite
- **Afiliere_Lipsa** — articolele pentru care nu s-a găsit o afiliere reală (text lipsă sau text care nu pare o afiliere validă)
- **Tari_Tinta** — articolele a căror afiliere conține Lithuania, Latvia, Estonia, Ukraine, Belarus sau Kazakhstan

## Probleme cunoscute

- Rulări repetate/foarte dese pot declanșa o limitare temporară (rate limiting) din partea PubMed
- Articolele foarte vechi din PubMed uneori nu au o secțiune structurată de affiliations
