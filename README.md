# Petice_RPOM_checker Manual
## Nástroj pro kontrolu podpisů petice ruceprycodmedii.cz

Eva Nedvědová 
https://www.linkedin.com/in/eva-nedvedova-4015b5154/
Dotazy, připomínky, pochvaly i nápady na zlepšení jsou vítané — klidně napište.

datum aktualizace 17.05.2026
---

## Co nástroj dělá

Nástroj se skládá ze dvou samostatných skriptů:

- **`petice_scraper.py`** — stáhne podpisy z webu a uloží je do CSV souboru
- **`petice_analyzer.py`** — načte CSV a vyhledá problematické záznamy, výsledek uloží do Excelu

Skripty se spouštějí zvlášť. Scraper může běžet i několik hodin — analyzer pak zpracuje stažená data za pár sekund.

---

## Požadavky

- Windows 10 nebo novější
- Python 3.10 nebo novější
  → stáhni na https://www.python.org/downloads/
  → při instalaci zaškrtni **"Add Python to PATH"**

---

## První spuštění — jednorázové nastavení

Toto proveď jen jednou při prvním použití.

### 1. Otevři složku se skripty v příkazovém řádku

Otevři Průzkumník souborů, přejdi do složky kde máš uložené skripty.
Klikni do adresního řádku nahoře, napiš `cmd` a stiskni Enter.

Otevře se příkazový řádek přímo ve správné složce.

### 2. Vytvoř virtuální prostředí

```
python -m venv venv
```

Vznikne podsložka `venv` — virtuální prostředí pro instalaci knihoven.

### 3. Aktivuj virtuální prostředí

```
venv\Scripts\activate
```

Poznáš, že je aktivní — na začátku řádku se objeví `(venv)`.

> ⚠️ Virtuální prostředí musíš aktivovat **pokaždé** když otevřeš nový příkazový řádek.
> Stačí kroky: otevři složku v cmd → spusť `venv\Scripts\activate` → spusť skript.

### 4. Nainstaluj potřebné knihovny

```
pip install playwright openpyxl
```

```
playwright install chromium
```

> Druhý příkaz stáhne neviditelný prohlížeč (~150 MB). Chvilku to potrvá.

---

## Každodenní použití

### Krok 1 — Otevři složku v příkazovém řádku

Průzkumník souborů → složka se skripty → klikni do adresního řádku → napiš `cmd` → Enter.

### Krok 2 — Aktivuj virtuální prostředí

```
venv\Scripts\activate
```

Na začátku řádku se musí zobrazit `(venv)`.

---

## Scraper — stahování dat

### Spuštění

```
python petice_scraper.py
```

### Co se stane

Skript se zeptá na rozsah čísel podpisů:

```
Od čísla podpisu (Enter = od 1):
Do čísla podpisu  (Enter = až do konce):
```

- Stiskni **Enter** pro stahování od začátku / do konce.
- Nebo zadej konkrétní čísla, např. `70001` a `140000`.

### Výstup

Vznikne soubor s názvem ve formátu:
```
petice_data_20260515_143022.csv
```

Datum a čas v názvu ti řekne kdy byl soubor vytvořen.

### Přerušení a obnovení

Pokud scraper přerušíš (Ctrl+C) nebo spadne, vznikne soubor `checkpoint_scraper.json`.
Při příštím spuštění se skript zeptá:

```
Nalezen checkpoint. Pokračovat? [Y/n]:
```

Stiskni **Enter** nebo napiš `y` — skript pokračuje přesně od místa kde skončil.
Napiš `n` — skript začne znovu od začátku.

### Soubory které skript vytváří

| Soubor | Popis |
|--------|-------|
| `petice_data_DATUM.csv` | Stažená data — tento soubor si uchovej |
| `checkpoint_scraper.json` | Průběžný stav — smaže se po dokončení |
| `petice_scraper.log` | Záznam průběhu stahování |

### Tipy

- Scraper stahuje přibližně **1 stránku za sekundu**. Celá petice (~150 000 podpisů) zabere 1–2 hodiny.
- Počítač může zůstat zamčený, scraper běží na pozadí.
- Pokud chceš stahovat postupně (po dávkách), zadávej rozsahy:
  - První run: `1` až `70000`
  - Druhý run: `70001` až `140000`
  - atd.

---

## Analyzer — analýza dat

### Spuštění

```
python petice_analyzer.py
```

### Co se stane

Skript nabídne tři možnosti:

```
[1] Analyzovat jeden soubor
[2] Sloučit všechny petice_data_*.csv ve složce a analyzovat
[3] Vybrat soubory ručně
```

**Možnost 1** — analyzuje jeden CSV soubor. Skript automaticky nabídne nejnovější soubor ve složce.

**Možnost 2** — doporučeno pro pravidelné použití. Najde všechny `petice_data_*.csv` soubory ve složce, sloučí je (každé číslo podpisu jen jednou) a analyzuje dohromady. Tím se zachytí i duplicity přes hranici souborů.

**Možnost 3** — zadáš cesty k souborům ručně, jeden po druhém. Prázdný řádek ukončí výběr.

### Výstup

Vznikne Excel soubor:
```
petice_problemy_20260515_143022.xlsx
```

Obsahuje dva listy:
- **Problematické podpisy** — nalezené záznamy s barevným rozlišením
- **Legenda** — vysvětlení barev

### Sloupce ve výstupu

| Sloupec | Popis |
|---------|-------|
| DONE? | Vyplňuje člověk — TRUE pokud záznam zpracován |
| Číslo podpisu | Číslo z petice (více čísel oddělených čárkou = duplicity) |
| Jméno a Příjmení | Jak bylo zadáno v petici |
| Email | Není veřejně zobrazován — prázdné |
| Co s tím? | Vyplňuje člověk — např. Smazat |
| Komentář | Důvod označení + nalezená slova |
| Město | Jak bylo zadáno v petici |
| Povolání | Jak bylo zadáno v petici |

### Barevné rozlišení

| Barva | Kategorie |
|-------|-----------|
| 🔴 Červená | Sprosté slovo |
| 🟠 Oranžová | Podezřelé slovo (historické postavy, hanlivé výrazy) |
| 🟢 Zelená | Sloveso v jméně nebo městě |
| 🔵 Modrá | Podezřelá délka jména nebo města |
| 🟡 Žlutá | Duplicita |
| 🟣 Fialová | Více kategorií najednou |

### Prahy podezřelé délky

Nastaveno v horní části `petice_analyzer.py`:

```python
MIN_JMENO = 9    # jméno kratší než 9 znaků → podezřelé
MAX_JMENO = 20   # jméno delší než 20 znaků → podezřelé
MAX_MESTO = 20   # město delší než 20 znaků → podezřelé
```

Hodnoty lze upravit podle výsledků měřícího skriptu.

---

## Doporučený pracovní postup

```
1.  Spusť petice_scraper.py
    → zadej rozsah nebo stáhni vše
    → vznikne petice_data_DATUM.csv

2.  Spusť petice_analyzer.py
    → zvol možnost 2 (sloučit vše)
    → vznikne petice_problemy_DATUM.xlsx

3.  Otevři Excel, projdi záznamy
    → sloupec "Co s tím?" vyplň ručně
    → sloupec "DONE?" označ jako TRUE po zpracování

4.  Při dalším kole stahování:
    → spusť scraper znovu s novým rozsahem čísel
    → vznikne nový petice_data_*.csv ve stejné složce
    → spusť analyzer s možností 2 — sloučí automaticky
```

---

## Časté problémy

**`ModuleNotFoundError: No module named 'playwright'`**
→ Virtuální prostředí není aktivované, nebo knihovny nejsou nainstalované.
Zkontroluj že vidíš `(venv)` na začátku řádku. Pokud ne, spusť:
```
venv\Scripts\activate
pip install playwright openpyxl
playwright install chromium
```

**`❌ Soubor nenalezen`**
→ Zkontroluj cestu k souboru. Nejjednodušší řešení: dej CSV soubor do stejné složky jako skripty.

**Scraper se zastaví po pár stránkách**
→ Pravděpodobně výpadek internetu nebo dočasná nedostupnost webu.
Znovu spusť scraper a pokračuj z checkpointu (stiskni Enter nebo y).

**Analyzer nenajde žádné CSV soubory (možnost 2)**
→ CSV soubory musí být ve stejné složce jako `petice_analyzer.py` a musí mít název začínající `petice_data_`.

---

## Přizpůsobení — kde co měnit

Vše se mění přímo v souboru `petice_analyzer.py` v textovém editoru (např. Notepad++, VS Code).

| Co chceš změnit | Kde to najdeš v kódu |
|-----------------|----------------------|
| Přidat sprosté slovo | Seznam `PROFANITY_LIST` |
| Přidat podezřelé slovo | Seznam `SUSPICIOUS_LIST` |
| Přidat sloveso | Seznam `VERB_WORDS` |
| Přidat výjimku pro město | Slovník `PROFANITY_CITY_WHITELIST` |
| Přidat výjimku pro příjmení | Slovník `PROFANITY_NAME_WHITELIST` |
| Změnit prahy délky | Proměnné `MIN_JMENO`, `MAX_JMENO`, `MAX_MESTO` |
