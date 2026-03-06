# CSV Uren naar Nmbrs

Automatisch tijdregistraties invoeren in Nmbrs vanuit een CSV bestand.

---

## Vereisten

- Windows 10 of hoger
- Python 3.10 of hoger
- Internetverbinding

---

## Stap 1 — Python installeren (alleen als je het nog niet hebt)

1. Ga naar https://www.python.org/downloads/
2. Klik op de grote gele knop **"Download Python 3.x.x"**
3. Open het gedownloade bestand
4. ⚠️ **Vink onderaan "Add Python to PATH" aan** — dit is belangrijk!
5. Klik op **"Install Now"**
6. Na installatie: open CMD en typ `python --version`
   Je ziet dan zoiets als `Python 3.12.6` — dan werkt het

---

## Stap 2 — Bestanden downloaden

1. Klik rechtsboven op de groene knop **"Code"**
2. Klik op **"Download ZIP"**
3. Pak de ZIP uit naar een map, bijvoorbeeld `C:\Users\JouwNaam\Apps\uren-naar-nmbrs\`

---

## Stap 3 — Het programma starten

Dubbelklik op **`start.bat`**

De app installeert automatisch alles wat ontbreekt (playwright, python-dotenv, Chromium) en start daarna direct op. Dit kan de eerste keer een paar minuten duren.

> Wil je liever handmatig starten via CMD?
> ```
> cd C:\Users\JouwNaam\Apps\uren-naar-nmbrs
> pip install playwright python-dotenv
> playwright install chromium
> python nmbrs_uren_invullen.py
> ```

---

## Stap 4 — CSV exporteren vanuit Google Sheets of Excel

**Google Sheets:**
- Ga naar Bestand → Downloaden → Kommagescheiden waarden (.csv)

**Excel:**
- Ga naar Bestand → Opslaan als → Kies als type "CSV (door komma's gescheiden)"

Het CSV bestand moet deze kolommen hebben:

| DATUM | VAN | TOT | PAUZE | TOTAAL |
|-------|-----|-----|-------|--------|
| maandag 2 februari 2026 | 08:30 | 17:30 | | 09:00 |

> Komma's én puntkomma's als scheidingsteken worden automatisch herkend.

Rijen zonder tijden (lege VAN/TOT) worden beschouwd als vrije dagen en eventueel uit Nmbrs verwijderd als daar al een registratie stond.

---

## Gebruik

1. Vul je e-mailadres en wachtwoord in
2. Vink **"Onthoud inloggegevens"** aan als je ze wilt bewaren voor de volgende keer
3. Klik op **Bladeren** en selecteer je CSV bestand
4. Klik op **Start verwerking**
5. Er opent automatisch een browservenster — laat dit open staan

Na afloop:
- Nieuwe dagen worden **ingevoerd**
- Dagen die al bestonden maar anders zijn worden **bijgewerkt**
- Dagen zonder tijden die al in Nmbrs stonden worden **verwijderd**
- De CSV wordt **gekopieerd** naar de map `Archief/`

> **Inloggegevens opslaan via `.env`** (alternatief voor de checkbox):
> Kopieer `.env.example` naar `.env`, open het met Kladblok en vul in:
> ```
> EMAIL=jouwemailadres@gmail.com
> WACHTWOORD=jouwwachtwoord
> ```
> Bevat je wachtwoord een `#`? Zet het dan tussen aanhalingstekens: `WACHTWOORD="abc#123"`
> Let op: push `.env` nooit naar GitHub — dit is al geblokkeerd via `.gitignore`.

---

## Problemen?

| Probleem | Oplossing |
|----------|-----------|
| `python` werkt niet in CMD | Herinstalleer Python en vink "Add to PATH" aan |
| `pip` werkt niet | Typ `python -m pip install playwright python-dotenv` |
| Browser opent maar logt niet in | Controleer je e-mailadres en wachtwoord |
| Uren worden niet zichtbaar in Nmbrs | Navigeer in Nmbrs handmatig naar de juiste maand |
| CSV wordt niet herkend | Zorg dat de kolomnamen exact zijn: DATUM, VAN, TOT, PAUZE, TOTAAL |
