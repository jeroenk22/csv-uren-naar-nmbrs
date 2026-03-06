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

## Stap 3 — Installatie (eenmalig)

Open CMD (druk op Windows-toets, typ `cmd`, druk Enter) en voer deze commando's één voor één uit:

```
cd C:\Users\JouwNaam\Apps\uren-naar-nmbrs
pip install playwright python-dotenv
playwright install chromium
```

Dit kan een paar minuten duren.

---

## Stap 4 — Inloggegevens instellen (optioneel)

Wil je je gegevens opslaan zodat je ze niet elke keer hoeft in te typen?

1. Kopieer het bestand `.env.example` en hernoem de kopie naar `.env`
2. Open `.env` met Kladblok
3. Vul je gegevens in:

```
EMAIL=jouwemailadres@gmail.com
WACHTWOORD=jouwwachtwoord
```

> **Bevat je wachtwoord een `#`?** Zet het dan tussen aanhalingstekens, anders wordt alles na `#` genegeerd:
> ```
> WACHTWOORD="abc#123"
> ```

4. Sla op en sluit

> **Let op:** Laat je `.env` bestand nooit naar GitHub pushen. Dit is al geblokkeerd via `.gitignore`.  
> Heb je geen `.env` bestand? Dan vraagt de app gewoon om je gegevens bij het opstarten.

---

## Stap 5 — CSV exporteren vanuit Google Sheets of Excel

**Google Sheets:**
- Ga naar Bestand → Downloaden → Kommagescheiden waarden (.csv)

**Excel:**
- Ga naar Bestand → Opslaan als → Kies als type "CSV (door komma's gescheiden)"

Het CSV bestand moet deze kolommen hebben:

| DATUM | VAN | TOT | PAUZE | TOTAAL |
|-------|-----|-----|-------|--------|
| maandag 2 februari 2026 | 08:30 | 17:30 | | 09:00 |

---

## Stap 6 — Het programma starten

Dubbelklik op `nmbrs_uren_invullen.py`  
— of open CMD en typ:

```
python nmbrs_uren_invullen.py
```

---

## Gebruik

1. Vul je e-mailadres en wachtwoord in (als je geen `.env` hebt ingesteld)
2. Klik op **Bladeren** en selecteer je CSV bestand
3. Klik op **Start verwerking**
4. Er opent automatisch een browservenster — laat dit open staan
5. Na afloop wordt de CSV automatisch naar de map `Archief/` verplaatst

---

## Problemen?

| Probleem | Oplossing |
|----------|-----------|
| `python` werkt niet in CMD | Herinstalleer Python en vink "Add to PATH" aan |
| `pip` werkt niet | Typ `python -m pip install playwright python-dotenv` |
| Browser opent maar logt niet in | Controleer je e-mailadres en wachtwoord in `.env` |
| Uren worden niet zichtbaar in Nmbrs | Navigeer in Nmbrs handmatig naar de juiste maand |
| CSV wordt niet herkend | Zorg dat de kolomnamen exact zijn: DATUM, VAN, TOT, PAUZE, TOTAAL |
