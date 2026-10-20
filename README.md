# CSV Uren naar Nmbrs

Automatisch tijdregistraties invoeren in Nmbrs vanuit een CSV bestand.

---

## Vereisten

- Windows 10 of hoger
- Internetverbinding

---

## Stap 1 — Bestanden downloaden

1. Klik rechtsboven op de groene knop **"Code"**
2. Klik op **"Download ZIP"**
3. Pak de ZIP uit naar een map, bijvoorbeeld `C:\Users\JouwNaam\Apps\uren-naar-nmbrs\`

---

## Stap 2 — Het programma starten

Dubbelklik op **`start.bat`**

Is Python nog niet geïnstalleerd? Dan vraagt de app of je het automatisch wilt installeren (via winget). Daarna installeert het automatisch alles wat verder ontbreekt (playwright, python-dotenv, Chromium) en start de app op. Dit kan de eerste keer een paar minuten duren.

> Wil je liever handmatig starten via CMD?
> ```
> cd C:\Users\JouwNaam\Apps\uren-naar-nmbrs
> pip install playwright python-dotenv
> playwright install chromium
> python nmbrs_uren_invullen.py
> ```

---

## Stap 3 — CSV exporteren vanuit Google Sheets of Excel

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
6. Vraagt de app om een **2FA-code**? Vul de 6-cijferige code uit je Authenticator-app in (of klik Annuleren en typ hem zelf in het browservenster)

> **Tweestapsverificatie:** het apparaat wordt 30 dagen onthouden, dus je hoeft de code niet elke keer in te vullen.
> Dit wordt bewaard in de map `.browserprofiel/` (staat niet op GitHub). Verwijder die map om opnieuw in te loggen met code.

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

> **2FA-code automatisch laten invullen** (optioneel):
> Zet de geheime sleutel van je Authenticator-koppeling in `.env`:
> ```
> TOTP_SECRET=ABCD EFGH IJKL MNOP
> ```
> Die sleutel krijg je bij het instellen van tweestapsverificatie in Visma (de tekstcode onder de QR-code, "kan je de QR-code niet scannen?"). Achteraf is hij meestal niet meer op te vragen; stel 2FA dan opnieuw in en voeg de sleutel óók toe aan je Authenticator-app.
> Let op: met deze sleutel én je wachtwoord in `.env` kan iedereen met toegang tot dit bestand inloggen zonder je telefoon. Gebruik dit alleen op je eigen pc.
> Is de code onjuist, dan vraagt de app hem alsnog handmatig.

---

## Problemen?

| Probleem | Oplossing |
|----------|-----------|
| `python` werkt niet in CMD | Herinstalleer Python en vink "Add to PATH" aan |
| `pip` werkt niet | Typ `python -m pip install playwright python-dotenv` |
| Browser opent maar logt niet in | Controleer je e-mailadres en wachtwoord |
| Uren worden niet zichtbaar in Nmbrs | Navigeer in Nmbrs handmatig naar de juiste maand |
| CSV wordt niet herkend | Zorg dat de kolomnamen exact zijn: DATUM, VAN, TOT, PAUZE, TOTAAL |
