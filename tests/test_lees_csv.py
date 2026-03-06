import textwrap
import pytest
from nmbrs_uren_invullen import lees_csv, parse_datum
from datetime import datetime


# ── parse_datum ───────────────────────────────────────────────────────────────

def test_parse_datum_normaal():
    assert parse_datum("maandag 2 maart 2026") == datetime(2026, 3, 2)

def test_parse_datum_alle_maanden():
    maanden = [
        ("januari", 1), ("februari", 2), ("maart", 3), ("april", 4),
        ("mei", 5), ("juni", 6), ("juli", 7), ("augustus", 8),
        ("september", 9), ("oktober", 10), ("november", 11), ("december", 12),
    ]
    for naam, nummer in maanden:
        assert parse_datum(f"maandag 1 {naam} 2026").month == nummer

def test_parse_datum_hoofdletter():
    assert parse_datum("Dinsdag 3 Maart 2026") == datetime(2026, 3, 3)


# ── lees_csv ──────────────────────────────────────────────────────────────────

@pytest.fixture
def csv_bestand(tmp_path):
    def _maak(inhoud):
        pad = tmp_path / "test.csv"
        pad.write_text(inhoud, encoding='utf-8-sig')
        return str(pad)
    return _maak


def test_komma_scheidingsteken(csv_bestand):
    pad = csv_bestand("DATUM,VAN,TOT,PAUZE,TOTAAL\nmaandag 2 maart 2026,08:30,17:00,,\n")
    rijen, te_verwijderen = lees_csv(pad)
    assert len(rijen) == 1
    assert te_verwijderen == []
    assert rijen[0]['datum'] == '02-03-2026'

def test_puntkomma_scheidingsteken(csv_bestand):
    pad = csv_bestand("DATUM;VAN;TOT;PAUZE;TOTAAL\nmaandag 2 maart 2026;08:30;17:00;;\n")
    rijen, te_verwijderen = lees_csv(pad)
    assert len(rijen) == 1
    assert rijen[0]['datum'] == '02-03-2026'

def test_lege_regel_naar_te_verwijderen(csv_bestand):
    pad = csv_bestand(
        "DATUM,VAN,TOT,PAUZE,TOTAAL\n"
        "maandag 2 maart 2026,08:30,17:00,,\n"
        "dinsdag 3 maart 2026,,,,\n"
    )
    rijen, te_verwijderen = lees_csv(pad)
    assert len(rijen) == 1
    assert '03-03-2026' in te_verwijderen

def test_voorloopnul_van_uur_verwijderd(csv_bestand):
    pad = csv_bestand("DATUM,VAN,TOT,PAUZE,TOTAAL\nmaandag 2 maart 2026,08:30,09:45,,\n")
    rijen, _ = lees_csv(pad)
    assert rijen[0]['van_uur'] == '8'
    assert rijen[0]['van_min'] == '30'
    assert rijen[0]['tot_uur'] == '9'
    assert rijen[0]['tot_min'] == '45'

def test_middernacht_uur_nul(csv_bestand):
    pad = csv_bestand("DATUM,VAN,TOT,PAUZE,TOTAAL\nmaandag 2 maart 2026,00:00,08:00,,\n")
    rijen, _ = lees_csv(pad)
    assert rijen[0]['van_uur'] == '0'

def test_meerdere_dagen(csv_bestand):
    pad = csv_bestand(
        "DATUM,VAN,TOT,PAUZE,TOTAAL\n"
        "maandag 2 maart 2026,08:30,17:00,,\n"
        "dinsdag 3 maart 2026,09:00,18:00,,\n"
        "woensdag 4 maart 2026,,,,\n"
    )
    rijen, te_verwijderen = lees_csv(pad)
    assert len(rijen) == 2
    assert len(te_verwijderen) == 1
    assert rijen[0]['datum'] == '02-03-2026'
    assert rijen[1]['datum'] == '03-03-2026'
    assert te_verwijderen[0] == '04-03-2026'

def test_lege_csv(csv_bestand):
    pad = csv_bestand("DATUM,VAN,TOT,PAUZE,TOTAAL\n")
    rijen, te_verwijderen = lees_csv(pad)
    assert rijen == []
    assert te_verwijderen == []

def test_ongeldige_datum_overgeslagen(csv_bestand):
    pad = csv_bestand(
        "DATUM,VAN,TOT,PAUZE,TOTAAL\n"
        "geen datum,08:30,17:00,,\n"
        "maandag 2 maart 2026,08:30,17:00,,\n"
    )
    rijen, _ = lees_csv(pad)
    assert len(rijen) == 1

def test_datum_zonder_inhoud_overgeslagen(csv_bestand):
    pad = csv_bestand(
        "DATUM,VAN,TOT,PAUZE,TOTAAL\n"
        ",08:30,17:00,,\n"
        "maandag 2 maart 2026,08:30,17:00,,\n"
    )
    rijen, _ = lees_csv(pad)
    assert len(rijen) == 1
