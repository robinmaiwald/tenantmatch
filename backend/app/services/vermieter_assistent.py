# =============================================================================
#  Vermieter-Assistent: den passenden Mieter aus vielen Bewerbungen finden
# =============================================================================
#
#  WAS MACHT DIESES SKRIPT?
#    Vermieter bekommen oft Hunderte Bewerbungen auf eine Wohnung.
#    Dieser Assistent hilft beim Auswählen:
#
#      1. Er fragt den Vermieter in einem lockeren Gespräch nach der WOHNUNG
#         (Lage, Zimmer, Miete, Fahrstuhl ...) und nach seinen ANFORDERUNGEN
#         an Mieter (Rauchen, Haustiere, Mindesteinkommen, Unterlagen ...).
#         Er antwortet immer in der Sprache, in der der Vermieter schreibt.
#      2. Er zeigt eine Zusammenfassung; der Vermieter bestätigt oder korrigiert.
#      3. Er bewertet ALLE Bewerber aus der Datenbank nach festen, nachvoll-
#         ziehbaren Regeln (reines Python, keine KI -> gleiche Eingabe gibt
#         immer das gleiche Ergebnis).
#      4. Er zeigt die besten Kandidaten mit einer kurzen KI-Einschätzung und
#         speichert eine Excel-Datei mit ALLEN Bewerbern, die besten oben.
#
#  STARTEN:
#      .venv\Scripts\python.exe vermieter_assistent.py
#      Beenden jederzeit mit: exit / beenden / quit
#
#  EINBAU IN EINE APP:
#    - Datenbank anbinden: nur den Abschnitt DATENQUELLE ändern.
#    - Bewertung ohne Gespräch nutzen: bewerber_bewerten(bewerber, wohnung,
#      anforderungen) aufrufen – liefert die sortierte Liste.
#    - Web-Oberfläche: gespraech_schritt() führt genau eine Gesprächsrunde;
#      der Zustand ist ein einfaches Dict (in Session/Datenbank speicherbar).
#
#  FAIRNESS: Kriterien wie Alter, Geschlecht, Herkunft, Name oder Religion
#  sind nach dem Allgemeinen Gleichbehandlungsgesetz (AGG) unzulässig. Der
#  Assistent nimmt sie nicht an, und sie fließen nie in die Bewertung ein.
# =============================================================================

import json
import os
import re
import sys
import time
from datetime import date, datetime

import httpx                      # wird von google-genai mitinstalliert
from dotenv import load_dotenv
from google import genai
from google.genai import errors, types
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pydantic import BaseModel, Field, ValidationError

# Ordner dieses Skripts – Standardpfade werden relativ dazu gebildet.
SKRIPT_ORDNER = os.path.dirname(os.path.abspath(__file__))

# API-Key (GEMINI_API_KEY) aus der .env-Datei neben diesem Skript laden.
load_dotenv()


# =============================================================================
#  EINSTELLUNGEN
# =============================================================================

# Gemini-Modelle in Reihenfolge der Bevorzugung. Ist eines überlastet oder
# sein Tageskontingent aufgebraucht (kostenlos: ca. 20 Anfragen/Tag je
# Flash-Modell), wird automatisch das nächste genommen. Jedes Modell hat
# ein eigenes Kontingent – mehrere Einträge = mehr Anfragen pro Tag.
MODELLE = [
    os.getenv("GEMINI_MODELL", "gemini-flash-latest"),
    "gemini-3.8-flash",
    "gemini-3.7-flash",
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-flash-lite-latest",
]
VERSUCHE = 3                # Runden über alle Modelle, falls alle ausfallen
MAX_GESPRAECHSRUNDEN = 40   # Sicherheitsgrenze, damit das Gespräch endet
TOP_ANZAHL = 5              # so viele Kandidaten werden ausführlich gezeigt
ENDWOERTER = ("exit", "beenden", "quit")

# Gewichtung der Bewertung (Summe = 100 Punkte). Hier lässt sich festlegen,
# was dem Vermieter wie wichtig ist.
GEWICHTE = {
    "einkommen": 30,        # Nettoeinkommen im Verhältnis zur Gesamtmiete
    "schufa": 25,           # SCHUFA-Score und -Status
    "mietschulden": 20,     # Mietschuldenfreiheitsbescheinigung
    "beschaeftigung": 15,   # wie sicher ist das Einkommen?
    "einzug": 5,            # passt der Wunsch-Einzug zum Verfügbarkeitsdatum?
    "bewerbung": 5,         # Sorgfalt des Anschreibens
}

# Pfade für Bewerberdaten und Ergebnisse.
PROJEKT_ORDNER = os.path.abspath(
    os.path.join(SKRIPT_ORDNER, "../../..")
)

BEWERBER_DATEI = os.path.join(
    PROJEKT_ORDNER,
    "data",
    "sample_applicants.json",
)

AUSGABE_ORDNER = os.path.join(
    SKRIPT_ORDNER,
    "ausgabe",
)


# =============================================================================
#  DATENQUELLE  –  woher kommen die Bewerbungen?
# =============================================================================
#  Liefert eine Liste von Dicts (ein Dict pro Bewerbung, Felder wie in
#  applicants.json). Für eine echte Datenbank muss NUR diese Funktion
#  ausgetauscht werden – der Rest des Skripts bleibt gleich.
# =============================================================================

def bewerber_laden(pfad=BEWERBER_DATEI):
    """Lädt alle Bewerbungen aus der JSON-Datei."""
    if not os.path.isfile(pfad):
        raise FileNotFoundError(f"Bewerberdatei nicht gefunden: {pfad}")
    with open(pfad, encoding="utf-8") as datei:
        return json.load(datei)


# =============================================================================
#  DATENMODELLE  –  was wird beim Vermieter erfragt?
# =============================================================================
#  Alle Felder sind optional (None = noch nicht bekannt). Die Beschreibungen
#  bekommt auch die KI – so weiß sie, was in welches Feld gehört.
# =============================================================================

class Wohnung(BaseModel):
    """Angaben zur Wohnung."""
    adresse: str | None = Field(None, description="Adresse bzw. Lage (Straße, PLZ, Bezirk)")
    objektart: str | None = Field(None, description="Art: Wohnung, Haus, Apartment, WG-Zimmer ...")
    wohnflaeche_qm: float | None = Field(None, description="Wohnfläche in m²")
    zimmer: float | None = Field(None, description="Anzahl Zimmer (z. B. 2.5)")
    schlafzimmer: int | None = Field(None, description="Anzahl Schlafzimmer")
    etage: str | None = Field(None, description="Etage, z. B. 'EG', '3. OG', 'Dachgeschoss'")
    fahrstuhl: bool | None = Field(None, description="Fahrstuhl vorhanden?")
    verfuegbar_ab: str | None = Field(None, description="Verfügbar ab, Format TT.MM.JJJJ")
    kaltmiete: float | None = Field(None, description="Kaltmiete in Euro pro Monat")
    nebenkosten: float | None = Field(None, description="Nebenkosten in Euro pro Monat")
    gesamtmiete: float | None = Field(None, description="Warm-/Gesamtmiete in Euro pro Monat")
    kaution: float | None = Field(None, description="Kaution in Euro")
    moebliert: bool | None = Field(None, description="Möbliert?")
    balkon_terrasse: bool | None = Field(None, description="Balkon oder Terrasse vorhanden?")
    garten: bool | None = Field(None, description="Garten vorhanden?")
    stellplatz: bool | None = Field(None, description="Parkplatz/Stellplatz vorhanden?")
    keller_abstellraum: bool | None = Field(None, description="Keller oder Abstellraum vorhanden?")
    einbaukueche: bool | None = Field(None, description="Einbauküche vorhanden?")
    badezimmer: int | None = Field(None, description="Anzahl Badezimmer")
    energieausweis: str | None = Field(None, description="Energieausweis/Effizienzklasse, falls relevant")


class Anforderungen(BaseModel):
    """Anforderungen des Vermieters an die Mieter (nur AGG-zulässige Kriterien!)."""
    rauchen_erlaubt: bool | None = Field(None, description="Dürfen Mieter rauchen? False = nur Nichtraucher")
    hunde_erlaubt: bool | None = Field(None, description="Hunde erlaubt?")
    katzen_erlaubt: bool | None = Field(None, description="Katzen erlaubt?")
    kleintiere_erlaubt: bool | None = Field(None, description="Kleintiere (Hamster, Kaninchen ...) erlaubt?")
    haustiere_einzelfall: bool | None = Field(None, description="Nicht erlaubte Haustiere trotzdem im Einzelfall prüfen?")
    max_personen: int | None = Field(None, description="Maximale Anzahl Personen im Haushalt")
    min_nettoeinkommen: float | None = Field(None, description="Mindest-Nettoeinkommen des Haushalts in Euro/Monat")
    einkommen_miete_faktor: float | None = Field(None, description="Einkommen muss mind. X-mal die Gesamtmiete betragen, z. B. 3")
    einzug_fruehestens: str | None = Field(None, description="Frühester akzeptierter Einzug, TT.MM.JJJJ")
    einzug_spaetestens: str | None = Field(None, description="Spätester akzeptierter Einzug, TT.MM.JJJJ")
    mindestmietdauer_monate: int | None = Field(None, description="Mindestmietdauer in Monaten")
    mietschuldenfreiheit_erforderlich: bool | None = Field(None, description="Nachweis vom Vorvermieter (Mietschuldenfreiheit) nötig?")
    schufa_erforderlich: bool | None = Field(None, description="SCHUFA-Auskunft nötig?")
    min_schufa_score: int | None = Field(None, description="Mindest-SCHUFA-Score in Prozent (0-100)")
    beschaeftigungsnachweis_erforderlich: bool | None = Field(None, description="Nachweis über Arbeitgeber/Beschäftigung nötig?")


class ChatAntwort(BaseModel):
    """Antwort der KI in jeder Gesprächsrunde (strukturiertes JSON)."""
    sprache: str = Field(description="Sprache der letzten Nachricht des Vermieters, z. B. 'Deutsch'")
    nachricht: str = Field(description="Text an den Vermieter – in genau dieser Sprache")
    wohnung: Wohnung = Field(description="NUR neu genannte oder geänderte Wohnungsangaben, Rest None")
    anforderungen: Anforderungen = Field(description="NUR neu genannte oder geänderte Anforderungen, Rest None")
    nicht_zutreffend: list[str] = Field(description="Feldnamen, zu denen der Vermieter nichts sagen will/kann")
    alles_erfragt: bool = Field(description="True, wenn nichts Sinnvolles mehr zu fragen ist")
    abbruch_gewuenscht: bool = Field(description="True nur, wenn der Vermieter aufhören möchte")


class Korrektur(BaseModel):
    """Antwort beim Einarbeiten einer Korrektur: jeweils der KOMPLETTE neue Stand."""
    sprache: str
    nachricht: str = Field(description="Kurze Bestätigung der Änderung in der Sprache des Vermieters")
    wohnung: Wohnung
    anforderungen: Anforderungen


class Einschaetzung(BaseModel):
    rang: int
    text: str = Field(description="2-3 Sätze: Stärken und Risiken dieses Kandidaten")


class Einschaetzungen(BaseModel):
    """KI-Einschätzung der besten Kandidaten."""
    einschaetzungen: list[Einschaetzung]


# =============================================================================
#  VERBINDUNG ZU GEMINI (Ausweich-Modelle, Tageslimit, Netzwerkfehler)
# =============================================================================

LETZTES_MODELL = None        # welches Modell zuletzt geantwortet hat (nur zur Info)
TAGESLIMIT_ERREICHT = set()  # Modelle, deren Tageskontingent aufgebraucht ist


def ki_client():
    """Erstellt die Verbindung zu Gemini. Der API-Key kommt aus der .env-Datei."""
    if not os.getenv("GEMINI_API_KEY"):
        raise RuntimeError("Kein GEMINI_API_KEY gefunden. Bitte .env anlegen (Vorlage: .env.beispiel).")
    return genai.Client()


def ki_anfrage(client, inhalte, system, antwortformat, temperatur=0.3):
    """Schickt eine Anfrage an Gemini und gibt die Antwort als Pydantic-Objekt zurück.

    Ist ein Modell überlastet (429/503), nicht vorhanden (404) oder das Netz
    kurz weg, wird sofort das nächste Modell probiert. Erst wenn alle
    scheitern, wird kurz gewartet und die Runde wiederholt.
    """
    global LETZTES_MODELL
    konfiguration = types.GenerateContentConfig(
        system_instruction=system,
        temperature=temperatur,
        response_mime_type="application/json",
        response_schema=antwortformat,
        automatic_function_calling=types.AutomaticFunctionCallingConfig(disable=True),
    )
    letzter_fehler = None
    wartezeit = 0               # von Google empfohlene Wartezeit (Minutenlimit)
    for runde in range(VERSUCHE):
        for modell in MODELLE:
            if modell in TAGESLIMIT_ERREICHT:
                continue                        # heute ohnehin gesperrt -> Zeit sparen
            try:
                antwort = client.models.generate_content(model=modell, contents=inhalte, config=konfiguration)
                ergebnis = antwortformat.model_validate_json(antwort.text)
                LETZTES_MODELL = modell
                return ergebnis
            except errors.APIError as fehler:
                letzter_fehler = fehler
                if fehler.code == 429 and "PerDay" in str(fehler):
                    TAGESLIMIT_ERREICHT.add(modell)
                elif fehler.code == 429:
                    # Minutenlimit: Google nennt, wie lange man warten soll ("retry in 39s").
                    treffer = re.search(r"retry in ([\d.]+)s", str(fehler))
                    if treffer:
                        wartezeit = max(wartezeit, float(treffer.group(1)))
                if fehler.code not in (404, 429, 500, 502, 503, 504):
                    raise                       # z. B. ungültiger Key: sofort melden
            except (ValidationError, ValueError, TypeError) as fehler:
                letzter_fehler = fehler         # kaputtes JSON -> nächster Versuch
            except (httpx.TransportError, ConnectionError) as fehler:
                letzter_fehler = fehler         # Netzwerk weg/Timeout -> nächster Versuch
        # Alle Modelle gescheitert: kurz warten (bei Minutenlimit so lange wie
        # von Google verlangt, höchstens 60 s) und die Runde wiederholen.
        if runde < VERSUCHE - 1:
            pause = min(max(2 ** runde, wartezeit + 1), 60)
            print(f"  (Gemini ist ausgelastet – neuer Versuch in {pause:.0f} s ...)")
            time.sleep(pause)
            wartezeit = 0
    raise RuntimeError(f"Gemini ist gerade nicht erreichbar: {letzter_fehler}")


# =============================================================================
#  SCHRITT 1: GESPRÄCH MIT DEM VERMIETER
# =============================================================================

def felder_beschreibung(modell_klasse, werte):
    """Listet die Felder eines Modells für die KI auf: Name, Beschreibung, Stand."""
    zeilen = []
    for name, info in modell_klasse.model_fields.items():
        stand = werte.get(name)
        zeilen.append(f"- {name}: {info.description} | aktuell: {'offen' if stand is None else stand}")
    return "\n".join(zeilen)


def offene_felder(zustand):
    """Feldnamen, die weder beantwortet noch als 'nicht zutreffend' markiert sind."""
    offen = []
    for bereich, klasse in (("wohnung", Wohnung), ("anforderungen", Anforderungen)):
        for name in klasse.model_fields:
            if zustand[bereich].get(name) is None and name not in zustand["nicht_zutreffend"]:
                offen.append(name)
    return offen


def gespraech_starten():
    """Legt den (JSON-fähigen) Gesprächszustand an."""
    return {
        "wohnung": {},             # bereits bekannte Wohnungsangaben
        "anforderungen": {},       # bereits bekannte Anforderungen
        "nicht_zutreffend": [],    # Felder, zu denen der Vermieter nichts sagt
        "verlauf": [],             # [{"rolle": "user"|"model", "text": "..."}]
        "sprache": "Deutsch",      # zuletzt verwendete Sprache des Vermieters
        "nachgehakt": False,
        "fertig": False,
    }


SYSTEM_GRUNDLAGE = """Du bist der digitale Concierge von KiezMove und hilfst Vermietern in
Berlin, aus sehr vielen Bewerbungen den passenden Mieter zu finden. Dafür
erfragst du Angaben zur Wohnung und die Anforderungen an Mieter.

SPRACHE (wichtig!): Bestimme die Sprache der LETZTEN Nachricht des Vermieters
und schreibe deine Nachricht GENAU in dieser Sprache (Feld "sprache"). Nur beim
Gesprächsbeginn Deutsch. Werte in den Datenfeldern: Datum als TT.MM.JJJJ,
Zahlen als Zahlen (Euro ohne Währungszeichen), Texte auf Deutsch.

FAIRNESS (AGG): Kriterien zu Alter, Geschlecht, ethnischer Herkunft,
Nationalität, Name, Religion, Weltanschauung, Behinderung oder sexueller
Identität sind unzulässig. Nennt der Vermieter so etwas, erkläre freundlich
in einem Satz, dass das nach dem Allgemeinen Gleichbehandlungsgesetz nicht
berücksichtigt werden darf, und übernimm es NICHT. Bezug von Sozialleistungen
(Jobcenter) ist kein Ausschlussgrund – entscheidend ist die gesicherte Miete.

Nichts erfinden oder raten: Nur übernehmen, was der Vermieter gesagt hat oder
eindeutig daraus folgt (Gesamtmiete = Kaltmiete + Nebenkosten darfst du
ausrechnen; "Nichtraucher gesucht" -> rauchen_erlaubt=false).
"""


def gespraech_schritt(client, zustand, nutzer_text=None):
    """Führt EINE Gesprächsrunde: Vermieter-Nachricht rein, Antwort raus.

    nutzer_text=None startet das Gespräch (die KI begrüßt zuerst).
    Rückgabe: (nachricht_an_vermieter, neuer_zustand)
    """
    offen = offene_felder(zustand)
    # Die Anweisung wird jede Runde neu gebaut: So kennt die KI immer den
    # aktuellen Stand und fragt nichts doppelt.
    system = SYSTEM_GRUNDLAGE + f"""
GESPRÄCHSSTIL:
- Freundlich, klar, kurz (höchstens 3 Sätze). Beginne mit einer Begrüßung und
  einer offenen Frage zur Wohnung.
- Frag GEBÜNDELT nach zusammengehörigen Themen (z. B. Größe, Zimmer, Etage und
  Fahrstuhl in einer Frage; Miete, Nebenkosten und Kaution in einer Frage;
  Rauchen und Haustiere in einer Frage) statt jedes Feld einzeln abzufragen.
- Zieh aus jeder Antwort ALLE Angaben heraus. Sagt der Vermieter zu einem Punkt
  "egal", "keine Anforderung" o. ä., trage den passenden Wert ein
  (z. B. rauchen_erlaubt=true, schufa_erforderlich=false). Will oder kann er
  etwas nicht sagen, setze den Feldnamen auf "nicht_zutreffend".
- Mindestmietdauer darfst du notieren, sag aber ehrlich, dass die Bewerbungen
  dazu keine Angaben enthalten und danach nicht gefiltert werden kann.
- Solange Felder offen sind, endet deine Nachricht mit einer Frage. Ist danach
  nichts mehr offen, fasse kurz zusammen und setze alles_erfragt=true.
- abbruch_gewuenscht=true NUR, wenn der Vermieter ausdrücklich aufhören will.

WOHNUNG (Stand):
{felder_beschreibung(Wohnung, zustand["wohnung"])}

ANFORDERUNGEN (Stand):
{felder_beschreibung(Anforderungen, zustand["anforderungen"])}

NOCH OFFEN: {", ".join(offen) or "nichts"}

LETZTE NACHRICHT DES VERMIETERS (antworte in DEREN Sprache!):
{nutzer_text or "(Gesprächsbeginn – antworte auf Deutsch)"}"""

    verlauf = list(zustand["verlauf"])
    verlauf.append({"rolle": "user", "text": nutzer_text or "(Gesprächsbeginn – bitte begrüße mich.)"})
    inhalte = [types.Content(role=e["rolle"], parts=[types.Part(text=e["text"])]) for e in verlauf]

    antwort = ki_anfrage(client, inhalte, system, ChatAntwort, temperatur=0.5)
    zustand = antwort_uebernehmen(zustand, antwort)

    # NACHHAKEN: Meint die KI, sie sei fertig, obwohl noch Felder offen sind,
    # bekommt sie EINMAL die Liste und soll gebündelt danach fragen.
    # Der Vermieter sieht davon nur die neue Nachricht.
    offen = offene_felder(zustand)
    if antwort.alles_erfragt and offen and not antwort.abbruch_gewuenscht and not zustand["nachgehakt"]:
        zustand["nachgehakt"] = True
        intern = ("(Interner Hinweis vom System, nicht erwähnen: Noch offen sind: "
                  + ", ".join(offen) + ". Frag in EINER lockeren Nachricht gebündelt danach.)")
        inhalte += [types.Content(role="model", parts=[types.Part(text=antwort.nachricht)]),
                    types.Content(role="user", parts=[types.Part(text=intern)])]
        antwort = ki_anfrage(client, inhalte, system, ChatAntwort, temperatur=0.5)
        zustand = antwort_uebernehmen(zustand, antwort)
        offen = offene_felder(zustand)

    verlauf.append({"rolle": "model", "text": antwort.nachricht})
    zustand["verlauf"] = verlauf
    zustand["sprache"] = antwort.sprache
    # Fertig, wenn alles geklärt ist, der Vermieter aufhören will oder die KI
    # nach dem Nachhaken erneut sagt, dass nichts Sinnvolles mehr fehlt.
    zustand["fertig"] = (not offen or antwort.abbruch_gewuenscht
                         or (antwort.alles_erfragt and zustand["nachgehakt"]))
    return antwort.nachricht, zustand


def antwort_uebernehmen(zustand, antwort):
    """Überträgt neu erkannte Werte aus der KI-Antwort in den Zustand."""
    zustand = dict(zustand)
    for bereich in ("wohnung", "anforderungen"):
        neu = getattr(antwort, bereich).model_dump(exclude_none=True)
        zustand[bereich] = {**zustand[bereich], **neu}
    # Nur echte Feldnamen als "nicht zutreffend" merken (KI kann sich verschreiben).
    gueltig = set(Wohnung.model_fields) | set(Anforderungen.model_fields)
    zustand["nicht_zutreffend"] = sorted(set(zustand["nicht_zutreffend"])
                                         | {n for n in antwort.nicht_zutreffend if n in gueltig})
    return gesamtmiete_ergaenzen(zustand)


def gesamtmiete_ergaenzen(zustand):
    """Rechnet die Gesamtmiete aus, falls nur Kalt- und Nebenkosten bekannt sind."""
    w = zustand["wohnung"]
    if w.get("gesamtmiete") is None and w.get("kaltmiete") is not None and w.get("nebenkosten") is not None:
        w["gesamtmiete"] = w["kaltmiete"] + w["nebenkosten"]
    return zustand


def korrektur_anwenden(client, zustand, korrektur_text):
    """Arbeitet eine frei formulierte Änderung ein (in jeder Sprache).
    Rückgabe: (bestaetigung_an_vermieter, neuer_zustand)"""
    system = SYSTEM_GRUNDLAGE + f"""
Der Vermieter möchte die erfassten Angaben ändern. Gib den KOMPLETTEN neuen
Stand von "wohnung" und "anforderungen" zurück: unveränderte Werte übernehmen,
genannte Änderungen einarbeiten. Soll eine Anforderung wegfallen, setze sie auf
den Wert "keine Einschränkung" (z. B. rauchen_erlaubt=true) oder None.

AKTUELLE WOHNUNG: {json.dumps(zustand["wohnung"], ensure_ascii=False)}
AKTUELLE ANFORDERUNGEN: {json.dumps(zustand["anforderungen"], ensure_ascii=False)}"""
    antwort = ki_anfrage(client, korrektur_text, system, Korrektur, temperatur=0)
    zustand = dict(zustand)
    zustand["wohnung"] = antwort.wohnung.model_dump(exclude_none=True)
    zustand["anforderungen"] = antwort.anforderungen.model_dump(exclude_none=True)
    zustand["sprache"] = antwort.sprache
    return antwort.nachricht, gesamtmiete_ergaenzen(zustand)


# =============================================================================
#  SCHRITT 2: BEWERBER BEWERTEN (reines Python, keine KI)
# =============================================================================
#  Ablauf pro Bewerber:
#    a) Harte Kriterien prüfen -> bei eindeutigem Verstoß "ausgeschlossen".
#    b) Fehlende Angaben führen NICHT zum Ausschluss, kosten aber Punkte
#       und werden als "unvollständig" markiert.
#    c) Punkte 0–100 nach GEWICHTE, daraus eine Schulnote 1,0–6,0.
#  Name, Alter, Geburtsdatum und Bezirk fließen bewusst NICHT ein (AGG).
# =============================================================================

def datum_lesen(text):
    """Wandelt '01.11.2026' oder '2026-11-01' in ein Datum um (sonst None)."""
    if not text:
        return None
    for format_ in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%y"):
        try:
            return datetime.strptime(str(text).strip(), format_).date()
        except ValueError:
            continue
    return None


def zwischen(wert, von, bis):
    """Begrenzt einen Anteil auf 0..1, linear zwischen von (0) und bis (1)."""
    if bis == von:
        return 1.0
    return max(0.0, min(1.0, (wert - von) / (bis - von)))


def bewerber_bewerten(bewerber, wohnung, anforderungen):
    """Bewertet alle Bewerber und gibt sie sortiert zurück (beste zuerst,
    ausgeschlossene am Ende). Jeder Eintrag ist ein Dict mit dem Original-
    Bewerber sowie punkte, note, status, gruende und fehlend."""
    miete = wohnung.get("gesamtmiete") or wohnung.get("kaltmiete")
    verfuegbar = datum_lesen(wohnung.get("verfuegbar_ab"))
    a = anforderungen
    fruehestens = datum_lesen(a.get("einzug_fruehestens"))
    spaetestens = datum_lesen(a.get("einzug_spaetestens"))

    ergebnisse = []
    for b in bewerber:
        gruende = []    # Gründe für einen Ausschluss
        fehlend = []    # fehlende Angaben/Unterlagen
        hinweise = []   # Auffälligkeiten ohne Ausschluss
        einkommen = b.get("monthly_net_income_eur")
        schufa = b.get("schufa_score_percent")
        schulden = b.get("rent_arrears_document") or {}
        wunsch = datum_lesen(b.get("desired_move_in_date"))
        jobcenter = bool(b.get("rent_paid_by_jobcenter"))

        # --- a) Harte Kriterien --------------------------------------------
        if a.get("rauchen_erlaubt") is False and b.get("smoker"):
            gruende.append("Raucher")

        if b.get("has_pets"):
            tier = b.get("pet_type")
            erlaubt = {"dog": a.get("hunde_erlaubt"), "cat": a.get("katzen_erlaubt"),
                       "small_animal": a.get("kleintiere_erlaubt")}.get(tier)
            namen = {"dog": "Hund", "cat": "Katze", "small_animal": "Kleintier"}
            if erlaubt is False:
                if a.get("haustiere_einzelfall"):
                    hinweise.append(f"{namen.get(tier, 'Haustier')}: Einzelfallprüfung")
                else:
                    gruende.append(f"{namen.get(tier, 'Haustier')} nicht erlaubt")

        if a.get("max_personen") and (b.get("household_size") or 0) > a["max_personen"]:
            gruende.append(f"zu viele Personen ({b['household_size']}, max. {a['max_personen']})")

        if wunsch and fruehestens and wunsch < fruehestens:
            gruende.append("Einzug zu früh")
        if wunsch and spaetestens and wunsch > spaetestens:
            gruende.append("Einzug zu spät")

        # Einkommen: Übernimmt das Jobcenter die Miete, gilt sie als gesichert.
        if einkommen is None and not jobcenter:
            fehlend.append("Einkommen")
        elif einkommen is not None and not jobcenter:
            if a.get("min_nettoeinkommen") and einkommen < a["min_nettoeinkommen"]:
                gruende.append(f"Einkommen unter {a['min_nettoeinkommen']:.0f} €")
            elif a.get("einkommen_miete_faktor") and miete and einkommen < a["einkommen_miete_faktor"] * miete:
                gruende.append(f"Einkommen unter {a['einkommen_miete_faktor']:g}× Miete")

        if schulden.get("status") == "arrears_reported" and a.get("mietschuldenfreiheit_erforderlich") is not False:
            gruende.append("Mietschulden gemeldet")
        if not schulden.get("exists") and a.get("mietschuldenfreiheit_erforderlich"):
            fehlend.append("Mietschuldenfreiheit")

        if a.get("schufa_erforderlich") or a.get("min_schufa_score"):
            if schufa is None or b.get("schufa_record_status") == "no_record":
                fehlend.append("SCHUFA")
            elif a.get("min_schufa_score") and schufa < a["min_schufa_score"]:
                gruende.append(f"SCHUFA unter {a['min_schufa_score']} %")

        if a.get("beschaeftigungsnachweis_erforderlich") and not b.get("employer_name"):
            fehlend.append("Beschäftigungsnachweis")

        # --- c) Punkte ---------------------------------------------------------
        teil = {}
        if jobcenter:
            teil["einkommen"] = 1.0
        elif einkommen is not None and miete:
            teil["einkommen"] = zwischen(einkommen / miete, 1.5, 3.5)   # ab 3,5× volle Punkte
        elif einkommen is not None:
            teil["einkommen"] = zwischen(einkommen, 1000, 3500)        # Miete unbekannt
        else:
            teil["einkommen"] = 0.0

        status_faktor = {"established": 1.0, "thin_file": 0.7}.get(b.get("schufa_record_status"), 0.0)
        teil["schufa"] = zwischen(schufa, 50, 100) * status_faktor if schufa is not None else 0.0

        teil["mietschulden"] = {"no_arrears_confirmed": 1.0, "minor_arrears_resolved": 0.5,
                                "arrears_reported": 0.0}.get(schulden.get("status"), 0.3)

        teil["beschaeftigung"] = {"permanent_contract": 1.0, "retired": 0.9, "fixed_term_contract": 0.7,
                                  "self_employed": 0.6, "student": 0.5, "unemployed": 0.4
                                  }.get(b.get("employment_status"), 0.3)
        if jobcenter:
            teil["beschaeftigung"] = max(teil["beschaeftigung"], 0.8)

        if verfuegbar and wunsch:
            teil["einzug"] = 1.0 - zwischen(abs((wunsch - verfuegbar).days), 14, 90)
        else:
            teil["einzug"] = 1.0

        teil["bewerbung"] = {"formal_detailed": 1.0, "semi_formal": 0.8, "casual_short": 0.5,
                             "careless": 0.2}.get(b.get("application_message_style"), 0.5)

        punkte = round(sum(GEWICHTE[k] * teil[k] for k in GEWICHTE), 1)
        note = round(1 + (100 - punkte) / 20, 1)    # 100 P. = 1,0 ... 0 P. = 6,0

        if gruende:
            status = "ausgeschlossen"
        elif fehlend:
            status = "unvollständig"
        else:
            status = "passt"

        ergebnisse.append({"bewerber": b, "punkte": punkte, "note": note, "status": status,
                           "gruende": gruende, "fehlend": fehlend, "hinweise": hinweise,
                           "einkommen_faktor": round(einkommen / miete, 2) if einkommen and miete else None})

    # Sortierung: erst alle Nicht-Ausgeschlossenen, jeweils nach Punkten absteigend.
    ergebnisse.sort(key=lambda e: (e["status"] == "ausgeschlossen", -e["punkte"]))
    for rang, e in enumerate(ergebnisse, start=1):
        e["rang"] = rang
    return ergebnisse


def statistik(ergebnisse):
    """Kurze Zusammenfassung: wie viele passen, woran scheitern die anderen?"""
    anzahl = {"passt": 0, "unvollständig": 0, "ausgeschlossen": 0}
    gruende = {}
    for e in ergebnisse:
        anzahl[e["status"]] += 1
        for g in e["gruende"]:
            # "Einkommen unter 3× Miete" usw. zu einer Kategorie zusammenfassen
            kategorie = g.split(" (")[0]
            gruende[kategorie] = gruende.get(kategorie, 0) + 1
    return anzahl, dict(sorted(gruende.items(), key=lambda x: -x[1]))


# =============================================================================
#  SCHRITT 3: KI-EINSCHÄTZUNG DER BESTEN KANDIDATEN
# =============================================================================

def top_einschaetzen(client, ergebnisse, wohnung, sprache, anzahl=TOP_ANZAHL):
    """Lässt Gemini die besten Kandidaten kurz einschätzen.

    Datenschutz: Es werden nur bewertungsrelevante Angaben geschickt –
    KEIN Name, keine E-Mail, keine Telefonnummer, kein Geburtsdatum.
    Rückgabe: {rang: text}
    """
    top = [e for e in ergebnisse if e["status"] != "ausgeschlossen"][:anzahl]
    if not top:
        return {}
    daten = [{
        "rang": e["rang"], "punkte": e["punkte"], "status": e["status"],
        "personen": e["bewerber"].get("household_size"),
        "haustier": e["bewerber"].get("pet_type"), "raucher": e["bewerber"].get("smoker"),
        "beschaeftigung": e["bewerber"].get("employment_status"),
        "einkommen_zu_miete": e["einkommen_faktor"],
        "miete_ueber_jobcenter": e["bewerber"].get("rent_paid_by_jobcenter"),
        "schufa_prozent": e["bewerber"].get("schufa_score_percent"),
        "schufa_status": e["bewerber"].get("schufa_record_status"),
        "mietschulden_nachweis": e["bewerber"].get("rent_arrears_document"),
        "wunsch_einzug": e["bewerber"].get("desired_move_in_date"),
        "anschreiben_stil": e["bewerber"].get("application_message_style"),
        "fehlend": e["fehlend"], "hinweise": e["hinweise"],
    } for e in top]
    system = (f"Du bewertest Mietbewerber für einen Vermieter. Schreibe pro Kandidat 2-3 "
              f"sachliche Sätze zu Stärken und Risiken, auf {sprache}. Nur die gegebenen "
              f"Daten verwenden, nichts erfinden, keine Bewertung nach Alter, Herkunft o. ä.")
    anfrage = ("Wohnung: " + json.dumps(wohnung, ensure_ascii=False)
               + "\nKandidaten: " + json.dumps(daten, ensure_ascii=False))
    try:
        antwort = ki_anfrage(client, anfrage, system, Einschaetzungen, temperatur=0.3)
        return {e.rang: e.text for e in antwort.einschaetzungen}
    except RuntimeError:
        return {}   # ohne KI-Text geht es auch – die Rangliste steht trotzdem


# =============================================================================
#  SCHRITT 4: EXCEL-DATEI SPEICHERN
# =============================================================================

TIER_NAMEN = {"dog": "Hund", "cat": "Katze", "small_animal": "Kleintier"}
BESCHAEFTIGUNG_NAMEN = {"permanent_contract": "unbefristet", "fixed_term_contract": "befristet",
                        "self_employed": "selbstständig", "student": "Studium",
                        "unemployed": "arbeitslos", "retired": "Rente"}
SCHULDEN_NAMEN = {"no_arrears_confirmed": "keine Mietschulden", "minor_arrears_resolved": "kleine, beglichen",
                  "arrears_reported": "Mietschulden gemeldet"}
STATUS_FARBEN = {"passt": "C6EFCE", "unvollständig": "FFEB9C", "ausgeschlossen": "FFC7CE"}


def excel_speichern(ergebnisse, wohnung, anforderungen, einschaetzungen, ordner=AUSGABE_ORDNER):
    """Speichert ALLE Bewerber sortiert (beste oben) plus die Kriterien als .xlsx.
    Rückgabe: Pfad der Datei."""
    os.makedirs(ordner, exist_ok=True)
    pfad = os.path.join(ordner, f"bewerber_ranking_{datetime.now():%Y-%m-%d_%H-%M}.xlsx")
    mappe = Workbook()

    # --- Blatt 1: Rangliste ------------------------------------------------
    blatt = mappe.active
    blatt.title = "Rangliste"
    spalten = ["Rang", "Punkte", "Note", "Status", "Gründe / fehlt / Hinweise", "Name", "E-Mail",
               "Telefon", "Personen", "Haustier", "Raucher", "Beschäftigung", "Netto (€)",
               "Einkommen/Miete", "Jobcenter", "SCHUFA %", "SCHUFA-Status", "Vorvermieter",
               "Wunsch-Einzug", "KI-Einschätzung", "Bewerbungs-ID"]
    blatt.append(spalten)
    for e in ergebnisse[:5]:
        b = e["bewerber"]
        details = "; ".join(e["gruende"] + [f"fehlt: {f}" for f in e["fehlend"]] + e["hinweise"])
        schulden = b.get("rent_arrears_document") or {}
        blatt.append([
            e["rang"], e["punkte"], e["note"], e["status"], details,
            f"{b.get('first_name', '')} {b.get('last_name', '')}".strip(),
            b.get("contact_email"), b.get("contact_phone"), b.get("household_size"),
            TIER_NAMEN.get(b.get("pet_type"), "–"), "ja" if b.get("smoker") else "nein",
            BESCHAEFTIGUNG_NAMEN.get(b.get("employment_status"), b.get("employment_status")),
            b.get("monthly_net_income_eur"), e["einkommen_faktor"],
            "ja" if b.get("rent_paid_by_jobcenter") else "nein",
            b.get("schufa_score_percent"), b.get("schufa_record_status"),
            SCHULDEN_NAMEN.get(schulden.get("status"), "kein Nachweis"),
            b.get("desired_move_in_date"), einschaetzungen.get(e["rang"], ""), b.get("applicant_id"),
        ])
        blatt.cell(row=blatt.max_row, column=4).fill = PatternFill(
            "solid", fgColor=STATUS_FARBEN[e["status"]])

    # Kopfzeile fett + fixiert, Filter aktiv, sinnvolle Spaltenbreiten.
    for zelle in blatt[1]:
        zelle.font = Font(bold=True)
    blatt.freeze_panes = "A2"
    blatt.auto_filter.ref = blatt.dimensions
    breiten = [6, 8, 6, 14, 40, 24, 30, 15, 9, 11, 8, 14, 10, 15, 10, 10, 14, 22, 14, 60, 12]
    for nummer, breite in enumerate(breiten, start=1):
        blatt.column_dimensions[blatt.cell(row=1, column=nummer).column_letter].width = breite
    for zeile in blatt.iter_rows(min_row=2, min_col=20, max_col=20):
        zeile[0].alignment = Alignment(wrap_text=True, vertical="top")

    # --- Blatt 2: Kriterien (damit nachvollziehbar bleibt, wie bewertet wurde)
    kriterien = mappe.create_sheet("Kriterien")
    kriterien.append(["WOHNUNG", ""])
    for name, info in Wohnung.model_fields.items():
        kriterien.append([info.description, _anzeigen(wohnung.get(name))])
    kriterien.append([])
    kriterien.append(["ANFORDERUNGEN", ""])
    for name, info in Anforderungen.model_fields.items():
        kriterien.append([info.description, _anzeigen(anforderungen.get(name))])
    kriterien.append([])
    kriterien.append(["GEWICHTUNG (Punkte)", ""])
    for name, gewicht in GEWICHTE.items():
        kriterien.append([name, gewicht])
    for zeile in kriterien.iter_rows():
        if zeile[0].value and str(zeile[0].value).isupper():
            zeile[0].font = Font(bold=True)
    kriterien.column_dimensions["A"].width = 60
    kriterien.column_dimensions["B"].width = 40

    mappe.save(pfad)
    return pfad


def _anzeigen(wert):
    """Werte für die Anzeige aufbereiten (ja/nein statt True/False)."""
    if wert is None:
        return "–"
    if isinstance(wert, bool):
        return "ja" if wert else "nein"
    return wert


# =============================================================================
#  KONSOLEN-ANZEIGE (eine App ersetzt diese Funktionen durch ihre Oberfläche)
# =============================================================================

def zusammenfassung_text(zustand):
    """Übersicht der erfassten Wohnung und Anforderungen."""
    zeilen = ["WOHNUNG:"]
    for name, info in Wohnung.model_fields.items():
        if name in zustand["wohnung"]:
            zeilen.append(f"  {info.description}: {_anzeigen(zustand['wohnung'][name])}")
    zeilen.append("ANFORDERUNGEN:")
    for name, info in Anforderungen.model_fields.items():
        if name in zustand["anforderungen"]:
            zeilen.append(f"  {info.description}: {_anzeigen(zustand['anforderungen'][name])}")
    return "\n".join(zeilen)


def ergebnis_text(ergebnisse, einschaetzungen):
    """Statistik und die besten Kandidaten als lesbarer Text."""
    anzahl, gruende = statistik(ergebnisse)
    zeilen = [f"{len(ergebnisse)} Bewerbungen: {anzahl['passt']} passen, "
              f"{anzahl['unvollständig']} unvollständig, {anzahl['ausgeschlossen']} ausgeschlossen."]
    if gruende:
        zeilen.append("Ausschlussgründe (Mehrfachnennung möglich): "
                      + ", ".join(f"{g} ({n})" for g, n in gruende.items()))
    top = [e for e in ergebnisse if e["status"] != "ausgeschlossen"][:TOP_ANZAHL]
    if not top:
        haupt = next(iter(gruende), None)
        zeilen.append(f"\nKein Bewerber erfüllt alle Kriterien. Am meisten schließt "
                      f"'{haupt}' aus – eine Lockerung dort bringt die meisten Treffer.")
    for e in top:
        b = e["bewerber"]
        zeilen.append(f"\n#{e['rang']}  {b.get('first_name')} {b.get('last_name')}  –  "
                      f"{e['punkte']} Punkte, Note {e['note']} ({e['status']})")
        zeilen.append(f"    {b.get('contact_email')} | {b.get('contact_phone')}")
        if e["fehlend"] or e["hinweise"]:
            zeilen.append("    " + "; ".join([f"fehlt: {f}" for f in e["fehlend"]] + e["hinweise"]))
        if e["rang"] in einschaetzungen:
            zeilen.append(f"    {einschaetzungen[e['rang']]}")
    return "\n".join(zeilen)


def eingabe(text="Du: "):
    """Liest eine Eingabe; bei exit/beenden/quit wird das Programm beendet."""
    antwort = ""
    while not antwort:
        antwort = input(text).strip()
    if antwort.lower() in ENDWOERTER:
        print("Chat beendet.")
        sys.exit(0)
    return antwort


# =============================================================================
#  HAUPTPROGRAMM
# =============================================================================

AUSFALL_HINWEIS = ("\n(Gemini ist gerade nicht erreichbar. Nichts ist verloren – "
                   "bitte die letzte Nachricht in einem Moment noch einmal senden.)")
JA_WOERTER = ("ja", "j", "yes", "y", "ok", "passt", "oui", "si", "sí", "tak", "так", "evet")


def main():
    client = ki_client()
    bewerber = bewerber_laden()
    print(f"{len(bewerber)} Bewerbungen geladen. (Beenden jederzeit mit 'exit')")

    # 1) Gespräch: Wohnung und Anforderungen erfragen.
    #    Schlägt ein KI-Aufruf fehl, bleibt der bisherige Stand erhalten und
    #    der Vermieter sendet seine Nachricht einfach noch einmal.
    zustand = gespraech_starten()
    nachricht, zustand = gespraech_schritt(client, zustand)
    print(f"\nConcierge: {nachricht}")
    for _ in range(MAX_GESPRAECHSRUNDEN):
        if zustand["fertig"]:
            break
        try:
            nachricht, zustand = gespraech_schritt(client, zustand, eingabe())
            print(f"\nConcierge: {nachricht}")
        except RuntimeError:
            print(AUSFALL_HINWEIS)

    while True:
        # 2) Zusammenfassung zeigen – bestätigen oder frei korrigieren.
        while True:
            print("\n" + "=" * 70 + "\n" + zusammenfassung_text(zustand) + "\n" + "=" * 70)
            antwort = eingabe("Passt alles? 'ja' = Bewerber suchen, sonst einfach schreiben, "
                              "was geändert werden soll: ")
            if antwort.lower().strip("!. ") in JA_WOERTER:
                break
            try:
                bestaetigung, zustand = korrektur_anwenden(client, zustand, antwort)
                print(f"\nConcierge: {bestaetigung}")
            except RuntimeError:
                print(AUSFALL_HINWEIS)

        # 3) Bewerten (ohne KI), Einschätzung der Besten (KI, optional), Excel.
        ergebnisse = bewerber_bewerten(bewerber, zustand["wohnung"], zustand["anforderungen"])
        einschaetzungen = top_einschaetzen(client, ergebnisse, zustand["wohnung"], zustand["sprache"])
        print("\n" + ergebnis_text(ergebnisse, einschaetzungen))
        pfad = excel_speichern(ergebnisse, zustand["wohnung"], zustand["anforderungen"], einschaetzungen)
        print(f"\nExcel gespeichert: {pfad}")

        # 4) Nachschärfen und neu suchen – oder beenden.
        while True:
            antwort = eingabe("\nKriterien anpassen und neu suchen? Schreib die Änderung – oder 'exit': ")
            try:
                bestaetigung, zustand = korrektur_anwenden(client, zustand, antwort)
                print(f"\nConcierge: {bestaetigung}")
                break
            except RuntimeError:
                print(AUSFALL_HINWEIS)


if __name__ == "__main__":
    # Sonderzeichen/Emojis dürfen die Konsole nicht zum Absturz bringen.
    sys.stdout.reconfigure(errors="replace")
    try:
        main()
    except (FileNotFoundError, RuntimeError, errors.APIError) as fehler:
        print(f"\nFehler: {fehler}")
        sys.exit(1)
    except (KeyboardInterrupt, EOFError):
        print("\nAbgebrochen.")
        sys.exit(1)
