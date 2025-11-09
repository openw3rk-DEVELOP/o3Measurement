#                            ____     __      _____  ___   _______  ________
#  ___  ___  ___ ___ _    __|_  /____/ /__   /  _/ |/ / | / / __/ |/ /_  __/
# / _ \/ _ \/ -_) _ \ |/|/ //_ </ __/  '_/  _/ //    /| |/ / _//    / / /   
# \___/ .__/\__/_//_/__,__/____/_/ /_/\_\  /___/_/|_/ |___/___/_/|_/ /_/    
#    /_/          - openw3rk INVENT - Vehicle Solutions -
# ****************************************************************************
# o3Measurement 12.5.1 | GERMAN VERSION
# -------------------------------------
#
# *********************************************************
# Copyright (c) 2025 openw3rk INVENT, All rights reserved.
# Licensed under GNU General Public License v3.0 (GPLv3)
# FULL LICENSE SEE: LICENSE.txt
# o3Measurement comes with ABSOLUTELY NO WARRANTY.
# *********************************************************
# Web:
# https://openw3rk.de
# https://o3measurement.openw3rk.de
# ----------------------------------
# For requests and feedback:
# develop@openw3rk.de
# ----------------------------------

import os
import uuid
import zipfile
import tempfile
import sys
import json
import csv
import base64
import shutil
from datetime import datetime, timedelta
from pathlib import Path
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog

o3NAME = "o3Measurement"
o3VERSION = "12.5.1"
o3COPYRIGHT = "openw3rk INVENT - Vehicle Solutions"

# Dateipfade/vars für komponenten/datenerfassung/logos
TEMPLATES_DIR = Path.home() / ".o3measurement" / "templates"
DATA_DIR = Path.home() / ".o3measurement" / "data"
VEHICLES_BASE_DIR = Path.home() / ".o3measurement" / "vehicles"
STORAGE_DIR = Path.home() / ".o3measurement" / "storage"
BACKUP_DIR_SHOW = Path.home() / ".o3measurement" / "backups" # nur als var zur ansicht in Über. Richtige pfad cfg in Backup klasse.
STORAGE_DIR.mkdir(parents=True, exist_ok=True)
TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
DATA_DIR.mkdir(parents=True, exist_ok=True)
VEHICLES_BASE_DIR.mkdir(parents=True, exist_ok=True)

# Logos
LOGO_PATH = Path(__file__).parent / "o3assets" / "o3_logo.png"
ICON_PATH = Path(__file__).parent / "o3assets" / "o3_icon.png"

# -------------------------------------------------------------------
# Masseinheiten ab hier
# einige einheiten sind mehrfach unter versch. kategorien verzeichnet, dies ist so vorgesehen.
# Einheiten sind im script vorgegeben ohne extrene datei.

AVAILABLE_UNITS = {
    "Längenmaße": [
        "", 
        "μm", "µm (Mikrometer)", "mm", "mm Hub", "cm", "m", 
        "km",                     # für  Messstrecken
        "inch", "zoll",           # Reifen-, Felgen-, Gewindemaße etc.
        "ft", "yard",             # Importfahrzeuge (z. B. US)
        "nm",                     # Nanometer, z. B. Oberflächenmessung
        "mil", "thou (0.001 inch)",  # Feinmechanik, Kabelisolierung
        "mm²/mm Länge",           # spezielle Leitungslängen-Angabe
        "m Schlauchlänge",        # Hydraulik-, Kraftstoffleitungen
        "mm Gewindesteigung",     # Gewindemaße
        "mm Kolbenhub",           # Motorteile, Bremszylinder etc.
        "mm Spurweite",           # Achsmaße
        "mm Ventilhub",           # Ventiltrieb
        "mm Zahnriemen",          # Riementrieb
        "mm Achsabstand",         # Fahrzeugrahmen
        "mm Radstand",            # Fahrzeugbasis
        "mm Federweg",            # Fahrwerk
        "mm Bremsscheibendicke",  # Verschleißmaß
        "mm Dichtungstiefe",      # Motordichtungen
        "mm Wandstärke",          # Rohr-/Leitungssysteme
        "mm Karosserielänge",     # Gesamtfahrzeug
        "mm Breite", "mm Höhe"    # Gesamtfahrzeug oder Bauteile
    ],

    "Flächen": [
        "",
        "mm²", "cm²", "dm²", "m²",          # metrische Standardflächen
        "in²", "ft²", "yd²",                # imperial (USA, UK)
        "μm² (Quadratmikrometer)",          # Mikrostrukturen, Beschichtungen
        "mm² Kolbenfläche",                 # Hydraulik- und Bremssysteme
        "cm² Dichtfläche",                  # Motordichtungen
        "m² Lackfläche",                    # Karosserie/Lackierarbeiten
        "m² Kühlfläche",                    # Wärmetauscher, Kühler
        "m² Frontfläche",                   # Aerodynamik, Luftwiderstand
        "m² Aufstandsfläche",               # Reifen, Fahrzeugaufstand
        "m² Bodenfläche",                   # Laderaum, Aufbauten
        "mm² Kabelquerschnitt",             # Elektrik/Leitungen
        "cm² Filterfläche",                 # Luft-/Öl-/Kraftstofffilter
        "m² Glasfläche"                     # Scheiben, Fenster
    ],

    "Volumen": [
        "mm³",          # Kubikmillimeter – z. B. für Einspritzmengen
        "cm³",          # Kubikzentimeter – Standard für Hubraum
        "cc",           # synonym zu cm³
        "dm³",          # Kubikdezimeter (entspricht 1 Liter)
        "m³",           # Kubikmeter – Tankvolumen, Luftdurchsatz
        "µm³",          # Mikrokubikmeter – sehr kleine Messbereiche (z. B. Tropfenanalyse)
        "ml/s",         # Milliliter pro Sekunde – Labor, Einspritzung
        "ml/min",       # Milliliter pro Minute
        "L/s",          # Liter pro Sekunde – z. B. Kühlmittel, Luftmassenstrom
        "L/min",        # Liter pro Minute – Hydraulik, Kraftstoffpumpen
        "L/h",          # Liter pro Stunde – Verbrauch, Klimaanlage, Dieselrücklauf
        "m³/s",         # Kubikmeter pro Sekunde – große Volumenströme (z. B. Windkanal)
        "m³/min",       # Kubikmeter pro Minute
        "m³/h",         # Kubikmeter pro Stunde – Klima, Lüfterleistung
        "dm³/s",        # Dezimeter³ pro Sekunde – selten, aber rechnerisch korrekt
        "cm³/s",        # Kubikzentimeter pro Sekunde – Feinmessung Einspritzsysteme
        "in³",          # cubic inch – z. B. Motorangabe (350 ci)
        "ft³",          # cubic foot – Luftdurchsatz, Kabinenvolumen
        "yd³",          # cubic yard – selten, eher Bau-/Transportvolumen
        "gal",          # gallon – generisch
        "gal (US)",     # US gallon (3.785 L)
        "gal (UK)",     # Imperial gallon (4.546 L)
        "qt",           # quart (¼ gallon)
        "pt",           # pint (⅛ gallon)
        "fl oz",        # fluid ounce – kleine Mengen (z. B. Additive)
        "cfm",          # cubic feet per minute – typ. für Luftdurchsatz (Turbo, Filter)
        "cfh",          # cubic feet per hour – Gasverbrauch / Luftsysteme
        "scfm",         # standard cubic feet per minute – normierter Volumenstrom
        "acfm",         # actual cubic feet per minute – realer Durchsatz
        "gpm",          # gallons per minute – Hydrauliksysteme
        "gph",          # gallons per hour – Verbrauchsangabe
        "Nm³/h",        # Normkubikmeter pro Stunde (auf Normdruck/Temp bezogen)
        "Nm³/s",        # Normkubikmeter pro Sekunde
        "Sm³/h",        # Standardkubikmeter pro Stunde (z. B. Gasanalytik)
        "ACFM",         # Actual Cubic Feet per Minute (bei realen Bedingungen)
        "cc/min",       # Einspritzmengen (z. B. Injektorfluss)
        "cc/stroke",    # Einspritzvolumen pro Hub
        "cm³/rev",      # Fördervolumen pro Umdrehung (Hydraulikpumpen)
        "L/rev",        # Fördervolumen großer Pumpen/Kompressoren
        "m³/rev",       # große Gebläse
        "L/100 km",     # Verbrauchskennzahl (Kraftstoffverbrauch)
        "ml/Hub",       # Mikroangabe Einspritzsysteme (Hubvolumen)
        "cc/Hub",       # alternative Schreibweise
        "cm³/Hub",      # wie oben auch
        "µl",           # Mikroliter
        "nl",           # Nanoliter
        "pl",           # Pikoliter
        "ml/Imp",       # Volumen pro Impuls (bei Sensoren)
        "% Volumen",    # Volumenanteil (z. B. bei Gasanalysen)
        "Vol%",         # Volumenprozent (Abgas, Kraftstoff, Luftfeuchte)
        "rel. L",       # relatives Volumen
        "abs. L"        # absolutes Volumen
    ],

    "Masse": [
        "Megatonne (Mt)",
        "mg", "g", "kg", "t",                # metrisch
        "oz", "lb", "st", "ton (US)",        # imperial
        "slug",                              # US-Einheit (seltener, Dynamik)
        "g/s", "g/min", "g/h", "kg/s", "kg/h", 
        "mg/Hub", "mg/Takt",                 # Einspritzsysteme
        "g/kWh", "g/PS·h",                   # spezifischer Verbrauch
        "kg/100km",                          # CO₂ / Verbrauchswerte
        "kg/m³", "g/cm³", "kg/L", "g/L",     # Flüssigkeiten, Kraftstoffe, Öle
        "g/mm³",                             # Labor-/Materialdaten
        "kg/dm³",                            # technische Angabe (z. B. Kühlmittel)
        "kg·m²", "g·cm²", "kg·mm²",
        "kg Achslast", "t Gesamtgewicht",    # Fahrwerk / Fahrzeugdaten
        "kg Nutzlast", "kg Fahrzeuggewicht", # Fahrzeugmassen
        "kg Schwungmasse",                   # Motor, Kupplung, Antriebsstrang
        "kg Gegengewicht"                    # Nutzfahrzeuge, Kräne, Baumaschinen
    ],

    "Kräfte": [
        "N",            # Newton
        "mN",           # Millinewton
        "kN",           # Kilonewton
        "MN",           # Meganewton (z. B. Crashsimulation)
        "N/m",          # Federsteifigkeit
        "N/mm",         # Federkonstante
        "N·s/m",        # Dämpfungskoeffizient
        "N·s/mm",       # Dämpfung auf kleiner Skala
        "N/m²",         # Druck bzw. Flächenlast
        "N/mm²",        # Spannung (Materialbelastung)
        "Nm",           # Newtonmeter (Drehmoment)
        "N·m",          # alternative Schreibweise
        "kNm",          # Kilonewtonmeter
        "Nm/s",         # Drehmomentänderung über Zeit
        "Nm·s",         # Drehimpuls (Trägheitseinheit)
        "Nm/°",         # Torsionssteifigkeit
        "Nm/rad",       # Torsionssteifigkeit pro Bogenmaß
        "N·s",          # Impuls (Kraft * Zeit)
        "kg·m/s",       # alternative Impulseinheit
        "kg·m²/s²",     # Energieeinheit, äquivalent zu Joule
        "kgf",          # Kilopond (technische Gewichtskraft)
        "kp",           # ältere Schreibweise von Kilopond
        "gf",           # Grammkraft
        "tf",           # Tonnenkraft
        "lbf",          # Pound-force (US)
        "ozf",          # Ounce-force (selten)
        "kgf·m",        # technisches Drehmoment
        "kgf·cm",       # Drehmoment in Kleinmechanik
        "lbf·ft",       # Pound-foot torque (US)
        "lbf·in",       # Pound-inch torque (US)
        "dyn",          # Dyne (cgs-System)
        "kp/cm²",       # spezifische Kraft (veraltet, Prüfstände)
        "N/l",          # Kraft pro Längeneinheit (z. B. Reifenprofil)
        "N/°",          # Lenkkraftgradient
        "N/g",          # Normierte Kraft auf Gewichtseinheit
        "N/(m/s²)",     # Trägheitskraft pro Beschleunigungseinheit
        "N/%",          # z. B. pro Steigungsprozent (Fahrleistungsdiagramm)
        "Downforce (N)",# Abtriebskraft
        "Lift (N)",     # Auftriebskraft
        "Drag (N)",     # Luftwiderstandskraft
        "dN/dt",        # Kraftänderungsrate
        "N/s",          # Kraftanstieg pro Zeit
        "N/°C",         # Kraftänderung pro Temperatur (z. B. bei Elastomer-Prüfung)
        "N/bar",        # Kraft pro Druckeinheit (z. B. Aktuatoren)
        "kg·m²",        # Massenträgheitsmoment
        "g·cm²",        # kleinere Skala (z. B. Elektromotor)
        "lb·ft²",       # US-Einheit für Trägheitsmoment
        "% Kraft",      # Prozentualer Kraftanteil (z. B. pro Achse)
        "rel. N",       # relative Kraft
        "abs. N"        # absolute Kraft
    ],

    "Druck": [
        # SI-Einheiten
        "Pa",          # Pascal
        "kPa",         # Kilopascal
        "MPa",         # Megapascal
        "GPa",         # Gigapascal (selten, z. B. Werkstofftests)

        "Bar",         # Standard in KFZ-Technik
        "mbar",        # Millibar, z. B. für Ladedrucksensoren
        "µbar",        # Mikrobar, extrem kleiner Druck (Labor)
        "hPa",         # Hektopascal, entspricht 1 mbar

        "psi",         # Pounds per square inch (allgemein)
        "psi(g)",      # Gauge pressure (Überdruck)
        "psi(a)",      # Absolute pressure
        "psf",         # Pounds per square foot
        "inHg",        # Inches of mercury – Saugrohrunterdruck
        "mmHg",        # Millimeter Quecksilbersäule
        "Torr",        # Alternative zur mmHg, 1 Torr ≈ 133,3 Pa
        "atm",         # Atmosphäre (Standarddruck)
        "kg/cm²",      # Kilogramm pro Quadratzentimeter – Brems/Hydraulik
        "mmH₂O",       # Millimeter Wassersäule – Feindruckmessung
        "inH₂O",       # Inches of water – bei Filter- und Saugdrucksensoren
        "kp/cm²",      # Kilopond pro Quadratzentimeter (veraltet, noch in alten Normen)

        "Pa/s",        # Druckänderungsrate – z. B. in Prüfständen
        "Bar/s",       # Druckanstieg pro Sekunde
        "kPa/min",     # für Langzeitmessungen / Prüfzyklen
        "psi/s",       # z. B. bei Turboladerregelung
        "mmHg/s",      # z. B. bei Bremsvakuumtest

        "% Druck",     # Prozent vom Nenndruck (z. B. Ladedruckregelung)
        "rel. Bar",    # relativer Druck
        "abs. Bar"     # absoluter Druck
    ],

    "Temperatur": [
        "°C", "K", "°F", "°R", "°C/s", "K/min", "°C/min"
    ],
    "Zeit": [
        "h", "s", "min", "ms", "µs", "μs (Mikrosekunde)", "ns", "Sektorenzeit (s)", "Frames/s"
    ],
    "Drehzahl/Geschwindigkeit": [
        "U/min", "RPM", "rad/s", "rad", "km/h", "m/s", "m/min", "ft/s", "mph", "°/s", "°/ms", "°/s (Gierrate)"
    ],
    "Leistung/Energie": [
        "kW", "MW", "PS", "W", "kWh", "Wh", "kj", "J", "J/kg", "W/m²K", "BTU/h", "kcal/h", "Nm/s"
    ],
    "Elektrik": [
        "V", "V DC", "V AC", "~", "A", "mA", "µA", "mV", "kV", "mΩ", "Ω", "kΩ", "MΩ", "W", "VA", "VAR",
        "Ah", "Wh", "mV/ms", "lm", "lux", "C", "F", "pF", "nF", "µF", "H", "dBV", "dBu", "Ruhestrom"
    ],
    "Frequenz": [
        "Hz", "kHz", "MHz", "GHz", "1/s"
    ],
    "Daten/Digital": [
        "Bit", 
        "Byte", 
        "kB", 
        "MB", 
        "GB", 
        "TB", 
        "KiB", 
        "MiB", 
        "GiB", 
        "TiB",
        "bit/s", 
        "kbit/s", 
        "Mbit/s", 
        "Gbit/s", 
        "Tbit/s",
        "B/s", 
        "kB/s", 
        "MB/s", 
        "GB/s", 
        "TB/s",
        "KiB/s", 
        "MiB/s", 
        "GiB/s", 
        "TiB/s",
        "Baud",
        "dB", 
        "dBm", 
        "dBµV", 
        "dBi",
        "fps", 
        "Hz", 
        "kHz", 
        "MHz", 
        "GHz",
        "ms", 
        "µs", 
        "ns",
        "%", 
        "% PWM", 
        "% duty cycle", 
        "PWM Hz",
        "°KW", 
        "°CA",
        "bit/s²", 
        "B/s²",
        "CAN ID", 
        "J1939 PG", 
        "J1939 SP", 
        "OBD-II PID",
        "LIN", 
        "MOST", 
        "FlexRay", 
        "SENT", 
        "PSI5",
        "CRC", 
        "Checksum", 
        "Parity bit",
        "ISO-TP", 
        "UDS", 
        "KWP2000",
        "RPM", 
        "RPM/s",
        "V", 
        "mV", 
        "A", 
        "mA", 
        "Ω", 
        "kΩ", 
        "MΩ",
        "°C", 
        "K", 
        "°F",
        "N",  
        "Ah", 
        "Wh", 
        "kWh"
    ],

    "Verbrauch/Emissionen": [
        "L/100 km",
        "km/L",
        "mpg (US)",
        "mpg (UK)",
        "L/h",
        "mL/h",
        "L/min",
        "kg/h",
        "g/h",
        "mg/h",
        
        "g/km",
        "mg/km",
        "g/mi",
        "g/kWh",
        "g/MJ",
        "g/kg",
        "g/L",
        "g/m³",
        
        "g/s",
        "mg/s",
        "µg/s",
        "g/min",
        "mg/min",
        
        "mg/Hub",
        "mg/str",
        "µg/Hub",
        "µg/str",
        "g/Zyl",
        "mg/Zyl",
        
        "ppm",
        "ppb",
        "ppt",
        "vol%",
        "mol%",
        "mol/m³",
        "mg/m³",
        "µg/m³",
        "ng/m³",
        
        "AFR",
        "λ (Lambda)",
        "λ-Sonde V",
        "λ-Sonde mV",
        "Strecke λ",
        "Fett/Lean",
        
        "CO g/km",
        "CO2 g/km",
        "HC g/km",
        "NOx g/km",
        "PM g/km",
        "PN #/km",
        
        "CO vol%",
        "CO2 vol%",
        "O2 vol%",
        "HC ppm",
        "NOx ppm",
        "SO2 ppm",
        
        "NH3 ppm",
        "N2O ppm",
        "H2 vol%",
        "CH4 vol%",
        "H2O vol%",
        
        "Opazität %",
        "Trübung FSN",
        "Rauchzahl BSU",
        "Rauchzahl HSU",
        "Partikel mg/m³",
        
        "ECE R49",
        "WHSC",
        "WHTC",
        "NEDC",
        "WLTP",
        "FTP-75",
        "JC08",
        
        "Euro 1-6",
        "Euro 1",
        "Euro 2",
        "Euro 3",
        "Euro 4",
        "Euro 5",
        "Euro 6",
        "EPA Tier",
        "CARB",
        "SULEV",
        "LEV",
        "ZEV",
        
        "AdBlue L/100 km",
        "DEF mL/h",
        "Harnstoff %",
        "NOx-Konverter %",
        "SCR Effizienz %",
        "DPF Differenzdruck mbar",
        "GPF Differenzdruck mbar",
        
        "BSFC g/kWh",
        "ISFC g/kWh",
        "TSFC g/kNh",
        "PSFC g/PSh",
        
        "Kraftstoff CZ",
        "Kraftstoff FAME %",
        "Biodiesel %",
        "Ethanol %",
        "MTBE %",
        "Oktanzahl ROZ",
        "Oktanzahl MOZ",
        "Cetanzahl CZ",
        "Cetanzahl ICN",
        
        "Evap. g/test",
        "Evap. g/h",
        "SHED Test",
        "ORVR Effizienz %",
        
        "CO2 Äquivalent g/km",
        "Treibhauspotential GWP",
        "CO2e g/km",
        "CH4e g/km",
        "N2Oe g/km"
    ],
    "Materialeigenschaften": [
        "kg/m³", "g/cm³", "W/m²K", "μ (Reibkoeffizient)", "Shore A", "Shore D", "Pa·s", "MPa", "GPa", "HV", "HB", "HRc", "°Shore"
    ],
    "Prozent/Winkel": [
        "%", "‰", "%-PWM", "% Slip", "%/s", "% Load", "°", "' (Bogenminute)", "°C/s", "K/min", "rad", "°KW", "deg BTDC"
    ],
    "Sonstige": [
        "Downforce (N)",
        "Downforce (kg)",
        "Downforce (lbf)",
        "Lift (N)",
        "Lift (kg)",
        "Aerodynamischer Widerstand (N)",
        "cw-Wert",
        "A-Wert (m²)",
        
        "kg·m²",
        "g·cm²",
        "lb·ft²",
        "N·m·s²",
        "Trägheitsmoment (kg·m²)",
        "Massenträgheitsmoment",
        
        "mV/ms",
        "V/s",
        "V/µs",
        "A/ms",
        "A/µs",
        "Ω/s",
        "Hz/s",
        
        "DIN",
        "SAE",
        "ISO",
        "Viskosität",
        "ECE",
        "JIS",
        "ANSI",
        "ASTM",
        "DIN/ISO",
        "DIN EN ISO",

        "Kunde",
        "Kundennummer",
        "Eigentümer",

        "FIN",
        "ABE",
        "Zulassungsnummer",
        "Baujahr", 
        "Listennummer",

        "Baugruppe",
        "Hersteller",
        "Ph",

        "Kennzeichen",
        "Sonderzulassung",
        "E-KFZ",
        "Historisch (>30 Jahre)",

        "Motorcode",
        "Teilenummer", 
        "OE-Nummer",
        "Klasse", # z.b. gefahrstoff oder kannzeichnungsklasse
        "KFZ",
        "LKW",
        "Kraftrad",
        "E-Kleinstfahrzeug",
        "Öl",
        "Gefahrgut",
        "Gefahrenklasse",
        "§"
        "(NACH §)",
        "oder früher",
        "oder später",
        "ab Baujahr",
        "nach Baujahr",
        "Kundenwunsch",
        "Beanstandung",
        "Nummer",
        "Fall", 
        "n.n.b.", # nicht näher bezeichnet
        "n/a", # nicht angegeben
        "BLD", # unterhalb der nachweisgrenze

        "AW (Arbeitswerte)",
        "IE (Internationale Einheiten)",
        "KE (Katalytische Einheiten)",
        "U (Enzymatische Einheiten)",
        
        "G (GPS)",
        "g (Erdbeschleunigung)",
        "m/s²",
        "ft/s²",
        "Gal (cm/s²)",
        "mG (Milligal)",
        "μg (Mikrog)",
        
        "mm/s",
        "cm/s",
        "m/s",
        "km/s",
        "in/s",
        "ft/s",
        "Kn (Knoten)",
        
        "rpm²",
        "rad/s²",
        "°/s²",
        "U/min²",
        "Drehbeschleunigung",
        "Winkelbeschleunigung",
        
        "A/m",
        "kA/m",
        "Oersted (Oe)",
        "Gauss (G)",
        "Tesla (T)",
        "mT",
        "μT",
        "nT",
        
        "Shore A",
        "Shore D",
        "IRHD",
        "Brinell (HB)",
        "Rockwell (HRC)",
        "Vickers (HV)",
        "Knoop (HK)",
        
        "N·m/°",
        "Nm/rad",
        "lb·ft/°",
        "Torsionssteifigkeit",
        "Biegesteifigkeit (N/m)",
        
        "PSI·in",
        "J/m",
        "ft·lb/in",
        "Kerbzähigkeit",
        "Bruchzähigkeit",
        
        "SAE J",
        "ISO/TS",
        "DIN SPEC",
        "VDA",
        "OEM SPEC",
        
        "R-value",
        "U-value (W/m²K)",
        "K-value",
        "λ-value (W/mK)",
        "Wärmeleitfähigkeit",
        
        "SHU (Scoville)",
        "ASTM D",
        "IP (Ingenieurpraxis)",
        "NORSOK",
        "DNV GL",
        
        "Munsell",
        "RAL",
        "Pantone",
        "NCS",
        "RGB",
        "HEX",
        "CMYK",
        
        "Bé (°Bé)",
        "Baumé",
        "Twaddell (°Tw)",
        "Brix (°Bx)",
        "Plato (°P)",
        "API Gravity (°API)",
        
        "FOG (Fiber Optic Grade)",
        "dB/km",
        "attenuation",
        "NA (Numerical Aperture)",
        
        "IP Code (IP67)",
        "IK Code (IK08)",
        "NEMA",
        "ATEX",
        "IECEx",
        
        "MIL-STD",
        "DEF STAN",
        "RTCA DO",
        "EUROCAE",
        
        "FAR",
        "CS",
        "EASA",
        "FAA",
        "CAA",
        
        " scleroscope",
        
        "Mache (ME)",
        "Curie (Ci)",
        "Rutherford (Rd)",
        "Einstein (E)",
        
        "Darcy (D)",
        "millidarcy (mD)",
        "Permeabilität",
        "Porosität (%)",
        
        "Strouhal Zahl",
        "Reynolds Zahl",
        "Mach Zahl",
        "Froude Zahl",
        "Weber Zahl",
        "Euler Zahl",
        "Knudsen Zahl",
        
        "Prandtl Zahl",
        "Nusselt Zahl",
        "Grashof Zahl",
        "Rayleigh Zahl",
        "Sherwood Zahl",
        "Schmidt Zahl",
        "Lewis Zahl",
        
        "Stokes (St)",
        "Centistokes (cSt)",
        "m²/s",
        "Kinematische Viskosität",
        "Dynamische Viskosität (Pa·s)",
        
        "Poise (P)",
        "Centipoise (cP)",
        "Rhe",
        "Saybolt Universal Seconds (SUS)",
        "Redwood Seconds",
        "Engler Grad (°E)",
        
        "Mercalli Skala",
        "Richter Skala",
        "Momente Magnitude",
        "EMS-98",
        
        "Sone",
        "Phon",
        "Mel",
        "Bark",
        "ERB",
        
        "Standard Atmosphere",
        "ISA",
        "Geopotential Meter",
        
        "FADE",
        "MIL-L",
        "MIL-PRF",
        "MIL-DTL",
        "MIL-HDBK",
        
        "AN",
        "MS",
        "NAS",
        "AS",
        "BAC",
        "BMS",
        "ESPEC",
        
        "GMW",
        "GMN",
        "GMI",
        "GME",
        "TL",
        "PV",
        "LV",
        
        "JASO",
        "JIS D",
        "JIS B",
        "JIS C",
        "JIS G",
        "JIS Z",
        
        "ADR",
        "RID",
        "ADNR",
        "IMDG",
        "IATA DGR",
        
        "ATA",
        "ETRTO",
        "DOT",
        "ECE-R30",
        "ECE-R54",
        "ECE-R75",
        "ECE-R117",
        
        "INMETRO",
        "GCC",
        "GSO",
        "Summe",
        "ARAI",
        "AIS",
        "CMVR"
    ]
}

# Masseinheiten ende
# -------------------------------------------------------------------

# Flache liste
ALL_UNITS = []
for category_units in AVAILABLE_UNITS.values():
    ALL_UNITS.extend(category_units)

# Service intervalle mit standards
SERVICE_TYPES = {
    "Ölwechsel": {"default_interval": 15000, "unit": "km"},
    "Inspektion": {"default_interval": 30000, "unit": "km"},
    "Luftfilter": {"default_interval": 60000, "unit": "km"},
    "Kraftstofffilter": {"default_interval": 120000, "unit": "km"},
    "Bremsflüssigkeit": {"default_interval": 2, "unit": "Jahre"},
    "Bremsbeläge": {"default_interval": 60000, "unit": "km"},
    "Zahnriemen": {"default_interval": 120000, "unit": "km"},
    "Kühlflüssigkeit": {"default_interval": 4, "unit": "Jahre"},
    "Getriebeöl": {"default_interval": 100000, "unit": "km"},
    "Achsmanschetten": {"default_interval": 100000, "unit": "km"}
}


class SavedTable:
    def __init__(self, name: str = "", description: str = "", template_data: dict = None, 
                 table_data: list = None, vehicle_name: str = "", created_date: str = ""):
        self.name = name
        self.description = description
        self.template_data = template_data if template_data is not None else {}
        self.table_data = table_data if table_data is not None else []
        self.vehicle_name = vehicle_name
        self.created_date = created_date or datetime.now().strftime("%Y-%m-%d %H:%M")
        self.modified_date = datetime.now().strftime("%Y-%m-%d %H:%M")
        self.table_id = str(uuid.uuid4())[:8]  # Kurze eindeutige ID

    def to_dict(self):
        return {
            "name": self.name,
            "description": self.description,
            "template_data": self.template_data,
            "table_data": self.table_data,
            "vehicle_name": self.vehicle_name,
            "created_date": self.created_date,
            "modified_date": self.modified_date,
            "table_id": self.table_id
        }

    @staticmethod
    def from_dict(d):
        return SavedTable(
            d.get("name", ""),
            d.get("description", ""),
            d.get("template_data", {}),
            d.get("table_data", []),
            d.get("vehicle_name", ""),
            d.get("created_date", "")
        )


class Vehicle:
    def __init__(self, name: str = "", description: str = "", specifications: dict = None):
        self.name = name
        self.description = description
        self.specifications = specifications if specifications is not None else {
            "fin": "", "hsn": "", "tsn": "", "farbe": "", "baujahr": "", 
            "motor": "", "leistung": "", "getriebe": "", "antrieb": "", "fluessigkeiten": {}
        }
        self.attachments = []
        self.service_history = []
        self.last_service = {"ölwechsel_km": 0, "ölwechsel_datum": ""}
        self.service_intervals = {}
        self.include_attachments_in_export = True
        self.parts = []
        self.defect_reports = []
        self.saved_tables = []
        self.saved_tables_dir = self.get_vehicle_dir() / "saved_tables"

    def to_dict(self):
        return {
            "name": self.name,
            "description": self.description,
            "specifications": self.specifications,
            "attachments": self.attachments,
            "service_history": self.service_history,
            "last_service": self.last_service,
            "service_intervals": self.service_intervals,
            "include_attachments_in_export": self.include_attachments_in_export,
            "parts": self.parts,
            "defect_reports": self.defect_reports,
            # WICHTIG: saved_tables muss hier serialisiert werden
            "saved_tables": [table.to_dict() for table in self.saved_tables]  # Diese Zeile war möglicherweise falsch
        }

    @staticmethod
    def from_dict(d):
        vehicle = Vehicle(
            d.get("name", ""),
            d.get("description", ""),
            d.get("specifications", {})
        )
        vehicle.attachments = d.get("attachments", [])
        vehicle.service_history = d.get("service_history", [])
        vehicle.last_service = d.get("last_service", {})
        vehicle.service_intervals = d.get("service_intervals", {})
        vehicle.include_attachments_in_export = d.get("include_attachments_in_export", True)
        vehicle.parts = d.get("parts", [])
        vehicle.defect_reports = d.get("defect_reports", [])
        
        # WICHTIG: saved_tables korrekt laden
        saved_tables_data = d.get("saved_tables", [])
        vehicle.saved_tables = []
        for table_data in saved_tables_data:
            try:
                table = SavedTable.from_dict(table_data)
                vehicle.saved_tables.append(table)
            except Exception as e:
                print(f"Fehler beim Laden der Tabelle: {e}")
        
        print(f"DEBUG: Vehicle.from_dict: {len(vehicle.saved_tables)} Tabellen geladen")
        
        if not vehicle.service_history and vehicle.service_intervals:
            vehicle._generate_initial_services()
        
        return vehicle

    def _generate_initial_services(self):
        current_km = self.last_service.get('ölwechsel_km', 0)
        current_date = datetime.now()
        
        for service_name, interval in self.service_intervals.items():
            next_km = current_km + interval
            next_date = (current_date + timedelta(days=180)).strftime("%Y-%m-%d")
            
            service_entry = {
                'datum': current_date.strftime("%Y-%m-%d"),
                'service': service_name,
                'kilometerstand': next_km,
                'Anmerkung': f"Geplant - Fällig bei {next_km} km oder {next_date}",
                'status': 'offen',
                'naechster_service_km': next_km,
                'naechster_service_datum': next_date
            }
            self.service_history.append(service_entry)
# pfad init

    def get_vehicle_dir(self):
        return VEHICLES_BASE_DIR / self.name

    def get_scans_dir(self):
        return self.get_vehicle_dir() / "scans"

    def get_saved_tables_dir(self):
        return self.get_vehicle_dir() / "saved_tables"

    def ensure_directories(self):
        self.get_vehicle_dir().mkdir(parents=True, exist_ok=True)
        self.get_scans_dir().mkdir(parents=True, exist_ok=True)
        self.get_saved_tables_dir().mkdir(parents=True, exist_ok=True)

    def save_table(self, table: SavedTable):
        table.vehicle_name = self.name
        table.modified_date = datetime.now().strftime("%Y-%m-%d %H:%M")
        
        # Prüfen ob Name bereits existiert
        for existing_table in self.saved_tables:
            if existing_table.name == table.name and existing_table.table_id != table.table_id:
                reply = QtWidgets.QMessageBox.question(
                    None, 
                    "Tabelle existiert bereits", 
                    f"Eine Tabelle mit dem Namen '{table.name}' existiert bereits.\nMöchten Sie sie überschreiben?",
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
                )
                if reply == QtWidgets.QMessageBox.No:
                    return False
        
        self.saved_tables = [t for t in self.saved_tables if t.table_id != table.table_id]        
        # Neue Tabelle hinzufügen
        self.saved_tables.append(table)
        
        # Als separate Datei speichern
        self._save_table_to_file(table)
        
        # Fahrzeug speichern
        self._save_vehicle_immediately()
        
        return True


    def _save_table_to_file(self, table: SavedTable):
        try:
            # Verzeichnis sicherstellen
            self.ensure_directories()
            
            table_file = self.get_saved_tables_dir() / f"{table.table_id}.json"
            with open(table_file, 'w', encoding='utf-8') as f:
                json.dump(table.to_dict(), f, indent=2, ensure_ascii=False)
        except Exception as e:
            QtWidgets.QMessageBox.warning(None, "Speicherfehler", f"Tabelle konnte nicht gespeichert werden: {e}")

    def load_table(self, table_id: str) -> SavedTable:
        # Zuerst aus der Liste suchen
        for table in self.saved_tables:
            if table.table_id == table_id:
                # Dann versuchen aus Datei zu laden falls vorhanden
                table_file = self.get_saved_tables_dir() / f"{table_id}.json"
                if table_file.exists():
                    try:
                        with open(table_file, 'r', encoding='utf-8') as f:
                            table_data = json.load(f)
                            return SavedTable.from_dict(table_data)
                    except Exception:
                        # Falls Datei nicht geladen werden kann, Rückfall auf Listeneintrag
                        return table
                return table
        return None

    def delete_table(self, table_id: str):
        # Aus Liste entfernen
        initial_count = len(self.saved_tables)
        self.saved_tables = [t for t in self.saved_tables if t.table_id != table_id]
        
        # Datei löschen falls vorhanden
        table_file = self.get_saved_tables_dir() / f"{table_id}.json"
        if table_file.exists():
            try:
                table_file.unlink()
            except Exception:
                pass
        
        # Fahrzeug speichern nur wenn etwas geändert wurde
        if len(self.saved_tables) < initial_count:
            self._save_vehicle_immediately()
            return True
        return False

    def _save_vehicle_immediately(self):
        try:
            self.ensure_directories()
            vehicle_file = self.get_vehicle_dir() / "vehicle.json"
            with open(vehicle_file, "w", encoding="utf-8") as f:
                json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            QtWidgets.QMessageBox.warning(None, "Speicherfehler", f"Fahrzeug konnte nicht gespeichert werden: {e}")
            return False

    def get_table_by_name(self, name: str) -> SavedTable:
        for table in self.saved_tables:
            if table.name == name:
                return self.load_table(table.table_id)
        return None

    def get_existing_table_name(self):
        if self.saved_tables:
            latest_table = max(self.saved_tables, key=lambda x: x.modified_date)
            return latest_table.name
        return None


class Template:
    def __init__(self, name: str, columns: list = None, description: str = ""):
        self.name = name
        self.columns = columns if columns is not None else []
        self.description = description

    def to_dict(self):
        return {"name": self.name, "columns": self.columns, "description": self.description}

    @staticmethod
    def from_dict(d):
        return Template(d.get("name", "untitled"), d.get("columns", []), d.get("description", ""))


class UeberWindow(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Über {o3NAME}")
        self.setMinimumSize(570, 360)  # 600x200 optimal, Über fenster
        layout = QtWidgets.QVBoxLayout(self)

        banner = QtWidgets.QFrame()
        banner_layout = QtWidgets.QHBoxLayout(banner)
        banner_layout.setContentsMargins(6, 6, 6, 6)

        self.logo_lbl = QtWidgets.QLabel()
        self.logo_lbl.setAlignment(QtCore.Qt.AlignCenter)
        self.logo_lbl.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        banner_layout.addWidget(self.logo_lbl, 0)

        # überschr.
        info_layout = QtWidgets.QVBoxLayout()
        title = QtWidgets.QLabel(f"{o3NAME} - Version {o3VERSION}")
        title.setFont(QtGui.QFont(None, 12, QtGui.QFont.Bold))
        info_layout.addWidget(title)
        
        # (c)
        copyright_label = QtWidgets.QLabel("Copyright © openw3rk INVENT, All Rechte vorbehalten.\n")
        info_layout.addWidget(copyright_label)

        # o3.de
        url1_label = QtWidgets.QLabel()
        url1_label.setText('<a href="https://openw3rk.de" style="color: #0088cc; text-decoration: none;">https://openw3rk.de</a>')
        url1_label.setOpenExternalLinks(True)
        url1_label.setTextInteractionFlags(QtCore.Qt.TextBrowserInteraction)
        info_layout.addWidget(url1_label)
        
        # o3m.o3.de
        url2_label = QtWidgets.QLabel()
        url2_label.setText('<a href="https://o3measurement.openw3rk.de" style="color: #0088cc; text-decoration: none;">https://o3measurement.openw3rk.de</a>')
        url2_label.setOpenExternalLinks(True)
        url2_label.setTextInteractionFlags(QtCore.Qt.TextBrowserInteraction)
        info_layout.addWidget(url2_label)
        
        # develop@openw3rk.de mailto
        url_mail = QtWidgets.QLabel()
        url_mail.setText(
            f'<a href="mailto:develop@openw3rk.de?subject=Anfrage/Rückfrage zu o3Measurement {o3VERSION}" '
            'style="color: #0088cc; text-decoration: none;">'
            'develop@openw3rk.de'
            '</a>'
        )
        url_mail.setOpenExternalLinks(True)
        url_mail.setTextInteractionFlags(QtCore.Qt.TextBrowserInteraction)
        info_layout.addWidget(url_mail)

        # LICENSE Info
        spacer_label = QtWidgets.QLabel("\nLIZENZ: GPLv3, Copyright © 2025 openw3rk INVENT\nSource Code:")
        info_layout.addWidget(spacer_label)
        # Source Code URL
        url2_label = QtWidgets.QLabel()
        url2_label.setText('<a href="https://github.com/openw3rk-DEVELOP/o3Measurement" style="color: #0088cc; text-decoration: none;">https://github.com/openw3rk-DEVELOP/o3Measurement</a>')
        url2_label.setOpenExternalLinks(True)
        url2_label.setTextInteractionFlags(QtCore.Qt.TextBrowserInteraction)
        info_layout.addWidget(url2_label)

        # dateipfade
        paths = QtWidgets.QLabel(f"\nPfade:\nVorlagen: {TEMPLATES_DIR}\nData: {DATA_DIR}\nFahrzeuge: {VEHICLES_BASE_DIR}\nLagerbestand: {STORAGE_DIR}\nBackups: {BACKUP_DIR_SHOW}")
        paths.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        info_layout.addWidget(paths)
        info_layout.addStretch()

        banner_layout.addLayout(info_layout, 1)

        layout.addWidget(banner)

        self._load_logo()

        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)
# logo wenn verfügbar
    def _load_logo(self): 
        if LOGO_PATH.exists():
            pix = QtGui.QPixmap(str(LOGO_PATH))
            self.logo_lbl.setPixmap(pix.scaled(240, 85, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
        else:
            self.logo_lbl.setText("[Logo nicht gefunden]")

class ActivityDocumentationDialog(QtWidgets.QDialog):
    def __init__(self, vehicle: Vehicle, activity_type: str = "", parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
        QDialog {
            background-color: #0a0a0a;
            color: #e0e0e0;
        }
        QGroupBox {
            background-color: #0f0f0f;
            color: #e0e0e0;
            border: 1px solid #555555;
            border-radius: 6px;
            margin-top: 10px;
            font-weight: bold;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 8px;
            padding: 0px 4px;
        }
        QLabel {
            color: #e0e0e0;
        }
        QLineEdit, QTextEdit, QComboBox, QSpinBox, QDoubleSpinBox, QDateEdit {
            background-color: #0a0a0a;
            color: #e0e0e0;
            border: 1px solid #555555;
            selection-background-color: #333333;
            selection-color: #ffffff;
        }
        QComboBox QAbstractItemView {
            background-color: #0a0a0a;
            color: #e0e0e0;
            selection-background-color: #333333;
            selection-color: #ffffff;
            border: 1px solid #555555;
        }
        QPushButton {
            background-color: #0f0f0f;
            color: #e0e0e0;
            border: 1px solid #555555;
            padding: 5px;
            border-radius: 4px;
        }
        QPushButton:hover { 
            background-color: #1a1a1a; 
        }
        QPushButton:pressed { 
            background-color: #222222; 
        }
        QTableWidget, QHeaderView::section {
            background-color: #0a0a0a;
            color: #e0e0e0;
            border: 1px solid #555555;
            gridline-color: #555555;
        }
        QHeaderView::section {
            background-color: #0d0d0d;
            color: #e0e0e0;
            border: 1px solid #555555;
        }
        QTableWidget::item:selected {
            background-color: #333333;
            color: #ffffff;
        }
        QScrollBar:vertical, QScrollBar:horizontal {
            background: #0f0f0f;
            width: 12px;
            margin: 0px;
        }
        QScrollBar::handle {
            background: #555555;
            border-radius: 6px;
        }
        QScrollBar::handle:hover {
            background: #777777;
        }
        QScrollBar::add-line, QScrollBar::sub-line {
            background: none;
        }
        """)
# ende css

        self.vehicle = vehicle
        self.activity_type = activity_type
        self.storage_manager = StorageManager()
        self.setWindowTitle(f"Aktivitätsdokumentation - {vehicle.name}")
        
        # Autoscaling für Laptops
        screen = QtWidgets.QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        self.setMinimumSize(1000, 650)
        self.resize(int(screen_geometry.width() * 0.85), int(screen_geometry.height() * 0.75))
        
        layout = QtWidgets.QVBoxLayout(self)
        
        # Haupt-Container mit horizontaler Ausrichtung
        main_container = QtWidgets.QHBoxLayout()
        
        # Linke Spalte - Basis-Informationen
        left_column = QtWidgets.QVBoxLayout()
        left_column.setContentsMargins(0, 0, 10, 0)
        
        # Fahrzeug-Informationen
        vehicle_group = QtWidgets.QGroupBox("Fahrzeug & Aktivität")
        vehicle_layout = QtWidgets.QFormLayout()
        
        self.activity_combo = QtWidgets.QComboBox()
        self.activity_combo.addItems([
            "Allgemeine Wartung", "Reparatur", "Diagnose", 
            "Inspektion", "Ölwechsel", "Bremsenarbeiten",
            "Elektrikarbeiten", "Motorarbeiten", "Getriebearbeiten",
            "Fahrwerksarbeiten", "Karosseriearbeiten", "Sonstiges"
        ])
        if activity_type and isinstance(activity_type, str):
            idx = self.activity_combo.findText(activity_type)
            if idx >= 0:
                self.activity_combo.setCurrentIndex(idx)
        vehicle_layout.addRow("Aktivitätstyp:", self.activity_combo)
        
        self.date_edit = QtWidgets.QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QtCore.QDate.currentDate())
        vehicle_layout.addRow("Datum:", self.date_edit)
        
        self.km_spin = QtWidgets.QSpinBox()
        self.km_spin.setRange(0, 1000000)
        self.km_spin.setValue(self.vehicle.last_service.get('ölwechsel_km', 0))
        vehicle_layout.addRow("KM-Stand:", self.km_spin)
        
        self.mechaniker_edit = QtWidgets.QLineEdit()
        self.mechaniker_edit.setPlaceholderText("Name des Mechanikers")
        vehicle_layout.addRow("Durchgeführt durch:", self.mechaniker_edit)
        
        vehicle_group.setLayout(vehicle_layout)
        left_column.addWidget(vehicle_group)
        
        # Beschreibung
        desc_group = QtWidgets.QGroupBox("Beschreibung")
        desc_layout = QtWidgets.QVBoxLayout()
        
        self.desc_edit = QtWidgets.QTextEdit()
        self.desc_edit.setMaximumHeight(120)
        self.desc_edit.setPlaceholderText("Beschreibung der durchgeführten Arbeiten, Probleme, Ergebnisse...")
        desc_layout.addWidget(self.desc_edit)
        
        desc_group.setLayout(desc_layout)
        left_column.addWidget(desc_group)
        
        # Verwendete Materialien
        material_group = QtWidgets.QGroupBox("Verwendete Materialien")
        material_layout = QtWidgets.QVBoxLayout()
        
        self.material_table = QtWidgets.QTableWidget()
        self.material_table.setColumnCount(4)
        self.material_table.setHorizontalHeaderLabels(["Teilenummer", "Bezeichnung", "Menge", "Einheit"])
        self.material_table.setMaximumHeight(150)
        material_layout.addWidget(self.material_table)
        
        material_controls = QtWidgets.QHBoxLayout()
        add_material_btn = QtWidgets.QPushButton("+ Material aus Lager")
        add_material_btn.clicked.connect(self.add_material_from_storage)
        add_custom_btn = QtWidgets.QPushButton("+ Eigenes Material")
        add_custom_btn.clicked.connect(self.add_custom_material)
        rem_material_btn = QtWidgets.QPushButton("- Material")
        rem_material_btn.clicked.connect(self.remove_material_row)
        
        material_controls.addWidget(add_material_btn)
        material_controls.addWidget(add_custom_btn)
        material_controls.addWidget(rem_material_btn)
        material_controls.addStretch()
        
        material_layout.addLayout(material_controls)
        material_group.setLayout(material_layout)
        left_column.addWidget(material_group)
        
        left_column.addStretch()
        
        # Rechte Spalte
        right_column = QtWidgets.QVBoxLayout()
        right_column.setContentsMargins(10, 0, 0, 0)
        
        # Durchgeführte Arbeiten
        work_group = QtWidgets.QGroupBox("Durchgeführt")
        work_layout = QtWidgets.QVBoxLayout()
        
        self.work_table = QtWidgets.QTableWidget()
        self.work_table.setColumnCount(3)
        self.work_table.setHorizontalHeaderLabels(["Aktivität", "Dauer (Min)", "Bemerkungen"])
        work_layout.addWidget(self.work_table)
        
        work_controls = QtWidgets.QHBoxLayout()
        add_work_btn = QtWidgets.QPushButton("+ Aktivität")
        add_work_btn.clicked.connect(self.add_work_row)
        rem_work_btn = QtWidgets.QPushButton("- Aktivität")
        rem_work_btn.clicked.connect(self.remove_work_row)
        work_controls.addWidget(add_work_btn)
        work_controls.addWidget(rem_work_btn)
        work_controls.addStretch()
        
        work_layout.addLayout(work_controls)
        work_group.setLayout(work_layout)
        right_column.addWidget(work_group)
        
        # Prüfliste & Bemerkungen
        checklist_group = QtWidgets.QGroupBox("Prüfliste & Bemerkungen")
        checklist_layout = QtWidgets.QVBoxLayout()
        
        self.notes_edit = QtWidgets.QTextEdit()
        self.notes_edit.setPlaceholderText("Besondere Bemerkungen, Auffälligkeiten, Empfehlungen...")
        checklist_layout.addWidget(self.notes_edit)
        
        checklist_group.setLayout(checklist_layout)
        right_column.addWidget(checklist_group)
        
        right_column.addStretch()
        
        # Spalten zum Haupt-Container hinzufügen
        main_container.addLayout(left_column, 40)   # 40% Breite
        main_container.addLayout(right_column, 60)  # 60% Breite
        
        layout.addLayout(main_container)
        
        # Buttons
        button_layout = QtWidgets.QHBoxLayout()
        
        self.save_btn = QtWidgets.QPushButton("Aktivität speichern")
        self.save_btn.clicked.connect(self.save_activity)
        self.save_btn.setStyleSheet("background-color: #2e7d32; color: white; font-weight: bold; padding: 8px;")
        
        cancel_btn = QtWidgets.QPushButton("Abbrechen")
        cancel_btn.clicked.connect(self.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(cancel_btn)
        button_layout.addWidget(self.save_btn)
        
        layout.addLayout(button_layout)

    def add_work_row(self):
        row = self.work_table.rowCount()
        self.work_table.insertRow(row)

    def remove_work_row(self):
        current_row = self.work_table.currentRow()
        if current_row >= 0:
            self.work_table.removeRow(current_row)

    def add_material_from_storage(self):
        dlg = MaterialSelectionDialog(self.storage_manager, self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            selected_materials = dlg.get_selected_materials()
            for material in selected_materials:
                self._add_material_to_table(
                    material['teilenummer'],
                    material['bezeichnung'],
                    material['einheit'],
                    1.0
                )

    def add_custom_material(self):
        teilenummer, ok = QtWidgets.QInputDialog.getText(self, "Eigenes Material", "Teilenummer/Name:")
        if ok and teilenummer:
            bezeichnung, ok = QtWidgets.QInputDialog.getText(self, "Material Bezeichnung", "Bezeichnung:")
            if ok:
                self._add_material_to_table(teilenummer, bezeichnung, "Stück", 1.0)

    def _add_material_to_table(self, teilenummer: str, bezeichnung: str, einheit: str, menge: float):
        row = self.material_table.rowCount()
        self.material_table.insertRow(row)
        
        self.material_table.setItem(row, 0, QtWidgets.QTableWidgetItem(teilenummer))
        self.material_table.setItem(row, 1, QtWidgets.QTableWidgetItem(bezeichnung))
        
        # Menge mit SpinBox
        qty_spin = QtWidgets.QDoubleSpinBox()
        qty_spin.setRange(0.1, 1000.0)
        qty_spin.setDecimals(2)
        qty_spin.setValue(menge)
        self.material_table.setCellWidget(row, 2, qty_spin)
        
        self.material_table.setItem(row, 3, QtWidgets.QTableWidgetItem(einheit))

    def remove_material_row(self):
        current_row = self.material_table.currentRow()
        if current_row >= 0:
            self.material_table.removeRow(current_row)

    def save_activity(self):
        if not self.mechaniker_edit.text().strip():
            QtWidgets.QMessageBox.warning(self, "Fehler", "Bitte geben Sie den Mechaniker-Namen ein.")
            return
        
        # Aktivitäts-Daten sammeln
        activity_data = {
            'activity_id': f"{self.date_edit.date().year()}-{len(self.vehicle.service_history) + 1:03d}",
            'activity_type': self.activity_combo.currentText(),
            'datum': self.date_edit.date().toString("yyyy-MM-dd"),
            'fahrzeug': self.vehicle.name,
            'km_stand': self.km_spin.value(),
            'beschreibung': self.desc_edit.toPlainText(),
            'durchgeführte_arbeiten': [],
            'verwendete_materialien': [],
            'bemerkungen': self.notes_edit.toPlainText(),
            'status': 'abgeschlossen',
            'erstellt_am': datetime.now().strftime("%Y-%m-%d %H:%M"),
            'erstellt_durch': self.mechaniker_edit.text().strip()
        }
        
        # Arbeiten sammeln
        for row in range(self.work_table.rowCount()):
            arbeit_item = self.work_table.item(row, 0)
            dauer_item = self.work_table.item(row, 1)
            bemerkung_item = self.work_table.item(row, 2)
            
            if arbeit_item and arbeit_item.text().strip():
                arbeit_data = {
                    'arbeit': arbeit_item.text(),
                    'dauer_minuten': int(dauer_item.text()) if dauer_item and dauer_item.text().strip() else 0,
                    'bemerkungen': bemerkung_item.text() if bemerkung_item else ""
                }
                activity_data['durchgeführte_arbeiten'].append(arbeit_data)
        
        # Materialien sammeln und Lagerbestand anpassen
        for row in range(self.material_table.rowCount()):
            tn_item = self.material_table.item(row, 0)
            bez_item = self.material_table.item(row, 1)
            qty_widget = self.material_table.cellWidget(row, 2)
            einheit_item = self.material_table.item(row, 3)
            
            if tn_item and tn_item.text().strip() and qty_widget:
                menge = qty_widget.value()
                
                material_data = {
                    'teilenummer': tn_item.text(),
                    'bezeichnung': bez_item.text() if bez_item else "",
                    'menge_verbraucht': menge,
                    'einheit': einheit_item.text() if einheit_item else "Stück",
                    'aus_lager_entnommen': self._is_from_storage(tn_item.text())
                }
                
                # Lagerbestand anpassen wenn aus Lager
                if material_data['aus_lager_entnommen']:
                    self._update_stock(tn_item.text(), menge)
                
                activity_data['verwendete_materialien'].append(material_data)
        
        # Aktivität speichern
        self._save_activity(activity_data)
        
        QtWidgets.QMessageBox.information(self, "Aktivität gespeichert", 
                                        f"Aktivität wurde erfolgreich dokumentiert.\n"
                                        f"Aktivitäts-ID: {activity_data['activity_id']}")
        
        self.accept()
    
    def _is_from_storage(self, teilenummer: str) -> bool:
        return any(item.teilenummer == teilenummer for item in self.storage_manager.items)
    
    def _update_stock(self, teilenummer: str, menge: float):
        for i, item in enumerate(self.storage_manager.items):
            if item.teilenummer == teilenummer:
                item.anzahl -= menge
                if item.anzahl < 0:
                    item.anzahl = 0
                item.geaendert_am = datetime.now().strftime("%Y-%m-%d")
                self.storage_manager.update_item(i, item)
                break
    
    def _save_activity(self, activity_data):
        # Verzeichnis erstellen
        activity_dir = self.vehicle.get_vehicle_dir() / "activities" / str(datetime.now().year)
        activity_dir.mkdir(parents=True, exist_ok=True)
        
        # Aktivitäts-Datei speichern
        activity_file = activity_dir / f"{activity_data['activity_id']}_{activity_data['datum']}_{activity_data['activity_type'].lower().replace(' ', '_')}.json"
        with open(activity_file, 'w', encoding='utf-8') as f:
            json.dump(activity_data, f, indent=2, ensure_ascii=False)
        
        if activity_data['km_stand'] > self.vehicle.last_service.get('ölwechsel_km', 0):
            self.vehicle.last_service['ölwechsel_km'] = activity_data['km_stand']
        
        # Fahrzeug speichern
        self._auto_save_vehicle()
    
    def _auto_save_vehicle(self):
        try:
            self.vehicle.ensure_directories()
            vehicle_file = self.vehicle.get_vehicle_dir() / "vehicle.json"
            with open(vehicle_file, "w", encoding="utf-8") as f:
                json.dump(self.vehicle.to_dict(), f, indent=2, ensure_ascii=False)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Auto-Save Fehler", f"Fehler beim Speichern: {e}")

class ServiceHistoryDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Service eintragen")
        self.setMinimumSize(500, 350)
        layout = QtWidgets.QFormLayout(self)
        
        self.service_combo = QtWidgets.QComboBox()
        self.service_combo.addItems(SERVICE_TYPES.keys())
        layout.addRow("Service:", self.service_combo)
        
        self.date_edit = QtWidgets.QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QtCore.QDate.currentDate())
        layout.addRow("Datum:", self.date_edit)
        
        self.km_edit = QtWidgets.QSpinBox()
        self.km_edit.setRange(0, 1000000)
        layout.addRow("Kilometerstand:", self.km_edit)
        
        self.notes_edit = QtWidgets.QTextEdit()
        self.notes_edit.setMaximumHeight(100)
        layout.addRow("Anmerkung:", self.notes_edit)
        
        box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        layout.addRow(box)
    
    def get_service_data(self):
        return {
            "datum": self.date_edit.date().toString("yyyy-MM-dd"),
            "service": self.service_combo.currentText(),
            "kilometerstand": self.km_edit.value(),
            "Anmerkung": self.notes_edit.toPlainText(),
            "status": "offen"
        }

class ServiceDialog(QtWidgets.QDialog):
    def __init__(self, vehicle: Vehicle = None, parent=None):
        super().__init__(parent)
        self.vehicle = vehicle if vehicle else Vehicle()
        self.setWindowTitle("Service-Intervalle verwalten")
        self.setMinimumSize(1000, 700)
        layout = QtWidgets.QVBoxLayout(self)
        
        # Letzter ölwechsel
        oil_group = QtWidgets.QGroupBox("Letzter Ölwechsel")
        oil_layout = QtWidgets.QFormLayout()
        
        self.last_oil_km = QtWidgets.QSpinBox()
        self.last_oil_km.setRange(0, 1000000)
        self.last_oil_km.setValue(self.vehicle.last_service.get("ölwechsel_km", 0))
        oil_layout.addRow("Kilometerstand:", self.last_oil_km)
        
        self.last_oil_date = QtWidgets.QDateEdit()
        self.last_oil_date.setCalendarPopup(True)
        oil_date = self.vehicle.last_service.get("ölwechsel_datum", "")
        if oil_date:
            self.last_oil_date.setDate(QtCore.QDate.fromString(oil_date, "yyyy-MM-dd"))
        else:
            self.last_oil_date.setDate(QtCore.QDate.currentDate())
        oil_layout.addRow("Datum:", self.last_oil_date)
        
        oil_group.setLayout(oil_layout)
        layout.addWidget(oil_group)
        
        # Service intervalle
        intervals_group = QtWidgets.QGroupBox("Service-Intervalle")
        intervals_layout = QtWidgets.QVBoxLayout()
        
        self.intervals_table = QtWidgets.QTableWidget()
        self.intervals_table.setColumnCount(4)
        self.intervals_table.setHorizontalHeaderLabels(["Service", "Intervall", "Einheit", "Aktiv"])
        self._load_intervals_to_table()
        intervals_layout.addWidget(self.intervals_table)
        
        intervals_group.setLayout(intervals_layout)
        layout.addWidget(intervals_group)
        
        # INITIALE SERVICES GENERIEREN wenn keine vorhanden
        if not self.vehicle.service_history and self.vehicle.service_intervals:
            self._generate_initial_services()
        
        # Service historie MIT FILTER
        history_group = QtWidgets.QGroupBox("Service-Historie")
        history_layout = QtWidgets.QVBoxLayout()
        
        # Filter-Leiste - DIESE MUSS DEFINIERET WERDEN!
        filter_layout = QtWidgets.QHBoxLayout()
        filter_layout.addWidget(QtWidgets.QLabel("Suchen:"))
        
        # HIER search_edit DEFINIEREN!
        self.search_edit = QtWidgets.QLineEdit()
        self.search_edit.setPlaceholderText("Service, Anmerkung...")
        self.search_edit.textChanged.connect(self._load_history_to_table)
        filter_layout.addWidget(self.search_edit)
        
        filter_layout.addWidget(QtWidgets.QLabel("Status:"))
        self.status_combo = QtWidgets.QComboBox()
        self.status_combo.addItems(["Alle", "Offen", "Erledigt"])
        self.status_combo.currentTextChanged.connect(self._load_history_to_table)
        filter_layout.addWidget(self.status_combo)
        
        filter_layout.addStretch()
        
        # Statistik
        self.stats_label = QtWidgets.QLabel("Lade...")
        self.stats_label.setStyleSheet("color: #888;")
        filter_layout.addWidget(self.stats_label)
        
        history_layout.addLayout(filter_layout)
        
        # History-Tabelle
        self.history_table = QtWidgets.QTableWidget()
        self.history_table.setColumnCount(6)
        self.history_table.setHorizontalHeaderLabels([
            "Datum", "Service", "Kilometerstand", "Status", "Nächster Service", "Anmerkung"
        ])
        self.history_table.setAlternatingRowColors(False)
        self.history_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.history_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        history_layout.addWidget(self.history_table)
        
        # Buttons
        history_controls = QtWidgets.QHBoxLayout()
        
        add_history_btn = QtWidgets.QPushButton("+ Service eintragen")
        add_history_btn.clicked.connect(self.add_service_history)
        
        rem_history_btn = QtWidgets.QPushButton("- Eintrag entfernen")
        rem_history_btn.clicked.connect(self.remove_service_history)
        
        complete_service_btn = QtWidgets.QPushButton("Service erledigt")
        complete_service_btn.clicked.connect(self.add_service_completion)
        
        history_controls.addWidget(add_history_btn)
        history_controls.addWidget(rem_history_btn)
        history_controls.addWidget(complete_service_btn)
        history_controls.addStretch()
        
        history_layout.addLayout(history_controls)
        history_group.setLayout(history_layout)
        layout.addWidget(history_group)
        
        # Buttons
        box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)
        
        # DANN erst History laden
        self._load_history_to_table()
    
    def _load_intervals_to_table(self):
        self.intervals_table.setRowCount(len(SERVICE_TYPES))
        
        for row, (service_name, service_info) in enumerate(SERVICE_TYPES.items()):
            self.intervals_table.setItem(row, 0, QtWidgets.QTableWidgetItem(service_name))
            
            interval_spin = QtWidgets.QSpinBox()
            interval_spin.setRange(0, 1000000)
            interval_spin.setValue(self.vehicle.service_intervals.get(service_name, service_info["default_interval"]))
            self.intervals_table.setCellWidget(row, 1, interval_spin)
            
            self.intervals_table.setItem(row, 2, QtWidgets.QTableWidgetItem(service_info["unit"]))
            
            active_check = QtWidgets.QCheckBox()
            active_check.setChecked(service_name in self.vehicle.service_intervals)
            self.intervals_table.setCellWidget(row, 3, active_check)
    
    def _load_history_to_table(self):
        if not hasattr(self, 'search_edit') or not hasattr(self, 'status_combo'):
            return
            
        search_text = self.search_edit.text().lower()
        status_filter = self.status_combo.currentText()
        
        filtered_history = []
        for service in self.vehicle.service_history:
            # Text-Suche
            text_match = (not search_text or 
                        search_text in service.get('service', '').lower() or
                        search_text in service.get('Anmerkung', '').lower())
            
            # Status-Filter
            status_match = True
            if status_filter == "Offen":
                status_match = service.get('status', 'offen') == 'offen'
            elif status_filter == "Erledigt":
                status_match = service.get('status', 'offen') == 'erledigt'
            
            if text_match and status_match:
                filtered_history.append(service)
        
        # Tabelle mit 7 Spalten
        self.history_table.setColumnCount(7)
        self.history_table.setHorizontalHeaderLabels([
            "Datum", "Service", "Letzter KM", "Nächster KM", "Status", "Nächster Termin", "Anmerkung"
        ])
        
        self.history_table.setRowCount(len(filtered_history))
        
        for row, service in enumerate(filtered_history):
            self.history_table.setItem(row, 0, QtWidgets.QTableWidgetItem(service.get("datum", "")))
            self.history_table.setItem(row, 1, QtWidgets.QTableWidgetItem(service.get("service", "")))
            
            # Letzter KM-Stand
            last_km = str(service.get('letzter_service_km', ''))
            self.history_table.setItem(row, 2, QtWidgets.QTableWidgetItem(last_km))
            
            # Nächster KM-Stand
            next_km = str(service.get('kilometerstand', ''))
            self.history_table.setItem(row, 3, QtWidgets.QTableWidgetItem(next_km))
            
            # Status mit Farbe
            status = service.get("status", "offen")
            status_item = QtWidgets.QTableWidgetItem(status)
            if status == "erledigt":
                status_item.setBackground(QtGui.QColor('#1e3a1e'))
                status_item.setForeground(QtGui.QColor('#90ee90'))
            else:
                status_item.setBackground(QtGui.QColor('#3a1e1e'))
                status_item.setForeground(QtGui.QColor('#ff6b6b'))
            self.history_table.setItem(row, 4, status_item)
            
            # Nächster Termin
            next_date = service.get('naechster_service_datum', '')
            self.history_table.setItem(row, 5, QtWidgets.QTableWidgetItem(next_date))
            
            self.history_table.setItem(row, 6, QtWidgets.QTableWidgetItem(service.get("Anmerkung", "")))
        
        # Spaltenbreiten anpassen
        self.history_table.setColumnWidth(0, 100)  # Datum
        self.history_table.setColumnWidth(1, 120)  # Service
        self.history_table.setColumnWidth(2, 100)   # Letzter KM
        self.history_table.setColumnWidth(3, 100)   # Nächster KM
        self.history_table.setColumnWidth(4, 70)   # Status
        self.history_table.setColumnWidth(5, 100)  # Nächster Termin
        self.history_table.setColumnWidth(6, 200)  # Anmerkung
        
        # Statistik aktualisieren
        total = len(self.vehicle.service_history)
        filtered = len(filtered_history)
        erledigt = sum(1 for s in self.vehicle.service_history if s.get('status') == 'erledigt')
        offen = total - erledigt
        
        self.stats_label.setText(f"Gesamt: {total} | Erledigt: {erledigt} | Offen: {offen} | Gefiltert: {filtered}")
    
    def add_service_history(self):
        dlg = ServiceHistoryDialog(self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            service_data = dlg.get_service_data()
            self.vehicle.service_history.append(service_data)
            self._load_history_to_table()
    
    def remove_service_history(self):
        current_row = self.history_table.currentRow()
        if current_row >= 0:
            self.vehicle.service_history.pop(current_row)
            self._load_history_to_table()
    
    def add_service_completion(self):
        current_row = self.history_table.currentRow()
        if current_row < 0:
            QtWidgets.QMessageBox.warning(self, "Kein Service", "Bitte wählen Sie einen Service aus der Liste aus.")
            return
        
        # Gefilterte History holen
        filtered_history = self._get_filtered_history()
        if current_row >= len(filtered_history):
            return
            
        service_data = filtered_history[current_row]
        service_name = service_data.get('service', '')
        
        # Nur offene Services können erledigt werden
        if service_data.get('status') == 'erledigt':
            QtWidgets.QMessageBox.warning(self, "Bereits erledigt", "Dieser Service wurde bereits als erledigt markiert.")
            return
        
        # Service-Intervall holen
        interval = self.vehicle.service_intervals.get(service_name, 
                    SERVICE_TYPES.get(service_name, {}).get('default_interval', 15000))
        
        # 1. AKTUELLEN KM-STAND erfragen
        current_km, ok1 = QtWidgets.QInputDialog.getInt(
            self,
            "Service abschließen",
            f"Aktueller KM-Stand für {service_name}:\n(Wann wurde der Service durchgeführt?)",
            self.vehicle.last_service.get('ölwechsel_km', 0),
            0, 
            1000000
        )
        
        if not ok1:
            return
        
        # 2. NÄCHSTEN SERVICE berechnen und bestätigen lassen
        next_km = current_km + interval
        next_date = (datetime.now() + timedelta(days=180)).strftime("%Y-%m-%d")
        
        next_km, ok2 = QtWidgets.QInputDialog.getInt(
            self,
            "Nächsten Service planen",
            f"Nächster {service_name} fällig bei:\n(Berechnet: {next_km} km - können Sie anpassen)",
            next_km,
            current_km + 1,  # Mindestens 1km mehr als aktuell
            1000000
        )
        
        if not ok2:
            return
        
        # 3. Alten Service in der ORIGINAL-History als erledigt markieren
        for i, service in enumerate(self.vehicle.service_history):
            if (service.get('service') == service_name and 
                service.get('kilometerstand') == service_data.get('kilometerstand') and
                service.get('datum') == service_data.get('datum') and
                service.get('status') == 'offen'):
                
                # Status auf erledigt setzen
                self.vehicle.service_history[i]['status'] = 'erledigt'
                self.vehicle.service_history[i]['Anmerkung'] = f"Abgeschlossen bei {current_km} km"
                break
        
        # 4. NEUEN Service-Eintrag für nächstes Intervall anlegen
        new_service = {
            'datum': datetime.now().strftime("%Y-%m-%d"),
            'service': service_name,
            'kilometerstand': next_km,
            'Anmerkung': f"Geplant - Fällig bei {next_km} km oder {next_date}",
            'status': 'offen',
            'naechster_service_km': next_km,
            'naechster_service_datum': next_date,
            'letzter_service_km': current_km  # NEU: Letzter Service-KM
        }
        self.vehicle.service_history.append(new_service)
        
        # 5. KM-Stand aktualisieren
        self.vehicle.last_service['ölwechsel_km'] = current_km
        
        # 6. History aktualisieren
        self._load_history_to_table()
        
        QtWidgets.QMessageBox.information(
            self, 
            "Service erledigt", 
            f"{service_name} wurde als erledigt markiert.\n\n"
            f"Letzter Service: {current_km} km\n"
            f"Nächster Service: {next_km} km\n"
            f"oder am: {next_date}"
        )
    
    def _get_filtered_history(self):
        search_text = self.search_edit.text().lower()
        status_filter = self.status_combo.currentText()
        
        filtered_history = []
        for service in self.vehicle.service_history:
            # Text-Suche
            text_match = (not search_text or 
                         search_text in service.get('service', '').lower() or
                         search_text in service.get('Anmerkung', '').lower())
            
            # Status-Filter
            status_match = True
            if status_filter == "Offen":
                status_match = service.get('status', 'offen') == 'offen'
            elif status_filter == "Erledigt":
                status_match = service.get('status', 'offen') == 'erledigt'
            
            if text_match and status_match:
                filtered_history.append(service)
        
        return filtered_history
    
    def _generate_initial_services(self):
        current_km = self.vehicle.last_service.get('ölwechsel_km', 0)
        current_date = datetime.now()
        
        for service_name, interval in self.vehicle.service_intervals.items():
            next_km = current_km + interval
            next_date = (current_date + timedelta(days=180)).strftime("%Y-%m-%d")
            
            service_entry = {
                'datum': current_date.strftime("%Y-%m-%d"),
                'service': service_name,
                'kilometerstand': next_km,
                'Anmerkung': f"Geplant - Fällig bei {next_km} km oder {next_date}",
                'status': 'offen',
                'naechster_service_km': next_km,
                'naechster_service_datum': next_date
            }
            self.vehicle.service_history.append(service_entry)
        
        # Auto-save
        self._auto_save_vehicle()

    def _auto_save_vehicle(self):
        try:
            self.vehicle.ensure_directories()
            vehicle_file = self.vehicle.get_vehicle_dir() / "vehicle.json"
            with open(vehicle_file, "w", encoding="utf-8") as f:
                json.dump(self.vehicle.to_dict(), f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Auto-Save Fehler: {e}")
    
    def get_service_data(self):
        # Ölwechsel daten
        self.vehicle.last_service = {
            "ölwechsel_km": self.last_oil_km.value(),
            "ölwechsel_datum": self.last_oil_date.date().toString("yyyy-MM-dd")
        }
        
        # Service intervalle
        self.vehicle.service_intervals = {}
        for row in range(self.intervals_table.rowCount()):
            service_name = self.intervals_table.item(row, 0).text()
            interval_spin = self.intervals_table.cellWidget(row, 1)
            active_check = self.intervals_table.cellWidget(row, 3)
            
            if active_check.isChecked():
                self.vehicle.service_intervals[service_name] = interval_spin.value()
        
        return self.vehicle


class VehicleViewDialog(QtWidgets.QDialog):
    def __init__(self, vehicle: Vehicle, parent=None):
        super().__init__(parent)
        self.vehicle = vehicle
        self.setWindowTitle(f"Fahrzeug ansehen - {vehicle.name}")
        
        screen = QtWidgets.QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        self.setMinimumSize(800, 500)  # Minimale Größe für kleine Bildschirme, WICHTIG.
        self.resize(int(screen_geometry.width() * 0.85), int(screen_geometry.height() * 0.8))
        
        layout = QtWidgets.QVBoxLayout(self)       
        scroll_area = QtWidgets.QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        
        # Scrollbereich
        container = QtWidgets.QWidget()
        container_layout = QtWidgets.QVBoxLayout(container)
        container_layout.setContentsMargins(10, 10, 10, 10)
        
        # Name und beschreibung aber kurz
        basic_info_group = QtWidgets.QGroupBox("Allgemeine Informationen")
        basic_info_layout = QtWidgets.QVBoxLayout()
        
        name_layout = QtWidgets.QHBoxLayout()
        name_layout.addWidget(QtWidgets.QLabel("Fahrzeugname/Eigentümer:"))
        self.name_label = QtWidgets.QLabel(self.vehicle.name)
        self.name_label.setStyleSheet("font-weight: bold; font-size: 14px; padding: 6px; background-color: #1b1b20; border-radius: 4px;")
        name_layout.addWidget(self.name_label)
        name_layout.addStretch()
        basic_info_layout.addLayout(name_layout)
        
        desc_layout = QtWidgets.QVBoxLayout()
        desc_layout.addWidget(QtWidgets.QLabel("Beschreibung:"))
        self.desc_label = QtWidgets.QLabel(self.vehicle.description)
        self.desc_label.setWordWrap(True)
        self.desc_label.setStyleSheet("padding: 6px; background-color: #1b1b20; border-radius: 4px; margin-bottom: 5px;")
        self.desc_label.setMaximumHeight(60)  # Begrenzte Höhe
        desc_layout.addWidget(self.desc_label)
        basic_info_layout.addLayout(desc_layout)
        
        basic_info_group.setLayout(basic_info_layout)
        container_layout.addWidget(basic_info_group)

        specs_group = QtWidgets.QGroupBox("Fahrzeugspezifikationen")
        specs_layout = QtWidgets.QHBoxLayout()
        
        # Linke Spalte
        left_specs = QtWidgets.QFormLayout()
        # Rechte Spalte
        right_specs = QtWidgets.QFormLayout()
        
        specs_data = [
            ('FIN:', self.vehicle.specifications.get('fin', '')),
            ('HSN:', self.vehicle.specifications.get('hsn', '')),
            ('TSN:', self.vehicle.specifications.get('tsn', '')),
            ('Farbe:', self.vehicle.specifications.get('farbe', '')),
            ('Baujahr:', self.vehicle.specifications.get('baujahr', '')),
            ('Motor:', self.vehicle.specifications.get('motor', '')),
            ('Leistung:', self.vehicle.specifications.get('leistung', '')),
            ('Getriebe:', self.vehicle.specifications.get('getriebe', '')),
            ('Antrieb:', self.vehicle.specifications.get('antrieb', ''))
        ]
        
        # Alle Daten auf zwei Spalten
        for i, (label, value) in enumerate(specs_data):
            if value:
                value_label = QtWidgets.QLabel(value)
                value_label.setStyleSheet("padding: 4px; background-color: #1b1b20; border-radius: 4px; margin: 2px;")
                value_label.setWordWrap(True)
                
                if i < len(specs_data) // 2 + 1:
                    left_specs.addRow(label, value_label)
                else:
                    right_specs.addRow(label, value_label)
        
        specs_layout.addLayout(left_specs)
        specs_layout.addLayout(right_specs)
        specs_group.setLayout(specs_layout)
        container_layout.addWidget(specs_group)
        
        # Flüs. kompakt
        fluids = self.vehicle.specifications.get('fluessigkeiten', {})
        if fluids:
            fluids_group = QtWidgets.QGroupBox("Flüssigkeiten")
            fluids_layout = QtWidgets.QVBoxLayout()
            
            fluids_table = QtWidgets.QTableWidget()
            fluids_table.setColumnCount(3)
            fluids_table.setHorizontalHeaderLabels(["Flüssigkeit", "Spezifikation", "Menge"])
            fluids_table.setRowCount(len(fluids))
            fluids_table.setMaximumHeight(120)  # Begrenzte Höhe, laptops
            
            for row, (fluid_name, fluid_spec) in enumerate(fluids.items()):
                fluids_table.setItem(row, 0, QtWidgets.QTableWidgetItem(fluid_name))
                fluids_table.setItem(row, 1, QtWidgets.QTableWidgetItem(fluid_spec.get('spezifikation', '')))
                fluids_table.setItem(row, 2, QtWidgets.QTableWidgetItem(fluid_spec.get('menge', '')))
            
            fluids_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
            fluids_table.horizontalHeader().setStretchLastSection(True)
            fluids_table.verticalHeader().setVisible(False)
            fluids_layout.addWidget(fluids_table)
            fluids_group.setLayout(fluids_layout)
            container_layout.addWidget(fluids_group)
        
        # Bauteile
        if self.vehicle.parts:
            parts_group = QtWidgets.QGroupBox("Bauteile und Teilenummern")
            parts_layout = QtWidgets.QVBoxLayout()
            
            parts_table = QtWidgets.QTableWidget()
            parts_table.setColumnCount(3)
            parts_table.setHorizontalHeaderLabels(["Bauteil", "Teilenummer", "Motor Code"])
            parts_table.setRowCount(len(self.vehicle.parts))
            parts_table.setMaximumHeight(120)  
            
            for row, part in enumerate(self.vehicle.parts):
                parts_table.setItem(row, 0, QtWidgets.QTableWidgetItem(part.get('bauteil', '')))
                parts_table.setItem(row, 1, QtWidgets.QTableWidgetItem(part.get('teilenummer', '')))
                parts_table.setItem(row, 2, QtWidgets.QTableWidgetItem(part.get('motor_code', '')))
            
            parts_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
            parts_table.horizontalHeader().setStretchLastSection(True)
            parts_table.verticalHeader().setVisible(False)
            parts_layout.addWidget(parts_table)
            parts_group.setLayout(parts_layout)
            container_layout.addWidget(parts_group)
        
        # Service informationen
        if self.vehicle.service_history or self.vehicle.service_intervals:
            service_group = QtWidgets.QGroupBox("Service-Informationen")
            service_layout = QtWidgets.QVBoxLayout()
            
            if self.vehicle.last_service.get('ölwechsel_km'):
                oil_info = QtWidgets.QLabel(
                    f"Letzter Ölwechsel: {self.vehicle.last_service.get('ölwechsel_km')} km "
                    f"am {self.vehicle.last_service.get('ölwechsel_datum', '')}"
                )
                oil_info.setStyleSheet("padding: 6px; background-color: #1b1b20; border-radius: 4px; margin: 2px;")
                service_layout.addWidget(oil_info)
            
            if self.vehicle.service_intervals:
                intervals_layout = QtWidgets.QVBoxLayout()
                intervals_label = QtWidgets.QLabel("Aktive Service-Intervalle:")
                intervals_label.setStyleSheet("font-weight: bold; margin-top: 5px;")
                intervals_layout.addWidget(intervals_label)
                
                for service, interval in self.vehicle.service_intervals.items():
                    unit = SERVICE_TYPES.get(service, {}).get('unit', 'km')
                    interval_label = QtWidgets.QLabel(f"• {service}: alle {interval} {unit}")
                    interval_label.setStyleSheet("margin-left: 10px;")
                    intervals_layout.addWidget(interval_label)
                
                service_layout.addLayout(intervals_layout)
            
            service_group.setLayout(service_layout)
            container_layout.addWidget(service_group)
        
        if self.vehicle.attachments:
            attachments_group = QtWidgets.QGroupBox("Fahrzeugpapiere & Scans")
            attachments_layout = QtWidgets.QVBoxLayout()
            
            self.attachments_list = QtWidgets.QListWidget()
            self.attachments_list.setMaximumHeight(100) 
            for attachment in self.vehicle.attachments:
                item = QtWidgets.QListWidgetItem(attachment['filename'])
                item.setData(QtCore.Qt.UserRole, attachment)
                self.attachments_list.addItem(item)
            
            self.attachments_list.itemDoubleClicked.connect(self.view_attachment)
            attachments_layout.addWidget(self.attachments_list)
            
            view_btn = QtWidgets.QPushButton("Ausgewählte Datei ansehen")
            view_btn.clicked.connect(self.view_attachment)
            attachments_layout.addWidget(view_btn)
            
            attachments_group.setLayout(attachments_layout)
            container_layout.addWidget(attachments_group)
        
        container_layout.addStretch(1)
        
        scroll_area.setWidget(container)
        layout.addWidget(scroll_area)
        
        # Schließen button
        btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)
    
    def view_attachment(self):
        current_item = self.attachments_list.currentItem()
        if not current_item:
            QtWidgets.QMessageBox.information(self, "Info", "Bitte wählen Sie zuerst eine Datei aus der Liste aus.")
            return
            
        attachment = current_item.data(QtCore.Qt.UserRole)
        
        if attachment['mimetype'] == 'image':
            dlg = ImageViewDialog(attachment, self)
            dlg.exec_()
        elif attachment['mimetype'] == 'application/pdf':
            QtWidgets.QMessageBox.information(self, "PDF", f"PDF-Datei: {attachment['filename']}\n\nFür die PDF-Anzeige wird ein externer Viewer benötigt.")
        else:
            QtWidgets.QMessageBox.information(self, "Datei", f"Datei: {attachment['filename']}\n\nMIME-Type: {attachment['mimetype']}")


class VehicleDialog(QtWidgets.QDialog):
    def __init__(self, vehicle: Vehicle = None, parent=None):
        super().__init__(parent)
        self.vehicle = vehicle if vehicle else Vehicle()
        self.setWindowTitle("Fahrzeug erstellen/bearbeiten")
        
        screen = QtWidgets.QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        self.setMinimumSize(800, 500) # Minimale Größe für kleine Bildschirme und oder laptops
        self.resize(int(screen_geometry.width() * 0.8), int(screen_geometry.height() * 0.7))  
        
        layout = QtWidgets.QVBoxLayout(self)
        
        scroll_area = QtWidgets.QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)

        container = QtWidgets.QWidget()
        container_layout = QtWidgets.QVBoxLayout(container)
        container_layout.setContentsMargins(5, 5, 5, 5)
        
        basic_info_group = QtWidgets.QGroupBox("Allgemeine Informationen")
        basic_info_layout = QtWidgets.QFormLayout()
        
        self.name_edit = QtWidgets.QLineEdit(self.vehicle.name)
        basic_info_layout.addRow('Fahrzeug/Eigentümer:', self.name_edit)
        
        self.desc_edit = QtWidgets.QTextEdit()
        self.desc_edit.setMaximumHeight(60)
        self.desc_edit.setPlainText(self.vehicle.description)
        basic_info_layout.addRow('Beschreibung:', self.desc_edit)
        
        basic_info_group.setLayout(basic_info_layout)
        container_layout.addWidget(basic_info_group)
        
        specs_group = QtWidgets.QGroupBox("Fahrzeugspezifikationen")
        specs_layout = QtWidgets.QHBoxLayout()
        
        # Linke Spalte
        left_specs = QtWidgets.QFormLayout()
        self.fin_edit = QtWidgets.QLineEdit(self.vehicle.specifications.get('fin', ''))
        left_specs.addRow('FIN:', self.fin_edit)
        
        self.hsn_edit = QtWidgets.QLineEdit(self.vehicle.specifications.get('hsn', ''))
        left_specs.addRow('HSN:', self.hsn_edit)
        
        self.tsn_edit = QtWidgets.QLineEdit(self.vehicle.specifications.get('tsn', ''))
        left_specs.addRow('TSN:', self.tsn_edit)
        
        self.color_edit = QtWidgets.QLineEdit(self.vehicle.specifications.get('farbe', ''))
        left_specs.addRow('Farbe:', self.color_edit)
        
        # Rechte Spalte
        right_specs = QtWidgets.QFormLayout()
        self.year_edit = QtWidgets.QLineEdit(self.vehicle.specifications.get('baujahr', ''))
        right_specs.addRow('Baujahr:', self.year_edit)
        
        self.engine_edit = QtWidgets.QLineEdit(self.vehicle.specifications.get('motor', ''))
        right_specs.addRow('Motor:', self.engine_edit)
        
        self.power_edit = QtWidgets.QLineEdit(self.vehicle.specifications.get('leistung', ''))
        right_specs.addRow('Leistung:', self.power_edit)
        
        self.gear_edit = QtWidgets.QLineEdit(self.vehicle.specifications.get('getriebe', ''))
        right_specs.addRow('Getriebe:', self.gear_edit)
        
        self.drive_edit = QtWidgets.QLineEdit(self.vehicle.specifications.get('antrieb', ''))
        right_specs.addRow('Antrieb:', self.drive_edit)
        
        specs_layout.addLayout(left_specs)
        specs_layout.addLayout(right_specs)
        specs_group.setLayout(specs_layout)
        container_layout.addWidget(specs_group)
        
        # Flüssigkeiten
        fluids_group = QtWidgets.QGroupBox("Flüssigkeiten")
        fluids_layout = QtWidgets.QVBoxLayout()
        
        self.fluids_table = QtWidgets.QTableWidget()
        self.fluids_table.setColumnCount(3)
        self.fluids_table.setHorizontalHeaderLabels(["Flüssigkeit", "Spezifikation", "Menge"])
        self.fluids_table.setMaximumHeight(150)  
        self._load_fluids_to_table()
        self.fluids_table.horizontalHeader().setStretchLastSection(True)
        fluids_layout.addWidget(self.fluids_table)
        
        fluid_controls = QtWidgets.QHBoxLayout()
        add_fluid_btn = QtWidgets.QPushButton("+ Flüssigkeit")
        add_fluid_btn.clicked.connect(self.add_fluid_row)
        rem_fluid_btn = QtWidgets.QPushButton("- Flüssigkeit")
        rem_fluid_btn.clicked.connect(self.remove_fluid_row)
        fluid_controls.addWidget(add_fluid_btn)
        fluid_controls.addWidget(rem_fluid_btn)
        fluid_controls.addStretch()
        fluids_layout.addLayout(fluid_controls)
        
        fluids_group.setLayout(fluids_layout)
        container_layout.addWidget(fluids_group)
        
        parts_group = QtWidgets.QGroupBox("Bauteile und Teilenummern")
        parts_layout = QtWidgets.QVBoxLayout()
        
        self.parts_table = QtWidgets.QTableWidget()
        self.parts_table.setColumnCount(3)
        self.parts_table.setHorizontalHeaderLabels(["Bauteil", "Teilenummer", "Motor Code"])
        self.parts_table.setMaximumHeight(150)
        self._load_parts_to_table()
        self.parts_table.horizontalHeader().setStretchLastSection(True)
        parts_layout.addWidget(self.parts_table)
        
        parts_controls = QtWidgets.QHBoxLayout()
        add_part_btn = QtWidgets.QPushButton("+ Bauteil")
        add_part_btn.clicked.connect(self.add_part_row)
        rem_part_btn = QtWidgets.QPushButton("- Bauteil")
        rem_part_btn.clicked.connect(self.remove_part_row)
        parts_controls.addWidget(add_part_btn)
        parts_controls.addWidget(rem_part_btn)
        parts_controls.addStretch()
        parts_layout.addLayout(parts_controls)
        
        parts_group.setLayout(parts_layout)
        container_layout.addWidget(parts_group)

        attachments_group = QtWidgets.QGroupBox("Fahrzeugpapiere & Scans")
        attachments_layout = QtWidgets.QVBoxLayout()
        
        self.attachments_list = QtWidgets.QListWidget()
        self.attachments_list.setMaximumHeight(120)  
        self._load_attachments_to_list()
        attachments_layout.addWidget(self.attachments_list)
        
        attachment_controls = QtWidgets.QHBoxLayout()
        add_attachment_btn = QtWidgets.QPushButton("+ Datei")
        add_attachment_btn.clicked.connect(self.add_attachment)
        rem_attachment_btn = QtWidgets.QPushButton("- Datei")
        rem_attachment_btn.clicked.connect(self.remove_attachment)
        view_attachment_btn = QtWidgets.QPushButton("Ansehen")
        view_attachment_btn.clicked.connect(self.view_attachment)
        attachment_controls.addWidget(add_attachment_btn)
        attachment_controls.addWidget(rem_attachment_btn)
        attachment_controls.addWidget(view_attachment_btn)
        attachment_controls.addStretch()
        
        self.include_attachments_chk = QtWidgets.QCheckBox("Anhänge im Export einbinden")
        self.include_attachments_chk.setChecked(self.vehicle.include_attachments_in_export)
        attachment_controls.addWidget(self.include_attachments_chk)
        
        attachments_layout.addLayout(attachment_controls)
        attachments_group.setLayout(attachments_layout)
        container_layout.addWidget(attachments_group)
        
        container_layout.addStretch(1)
        
        scroll_area.setWidget(container)
        layout.addWidget(scroll_area)
        
        # Buttons unten
        box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)
        
    def _load_fluids_to_table(self):
        fluids = self.vehicle.specifications.get('fluessigkeiten', {})
        self.fluids_table.setRowCount(len(fluids))
        
        for row, (fluid_name, fluid_spec) in enumerate(fluids.items()):
            self.fluids_table.setItem(row, 0, QtWidgets.QTableWidgetItem(fluid_name))
            self.fluids_table.setItem(row, 1, QtWidgets.QTableWidgetItem(fluid_spec.get('spezifikation', '')))
            self.fluids_table.setItem(row, 2, QtWidgets.QTableWidgetItem(fluid_spec.get('menge', '')))
    
    def _load_parts_to_table(self):
        self.parts_table.setRowCount(len(self.vehicle.parts))
        
        for row, part in enumerate(self.vehicle.parts):
            self.parts_table.setItem(row, 0, QtWidgets.QTableWidgetItem(part.get('bauteil', '')))
            self.parts_table.setItem(row, 1, QtWidgets.QTableWidgetItem(part.get('teilenummer', '')))
            self.parts_table.setItem(row, 2, QtWidgets.QTableWidgetItem(part.get('motor_code', '')))
    
    def _load_attachments_to_list(self):
        self.attachments_list.clear()
        for attachment in self.vehicle.attachments:
            item = QtWidgets.QListWidgetItem(attachment['filename'])
            item.setData(QtCore.Qt.UserRole, attachment)
            self.attachments_list.addItem(item)
    
    def add_fluid_row(self):
        row = self.fluids_table.rowCount()
        self.fluids_table.insertRow(row)
    
    def remove_fluid_row(self):
        current_row = self.fluids_table.currentRow()
        if current_row >= 0:
            self.fluids_table.removeRow(current_row)
    
    def add_part_row(self):
        row = self.parts_table.rowCount()
        self.parts_table.insertRow(row)
    
    def remove_part_row(self):
        current_row = self.parts_table.currentRow()
        if current_row >= 0:
            self.parts_table.removeRow(current_row)
    
    def add_attachment(self):
        fname, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Datei auswählen", "", "Alle Dateien (*);;Bilder (*.png *.jpg *.jpeg *.bmp);;PDF (*.pdf)")
        if fname:
            # Erstellen von Fahrzeugverzeichnissen
            self.vehicle.ensure_directories()
            
            # Datei in Fahrzeug-Scans-Verzeichnis kopieren
            scan_filename = Path(fname).name
            scan_path = self.vehicle.get_scans_dir() / scan_filename
            shutil.copy2(fname, scan_path)
            
            with open(fname, 'rb') as f:
                file_data = base64.b64encode(f.read()).decode('utf-8')
            
            attachment = {
                'filename': scan_filename,
                'scan_path': str(scan_path),
                'data': file_data,
                'mimetype': 'application/octet-stream'
            }
            
            if fname.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                attachment['mimetype'] = 'image'
            elif fname.lower().endswith('.pdf'):
                attachment['mimetype'] = 'application/pdf'
            
            self.vehicle.attachments.append(attachment)
            self._load_attachments_to_list()
    
    def remove_attachment(self):
        current_row = self.attachments_list.currentRow()
        if current_row >= 0:
            
            # Datei aus Fahrzeug scansverzeichnis löschen
            attachment = self.vehicle.attachments[current_row]
            if 'scan_path' in attachment:
                scan_path = Path(attachment['scan_path'])
                if scan_path.exists():
                    scan_path.unlink()
            
            self.vehicle.attachments.pop(current_row)
            self._load_attachments_to_list()
    
    def view_attachment(self):
        current_row = self.attachments_list.currentRow()
        if current_row >= 0:
            attachment = self.vehicle.attachments[current_row]
            
            if attachment['mimetype'] == 'image':
                # Bild anzeigen
                dlg = ImageViewDialog(attachment, self)
                dlg.exec_()
            elif attachment['mimetype'] == 'application/pdf':
                QtWidgets.QMessageBox.information(self, "PDF", f"PDF-Datei: {attachment['filename']}\n\nFür die PDF-Anzeige wird ein externer Viewer benötigt.")
            else:
                # Andere dateitypen
                QtWidgets.QMessageBox.information(self, "Datei", f"Datei: {attachment['filename']}\n\nMIME-Type: {attachment['mimetype']}")
    
    def get_vehicle(self):
        vehicle = Vehicle()
        vehicle.name = self.name_edit.text().strip()
        vehicle.description = self.desc_edit.toPlainText().strip()
        
        # Spezifikationen des kfz
        vehicle.specifications = {
            'fin': self.fin_edit.text().strip(),
            'hsn': self.hsn_edit.text().strip(),
            'tsn': self.tsn_edit.text().strip(),
            'farbe': self.color_edit.text().strip(),
            'baujahr': self.year_edit.text().strip(),
            'motor': self.engine_edit.text().strip(),
            'leistung': self.power_edit.text().strip(),
            'getriebe': self.gear_edit.text().strip(),
            'antrieb': self.drive_edit.text().strip(),
            'fluessigkeiten': {}
        }
        
        # Flüssigkeiten
        for row in range(self.fluids_table.rowCount()):
            fluid_item = self.fluids_table.item(row, 0)
            spec_item = self.fluids_table.item(row, 1)
            amount_item = self.fluids_table.item(row, 2)
            
            if fluid_item and fluid_item.text().strip():
                fluid_name = fluid_item.text().strip()
                vehicle.specifications['fluessigkeiten'][fluid_name] = {
                    'spezifikation': spec_item.text().strip() if spec_item else '',
                    'menge': amount_item.text().strip() if amount_item else ''
                }
        
        vehicle.parts = []
        for row in range(self.parts_table.rowCount()):
            bauteil_item = self.parts_table.item(row, 0)
            teilenummer_item = self.parts_table.item(row, 1)
            motor_code_item = self.parts_table.item(row, 2)
            
            if bauteil_item and bauteil_item.text().strip():
                part_data = {
                    'bauteil': bauteil_item.text().strip(),
                    'teilenummer': teilenummer_item.text().strip() if teilenummer_item else '',
                    'motor_code': motor_code_item.text().strip() if motor_code_item else ''
                }
                vehicle.parts.append(part_data)
# Sicherung
        vehicle.attachments = self.vehicle.attachments
        vehicle.include_attachments_in_export = self.include_attachments_chk.isChecked()
        vehicle.service_history = self.vehicle.service_history
        vehicle.last_service = self.vehicle.last_service
        vehicle.service_intervals = self.vehicle.service_intervals
        
        return vehicle


class ImageViewDialog(QtWidgets.QDialog):
    def __init__(self, attachment, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Bild ansehen - {attachment['filename']}")
        self.setMinimumSize(800, 600)  #
        layout = QtWidgets.QVBoxLayout(self)

        self.image_label = QtWidgets.QLabel()
        self.image_label.setAlignment(QtCore.Qt.AlignCenter)
        self.image_label.setScaledContents(False)
        
# Bild
        if 'scan_path' in attachment and Path(attachment['scan_path']).exists():
            pixmap = QtGui.QPixmap(attachment['scan_path'])
        else:
            image_data = base64.b64decode(attachment['data'])
            pixmap = QtGui.QPixmap()
            pixmap.loadFromData(image_data)
        
        if not pixmap.isNull():
            scaled_pixmap = pixmap.scaled(1000, 700, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation) 
            self.image_label.setPixmap(scaled_pixmap)
        else:
            self.image_label.setText("Bild konnte nicht geladen werden")
        
        layout.addWidget(self.image_label)

        btn = QtWidgets.QPushButton("Schließen")
        btn.clicked.connect(self.accept)
        layout.addWidget(btn)


class UnitComboBox(QtWidgets.QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self.completer = QtWidgets.QCompleter(ALL_UNITS)
        self.completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.completer.setFilterMode(QtCore.Qt.MatchContains)
        self.setCompleter(self.completer)
        
        # Gruppierte Einheiten hinzufügen
        for category, units in AVAILABLE_UNITS.items():
            if units:  # Nur Kategorien mit Einheiten anzeigen
                self.insertSeparator(self.count())
                for unit in units:
                    self.addItem(unit)


class ColumnPropertiesDialog(QtWidgets.QDialog):
    def __init__(self, col: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Spalteneigenschaften")
        self.setMinimumSize(600, 500)  
        self.col = dict(col)  # copy
        layout = QtWidgets.QFormLayout(self)

        self.name_edit = QtWidgets.QLineEdit(self.col.get('name', ''))
        layout.addRow('Name:', self.name_edit)

        self.type_cmb = QtWidgets.QComboBox()
        self.type_cmb.addItems(['Text', 'Zahl', 'Datum', 'Bool'])
        idx = self.type_cmb.findText(self.col.get('type', 'Text'))
        if idx >= 0: self.type_cmb.setCurrentIndex(idx)
        layout.addRow('Typ:', self.type_cmb)

        self.default_edit = QtWidgets.QLineEdit(str(self.col.get('default', '') or ''))
        layout.addRow('Default/Wert:', self.default_edit)

        self.setpoint_edit = QtWidgets.QLineEdit(str(self.col.get('setpoint', '') or ''))
        layout.addRow('Sollwert:', self.setpoint_edit)

        # Maßeinheit Auswahl mit Suche
        layout.addWidget(QtWidgets.QLabel("Maßeinheit:"))
        self.unit_cmb = UnitComboBox()
        current_unit = self.col.get('unit', '')
        idx = self.unit_cmb.findText(current_unit)
        if idx >= 0: 
            self.unit_cmb.setCurrentIndex(idx)
        else:
            self.unit_cmb.setCurrentIndex(0)
        layout.addRow(self.unit_cmb)

        # Toleranz pro Spalte
        self.tolerance_spin = QtWidgets.QDoubleSpinBox()
        self.tolerance_spin.setRange(0.0, 1000.0)
        self.tolerance_spin.setDecimals(3)
        self.tolerance_spin.setSuffix(" %")
        self.tolerance_spin.setValue(float(self.col.get('tolerance', 0.0)))
        self.tolerance_spin.setToolTip("Toleranz für diese spezifische Spalte")
        layout.addRow('Spalten-Toleranz:', self.tolerance_spin)

        self.readonly_chk = QtWidgets.QCheckBox('Sollwert-Zelle schreibgeschützt')
        self.readonly_chk.setChecked(bool(self.col.get('readonly', True)))
        layout.addRow(self.readonly_chk)

        # Immer farblich kennzeichnen
        self.always_highlight_chk = QtWidgets.QCheckBox('Immer farblich kennzeichnen')
        self.always_highlight_chk.setChecked(bool(self.col.get('always_highlight', False)))
        self.always_highlight_chk.setToolTip("Aktiviert farbliche Kennzeichnung auch bei Toleranz = 0")
        layout.addRow(self.always_highlight_chk)

        self.width_spin = QtWidgets.QSpinBox()
        self.width_spin.setRange(50, 1000)
        self.width_spin.setValue(int(self.col.get('width', 120)))
        layout.addRow('Breite (px):', self.width_spin)

        box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        layout.addRow(box)

    def accept(self):
        self.col['name'] = self.name_edit.text().strip() or self.col.get('name', 'unnamed')
        self.col['type'] = self.type_cmb.currentText()
        self.col['default'] = self.default_edit.text().strip() or None
        sp = self.setpoint_edit.text().strip()
        self.col['setpoint'] = sp if sp != '' else None
        self.col['unit'] = self.unit_cmb.currentText()
        self.col['tolerance'] = float(self.tolerance_spin.value())
        self.col['readonly'] = self.readonly_chk.isChecked()
        self.col['always_highlight'] = self.always_highlight_chk.isChecked() 
        self.col['width'] = int(self.width_spin.value())
        super().accept()


class ExportDescriptionDialog(QtWidgets.QDialog):
    def __init__(self, short_description: str = "", long_description: str = "", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Export-Beschreibungen")
        self.setMinimumSize(700, 500)  
        layout = QtWidgets.QVBoxLayout(self)
        
        # Kurzbeschreibung
        layout.addWidget(QtWidgets.QLabel("Kurzbeschreibung (erscheint über der Tabelle):"))
        self.short_desc_edit = QtWidgets.QTextEdit()
        self.short_desc_edit.setMaximumHeight(100)  
        self.short_desc_edit.setPlainText(short_description)
        self.short_desc_edit.setPlaceholderText("Kurze Beschreibung für den Header...")
        layout.addWidget(self.short_desc_edit)
        
        layout.addSpacing(10)
        
        # Langbeschreibung
        layout.addWidget(QtWidgets.QLabel("Ausführliche Beschreibung (erscheint unter der Tabelle):"))
        self.long_desc_edit = QtWidgets.QTextEdit()
        self.long_desc_edit.setPlainText(long_description)
        self.long_desc_edit.setPlaceholderText("Ausführliche Beschreibung für den Footer...")
        layout.addWidget(self.long_desc_edit)
        
        box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

    def get_descriptions(self):
        return self.short_desc_edit.toPlainText().strip(), self.long_desc_edit.toPlainText().strip()


class PendingServicesDialog(QtWidgets.QDialog):
    def __init__(self, vehicles, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Anstehende Services")
        self.setMinimumSize(800, 600)  # Breiter gemacht
        layout = QtWidgets.QVBoxLayout(self)
        
        self.services_table = QtWidgets.QTableWidget()
        self.services_table.setColumnCount(5)
        self.services_table.setHorizontalHeaderLabels(["Fahrzeug", "Service", "Fällig in", "Fällig am", "Status"])
        self.services_table.horizontalHeader().setStretchLastSection(True)
        self._load_pending_services(vehicles)
        layout.addWidget(self.services_table)
        
        btn = QtWidgets.QPushButton("Schließen")
        btn.clicked.connect(self.accept)
        layout.addWidget(btn)
    
    def _load_pending_services(self, vehicles):
        pending_services = []
        
        for vehicle in vehicles:
            for service in vehicle.service_history:
                # Nur OFFENE Services prüfen
                if service.get('status') == 'offen':
                    service_name = service.get('service', '')
                    next_km = service.get('naechster_service_km', 0)
                    current_km = vehicle.last_service.get('ölwechsel_km', 0)
                    km_remaining = next_km - current_km
                    
                    # Wenn weniger als 500km bis zum Service
                    if km_remaining <= 500:
                        status = "BALD FÄLLIG" if km_remaining > 0 else "ÜBERFÄLLIG"
                        pending_services.append({
                            'vehicle': vehicle.name,
                            'service': service_name,
                            'remaining': f"{km_remaining} km",
                            'due_date': f"bei {next_km} km",
                            'status': status
                        })
        
        self.services_table.setRowCount(len(pending_services))
        for row, service in enumerate(pending_services):
            self.services_table.setItem(row, 0, QtWidgets.QTableWidgetItem(service['vehicle']))
            self.services_table.setItem(row, 1, QtWidgets.QTableWidgetItem(service['service']))
            self.services_table.setItem(row, 2, QtWidgets.QTableWidgetItem(service['remaining']))
            self.services_table.setItem(row, 3, QtWidgets.QTableWidgetItem(service['due_date']))
            
            status_item = QtWidgets.QTableWidgetItem(service['status'])
            if service['status'] == "ÜBERFÄLLIG":
                status_item.setBackground(QtGui.QColor('#ff6b6b'))
            else:
                status_item.setBackground(QtGui.QColor('#ffd700'))
            self.services_table.setItem(row, 4, status_item)


class ToleranceDialog(QtWidgets.QDialog):
    def __init__(self, current_tolerance=0.0, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Globale Toleranz einstellen")
        self.setMinimumSize(400, 150)  # Breiter gemacht
        layout = QtWidgets.QFormLayout(self)
        
        self.tolerance_spin = QtWidgets.QDoubleSpinBox()
        self.tolerance_spin.setRange(0.0, 100.0)
        self.tolerance_spin.setDecimals(2)
        self.tolerance_spin.setSuffix(" %")
        self.tolerance_spin.setValue(current_tolerance)
        layout.addRow("Globale Toleranz:", self.tolerance_spin)
        
        box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        layout.addRow(box)
    
    def get_tolerance(self):
        return self.tolerance_spin.value()

class UnitsViewDialog(QtWidgets.QDialog): # ansicht der maßeinheiten
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Maßeinheiten - Übersicht")
        self.setMinimumSize(600, 500)
        layout = QtWidgets.QVBoxLayout(self)
        
        # Suchfeld und Info in EINER Reihe
        search_layout = QtWidgets.QHBoxLayout()
        
        search_layout.addWidget(QtWidgets.QLabel("Suchen:"))
        self.search_edit = QtWidgets.QLineEdit()
        self.search_edit.setPlaceholderText("Maßeinheit oder Kategorie suchen...")
        self.search_edit.textChanged.connect(self.filter_units)
        search_layout.addWidget(self.search_edit)
        
        # Aus-/Einklappen Buttons
        expand_btn = QtWidgets.QPushButton("▶ Alle")
        expand_btn.setFixedWidth(60)
        expand_btn.clicked.connect(self.collapse_all)
        collapse_btn = QtWidgets.QPushButton("▼ Alle") 
        collapse_btn.setFixedWidth(60)
        collapse_btn.clicked.connect(self.expand_all)
        
        search_layout.addWidget(expand_btn)
        search_layout.addWidget(collapse_btn)
        
        # NEU: Anzahl der Maßeinheiten RECHTS
        total_units = sum(len(units) for units in AVAILABLE_UNITS.values())
        count_label = QtWidgets.QLabel(f"{total_units} Einheiten")
        count_label.setStyleSheet("color: #888; font-size: 11px; padding: 2px;")
        search_layout.addWidget(count_label)
        
        layout.addLayout(search_layout)
        
        # Tree Widget für Kategorien und Einheiten
        self.units_tree = QtWidgets.QTreeWidget()
        self.units_tree.setHeaderLabels(["Kategorie / Maßeinheit", "Einheiten"])
        self.units_tree.setColumnCount(2)
        self.units_tree.setAlternatingRowColors(False)
        
        # Style für schwarzen Hintergrund und weiße Schrift
        self.units_tree.setStyleSheet("""
            QTreeWidget {
                background-color: #0f0f12;
                color: #ffffff;
                border: 1px solid #2a2a33;
                outline: 0;
            }
            QTreeWidget::item {
                background-color: #0f0f12;
                color: #ffffff;
                border: none;
                padding: 4px;
            }
            QTreeWidget::item:selected {
                background-color: #2b6ea3;
                color: #ffffff;
            }
            QTreeWidget::item:hover {
                background-color: #1a1a1f;
            }
        """)
        
        self.units_tree.header().setStretchLastSection(True)
        
        # Header Style
        self.units_tree.header().setStyleSheet("""
            QHeaderView::section {
                background-color: #16161a;
                color: #ffffff;
                padding: 6px;
                border: none;
                font-weight: bold;
            }
        """)
        
        # Einheiten laden
        self._load_units_to_tree()
        
        layout.addWidget(self.units_tree)
        
        # Info Label
        info_label = QtWidgets.QLabel("--> Nur zur Ansicht - Einheiten können in den Spalteneigenschaften ausgewählt werden")
        info_label.setStyleSheet("color: #888; font-style: italic; padding: 5px;")
        layout.addWidget(info_label)
        
        # Schließen-Button
        btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)
    
    def _load_units_to_tree(self):
        self.units_tree.clear()
        
        for category, units in AVAILABLE_UNITS.items():
            if units:  # Nur Kategorien mit Einheiten anzeigen
                category_item = QtWidgets.QTreeWidgetItem([category, f"{len(units)} Einheiten"])
                category_item.setExpanded(True)  # Standardmäßig eingeklappt
                
                for unit in units:
                    unit_item = QtWidgets.QTreeWidgetItem([unit, ""])
                    category_item.addChild(unit_item)
                
                self.units_tree.addTopLevelItem(category_item)
        
        # Erste Spalte etwas breiter machen
        self.units_tree.setColumnWidth(0, 250)
    
    def filter_units(self, text):
        search_text = text.lower().strip()
        
        for i in range(self.units_tree.topLevelItemCount()):
            category_item = self.units_tree.topLevelItem(i)
            category_visible = False
            
            # Durch alle Unterelemente (Einheiten) gehen
            for j in range(category_item.childCount()):
                unit_item = category_item.child(j)
                unit_text = unit_item.text(0).lower()
                
                # Prüfen ob Einheit oder Kategorie dem Suchtext entspricht
                if search_text in unit_text or search_text in category_item.text(0).lower():
                    unit_item.setHidden(False)
                    category_visible = True
                else:
                    unit_item.setHidden(True)
            
            # Kategorie anzeigen/verstecken basierend auf Suchergebnissen
            category_item.setHidden(not category_visible)
            
            # Kategorie expandieren wenn Suchtext vorhanden
            category_item.setExpanded(bool(search_text))
    
    # NEU: Methoden für Aus-/Einklappen
    def expand_all(self):
        for i in range(self.units_tree.topLevelItemCount()):
            item = self.units_tree.topLevelItem(i)
            item.setExpanded(True)
    
    def collapse_all(self):
        for i in range(self.units_tree.topLevelItemCount()):
            item = self.units_tree.topLevelItem(i)
            item.setExpanded(False)

class DefectReportsDialog(QtWidgets.QDialog):
    def __init__(self, vehicle: Vehicle, parent=None):
        super().__init__(parent)
        self.vehicle = vehicle
        self.setWindowTitle(f"Beanstandungsaufnahme - {vehicle.name}")
        
        # Autoscaling
        screen = QtWidgets.QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        self.setMinimumSize(790, 500)
        self.resize(int(screen_geometry.width() * 0.9), int(screen_geometry.height() * 0.85))
        
        layout = QtWidgets.QVBoxLayout(self)
        
        # Fahrzeug info oben
        vehicle_info = QtWidgets.QGroupBox("Fahrzeuginformation")
        vehicle_layout = QtWidgets.QHBoxLayout()
        vehicle_layout.addWidget(QtWidgets.QLabel(f"Fahrzeug: {vehicle.name}"))
        vehicle_layout.addWidget(QtWidgets.QLabel(f"FIN: {vehicle.specifications.get('fin', '')}"))
        vehicle_layout.addWidget(QtWidgets.QLabel(f"Aktueller KM-Stand: {vehicle.last_service.get('ölwechsel_km', '0')} km"))
        vehicle_layout.addStretch()
        vehicle_info.setLayout(vehicle_layout)
        layout.addWidget(vehicle_info)
        
        # für hauptinhalt
        scroll_area = QtWidgets.QScrollArea()
        scroll_area.setWidgetResizable(True)
        container = QtWidgets.QWidget()
        container_layout = QtWidgets.QVBoxLayout(container)
        
        # Neue Beanstandung
        new_report_group = QtWidgets.QGroupBox("Neue Beanstandung erfassen")
        new_report_layout = QtWidgets.QFormLayout()
        
        # Erste Zeile: Datum und KM-Stand
        date_km_layout = QtWidgets.QHBoxLayout()
        
        self.date_edit = QtWidgets.QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QtCore.QDate.currentDate())
        date_km_layout.addWidget(QtWidgets.QLabel("Datum:"))
        date_km_layout.addWidget(self.date_edit)
        
        self.km_edit = QtWidgets.QSpinBox()
        self.km_edit.setRange(0, 1000000)
        self.km_edit.setValue(self.vehicle.last_service.get('ölwechsel_km', 0))
        date_km_layout.addWidget(QtWidgets.QLabel("KM-Stand:"))
        date_km_layout.addWidget(self.km_edit)
        date_km_layout.addStretch()
        
        new_report_layout.addRow(date_km_layout)
        
        # Beschreibung
        self.description_edit = QtWidgets.QTextEdit()
        self.description_edit.setMaximumHeight(100)
        self.description_edit.setPlaceholderText("Detaillierte Beschreibung des Defekts/Problems...")
        new_report_layout.addRow("Beschreibung:", self.description_edit)
        
        # Status, Priorität, Anmerkung
        status_priority_layout = QtWidgets.QHBoxLayout()
        
        self.status_combo = QtWidgets.QComboBox()
        self.status_combo.addItems(["Offen", "In Bearbeitung", "Behoben", "Storniert"])
        status_priority_layout.addWidget(QtWidgets.QLabel("Status:"))
        status_priority_layout.addWidget(self.status_combo)
        
        self.priority_combo = QtWidgets.QComboBox()
        self.priority_combo.addItems(["Niedrig", "Mittel", "Hoch", "Kritisch"])
        status_priority_layout.addWidget(QtWidgets.QLabel("Priorität:"))
        status_priority_layout.addWidget(self.priority_combo)
        
        self.solution_edit = QtWidgets.QLineEdit()
        self.solution_edit.setPlaceholderText("Anmerkung (optional)")
        status_priority_layout.addWidget(QtWidgets.QLabel("Anmerkung:"))
        status_priority_layout.addWidget(self.solution_edit)
        
        status_priority_layout.addStretch()
        new_report_layout.addRow(status_priority_layout)
        
        add_btn = QtWidgets.QPushButton("+ Beanstandung hinzufügen")
        add_btn.clicked.connect(self.add_defect_report)
        new_report_layout.addRow(add_btn)
        
        new_report_group.setLayout(new_report_layout)
        container_layout.addWidget(new_report_group)
        
        # Bestehende Beanstandungen, Tabelle
        existing_group = QtWidgets.QGroupBox("Bestehende Beanstandungen")
        existing_layout = QtWidgets.QVBoxLayout()
        
        self.defects_table = QtWidgets.QTableWidget()
        self.defects_table.setColumnCount(6)
        self.defects_table.setHorizontalHeaderLabels(["Datum", "KM-Stand", "Beschreibung", "Status", "Priorität", "Anmerkung"])
        self.defects_table.setAlternatingRowColors(False)
        self.defects_table.horizontalHeader().setStretchLastSection(True)
        self.defects_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.defects_table.doubleClicked.connect(self.show_report_details)
        self._load_defects_to_table()
        existing_layout.addWidget(self.defects_table)
        
        table_controls = QtWidgets.QHBoxLayout()
        remove_btn = QtWidgets.QPushButton("Ausgewählte entfernen")
        remove_btn.clicked.connect(self.remove_defect_report)
        edit_btn = QtWidgets.QPushButton("Bearbeiten")
        edit_btn.clicked.connect(self.edit_defect_report)
        details_btn = QtWidgets.QPushButton("Details anzeigen")
        details_btn.clicked.connect(self.show_report_details)
        export_btn = QtWidgets.QPushButton("HTML Export")
        export_btn.clicked.connect(self.export_defects_html)
        
        table_controls.addWidget(remove_btn)
        table_controls.addWidget(edit_btn)
        table_controls.addWidget(details_btn)
        table_controls.addWidget(export_btn)
        table_controls.addStretch()
        
        count_label = QtWidgets.QLabel(f"Gesamt: {len(self.vehicle.defect_reports)} Beanstandungen")
        count_label.setStyleSheet("color: #888;")
        table_controls.addWidget(count_label)
        
        existing_layout.addLayout(table_controls)
        existing_group.setLayout(existing_layout)
        container_layout.addWidget(existing_group)
        
        container_layout.addStretch()
        scroll_area.setWidget(container)
        layout.addWidget(scroll_area)
        
        # Buttons unten
        box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        box.rejected.connect(self.reject)
        layout.addWidget(box)
    
    def _load_defects_to_table(self):
        self.defects_table.setRowCount(len(self.vehicle.defect_reports))
        
        for row, report in enumerate(self.vehicle.defect_reports):
            self.defects_table.setItem(row, 0, QtWidgets.QTableWidgetItem(report.get('datum', '')))
            self.defects_table.setItem(row, 1, QtWidgets.QTableWidgetItem(str(report.get('km_stand', ''))))
            self.defects_table.setItem(row, 2, QtWidgets.QTableWidgetItem(report.get('beschreibung', '')))
            
            status_item = QtWidgets.QTableWidgetItem(report.get('status', ''))
            priority_item = QtWidgets.QTableWidgetItem(report.get('priorität', ''))
            solution_item = QtWidgets.QTableWidgetItem(report.get('Anmerkung', ''))
            
            # Farbliche Kennzeichnung
            status = report.get('status', '')
            if status == 'Behoben':
                status_item.setBackground(QtGui.QColor('#1e3a1e'))
            elif status == 'Storniert':
                status_item.setBackground(QtGui.QColor('#3a3a3a'))
                
            priority = report.get('priorität', '')
            if priority == 'Kritisch':
                priority_item.setBackground(QtGui.QColor('#3a1e1e'))
            elif priority == 'Hoch':
                priority_item.setBackground(QtGui.QColor('#5a3a1e'))
            
            self.defects_table.setItem(row, 3, status_item)
            self.defects_table.setItem(row, 4, priority_item)
            self.defects_table.setItem(row, 5, solution_item)
    
    def add_defect_report(self):
        beschreibung = self.description_edit.toPlainText().strip()
        if not beschreibung:
            QtWidgets.QMessageBox.warning(self, "Fehler", "Bitte eine Beschreibung eingeben.")
            return
        
        report = {
            'datum': self.date_edit.date().toString("yyyy-MM-dd"),
            'km_stand': self.km_edit.value(),
            'beschreibung': beschreibung,
            'status': self.status_combo.currentText(),
            'priorität': self.priority_combo.currentText(),
            'Anmerkung': self.solution_edit.text().strip()
        }
        
        self.vehicle.defect_reports.append(report)
        self._load_defects_to_table()
        
        # Automatisch speichern
        self._auto_save_vehicle()
        
        # Felder zurücksetzen
        self.description_edit.clear()
        self.solution_edit.clear()
        self.date_edit.setDate(QtCore.QDate.currentDate())
        self.status_combo.setCurrentIndex(0)
        self.priority_combo.setCurrentIndex(0)
        
        QtWidgets.QMessageBox.information(self, "Erfolg", "Beanstandung wurde hinzugefügt und gespeichert.")
    
    def _auto_save_vehicle(self):
        try:
            self.vehicle.ensure_directories()
            vehicle_file = self.vehicle.get_vehicle_dir() / "vehicle.json"
            with open(vehicle_file, "w", encoding="utf-8") as f:
                json.dump(self.vehicle.to_dict(), f, indent=2, ensure_ascii=False)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Auto-Save Fehler", f"Fehler beim Speichern: {e}")
    
    def remove_defect_report(self):
        current_row = self.defects_table.currentRow()
        if current_row >= 0:
            reply = QtWidgets.QMessageBox.question(self, "Löschen", 
                                                 "Beanstandung wirklich löschen?",
                                                 QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if reply == QtWidgets.QMessageBox.Yes:
                self.vehicle.defect_reports.pop(current_row)
                self._load_defects_to_table()
                self._auto_save_vehicle()
    
    def edit_defect_report(self):
        current_row = self.defects_table.currentRow()
        if current_row >= 0:
            # Erweiterte Bearbeitung
            report = self.vehicle.defect_reports[current_row]
            dlg = DefectReportDetailDialog(report, self)
            if dlg.exec_() == QtWidgets.QDialog.Accepted:
                self.vehicle.defect_reports[current_row] = dlg.get_report_data()
                self._load_defects_to_table()
                self._auto_save_vehicle()
    
    def show_report_details(self):
        current_row = self.defects_table.currentRow()
        if current_row >= 0:
            report = self.vehicle.defect_reports[current_row]
            dlg = DefectReportDetailDialog(report, self, read_only=True)
            dlg.exec_()
    
    def export_defects_html(self):
        if not self.vehicle.defect_reports:
            QtWidgets.QMessageBox.warning(self, "Fehler", "Keine Beanstandungen zum Exportieren vorhanden.")
            return
            
        fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Beanstandungen exportieren", 
                                                        str(VEHICLES_BASE_DIR / f"{self.vehicle.name}_beanstandungen.html"), 
                                                        "HTML Dateien (*.html)")
        if fname:
            html = self._generate_defects_html()
            with open(fname, 'w', encoding='utf-8') as f:
                f.write(html)
            QtWidgets.QMessageBox.information(self, "Erfolg", f"Beanstandungen exportiert: {fname}")
    
    def _generate_defects_html(self):
        html = f"""
        <html>
        <head>
            <meta charset='utf-8'>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                .header {{ background: #2c3e50; color: white; padding: 20px; border-radius: 5px; }}
                .table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                .table th, .table td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                .table th {{ background-color: #f2f2f2; }}
                .status-behoben {{ background-color: #d4edda; }}
                .status-storniert {{ background-color: #e2e3e5; }}
                .priority-kritisch {{ background-color: #f8d7da; font-weight: bold; }}
                .priority-hoch {{ background-color: #fff3cd; }}
                .report-detail {{ margin: 10px 0; padding: 10px; background: #f8f9fa; border-left: 4px solid #007bff; }}
            </style>
        </head>
        <body>
            <div class='header'>
                <h1>Beanstandungen - {self.vehicle.name}</h1>
                <p>FIN: {self.vehicle.specifications.get('fin', '')} | Exportiert am: {datetime.now().strftime("%d.%m.%Y %H:%M")}</p>
            </div>
            
            <h2>Übersichtstabelle</h2>
            <table class='table'>
                <tr>
                    <th>Datum</th>
                    <th>KM-Stand</th>
                    <th>Beschreibung</th>
                    <th>Status</th>
                    <th>Priorität</th>
                    <th>Anmerkung</th>
                </tr>
        """
        
        for report in self.vehicle.defect_reports:
            status_class = "status-behoben" if report['status'] == 'Behoben' else "status-storniert" if report['status'] == 'Storniert' else ""
            priority_class = "priority-kritisch" if report['priorität'] == 'Kritisch' else "priority-hoch" if report['priorität'] == 'Hoch' else ""
            
            html += f"""
                <tr>
                    <td>{report['datum']}</td>
                    <td>{report['km_stand']} km</td>
                    <td>{report['beschreibung']}</td>
                    <td class='{status_class}'>{report['status']}</td>
                    <td class='{priority_class}'>{report['priorität']}</td>
                    <td>{report.get('Anmerkung', '')}</td>
                </tr>
            """
        
        html += """
            </table>
            
            <h2>Detaillierte Berichte</h2>
        """
        
        for i, report in enumerate(self.vehicle.defect_reports, 1):
            html += f"""
            <div class='report-detail'>
                <h3>Beanstandung #{i} - {report['datum']}</h3>
                <p><strong>KM-Stand:</strong> {report['km_stand']} km</p>
                <p><strong>Status:</strong> {report['status']} | <strong>Priorität:</strong> {report['priorität']}</p>
                <p><strong>Beschreibung:</strong><br>{report['beschreibung'].replace(chr(10), '<br>')}</p>
                <p><strong>Anmerkung:</strong><br>{report.get('Anmerkung', 'Keine Anmerkung eingetragen').replace(chr(10), '<br>')}</p>
            </div>
            """
        
        html += f"""
            <div style='margin-top: 20px; color: #666;'>
                <p>Erstellt mit:<br>o3Measurement {o3VERSION}<br>{o3COPYRIGHT}</p>
            </div>
        </body>
        </html>
        """
        return html

class DefectReportDetailDialog(QtWidgets.QDialog):
    def __init__(self, report_data: dict, parent=None, read_only=False):
        super().__init__(parent)
        self.report_data = report_data.copy()
        self.read_only = read_only
        
        title = "Beanstandungs-Details" if read_only else "Beanstandung bearbeiten"
        self.setWindowTitle(title)
        self.setMinimumSize(600, 500)
        
        layout = QtWidgets.QFormLayout(self)
        
        # Datum
        self.date_edit = QtWidgets.QDateEdit()
        self.date_edit.setCalendarPopup(True)
        if report_data.get('datum'):
            self.date_edit.setDate(QtCore.QDate.fromString(report_data['datum'], "yyyy-MM-dd"))
        else:
            self.date_edit.setDate(QtCore.QDate.currentDate())
        self.date_edit.setReadOnly(read_only)
        layout.addRow("Datum:", self.date_edit)
        
        # KM-Stand
        self.km_edit = QtWidgets.QSpinBox()
        self.km_edit.setRange(0, 1000000)
        self.km_edit.setValue(report_data.get('km_stand', 0))
        self.km_edit.setReadOnly(read_only)
        layout.addRow("KM-Stand:", self.km_edit)
        
        # Beschreibung
        self.description_edit = QtWidgets.QTextEdit()
        self.description_edit.setPlainText(report_data.get('beschreibung', ''))
        self.description_edit.setReadOnly(read_only)
        layout.addRow("Beschreibung:", self.description_edit)
        
        # Status und Priorität
        status_priority_layout = QtWidgets.QHBoxLayout()
        
        self.status_combo = QtWidgets.QComboBox()
        self.status_combo.addItems(["Offen", "In Bearbeitung", "Behoben", "Storniert"])
        self.status_combo.setCurrentText(report_data.get('status', 'Offen'))
        self.status_combo.setEnabled(not read_only)
        status_priority_layout.addWidget(QtWidgets.QLabel("Status:"))
        status_priority_layout.addWidget(self.status_combo)
        
        self.priority_combo = QtWidgets.QComboBox()
        self.priority_combo.addItems(["Niedrig", "Mittel", "Hoch", "Kritisch"])
        self.priority_combo.setCurrentText(report_data.get('priorität', 'Mittel'))
        self.priority_combo.setEnabled(not read_only)
        status_priority_layout.addWidget(QtWidgets.QLabel("Priorität:"))
        status_priority_layout.addWidget(self.priority_combo)
        
        status_priority_layout.addStretch()
        layout.addRow(status_priority_layout)
        
        # Anmerkung
        self.solution_edit = QtWidgets.QTextEdit()
        self.solution_edit.setPlainText(report_data.get('Anmerkung', ''))
        self.solution_edit.setMaximumHeight(80)
        self.solution_edit.setReadOnly(read_only)
        layout.addRow("Anmerkung:", self.solution_edit)
        
        if not read_only:
            box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
            box.accepted.connect(self.accept)
            box.rejected.connect(self.reject)
            layout.addRow(box)
        else:
            box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
            box.rejected.connect(self.reject)
            layout.addRow(box)
    
    def get_report_data(self):
        return {
            'datum': self.date_edit.date().toString("yyyy-MM-dd"),
            'km_stand': self.km_edit.value(),
            'beschreibung': self.description_edit.toPlainText().strip(),
            'status': self.status_combo.currentText(),
            'priorität': self.priority_combo.currentText(),
            'Anmerkung': self.solution_edit.toPlainText().strip()
        }

class StorageItem:
    def __init__(self, teilenummer="", hersteller="", bezeichnung="", kategorie="", lagerplatz="", 
                 fach="", anzahl=0, mindestbestand=0, einheit="", zusatzinfos="", hersteller_teilenummer=""):
        self.teilenummer = teilenummer
        self.hersteller = hersteller
        self.bezeichnung = bezeichnung
        self.kategorie = kategorie
        self.lagerplatz = lagerplatz
        self.fach = fach
        self.anzahl = anzahl
        self.mindestbestand = mindestbestand
        self.einheit = einheit
        self.zusatzinfos = zusatzinfos
        self.hersteller_teilenummer = hersteller_teilenummer
        self.erstellt_am = datetime.now().strftime("%Y-%m-%d")
        self.geaendert_am = datetime.now().strftime("%Y-%m-%d")
    
    def to_dict(self):
        return {
            'teilenummer': self.teilenummer,
            'hersteller': self.hersteller,
            'bezeichnung': self.bezeichnung,
            'kategorie': self.kategorie,
            'lagerplatz': self.lagerplatz,
            'fach': self.fach,
            'anzahl': self.anzahl,
            'mindestbestand': self.mindestbestand,
            'einheit': self.einheit,
            'zusatzinfos': self.zusatzinfos,
            'hersteller_teilenummer': self.hersteller_teilenummer,
            'erstellt_am': self.erstellt_am,
            'geaendert_am': self.geaendert_am
        }
    
    @staticmethod
    def from_dict(d):
        item = StorageItem()
        item.teilenummer = d.get('teilenummer', '')
        item.hersteller = d.get('hersteller', '')
        item.bezeichnung = d.get('bezeichnung', '')
        item.kategorie = d.get('kategorie', '')
        item.lagerplatz = d.get('lagerplatz', '')
        item.fach = d.get('fach', '')
        item.anzahl = d.get('anzahl', 0)
        item.mindestbestand = d.get('mindestbestand', 0)
        item.einheit = d.get('einheit', '')
        item.zusatzinfos = d.get('zusatzinfos', '')
        item.hersteller_teilenummer = d.get('hersteller_teilenummer', '')
        item.erstellt_am = d.get('erstellt_am', '')
        item.geaendert_am = d.get('geaendert_am', '')
        return item

class StorageManager:
    def __init__(self):
        self.storage_file = STORAGE_DIR / "lagerbestand.json"
        self.items = []
        self.load_storage()
    
    def load_storage(self):
        if self.storage_file.exists():
            try:
                with open(self.storage_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.items = [StorageItem.from_dict(item) for item in data]
            except Exception as e:
                print(f"Fehler beim Laden des Lagerbestands: {e}")
                self.items = []
        else:
            self.items = []
    
    def save_storage(self):
        try:
            data = [item.to_dict() for item in self.items]
            with open(self.storage_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Fehler beim Speichern des Lagerbestands: {e}")
    
    def add_item(self, item):
        self.items.append(item)
        self.save_storage()
    
    def update_item(self, index, item):
        if 0 <= index < len(self.items):
            self.items[index] = item
            self.save_storage()
    
    def remove_item(self, index):
        if 0 <= index < len(self.items):
            self.items.pop(index)
            self.save_storage()
    
    def search_items(self, search_term):
        search_term = search_term.lower()
        results = []
        for item in self.items:
            if (search_term in item.teilenummer.lower() or 
                search_term in item.hersteller.lower() or 
                search_term in item.bezeichnung.lower() or 
                search_term in item.kategorie.lower() or 
                search_term in item.lagerplatz.lower() or
                search_term in item.zusatzinfos.lower()):
                results.append(item)
        return results
    
    def get_categories(self):
        return sorted(set(item.kategorie for item in self.items if item.kategorie))
    
    def get_locations(self):
        return sorted(set(item.lagerplatz for item in self.items if item.lagerplatz))
    
    def get_manufacturers(self):
        return sorted(set(item.hersteller for item in self.items if item.hersteller))

class StorageManagerDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.storage_manager = StorageManager()
        self.setWindowTitle("Lagerbestandsverwaltung")
        
        # Autoscaling
        screen = QtWidgets.QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        self.setMinimumSize(800, 550)
        self.resize(int(screen_geometry.width() * 0.9), int(screen_geometry.height() * 0.85))
        
        layout = QtWidgets.QVBoxLayout(self)
        
        self._create_search_filter_bar(layout)
        
        # Tabelle
        self._create_main_table(layout)
        
        # Buttons
        self._create_control_buttons(layout)
        
        self._update_filter_combos()
        
    def _create_search_filter_bar(self, layout):
        search_filter_group = QtWidgets.QGroupBox("Suche & Filter")
        search_filter_layout = QtWidgets.QHBoxLayout()
        
        # Suchfeld
        search_filter_layout.addWidget(QtWidgets.QLabel("Suche:"))
        self.search_edit = QtWidgets.QLineEdit()
        self.search_edit.setPlaceholderText("Teilenummer, Hersteller, Bezeichnung...")
        self.search_edit.textChanged.connect(self.filter_items)
        search_filter_layout.addWidget(self.search_edit)
        
        search_filter_layout.addWidget(QtWidgets.QLabel("Kategorie:"))
        self.category_combo = QtWidgets.QComboBox()
        self.category_combo.addItem("Alle Kategorien", "")
        self.category_combo.currentTextChanged.connect(self.filter_items)
        search_filter_layout.addWidget(self.category_combo)
        
        # Hersteller filter
        search_filter_layout.addWidget(QtWidgets.QLabel("Hersteller:"))
        self.manufacturer_combo = QtWidgets.QComboBox()
        self.manufacturer_combo.addItem("Alle Hersteller", "")
        self.manufacturer_combo.currentTextChanged.connect(self.filter_items)
        search_filter_layout.addWidget(self.manufacturer_combo)
        
        search_filter_layout.addWidget(QtWidgets.QLabel("Lagerplatz:"))
        self.location_combo = QtWidgets.QComboBox()
        self.location_combo.addItem("Alle Lagerplätze", "")
        self.location_combo.currentTextChanged.connect(self.filter_items)
        search_filter_layout.addWidget(self.location_combo)
        
        search_filter_layout.addStretch()
        
        # Anzahl
        self.count_label = QtWidgets.QLabel("0 Teile")
        self.count_label.setStyleSheet("color: #888; font-weight: bold;")
        search_filter_layout.addWidget(self.count_label)
        
        search_filter_group.setLayout(search_filter_layout)
        layout.addWidget(search_filter_group)
            
    def _create_main_table(self, layout):
        table_group = QtWidgets.QGroupBox("Lagerbestand")
        table_layout = QtWidgets.QVBoxLayout()
        
        self.storage_table = QtWidgets.QTableWidget()
        self.storage_table.setColumnCount(12)
        self.storage_table.setHorizontalHeaderLabels([
            "Teilenummer", "Hersteller", "Bezeichnung", "Kategorie", 
            "Lagerplatz", "Fach", "Anzahl", "Mindestbestand", "Einheit",
            "Hersteller-TNr", "Zusatzinfos", "Geändert am"
        ])
        self.storage_table.setAlternatingRowColors(False)
        self.storage_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.storage_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.storage_table.doubleClicked.connect(self.edit_item)
        self.storage_table.setSortingEnabled(True)
        
        # Spaltenbreite
        self.storage_table.setColumnWidth(0, 120)
        self.storage_table.setColumnWidth(1, 120)
        self.storage_table.setColumnWidth(2, 200)
        self.storage_table.setColumnWidth(3, 120)
        self.storage_table.setColumnWidth(4, 100)
        self.storage_table.setColumnWidth(5, 80)
        self.storage_table.setColumnWidth(6, 80)
        self.storage_table.setColumnWidth(7, 100)
        self.storage_table.setColumnWidth(8, 80)
        self.storage_table.setColumnWidth(9, 120)
        self.storage_table.setColumnWidth(10, 200)
        self.storage_table.setColumnWidth(11, 100)
        self.storage_table.horizontalHeader().sortIndicatorChanged.connect(self.on_sort_changed)
        
        table_layout.addWidget(self.storage_table)
        table_group.setLayout(table_layout)
        layout.addWidget(table_group)
        
        self._load_table_data()
    
    def on_sort_changed(self, logical_index, order):
        # Auswahl speichern
        current_selection = None
        if self.storage_table.currentRow() >= 0:
            current_item = self.storage_table.item(self.storage_table.currentRow(), 0)
            if current_item:
                current_selection = current_item.text()
        
        if current_selection:
            QtCore.QTimer.singleShot(50, lambda: self.restore_selection(current_selection))

    def restore_selection(self, teilenummer):
        for row in range(self.storage_table.rowCount()):
            if self.storage_table.item(row, 0) and self.storage_table.item(row, 0).text() == teilenummer:
                self.storage_table.setCurrentCell(row, 0)
                self.storage_table.scrollToItem(self.storage_table.item(row, 0))
                break

    def _create_control_buttons(self, layout):
        control_layout = QtWidgets.QHBoxLayout()
        
        # buttons
        add_btn = QtWidgets.QPushButton("+ Neu")
        add_btn.clicked.connect(self.add_item)
        edit_btn = QtWidgets.QPushButton("Bearbeiten")
        edit_btn.clicked.connect(self.edit_item)
        remove_btn = QtWidgets.QPushButton("X Löschen")
        remove_btn.clicked.connect(self.remove_item)
        increase_btn = QtWidgets.QPushButton("+ Bestand erhöhen")
        increase_btn.clicked.connect(self.increase_stock)
        decrease_btn = QtWidgets.QPushButton("- Bestand verringern")
        decrease_btn.clicked.connect(self.decrease_stock)
        export_btn = QtWidgets.QPushButton("Bestand Exportieren")
        export_btn.clicked.connect(self.export_storage)
        
        control_layout.addWidget(add_btn)
        control_layout.addWidget(edit_btn)
        control_layout.addWidget(remove_btn)
        control_layout.addWidget(increase_btn)
        control_layout.addWidget(decrease_btn)
        control_layout.addWidget(export_btn)
        control_layout.addStretch()
        
        layout.addLayout(control_layout)
        
        # Button
        box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        box.rejected.connect(self.reject)
        layout.addWidget(box)
    
    def _update_filter_combos(self):
        # Kategorien
        self.category_combo.clear()
        self.category_combo.addItem("Alle Kategorien", "")
        for category in self.storage_manager.get_categories():
            self.category_combo.addItem(category, category)
        
        # Hersteller
        self.manufacturer_combo.clear()
        self.manufacturer_combo.addItem("Alle Hersteller", "")
        for manufacturer in self.storage_manager.get_manufacturers():
            self.manufacturer_combo.addItem(manufacturer, manufacturer)
        
        # Lagerplätze
        self.location_combo.clear()
        self.location_combo.addItem("Alle Lagerplätze", "")
        for location in self.storage_manager.get_locations():
            self.location_combo.addItem(location, location)
    
    def _load_table_data(self, items=None):
        current_selection = None
        if self.storage_table.currentRow() >= 0:
            current_item = self.storage_table.item(self.storage_table.currentRow(), 0)
            if current_item:
                current_selection = current_item.text()
        
        if items is None:
            items = self.storage_manager.items
        
        self.storage_table.setRowCount(len(items))
        
        for row, item in enumerate(items):
            self.storage_table.setItem(row, 0, QtWidgets.QTableWidgetItem(item.teilenummer))
            self.storage_table.setItem(row, 1, QtWidgets.QTableWidgetItem(item.hersteller))
            self.storage_table.setItem(row, 2, QtWidgets.QTableWidgetItem(item.bezeichnung))
            self.storage_table.setItem(row, 3, QtWidgets.QTableWidgetItem(item.kategorie))
            self.storage_table.setItem(row, 4, QtWidgets.QTableWidgetItem(item.lagerplatz))
            self.storage_table.setItem(row, 5, QtWidgets.QTableWidgetItem(item.fach))
            
    # bei niedrigem bestand warnen, farblich
            anzahl_item = QtWidgets.QTableWidgetItem(str(item.anzahl))
            if item.anzahl <= item.mindestbestand:
                anzahl_item.setBackground(QtGui.QColor('#3a1e1e'))
            self.storage_table.setItem(row, 6, anzahl_item)

            self.storage_table.setItem(row, 7, QtWidgets.QTableWidgetItem(str(item.mindestbestand)))
            self.storage_table.setItem(row, 8, QtWidgets.QTableWidgetItem(item.einheit))
            self.storage_table.setItem(row, 9, QtWidgets.QTableWidgetItem(item.hersteller_teilenummer))
            self.storage_table.setItem(row, 10, QtWidgets.QTableWidgetItem(item.zusatzinfos))
            self.storage_table.setItem(row, 11, QtWidgets.QTableWidgetItem(item.geaendert_am))
        
        self.count_label.setText(f"{len(items)} Einträge")
        
        if current_selection:
            for row in range(self.storage_table.rowCount()):
                if self.storage_table.item(row, 0) and self.storage_table.item(row, 0).text() == current_selection:
                    self.storage_table.setCurrentCell(row, 0)
                    break

    def filter_items(self):
        self.storage_table.setSortingEnabled(False)
        
        search_term = self.search_edit.text().lower()
        category_filter = self.category_combo.currentData()
        manufacturer_filter = self.manufacturer_combo.currentData()
        location_filter = self.location_combo.currentData()
        
        filtered_items = []
        for item in self.storage_manager.items:
            # Text suche
            text_match = (not search_term or 
                        search_term in item.teilenummer.lower() or 
                        search_term in item.hersteller.lower() or 
                        search_term in item.bezeichnung.lower() or 
                        search_term in item.kategorie.lower() or 
                        search_term in item.lagerplatz.lower() or
                        search_term in item.zusatzinfos.lower() or
                        search_term in item.hersteller_teilenummer.lower())
            
            # Filter
            category_match = not category_filter or item.kategorie == category_filter
            manufacturer_match = not manufacturer_filter or item.hersteller == manufacturer_filter
            location_match = not location_filter or item.lagerplatz == location_filter
            
            if text_match and category_match and manufacturer_match and location_match:
                filtered_items.append(item)
        
        self._load_table_data(filtered_items)
        
        # Sortierung wieder aktivieren
        self.storage_table.setSortingEnabled(True)

    
    def add_item(self):
        dlg = StorageItemDialog(self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            new_item = dlg.get_item_data()
            self.storage_manager.add_item(new_item)
            self._load_table_data()
            self._update_filter_combos()
    
    def edit_item(self):
        current_row = self.storage_table.currentRow()
        if current_row >= 0:
            teilenummer_item = self.storage_table.item(current_row, 0)
            if teilenummer_item:
                teilenummer = teilenummer_item.text()
                for i, item in enumerate(self.storage_manager.items):
                    if item.teilenummer == teilenummer:
                        dlg = StorageItemDialog(self, item)
                        if dlg.exec_() == QtWidgets.QDialog.Accepted:
                            updated_item = dlg.get_item_data()
                            self.storage_manager.update_item(i, updated_item)
                            self.filter_items()
                            self._update_filter_combos()
                        break

    def remove_item(self):
        current_row = self.storage_table.currentRow()
        if current_row >= 0:
            reply = QtWidgets.QMessageBox.question(self, "Löschen", 
                                                "Teil wirklich löschen?",
                                                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if reply == QtWidgets.QMessageBox.Yes:
                teilenummer_item = self.storage_table.item(current_row, 0)
                if teilenummer_item:
                    teilenummer = teilenummer_item.text()
                    # Item im StorageManager
                    for i, item in enumerate(self.storage_manager.items):
                        if item.teilenummer == teilenummer:
                            self.storage_manager.remove_item(i)
                            # Filter immer beibehalten
                            self.filter_items()
                            self._update_filter_combos()
                            break

    def increase_stock(self):
        self._change_stock(1)

    def decrease_stock(self):
        self._change_stock(-1)

    def _change_stock(self, change):
        current_row = self.storage_table.currentRow()
        if current_row >= 0:
            teilenummer_item = self.storage_table.item(current_row, 0)
            if teilenummer_item:
                teilenummer = teilenummer_item.text()
                for i, item in enumerate(self.storage_manager.items):
                    if item.teilenummer == teilenummer:
                        new_amount = item.anzahl + change
                        if new_amount < 0:
                            QtWidgets.QMessageBox.warning(self, "Fehler", "Bestand kann nicht negativ sein.")
                            return
                        
                        item.anzahl = new_amount
                        item.geaendert_am = datetime.now().strftime("%Y-%m-%d")
                        self.storage_manager.update_item(i, item)
                        self.filter_items()
                        break

    def _get_current_filtered_items(self):
        search_term = self.search_edit.text().lower()
        category_filter = self.category_combo.currentData()
        manufacturer_filter = self.manufacturer_combo.currentData()
        location_filter = self.location_combo.currentData()
        
        filtered_items = []
        for item in self.storage_manager.items:
            text_match = (not search_term or 
                        search_term in item.teilenummer.lower() or 
                        search_term in item.hersteller.lower() or 
                        search_term in item.bezeichnung.lower() or 
                        search_term in item.kategorie.lower() or 
                        search_term in item.lagerplatz.lower() or
                        search_term in item.zusatzinfos.lower() or
                        search_term in item.hersteller_teilenummer.lower())
            
            # Filter
            category_match = not category_filter or item.kategorie == category_filter
            manufacturer_match = not manufacturer_filter or item.hersteller == manufacturer_filter
            location_match = not location_filter or item.lagerplatz == location_filter
            
            if text_match and category_match and manufacturer_match and location_match:
                filtered_items.append(item)
        
        return filtered_items

    # Export als html
    def export_storage(self):
        if not self.storage_manager.items:
            QtWidgets.QMessageBox.warning(self, "Fehler", "Keine Teile zum Exportieren vorhanden.")
            return
            
        fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Lagerbestand exportieren", 
                                                        str(STORAGE_DIR / "lagerbestand_export.html"), 
                                                        "HTML Dateien (*.html)")
        if fname:
            html = self._generate_export_html()
            with open(fname, 'w', encoding='utf-8') as f:
                f.write(html)
            QtWidgets.QMessageBox.information(self, "Erfolg", f"Lagerbestand exportiert: {fname}")

        # Html export für lager

    def _generate_export_html(self):
        # Daten für Filter
        categories = sorted(set(item.kategorie for item in self.storage_manager.items if item.kategorie))
        manufacturers = sorted(set(item.hersteller for item in self.storage_manager.items if item.hersteller))
        locations = sorted(set(item.lagerplatz for item in self.storage_manager.items if item.lagerplatz))
        
        # Zusammenfassungsdaten
        total_items = sum(item.anzahl for item in self.storage_manager.items)
        low_stock_count = sum(1 for item in self.storage_manager.items if item.anzahl <= item.mindestbestand)
        total_parts = len(self.storage_manager.items)
        
        html = f"""
        <!-- o3Measurement {o3VERSION} by openw3rk INVENT -->

        <html>
        <head>
            <meta charset='utf-8'>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                .header {{ background: #2c3e50; color: white; padding: 20px; border-radius: 5px; }}
                .filter-bar {{ 
                    background: #ecf0f1; 
                    padding: 15px; 
                    border-radius: 5px; 
                    margin: 20px 0;
                    display: flex;
                    flex-wrap: wrap;
                    gap: 15px;
                    align-items: end;
                }}
                .filter-group {{
                    display: flex;
                    flex-direction: column;
                    min-width: 150px;
                }}
                .filter-group label {{
                    font-weight: bold;
                    margin-bottom: 5px;
                    color: #2c3e50;
                }}
                .filter-input {{
                    padding: 8px;
                    border: 1px solid #bdc3c7;
                    border-radius: 4px;
                }}
                .stats {{
                    display: flex;
                    justify-content: space-between;
                    margin-bottom: 15px;
                    padding: 10px;
                    background: #34495e;
                    color: white;
                    border-radius: 4px;
                }}
                .table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                .table th, .table td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                .table th {{ 
                    background-color: #f2f2f2; 
                    cursor: pointer;
                    position: relative;
                }}
                .table th:hover {{ background-color: #e8e8e8; }}
                .sort-indicator {{ margin-left: 5px; }}
                .low-stock {{ background-color: #f8d7da; }}
                .table-row {{ transition: background-color 0.2s; }}
                .table-row:hover {{ background-color: #f5f5f5; }}
                .summary {{ margin-top: 20px; padding: 15px; background: #e9ecef; border-radius: 5px; }}
                .hidden {{ display: none; }}
                .results-info {{ 
                    margin: 10px 0; 
                    padding: 8px; 
                    background: #d4edda; 
                    border-radius: 4px;
                    color: #155724;
                }}
            </style>
        </head>
        <body>
            <div class='header'>
                <h1>Lagerbestand - Export</h1>
                <p>Erstellt am: {datetime.now().strftime("%d.%m.%Y %H:%M")} | Gesamt: {total_parts} im Bestand</p>
            </div>
            
            <!-- Filter- und Suchleiste -->
            <div class="filter-bar">
                <div class="filter-group">
                    <label for="searchInput">Suche:</label>
                    <input type="text" id="searchInput" class="filter-input" 
                        placeholder="Suchen...">
                </div>
                
                <div class="filter-group">
                    <label for="categoryFilter">Kategorie:</label>
                    <select id="categoryFilter" class="filter-input">
                        <option value="">Alle Kategorien</option>
        """

        for category in categories:
            html += f'<option value="{category}">{category}</option>'
        
        html += """
                    </select>
                </div>
                
                <div class="filter-group">
                    <label for="manufacturerFilter">Hersteller:</label>
                    <select id="manufacturerFilter" class="filter-input">
                        <option value="">Alle Hersteller</option>
        """

        for manufacturer in manufacturers:
            html += f'<option value="{manufacturer}">{manufacturer}</option>'
        
        html += """
                    </select>
                </div>
                
                <div class="filter-group">
                    <label for="locationFilter">Lagerplatz:</label>
                    <select id="locationFilter" class="filter-input">
                        <option value="">Alle Lagerplätze</option>
        """
        
        # Lagerplatz
        for location in locations:
            html += f'<option value="{location}">{location}</option>'
        
        html += f"""
                    </select>
                </div>
                
                <div class="filter-group">
                    <label for="stockFilter">Bestand:</label>
                    <select id="stockFilter" class="filter-input">
                        <option value="">Alle</option>
                        <option value="low">Niedriger Bestand</option>
                        <option value="normal">Normaler Bestand</option>
                    </select>
                </div>
            </div>
            
            <div class="stats">
                <div>Gesamt: <strong>{total_parts}</strong> Einträge</div>
                <div>Gesamtstückzahl: <strong>{total_items}</strong></div>
                <div>Niedriger Bestand: <strong>{low_stock_count}</strong></div>
            </div>
            
            <div id="resultsInfo" class="results-info">
                Zeige Bestand: {total_parts} Einträge
            </div>
            
            <table class='table' id="storageTable">
                <thead>
                    <tr>
                        <th onclick="sortTable(0)">Teilenummer <span class="sort-indicator">↕</span></th>
                        <th onclick="sortTable(1)">Hersteller <span class="sort-indicator">↕</span></th>
                        <th onclick="sortTable(2)">Bezeichnung <span class="sort-indicator">↕</span></th>
                        <th onclick="sortTable(3)">Kategorie <span class="sort-indicator">↕</span></th>
                        <th onclick="sortTable(4)">Lagerplatz <span class="sort-indicator">↕</span></th>
                        <th onclick="sortTable(5)">Fach <span class="sort-indicator">↕</span></th>
                        <th onclick="sortTable(6)">Anzahl <span class="sort-indicator">↕</span></th>
                        <th onclick="sortTable(7)">Mindestbestand <span class="sort-indicator">↕</span></th>
                        <th onclick="sortTable(8)">Einheit <span class="sort-indicator">↕</span></th>
                        <th onclick="sortTable(9)">Hersteller-TNr <span class="sort-indicator">↕</span></th>
                        <th onclick="sortTable(10)">Zusatzinfos <span class="sort-indicator">↕</span></th>
                    </tr>
                </thead>
                <tbody>
        """
        
        # Tabellendaten
        for item in self.storage_manager.items:
            row_class = "low-stock" if item.anzahl <= item.mindestbestand else ""
            
            html += f"""
                    <tr class='table-row {row_class}' 
                        data-teilenummer="{item.teilenummer}"
                        data-hersteller="{item.hersteller}"
                        data-bezeichnung="{item.bezeichnung}"
                        data-kategorie="{item.kategorie}"
                        data-lagerplatz="{item.lagerplatz}"
                        data-fach="{item.fach}"
                        data-anzahl="{item.anzahl}"
                        data-mindestbestand="{item.mindestbestand}"
                        data-einheit="{item.einheit}"
                        data-hersteller-tnr="{item.hersteller_teilenummer}"
                        data-zusatzinfos="{item.zusatzinfos}">
                        <td>{item.teilenummer}</td>
                        <td>{item.hersteller}</td>
                        <td>{item.bezeichnung}</td>
                        <td>{item.kategorie}</td>
                        <td>{item.lagerplatz}</td>
                        <td>{item.fach}</td>
                        <td>{item.anzahl}</td>
                        <td>{item.mindestbestand}</td>
                        <td>{item.einheit}</td>
                        <td>{item.hersteller_teilenummer}</td>
                        <td>{item.zusatzinfos}</td>
                    </tr>
            """
        
        html += f"""
                </tbody>
            </table>
            
            <div class='summary'>
                <h3>Zusammenfassung</h3>
                <p><strong>Gesamtanzahl Teile:</strong> {total_parts}</p>
                <p><strong>Gesamtstückzahl:</strong> {total_items}</p>
                <p><strong>Einträge mit niedrigem Bestand:</strong> {low_stock_count}</p>
                <p><strong>Lagerplätze:</strong> {len(locations)}</p>
                <p><strong>Hersteller:</strong> {len(manufacturers)}</p>
                <p><strong>Kategorien:</strong> {len(categories)}</p>
            </div>
            
            <div style='margin-top: 20px; color: #666;'>
                <p>Erstellt mit:<br>{o3NAME} {o3VERSION}<br>{o3COPYRIGHT}</p>
            </div>

            <script>
                let currentSortColumn = -1;
                let sortDirection = 1;
                const totalParts = {total_parts};
                
                // Filter
                function filterTable() {{
                    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
                    const categoryFilter = document.getElementById('categoryFilter').value;
                    const manufacturerFilter = document.getElementById('manufacturerFilter').value;
                    const locationFilter = document.getElementById('locationFilter').value;
                    const stockFilter = document.getElementById('stockFilter').value;
                    
                    const rows = document.querySelectorAll('#storageTable tbody tr');
                    let visibleCount = 0;
                    
                    rows.forEach(row => {{
                        const teilenummer = row.getAttribute('data-teilenummer').toLowerCase();
                        const hersteller = row.getAttribute('data-hersteller').toLowerCase();
                        const bezeichnung = row.getAttribute('data-bezeichnung').toLowerCase();
                        const kategorie = row.getAttribute('data-kategorie');
                        const lagerplatz = row.getAttribute('data-lagerplatz');
                        const anzahl = parseInt(row.getAttribute('data-anzahl'));
                        const mindestbestand = parseInt(row.getAttribute('data-mindestbestand'));
                        
                        // Text suche
                        const textMatch = !searchTerm || 
                            teilenummer.includes(searchTerm) || 
                            hersteller.includes(searchTerm) || 
                            bezeichnung.includes(searchTerm);
                        
                        // Filter
                        const categoryMatch = !categoryFilter || kategorie === categoryFilter;
                        const manufacturerMatch = !manufacturerFilter || hersteller === manufacturerFilter.toLowerCase();
                        const locationMatch = !locationFilter || lagerplatz === locationFilter;
                        
                        // Bestand
                        let stockMatch = true;
                        if (stockFilter === 'low') {{
                            stockMatch = anzahl <= mindestbestand;
                        }} else if (stockFilter === 'normal') {{
                            stockMatch = anzahl > mindestbestand;
                        }}
                        
                        if (textMatch && categoryMatch && manufacturerMatch && locationMatch && stockMatch) {{
                            row.classList.remove('hidden');
                            visibleCount++;
                        }} else {{
                            row.classList.add('hidden');
                        }}
                    }});
                    
                    // Ergebnisse ausgeben
                    const resultsInfo = document.getElementById('resultsInfo');
                    if (visibleCount === totalParts) {{
                        resultsInfo.textContent = 'Zeige alle ' + totalParts + ' Einträge';
                    }} else {{
                        resultsInfo.textContent = 'Zeige ' + visibleCount + ' von ' + totalParts + ' Einträgen';
                    }}
                }}
                
                // Sortierung
                function sortTable(columnIndex) {{
                    const table = document.getElementById('storageTable');
                    const tbody = table.querySelector('tbody');
                    const rows = Array.from(tbody.querySelectorAll('tr:not(.hidden)'));
                    
                    if (currentSortColumn === columnIndex) {{
                        sortDirection = -sortDirection;
                    }} else {{
                        currentSortColumn = columnIndex;
                        sortDirection = 1;
                    }}
                    
                    rows.sort((a, b) => {{
                        let aValue = a.cells[columnIndex].textContent;
                        let bValue = b.cells[columnIndex].textContent;
                        
                        if (columnIndex === 6 || columnIndex === 7) {{
                            aValue = parseInt(aValue) || 0;
                            bValue = parseInt(bValue) || 0;
                            return (aValue - bValue) * sortDirection;
                        }}
                        
                        return aValue.localeCompare(bValue) * sortDirection;
                    }});
                    
                    // Sortierte Reihenfolge
                    rows.forEach(row => tbody.appendChild(row));

                    updateSortIndicators(columnIndex);
                }}
                
                function updateSortIndicators(activeColumn) {{
                    const headers = document.querySelectorAll('#storageTable th');
                    headers.forEach((header, index) => {{
                        const indicator = header.querySelector('.sort-indicator');
                        if (index === activeColumn) {{
                            indicator.textContent = sortDirection === 1 ? ' ↑' : ' ↓';
                        }} else {{
                            indicator.textContent = ' ↕';
                        }}
                    }});
                }}
                
                // Event für Filter
                document.getElementById('searchInput').addEventListener('input', filterTable);
                document.getElementById('categoryFilter').addEventListener('change', filterTable);
                document.getElementById('manufacturerFilter').addEventListener('change', filterTable);
                document.getElementById('locationFilter').addEventListener('change', filterTable);
                document.getElementById('stockFilter').addEventListener('change', filterTable);
                
                updateSortIndicators(-1);
            </script>
        </body>
        </html>
        """
        return html

class StorageItemDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, item_data=None):
        super().__init__(parent)
        self.item_data = item_data if item_data else StorageItem()
        
        title = "Teil bearbeiten" if item_data else "Neues Teil"
        self.setWindowTitle(title)
        self.setMinimumSize(600, 500)
        
        layout = QtWidgets.QFormLayout(self)
        
        # Teilenummer
        self.part_number_edit = QtWidgets.QLineEdit(self.item_data.teilenummer)
        self.part_number_edit.setPlaceholderText("Interne Teilenummer")
        layout.addRow("Teilenummer*:", self.part_number_edit)
        
        # Hersteller teilenummer
        self.manufacturer_part_edit = QtWidgets.QLineEdit(self.item_data.hersteller_teilenummer)
        self.manufacturer_part_edit.setPlaceholderText("Hersteller-Teilenummer")
        layout.addRow("Hersteller-TNr:", self.manufacturer_part_edit)
        
        # Hersteller
        self.manufacturer_edit = QtWidgets.QLineEdit(self.item_data.hersteller)
        self.manufacturer_edit.setPlaceholderText("Hersteller Name")
        layout.addRow("Hersteller*:", self.manufacturer_edit)
        
        # Bezeichnung
        self.description_edit = QtWidgets.QLineEdit(self.item_data.bezeichnung)
        self.description_edit.setPlaceholderText("Bezeichnung des Teils")
        layout.addRow("Bezeichnung*:", self.description_edit)
        
        # Kategorie
        self.category_edit = QtWidgets.QLineEdit(self.item_data.kategorie)
        self.category_edit.setPlaceholderText("z.B. Motor, Getriebe, Elektrik...")
        layout.addRow("Kategorie:", self.category_edit)
        
        # Lagerplatz und Fach
        location_layout = QtWidgets.QHBoxLayout()
        self.location_edit = QtWidgets.QLineEdit(self.item_data.lagerplatz)
        self.location_edit.setPlaceholderText("Lagerplatz")
        location_layout.addWidget(self.location_edit)
        
        self.compartment_edit = QtWidgets.QLineEdit(self.item_data.fach)
        self.compartment_edit.setPlaceholderText("Fach")
        location_layout.addWidget(self.compartment_edit)
        layout.addRow("Lagerplatz / Fach:", location_layout)
        
        # Anzahl und Mindestbestand
        stock_layout = QtWidgets.QHBoxLayout()
        self.quantity_spin = QtWidgets.QSpinBox()
        self.quantity_spin.setRange(0, 1000000)
        self.quantity_spin.setValue(self.item_data.anzahl)
        stock_layout.addWidget(self.quantity_spin)
        
        self.min_stock_spin = QtWidgets.QSpinBox()
        self.min_stock_spin.setRange(0, 1000000)
        self.min_stock_spin.setValue(self.item_data.mindestbestand)
        stock_layout.addWidget(self.min_stock_spin)
        
        self.unit_combo = QtWidgets.QComboBox()
        self.unit_combo.addItems(["Stk", "m", "kg", "l", "Paar", "Set"])
        self.unit_combo.setEditable(True)
        self.unit_combo.setCurrentText(self.item_data.einheit if self.item_data.einheit else "Stk")
        stock_layout.addWidget(self.unit_combo)
        
        layout.addRow("Anzahl / Mindestbestand / Einheit:", stock_layout)
        
        # Zusatzinfos
        self.notes_edit = QtWidgets.QTextEdit()
        self.notes_edit.setPlainText(self.item_data.zusatzinfos)
        self.notes_edit.setMaximumHeight(100)
        self.notes_edit.setPlaceholderText("Weitere Informationen, Kompatibilität, Anmerkung...")
        layout.addRow("Zusatzinfos:", self.notes_edit)
        
        # Buttons
        box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        layout.addRow(box)
    
    def get_item_data(self):
        item = StorageItem()
        item.teilenummer = self.part_number_edit.text().strip()
        item.hersteller = self.manufacturer_edit.text().strip()
        item.bezeichnung = self.description_edit.text().strip()
        item.kategorie = self.category_edit.text().strip()
        item.lagerplatz = self.location_edit.text().strip()
        item.fach = self.compartment_edit.text().strip()
        item.anzahl = self.quantity_spin.value()
        item.mindestbestand = self.min_stock_spin.value()
        item.einheit = self.unit_combo.currentText()
        item.zusatzinfos = self.notes_edit.toPlainText().strip()
        item.hersteller_teilenummer = self.manufacturer_part_edit.text().strip()
        item.geaendert_am = datetime.now().strftime("%Y-%m-%d")
        
        return item
    
    def accept(self):
        if not self.part_number_edit.text().strip():
            QtWidgets.QMessageBox.warning(self, "Fehler", "Bitte eine Teilenummer eingeben.")
            return
        if not self.manufacturer_edit.text().strip():
            QtWidgets.QMessageBox.warning(self, "Fehler", "Bitte einen Hersteller eingeben.")
            return
        if not self.description_edit.text().strip():
            QtWidgets.QMessageBox.warning(self, "Fehler", "Bitte eine Bezeichnung eingeben.")
            return
            
        super().accept()


class hu_auManagerDialog(QtWidgets.QDialog):
    def __init__(self, vehicles, parent=None):
        super().__init__(parent)
        self.vehicles = vehicles
        self.setWindowTitle("HU/AU Verwaltung - Alle Fahrzeuge")
        
        screen = QtWidgets.QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        self.setMinimumSize(700, 600)
        self.resize(int(screen_geometry.width() * 0.9), int(screen_geometry.height() * 0.85))
        
        layout = QtWidgets.QVBoxLayout(self)
        
        # Erweiterte Filter-Leiste
        filter_group = QtWidgets.QGroupBox("Filter & Suche")
        filter_layout = QtWidgets.QHBoxLayout()
        
        # Suchfeld
        filter_layout.addWidget(QtWidgets.QLabel("Suche:"))
        self.search_edit = QtWidgets.QLineEdit()
        self.search_edit.setPlaceholderText("Fahrzeugname, Anmerkung...")
        self.search_edit.textChanged.connect(self.filter_hu_au_entries)
        filter_layout.addWidget(self.search_edit)
        
        # Status-Filter
        filter_layout.addWidget(QtWidgets.QLabel("Status:"))
        self.filter_combo = QtWidgets.QComboBox()
        self.filter_combo.addItems([
            "Alle anzeigen", 
            "Überfällig", 
            "Diesen Monat fällig", 
            "Nächsten Monat fällig",
            "In 3 Monaten fällig",
            "Erledigt",
            "Nachprüfung anstehend"
        ])
        self.filter_combo.currentTextChanged.connect(self.filter_hu_au_entries)
        filter_layout.addWidget(self.filter_combo)
        
        filter_layout.addStretch()
        
        # Aktualisieren Button
        refresh_btn = QtWidgets.QPushButton("Aktualisieren")
        refresh_btn.clicked.connect(self.load_hu_au_data)
        filter_layout.addWidget(refresh_btn)
        
        # Auto-Update Status Button
        auto_update_btn = QtWidgets.QPushButton("Status automatisch berechnen")
        auto_update_btn.clicked.connect(self.auto_update_status)
        auto_update_btn.setToolTip("Berechnet automatisch den Status basierend auf den aktuellen Daten")
        filter_layout.addWidget(auto_update_btn)
        
        filter_group.setLayout(filter_layout)
        layout.addWidget(filter_group)
        
        # Tabelle für HU/AU Übersicht
        self.hu_au_table = QtWidgets.QTableWidget()
        self.hu_au_table.setColumnCount(9)
        self.hu_au_table.setHorizontalHeaderLabels([
            "Fahrzeug", 
            "HU fällig", 
            "AU fällig", 
            "Status", 
            "Nächste Prüfung",
            "Letzte HU",
            "Letzte AU", 
            "Nachprüfung nötig",
            "Anmerkung"
        ])
        self.hu_au_table.setAlternatingRowColors(False)
        self.hu_au_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.hu_au_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.hu_au_table.doubleClicked.connect(self.edit_hu_au_entry)
        self.hu_au_table.setSortingEnabled(True)
        
        # Spaltenbreiten
        self.hu_au_table.setColumnWidth(0, 150)  # Fahrzeug
        self.hu_au_table.setColumnWidth(1, 100)  # HU fällig
        self.hu_au_table.setColumnWidth(2, 100)  # AU fällig
        self.hu_au_table.setColumnWidth(3, 120)  # Status
        self.hu_au_table.setColumnWidth(4, 120)  # Nächste Prüfung
        self.hu_au_table.setColumnWidth(5, 100)  # Letzte HU
        self.hu_au_table.setColumnWidth(6, 100)  # Letzte AU
        self.hu_au_table.setColumnWidth(7, 120)  # Nachprüfung
        self.hu_au_table.setColumnWidth(8, 200)  # Anmerkung
        
        layout.addWidget(self.hu_au_table, 1)
        
        # Statistik-Leiste
        stats_layout = QtWidgets.QHBoxLayout()
        
        self.stats_label = QtWidgets.QLabel("Lade Daten...")
        self.stats_label.setStyleSheet("color: #888; font-weight: bold; padding: 5px;")
        stats_layout.addWidget(self.stats_label)
        stats_layout.addStretch()
        
        layout.addLayout(stats_layout)
        
        # Buttons
        button_layout = QtWidgets.QHBoxLayout()
        
        add_btn = QtWidgets.QPushButton("+ Neuer Eintrag")
        add_btn.clicked.connect(self.add_hu_au_entry)
        edit_btn = QtWidgets.QPushButton("Bearbeiten")
        edit_btn.clicked.connect(self.edit_hu_au_entry)
        remove_btn = QtWidgets.QPushButton("- Löschen")
        remove_btn.clicked.connect(self.remove_hu_au_entry)
        export_btn = QtWidgets.QPushButton("HTML Export")
        export_btn.clicked.connect(self.export_hu_au_overview)
        
        # Nachprüfung Buttons
        retest_btn = QtWidgets.QPushButton("Als Nachprüfung markieren")
        retest_btn.clicked.connect(self.mark_as_retest)
        retest_btn.setToolTip("Markiert ausgewählten Eintrag als nicht bestanden und setzt Nachprüfungstermin")
        
        button_layout.addWidget(add_btn)
        button_layout.addWidget(edit_btn)
        button_layout.addWidget(remove_btn)
        button_layout.addWidget(retest_btn)
        button_layout.addWidget(export_btn)
        button_layout.addStretch()
        
        layout.addLayout(button_layout)
        
        # Schließen-Button
        close_btn = QtWidgets.QPushButton("Schließen")
        close_btn.clicked.connect(self.reject)
        close_btn.setMinimumHeight(35)
        layout.addWidget(close_btn)
        
        # Daten laden
        self.load_hu_au_data()
    
    def load_hu_au_data(self):
        self.hu_au_entries = []
        
        for vehicle in self.vehicles:
            vehicle_file = vehicle.get_vehicle_dir() / "vehicle.json"
            if vehicle_file.exists():
                try:
                    with open(vehicle_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        hu_au_data = data.get('hu_au_data', {})
                        
                        entry = {
                            'vehicle': vehicle.name,
                            'vehicle_obj': vehicle,
                            'hu_au_due': hu_au_data.get('hu_au_due', ''),
                            'au_due': hu_au_data.get('au_due', ''),
                            'status': hu_au_data.get('status', 'Offen'),
                            'next_check': hu_au_data.get('next_check', ''),
                            'notes': hu_au_data.get('notes', ''),
                            'last_hu_au': hu_au_data.get('last_hu_au', ''),
                            'last_au': hu_au_data.get('last_au', ''),
                            'needs_retest': hu_au_data.get('needs_retest', False),
                            'retest_date': hu_au_data.get('retest_date', ''),
                            'original_due_date': hu_au_data.get('original_due_date', '')
                        }
                        self.hu_au_entries.append(entry)
                        
                except Exception as e:
                    print(f"Fehler beim Laden von HU/AU Daten für {vehicle.name}: {e}")
        
        # Automatische Status-Berechnung
        self.auto_update_status()
        self.update_table()
    
    def auto_update_status(self):
        current_date = datetime.now().date()
        
        for entry in self.hu_au_entries:
            hu_au_due = self.parse_date(entry['hu_au_due'])
            au_due = self.parse_date(entry['au_due'])
            retest_date = self.parse_date(entry['retest_date'])
            
            # Prüfe zuerst Nachprüfungen
            if entry['needs_retest'] and retest_date:
                if retest_date < current_date:
                    entry['status'] = 'Nachprüfung überfällig'
                elif retest_date.month == current_date.month and retest_date.year == current_date.year:
                    entry['status'] = 'Nachprüfung diesen Monat'
                elif (retest_date.month == current_date.month + 1 and retest_date.year == current_date.year) or \
                     (retest_date.month == 1 and current_date.month == 12 and retest_date.year == current_date.year + 1):
                    entry['status'] = 'Nachprüfung nächsten Monat'
                else:
                    entry['status'] = 'Nachprüfung anstehend'
            
            # Normale AU prüfungen
            elif hu_au_due or au_due:
                due_date = hu_au_due or au_due 
                if au_due and hu_au_due:
                    due_date = min(hu_au_due, au_due)
                
                if due_date < current_date:
                    entry['status'] = 'Überfällig'
                elif due_date.month == current_date.month and due_date.year == current_date.year:
                    entry['status'] = 'Diesen Monat fällig'
                elif (due_date.month == current_date.month + 1 and due_date.year == current_date.year) or \
                     (due_date.month == 1 and current_date.month == 12 and due_date.year == current_date.year + 1):
                    entry['status'] = 'Nächsten Monat fällig'
                elif due_date <= current_date + timedelta(days=90):  # 3 Monate
                    entry['status'] = 'In 3 Monaten fällig'
                else:
                    entry['status'] = 'Zukunft'
            
            else:
                entry['status'] = 'Keine Daten'
        
        self.update_stats()
    
    def update_table(self):
        filtered_entries = self.get_filtered_entries()
        
        self.hu_au_table.setRowCount(len(filtered_entries))
        
        for row, entry in enumerate(filtered_entries):
            self.hu_au_table.setItem(row, 0, QtWidgets.QTableWidgetItem(entry['vehicle']))
            self.hu_au_table.setItem(row, 1, QtWidgets.QTableWidgetItem(entry['hu_au_due']))
            self.hu_au_table.setItem(row, 2, QtWidgets.QTableWidgetItem(entry['au_due']))
            
            status_item = QtWidgets.QTableWidgetItem(entry['status'])
            # Farbliche Kennzeichnung
            if 'überfällig' in entry['status'].lower():
                status_item.setBackground(QtGui.QColor('#3a1e1e'))  # Rot
                status_item.setForeground(QtGui.QColor('#ff6b6b'))
            elif 'diesen monat' in entry['status'].lower():
                status_item.setBackground(QtGui.QColor('#5a3a1e'))  # Orange
                status_item.setForeground(QtGui.QColor('#ffd700'))
            elif 'nächsten monat' in entry['status'].lower():
                status_item.setBackground(QtGui.QColor('#3a5a1e'))  # Gelb-Grün
                status_item.setForeground(QtGui.QColor('#e6ee9c'))
            elif 'nachprüfung' in entry['status'].lower():
                status_item.setBackground(QtGui.QColor('#1e3a5a'))  # Blau
                status_item.setForeground(QtGui.QColor('#90caf9'))
            elif entry['status'] == 'Erledigt':
                status_item.setBackground(QtGui.QColor('#1e3a1e'))  # Grün
                status_item.setForeground(QtGui.QColor('#90ee90'))
            
            self.hu_au_table.setItem(row, 3, status_item)
            self.hu_au_table.setItem(row, 4, QtWidgets.QTableWidgetItem(entry['next_check']))
            self.hu_au_table.setItem(row, 5, QtWidgets.QTableWidgetItem(entry['last_hu_au']))
            self.hu_au_table.setItem(row, 6, QtWidgets.QTableWidgetItem(entry['last_au']))
            
            # Nachprüfungs info
            retest_info = "Ja" if entry['needs_retest'] else "Nein"
            if entry['needs_retest'] and entry['retest_date']:
                retest_info = f"Ja ({entry['retest_date']})"
            retest_item = QtWidgets.QTableWidgetItem(retest_info)
            if entry['needs_retest']:
                retest_item.setBackground(QtGui.QColor('#1e3a5a'))
                retest_item.setForeground(QtGui.QColor('#90caf9'))
            self.hu_au_table.setItem(row, 7, retest_item)
            
            self.hu_au_table.setItem(row, 8, QtWidgets.QTableWidgetItem(entry['notes']))
        
        self.update_stats()
    
    def get_filtered_entries(self):
        filter_text = self.filter_combo.currentText()
        search_text = self.search_edit.text().lower()
        
        filtered = []
        for entry in self.hu_au_entries:
            # Text-Suche
            text_match = (not search_text or 
                         search_text in entry['vehicle'].lower() or 
                         search_text in entry['notes'].lower() or 
                         search_text in entry['status'].lower())
            
            # Status-Filter
            status_match = True
            if filter_text == "Überfällig":
                status_match = 'überfällig' in entry['status'].lower()
            elif filter_text == "Diesen Monat fällig":
                status_match = 'diesen monat' in entry['status'].lower()
            elif filter_text == "Nächsten Monat fällig":
                status_match = 'nächsten monat' in entry['status'].lower()
            elif filter_text == "In 3 Monaten fällig":
                status_match = entry['status'] == 'In 3 Monaten fällig'
            elif filter_text == "Erledigt":
                status_match = entry['status'] == 'Erledigt'
            elif filter_text == "Nachprüfung anstehend":
                status_match = 'nachprüfung' in entry['status'].lower()
            
            if text_match and status_match:
                filtered.append(entry)
        
        return filtered
    
    def filter_hu_au_entries(self):
        self.update_table()
    
    def update_stats(self):
        filtered_entries = self.get_filtered_entries()
        total_entries = len(self.hu_au_entries)
        
        stats = {
            'überfällig': 0,
            'diesen_monat': 0,
            'nächsten_monat': 0,
            'nachprüfung': 0,
            'erledigt': 0
        }
        
        for entry in self.hu_au_entries:
            if 'überfällig' in entry['status'].lower():
                stats['überfällig'] += 1
            elif 'diesen monat' in entry['status'].lower():
                stats['diesen_monat'] += 1
            elif 'nächsten monat' in entry['status'].lower():
                stats['nächsten_monat'] += 1
            elif 'nachprüfung' in entry['status'].lower():
                stats['nachprüfung'] += 1
            elif entry['status'] == 'Erledigt':
                stats['erledigt'] += 1
        
        self.stats_label.setText(
            f"Gesamt: {total_entries} | "
            f"Überfällig: {stats['überfällig']} | "
            f"Diesen Monat: {stats['diesen_monat']} | "
            f"Nächsten Monat: {stats['nächsten_monat']} | "
            f"Nachprüfungen: {stats['nachprüfung']} | "
            f"Erledigt: {stats['erledigt']} | "
            f"Gefiltert: {len(filtered_entries)}"
        )
    
    def parse_date(self, date_str):
        if not date_str:
            return None
        try:
            return datetime.strptime(date_str, "%Y-%m-%d").date()
        except:
            return None
    
    def add_hu_au_entry(self):
        dlg = hu_auEntryDialog(self.vehicles, self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            new_entry = dlg.get_hu_au_data()
            self.save_hu_au_entry(new_entry)
            self.load_hu_au_data()
    
    def edit_hu_au_entry(self):
        current_row = self.hu_au_table.currentRow()
        if current_row >= 0:
            filtered_entries = self.get_filtered_entries()
            if current_row < len(filtered_entries):
                entry = filtered_entries[current_row]
                dlg = hu_auEntryDialog(self.vehicles, self, entry)
                if dlg.exec_() == QtWidgets.QDialog.Accepted:
                    updated_entry = dlg.get_hu_au_data()
                    self.save_hu_au_entry(updated_entry)
                    self.load_hu_au_data()
    
    def mark_as_retest(self):
        current_row = self.hu_au_table.currentRow()
        if current_row >= 0:
            filtered_entries = self.get_filtered_entries()
            if current_row < len(filtered_entries):
                entry = filtered_entries[current_row]
                
                # Dialog für Nachprüfungsdetails
                dlg = RetestDialog(entry, self)
                if dlg.exec_() == QtWidgets.QDialog.Accepted:
                    retest_data = dlg.get_retest_data()
                    
                    # Aktualisiere den Eintrag mit Nachprüfungsdaten
                    entry.update(retest_data)
                    self.save_hu_au_entry(entry)
                    self.load_hu_au_data()
                    
                    QtWidgets.QMessageBox.information(self, "Nachprüfung", 
                                                    "Nachprüfungstermin wurde eingetragen.")
    
    def remove_hu_au_entry(self):
        current_row = self.hu_au_table.currentRow()
        if current_row >= 0:
            reply = QtWidgets.QMessageBox.question(self, "Löschen", 
                                                 "Eintrag wirklich löschen?",
                                                 QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if reply == QtWidgets.QMessageBox.Yes:
                filtered_entries = self.get_filtered_entries()
                if current_row < len(filtered_entries):
                    entry = filtered_entries[current_row]
                    self.save_hu_au_entry({
                        'vehicle': entry['vehicle'],
                        'hu_au_due': '',
                        'au_due': '',
                        'status': 'Offen',
                        'next_check': '',
                        'notes': '',
                        'last_hu_au': '',
                        'last_au': '',
                        'needs_retest': False,
                        'retest_date': '',
                        'original_due_date': ''
                    })
                    self.load_hu_au_data()
    
    def save_hu_au_entry(self, hu_au_data):
        vehicle_name = hu_au_data['vehicle']
        for vehicle in self.vehicles:
            if vehicle.name == vehicle_name:
                try:
                    vehicle_file = vehicle.get_vehicle_dir() / "vehicle.json"
                    if vehicle_file.exists():
                        with open(vehicle_file, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                        
                        data['hu_au_data'] = {
                            'hu_au_due': hu_au_data['hu_au_due'],
                            'au_due': hu_au_data['au_due'],
                            'status': hu_au_data['status'],
                            'next_check': hu_au_data['next_check'],
                            'notes': hu_au_data['notes'],
                            'last_hu_au': hu_au_data['last_hu_au'],
                            'last_au': hu_au_data['last_au'],
                            'needs_retest': hu_au_data.get('needs_retest', False),
                            'retest_date': hu_au_data.get('retest_date', ''),
                            'original_due_date': hu_au_data.get('original_due_date', '')
                        }
                        
                        with open(vehicle_file, 'w', encoding='utf-8') as f:
                            json.dump(data, f, indent=2, ensure_ascii=False)
                            
                except Exception as e:
                    QtWidgets.QMessageBox.warning(self, "Fehler", f"Fehler beim Speichern: {e}")
                break
    
    def export_hu_au_overview(self):
        if not self.hu_au_entries:
            QtWidgets.QMessageBox.warning(self, "Fehler", "Keine Daten zum Exportieren vorhanden.")
            return
            
        fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, "HU/AU Übersicht exportieren", 
                                                        str(VEHICLES_BASE_DIR / "hu_au_uebersicht.html"), 
                                                        "HTML Dateien (*.html)")
        if fname:
            html = self.generate_hu_au_html()
            with open(fname, 'w', encoding='utf-8') as f:
                f.write(html)
            QtWidgets.QMessageBox.information(self, "Erfolg", f"HU/AU Übersicht exportiert: {fname}")

# HU/AU html export
    def generate_hu_au_html(self):
        current_date = datetime.now()
        
        # Statistik berechnen
        stats = {
            'gesamt': len(self.hu_au_entries),
            'überfällig': 0,
            'diesen_monat': 0,
            'nächsten_monat': 0,
            'nachprüfung': 0,
            'erledigt': 0
        }
        
        for entry in self.hu_au_entries:
            if 'überfällig' in entry['status'].lower():
                stats['überfällig'] += 1
            elif 'diesen monat' in entry['status'].lower():
                stats['diesen_monat'] += 1
            elif 'nächsten monat' in entry['status'].lower():
                stats['nächsten_monat'] += 1
            elif 'nachprüfung' in entry['status'].lower():
                stats['nachprüfung'] += 1
            elif entry['status'] == 'Erledigt':
                stats['erledigt'] += 1

        html = f"""
        <html>
        <head>
            <meta charset='utf-8'>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                .header {{ background: #2c3e50; color: white; padding: 20px; border-radius: 5px; }}
                .table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                .table th, .table td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                .table th {{ background-color: #f2f2f2; }}
                .status-überfällig {{ background-color: #f8d7da; font-weight: bold; }}
                .status-diesen-monat {{ background-color: #fff3cd; }}
                .status-nächsten-monat {{ background-color: #e6ee9c; }}
                .status-nachprüfung {{ background-color: #d1ecf1; }}
                .status-erledigt {{ background-color: #d4edda; }}
                .summary {{ margin-top: 20px; padding: 15px; background: #e9ecef; border-radius: 5px; }}
                .stats-grid {{ 
                    display: grid; 
                    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); 
                    gap: 10px; 
                    margin-bottom: 20px; 
                }}
                .stat-item {{ 
                    padding: 15px; 
                    border-radius: 5px; 
                    text-align: center; 
                    color: white;
                    font-weight: bold;
                }}
                .stat-überfällig {{ background: #dc3545; }}
                .stat-diesen-monat {{ background: #fd7e14; }}
                .stat-nächsten-monat {{ background: #ffc107; color: #000; }}
                .stat-nachprüfung {{ background: #17a2b8; }}
                .stat-erledigt {{ background: #28a745; }}
            </style>
        </head>
        <body>
            <div class='header'>
                <h1>HU/AU Übersicht - Alle Fahrzeuge</h1>
                <p>Exportiert am: {current_date.strftime("%d.%m.%Y %H:%M")}</p>
            </div>
            
            <!-- Statistik HU/AU -->
            <div class='stats-grid'>
                <div class='stat-item stat-überfällig'>
                    <div style='font-size: 24px;'>{stats['überfällig']}</div>
                    <div>Überfällig</div>
                </div>
                <div class='stat-item stat-diesen-monat'>
                    <div style='font-size: 24px;'>{stats['diesen_monat']}</div>
                    <div>Diesen Monat</div>
                </div>
                <div class='stat-item stat-nächsten-monat'>
                    <div style='font-size: 24px;'>{stats['nächsten_monat']}</div>
                    <div>Nächsten Monat</div>
                </div>
                <div class='stat-item stat-nachprüfung'>
                    <div style='font-size: 24px;'>{stats['nachprüfung']}</div>
                    <div>Nachprüfungen</div>
                </div>
                <div class='stat-item stat-erledigt'>
                    <div style='font-size: 24px;'>{stats['erledigt']}</div>
                    <div>Erledigt</div>
                </div>
            </div>
            
            <table class='table'>
                <tr>
                    <th>Fahrzeug</th>
                    <th>HU fällig</th>
                    <th>AU fällig</th>
                    <th>Status</th>
                    <th>Nächste Prüfung</th>
                    <th>Letzte HU</th>
                    <th>Letzte AU</th>
                    <th>Nachprüfung nötig</th>
                    <th>Anmerkung</th>
                </tr>
        """
        
        for entry in self.hu_au_entries:
            # Status-Klasse für CSS
            status_class = ""
            if 'überfällig' in entry['status'].lower():
                status_class = 'status-überfällig'
            elif 'diesen monat' in entry['status'].lower():
                status_class = 'status-diesen-monat'
            elif 'nächsten monat' in entry['status'].lower():
                status_class = 'status-nächsten-monat'
            elif 'nachprüfung' in entry['status'].lower():
                status_class = 'status-nachprüfung'
            elif entry['status'] == 'Erledigt':
                status_class = 'status-erledigt'
            
            # Nachprüfungs-Info
            retest_info = "Nein"
            if entry['needs_retest']:
                if entry['retest_date']:
                    retest_info = f"Ja ({entry['retest_date']})"
                else:
                    retest_info = "Ja"
            
            html += f"""
                <tr>
                    <td>{entry['vehicle']}</td>
                    <td>{entry['hu_au_due']}</td>
                    <td>{entry['au_due']}</td>
                    <td class='{status_class}'>{entry['status']}</td>
                    <td>{entry['next_check']}</td>
                    <td>{entry['last_hu_au']}</td>
                    <td>{entry['last_au']}</td>
                    <td>{retest_info}</td>
                    <td>{entry['notes']}</td>
                </tr>
            """
        
        html += f"""
            </table>
            
            <div class='summary'>
                <h3>Zusammenfassung</h3>
                <p><strong>Gesamt Fahrzeuge:</strong> {stats['gesamt']}</p>
                <p><strong>Überfällig:</strong> {stats['überfällig']}</p>
                <p><strong>Diesen Monat fällig:</strong> {stats['diesen_monat']}</p>
                <p><strong>Nächsten Monat fällig:</strong> {stats['nächsten_monat']}</p>
                <p><strong>Nachprüfungen:</strong> {stats['nachprüfung']}</p>
                <p><strong>Erledigt:</strong> {stats['erledigt']}</p>
            </div>
            
            <div style='margin-top: 20px; color: #666;'>
                <p>Erstellt mit:<br>{o3NAME} {o3VERSION}<br>{o3COPYRIGHT}</p>
            </div>
        </body>
        </html>
        """
        return html

class RetestDialog(QtWidgets.QDialog):
    def __init__(self, entry_data, parent=None):
        super().__init__(parent)
        self.entry_data = entry_data
        self.setWindowTitle("Nachprüfung eintragen")
        self.setMinimumSize(500, 300)
        
        layout = QtWidgets.QFormLayout(self)
        
        layout.addWidget(QtWidgets.QLabel(f"<b>Nachprüfung für: {entry_data['vehicle']}</b>"))
        layout.addWidget(QtWidgets.QLabel("Die Hauptprüfung wurde nicht bestanden."))
        
        # Nachprüfungsdatum
        self.retest_date_edit = QtWidgets.QDateEdit()
        self.retest_date_edit.setCalendarPopup(True)
        self.retest_date_edit.setDate(QtCore.QDate.currentDate().addDays(14))  # Standard: 2 Wochen
        layout.addRow("Nachprüfungstermin:", self.retest_date_edit)
        
        # Ursprüngliches Datum speichern falls nicht vorhanden
        if not entry_data.get('original_due_date') and entry_data.get('hu_au_due'):
            self.original_due_edit = QtWidgets.QDateEdit()
            self.original_due_edit.setCalendarPopup(True)
            try:
                self.original_due_edit.setDate(QtCore.QDate.fromString(entry_data['hu_au_due'], "yyyy-MM-dd"))
            except:
                self.original_due_edit.setDate(QtCore.QDate.currentDate())
            self.original_due_edit.setEnabled(False)
            layout.addRow("Ursprünglicher Termin:", self.original_due_edit)
        
        # Anmerkung zur Nachprüfung
        self.retest_notes_edit = QtWidgets.QTextEdit()
        self.retest_notes_edit.setMaximumHeight(80)
        self.retest_notes_edit.setPlaceholderText("Gründe für die Nachprüfung, Mängel, etc...")
        layout.addRow("Nachprüfungs Anmerkung:", self.retest_notes_edit)
        
        # Buttons
        box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        layout.addRow(box)
    
    def get_retest_data(self):
        original_due = self.entry_data.get('original_due_date', '')
        if not original_due and self.entry_data.get('hu_au_due'):
            original_due = self.entry_data['hu_au_due']
        
        return {
            'vehicle': self.entry_data['vehicle'],
            'needs_retest': True,
            'retest_date': self.retest_date_edit.date().toString("yyyy-MM-dd"),
            'original_due_date': original_due,
            'notes': f"{self.entry_data.get('notes', '')} | Nachprüfung: {self.retest_notes_edit.toPlainText()}".strip(),
            'status': 'Nachprüfung anstehend'
        }


class hu_auEntryDialog(QtWidgets.QDialog):
    def __init__(self, vehicles, parent=None, entry_data=None):
        super().__init__(parent)
        self.vehicles = vehicles
        self.entry_data = entry_data if entry_data else {}
        
        title = "HU/AU Eintrag bearbeiten" if entry_data else "Neuer HU/AU Eintrag"
        self.setWindowTitle(title)
        self.setMinimumSize(500, 400)
        
        layout = QtWidgets.QFormLayout(self)
        
        # Fahrzeug-Auswahl
        self.vehicle_combo = QtWidgets.QComboBox()
        for vehicle in vehicles:
            self.vehicle_combo.addItem(vehicle.name)
        
        if entry_data and 'vehicle' in entry_data:
            idx = self.vehicle_combo.findText(entry_data['vehicle'])
            if idx >= 0:
                self.vehicle_combo.setCurrentIndex(idx)
        
        layout.addRow("Fahrzeug:", self.vehicle_combo)
        
        # HU Datum
        self.hu_au_date_edit = QtWidgets.QDateEdit()
        self.hu_au_date_edit.setCalendarPopup(True)
        if entry_data and entry_data.get('hu_au_due'):
            try:
                self.hu_au_date_edit.setDate(QtCore.QDate.fromString(entry_data['hu_au_due'], "yyyy-MM-dd"))
            except:
                self.hu_au_date_edit.setDate(QtCore.QDate.currentDate().addMonths(6))
        else:
            self.hu_au_date_edit.setDate(QtCore.QDate.currentDate().addMonths(6))
        layout.addRow("HU fällig am:", self.hu_au_date_edit)
        
        # AU Datum
        self.au_date_edit = QtWidgets.QDateEdit()
        self.au_date_edit.setCalendarPopup(True)
        if entry_data and entry_data.get('au_due'):
            try:
                self.au_date_edit.setDate(QtCore.QDate.fromString(entry_data['au_due'], "yyyy-MM-dd"))
            except:
                self.au_date_edit.setDate(QtCore.QDate.currentDate().addMonths(6))
        else:
            self.au_date_edit.setDate(QtCore.QDate.currentDate().addMonths(6))
        layout.addRow("AU fällig am:", self.au_date_edit)
        
        # Status
        self.status_combo = QtWidgets.QComboBox()
        self.status_combo.addItems(["Offen", "Diesen Monat", "Nächsten Monat", "Überfällig", "Erledigt"])
        if entry_data and 'status' in entry_data:
            idx = self.status_combo.findText(entry_data['status'])
            if idx >= 0:
                self.status_combo.setCurrentIndex(idx)
        layout.addRow("Status:", self.status_combo)
        
        # Nächste Prüfung
        self.next_check_edit = QtWidgets.QLineEdit(entry_data.get('next_check', '') if entry_data else '')
        layout.addRow("Nächste Prüfung:", self.next_check_edit)
        
        # Letzte HU
        self.last_hu_au_edit = QtWidgets.QLineEdit(entry_data.get('last_hu_au', '') if entry_data else '')
        layout.addRow("Letzte HU:", self.last_hu_au_edit)
        
        # Letzte AU
        self.last_au_edit = QtWidgets.QLineEdit(entry_data.get('last_au', '') if entry_data else '')
        layout.addRow("Letzte AU:", self.last_au_edit)
        
        # Anmerkung
        self.notes_edit = QtWidgets.QTextEdit()
        self.notes_edit.setMaximumHeight(80)
        self.notes_edit.setPlainText(entry_data.get('notes', '') if entry_data else '')
        layout.addRow("Anmerkung:", self.notes_edit)
        
        # Buttons
        box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        layout.addRow(box)
    
    def get_hu_au_data(self):
        return {
            'vehicle': self.vehicle_combo.currentText(),
            'hu_au_due': self.hu_au_date_edit.date().toString("yyyy-MM-dd"),
            'au_due': self.au_date_edit.date().toString("yyyy-MM-dd"),
            'status': self.status_combo.currentText(),
            'next_check': self.next_check_edit.text().strip(),
            'notes': self.notes_edit.toPlainText().strip(),
            'last_hu_au': self.last_hu_au_edit.text().strip(),
            'last_au': self.last_au_edit.text().strip()
        }


class ActivitiesHistoryDialog(QtWidgets.QDialog):
    def __init__(self, vehicle: Vehicle, parent=None):
        super().__init__(parent)
        self.vehicle = vehicle
        self.setWindowTitle(f"Aktivitäten Historie - {vehicle.name}")
        
        # Autoscaling
        screen = QtWidgets.QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        self.setMinimumSize(800, 600)
        self.resize(int(screen_geometry.width() * 0.9), int(screen_geometry.height() * 0.8))
        
        layout = QtWidgets.QVBoxLayout(self)
        
        # Such- und Filter-Leiste
        filter_group = QtWidgets.QGroupBox("Filter & Suche")
        filter_layout = QtWidgets.QHBoxLayout()
        
        filter_layout.addWidget(QtWidgets.QLabel("Suchen:"))
        self.search_edit = QtWidgets.QLineEdit()
        self.search_edit.setPlaceholderText("Aktivität, Beschreibung, Material...")
        self.search_edit.textChanged.connect(self.filter_activities)
        filter_layout.addWidget(self.search_edit)
        
        filter_layout.addWidget(QtWidgets.QLabel("Aktivitätstyp:"))
        self.type_combo = QtWidgets.QComboBox()
        self.type_combo.addItems(["Alle Typen", "Allgemeine Wartung", "Reparatur", "Diagnose", "Inspektion", "Ölwechsel", "Bremsenarbeiten", "Elektrikarbeiten", "Motorarbeiten", "Getriebearbeiten", "Fahrwerksarbeiten", "Karosseriearbeiten", "Sonstiges"])
        self.type_combo.currentTextChanged.connect(self.filter_activities)
        filter_layout.addWidget(self.type_combo)
        
        filter_layout.addStretch()
        
        # Statistik-Label
        self.stats_label = QtWidgets.QLabel("Lade Aktivitäten...")
        self.stats_label.setStyleSheet("color: #888; font-weight: bold;")
        filter_layout.addWidget(self.stats_label)
        
        filter_group.setLayout(filter_layout)
        layout.addWidget(filter_group)
        
        # Aktivitäten-Tabelle
        self.activities_table = QtWidgets.QTableWidget()
        self.activities_table.setColumnCount(6)
        self.activities_table.setHorizontalHeaderLabels(["Datum", "Aktivitätstyp", "KM-Stand", "Beschreibung", "Materialien", "Durchgeführt durch"])
        self.activities_table.setAlternatingRowColors(False)
        self.activities_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.activities_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.activities_table.doubleClicked.connect(self.show_activity_details)
        self.activities_table.setSortingEnabled(True)
        
        # Spaltenbreiten
        self.activities_table.setColumnWidth(0, 100)  # Datum
        self.activities_table.setColumnWidth(1, 120)  # Typ
        self.activities_table.setColumnWidth(2, 80)   # KM
        self.activities_table.setColumnWidth(3, 250)  # Beschreibung
        self.activities_table.setColumnWidth(4, 150)  # Materialien
        self.activities_table.setColumnWidth(5, 120)  # Durchgeführt
        
        layout.addWidget(self.activities_table, 1)
        
        # Buttons
        button_layout = QtWidgets.QHBoxLayout()
        
        view_btn = QtWidgets.QPushButton("Details anzeigen")
        view_btn.clicked.connect(self.show_activity_details)
        
        export_btn = QtWidgets.QPushButton("HTML Export")
        export_btn.clicked.connect(self.export_activities_html)
        
        refresh_btn = QtWidgets.QPushButton("Aktualisieren")
        refresh_btn.clicked.connect(self.load_activities)
        
        button_layout.addWidget(view_btn)
        button_layout.addWidget(export_btn)
        button_layout.addWidget(refresh_btn)
        button_layout.addStretch()
        
        close_btn = QtWidgets.QPushButton("Schließen")
        close_btn.clicked.connect(self.reject)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        
        # Aktivitäten laden
        self.activities = []
        self.load_activities()
    
    def load_activities(self):
        self.activities = []
        activities_dir = self.vehicle.get_vehicle_dir() / "activities"
        
        if activities_dir.exists():
            for year_dir in activities_dir.iterdir():
                if year_dir.is_dir():
                    for activity_file in year_dir.glob("*.json"):
                        try:
                            with open(activity_file, 'r', encoding='utf-8') as f:
                                activity_data = json.load(f)
                                self.activities.append(activity_data)
                        except Exception as e:
                            print(f"Fehler beim Laden von {activity_file}: {e}")
        
        # Nach Datum sortieren (neueste zuerst)
        self.activities.sort(key=lambda x: x.get('datum', ''), reverse=True)
        self.update_table()
    
    def update_table(self):
        filtered_activities = self.get_filtered_activities()
        
        self.activities_table.setRowCount(len(filtered_activities))
        
        for row, activity in enumerate(filtered_activities):
            # Materialien-Text erstellen
            materialien_text = ""
            if activity.get('verwendete_materialien'):
                material_count = len(activity['verwendete_materialien'])
                material_names = [mat.get('bezeichnung', mat.get('teilenummer', '')) for mat in activity['verwendete_materialien'][:2]]
                materialien_text = f"{material_count} Materialien: {', '.join(material_names)}"
                if material_count > 2:
                    materialien_text += " ..."
            
            # Beschreibung kürzen
            beschreibung = activity.get('beschreibung', '')
            if len(beschreibung) > 80:
                beschreibung = beschreibung[:77] + "..."
            
            self.activities_table.setItem(row, 0, QtWidgets.QTableWidgetItem(activity.get('datum', '')))
            self.activities_table.setItem(row, 1, QtWidgets.QTableWidgetItem(activity.get('activity_type', '')))
            self.activities_table.setItem(row, 2, QtWidgets.QTableWidgetItem(str(activity.get('km_stand', ''))))
            self.activities_table.setItem(row, 3, QtWidgets.QTableWidgetItem(beschreibung))
            self.activities_table.setItem(row, 4, QtWidgets.QTableWidgetItem(materialien_text))
            self.activities_table.setItem(row, 5, QtWidgets.QTableWidgetItem(activity.get('erstellt_durch', '')))
        
        # Statistik aktualisieren
        total = len(self.activities)
        filtered = len(filtered_activities)
        self.stats_label.setText(f"Aktivitäten: {filtered}/{total}")
    
    def get_filtered_activities(self):
        search_text = self.search_edit.text().lower()
        type_filter = self.type_combo.currentText()
        
        filtered = []
        for activity in self.activities:
            # Text-Suche
            text_match = (not search_text or 
                         search_text in activity.get('activity_type', '').lower() or
                         search_text in activity.get('beschreibung', '').lower() or
                         search_text in activity.get('erstellt_durch', '').lower() or
                         any(search_text in mat.get('bezeichnung', '').lower() or 
                             search_text in mat.get('teilenummer', '').lower() 
                             for mat in activity.get('verwendete_materialien', [])))
            
            # Typ-Filter
            type_match = (type_filter == "Alle Typen" or 
                         activity.get('activity_type', '') == type_filter)
            
            if text_match and type_match:
                filtered.append(activity)
        
        return filtered
    
    def filter_activities(self):
        self.update_table()
    
    def show_activity_details(self):
        current_row = self.activities_table.currentRow()
        if current_row >= 0:
            filtered_activities = self.get_filtered_activities()
            if current_row < len(filtered_activities):
                activity = filtered_activities[current_row]
                dlg = ActivityDetailsDialog(activity, self)
                dlg.exec_()
    
    def export_activities_html(self):
        if not self.activities:
            QtWidgets.QMessageBox.warning(self, "Fehler", "Keine Aktivitäten zum Exportieren vorhanden.")
            return
            
        fname, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, 
            "Aktivitäten exportieren", 
            str(self.vehicle.get_vehicle_dir() / f"{self.vehicle.name}_aktivitäten.html"), 
            "HTML Dateien (*.html)"
        )
        
        if fname:
            html = self._generate_activities_html()
            with open(fname, 'w', encoding='utf-8') as f:
                f.write(html)
            QtWidgets.QMessageBox.information(self, "Erfolg", f"Aktivitäten exportiert: {fname}")
    
    def _generate_activities_html(self):
        activities_to_export = self.get_filtered_activities()
        
        html = f"""
        <!DOCTYPE html>
        <html lang="de">
        <head>
            <meta charset="utf-8">
            <title>Aktivitäten-Historie - {self.vehicle.name}</title>
            <style>
                body {{ 
                    font-family: Arial, sans-serif; 
                    margin: 20px; 
                    background-color: #0f0f12; 
                    color: #e6eef8; 
                }}
                .header {{ 
                    background: #2c3e50; 
                    color: white; 
                    padding: 20px; 
                    border-radius: 5px; 
                    margin-bottom: 20px;
                }}
                .activity {{ 
                    background: #1b1b20; 
                    border: 1px solid #2a2a33; 
                    border-radius: 5px; 
                    padding: 15px; 
                    margin-bottom: 15px;
                }}
                .activity-header {{ 
                    display: flex; 
                    justify-content: space-between; 
                    border-bottom: 1px solid #2a2a33; 
                    padding-bottom: 10px; 
                    margin-bottom: 10px;
                }}
                .activity-type {{ 
                    font-weight: bold; 
                    color: #90caf9; 
                    font-size: 18px;
                }}
                .activity-date {{ 
                    color: #9aa; 
                }}
                .material-list {{ 
                    background: #16161a; 
                    padding: 10px; 
                    border-radius: 3px; 
                    margin: 10px 0;
                }}
                .work-list {{ 
                    margin: 10px 0;
                }}
                .work-item {{ 
                    padding: 5px 0; 
                    border-bottom: 1px dashed #2a2a33;
                }}
                .stats {{ 
                    background: #34495e; 
                    padding: 15px; 
                    border-radius: 5px; 
                    margin: 20px 0;
                }}
                table {{ 
                    width: 100%; 
                    border-collapse: collapse; 
                    margin: 20px 0;
                }}
                th, td {{ 
                    border: 1px solid #2a2a33; 
                    padding: 8px; 
                    text-align: left;
                }}
                th {{ 
                    background: #2c3e50; 
                    color: white;
                }}
            </style>
        </head>
        <body>
            <div class="header">
                <h1>Aktivitäten-Historie - {self.vehicle.name}</h1>
                <p>FIN: {self.vehicle.specifications.get('fin', '')} | Exportiert am: {datetime.now().strftime("%d.%m.%Y %H:%M")}</p>
            </div>
            
            <div class="stats">
                <h3>Statistik</h3>
                <p><strong>Gesamtaktivitäten:</strong> {len(activities_to_export)}</p>
                <p><strong>Zeitraum:</strong> {activities_to_export[-1].get('datum', '') if activities_to_export else ''} bis {activities_to_export[0].get('datum', '') if activities_to_export else ''}</p>
                <p><strong>Letzter KM-Stand:</strong> {activities_to_export[0].get('km_stand', '') if activities_to_export else ''} km</p>
            </div>
        """
        
        # Aktivitäten auflisten
        for activity in activities_to_export:
            html += f"""
            <div class="activity">
                <div class="activity-header">
                    <div class="activity-type">{activity.get('activity_type', '')}</div>
                    <div class="activity-date">{activity.get('datum', '')} | {activity.get('km_stand', '')} km</div>
                </div>
                
                <p><strong>Beschreibung:</strong><br>{activity.get('beschreibung', '').replace(chr(10), '<br>')}</p>
                
                <p><strong>Durchgeführt durch:</strong> {activity.get('erstellt_durch', '')}</p>
            """
            
            # Aktivitäz anzeigen
            if activity.get('durchgeführte_arbeiten'):
                html += """
                <div class="work-list">
                    <strong>Durchgeführte Arbeiten:</strong>
                """
                for arbeit in activity['durchgeführte_arbeiten']:
                    html += f"""
                    <div class="work-item">
                        • {arbeit.get('arbeit', '')} 
                        ({arbeit.get('dauer_minuten', 0)} Min)
                        {f"- {arbeit.get('bemerkungen', '')}" if arbeit.get('bemerkungen') else ""}
                    </div>
                    """
                html += "</div>"
            
            # Materialien anzeigen
            if activity.get('verwendete_materialien'):
                html += """
                <div class="material-list">
                    <strong>Verwendete Materialien:</strong>
                    <table>
                        <tr><th>Teilenummer</th><th>Bezeichnung</th><th>Menge</th><th>Einheit</th></tr>
                """
                for material in activity['verwendete_materialien']:
                    html += f"""
                    <tr>
                        <td>{material.get('teilenummer', '')}</td>
                        <td>{material.get('bezeichnung', '')}</td>
                        <td>{material.get('menge_verbraucht', '')}</td>
                        <td>{material.get('einheit', '')}</td>
                    </tr>
                    """
                html += "</table></div>"
            
            html += "</div>"
        
        html += f"""
            <div style="margin-top: 30px; color: #666; border-top: 1px solid #2a2a33; padding-top: 15px;">
                <p>Erstellt mit {o3NAME} {o3VERSION}<br>
                {o3COPYRIGHT}<br>
                https://openw3rk.de | https://o3measurement.openw3rk.de</p>
            </div>
        </body>
        </html>
        """
        
        return html


class ActivityDetailsDialog(QtWidgets.QDialog):
    def __init__(self, activity_data: dict, parent=None):
        super().__init__(parent)
        self.activity_data = activity_data
        self.setWindowTitle(f"Aktivitäts-Details - {activity_data.get('activity_type', '')}")
        self.setMinimumSize(600, 500)
        
        layout = QtWidgets.QVBoxLayout(self)
        
        # Scroll Area
        scroll_area = QtWidgets.QScrollArea()
        scroll_area.setWidgetResizable(True)
        container = QtWidgets.QWidget()
        container_layout = QtWidgets.QVBoxLayout(container)
        
        # Basis-Informationen
        info_group = QtWidgets.QGroupBox("Basis-Informationen")
        info_layout = QtWidgets.QFormLayout()
        
        info_layout.addRow("Aktivitätstyp:", QtWidgets.QLabel(activity_data.get('activity_type', '')))
        info_layout.addRow("Datum:", QtWidgets.QLabel(activity_data.get('datum', '')))
        info_layout.addRow("KM-Stand:", QtWidgets.QLabel(str(activity_data.get('km_stand', '')) + " km"))
        info_layout.addRow("Durchgeführt durch:", QtWidgets.QLabel(activity_data.get('erstellt_durch', '')))
        
        info_group.setLayout(info_layout)
        container_layout.addWidget(info_group)
        
        # Beschreibung
        desc_group = QtWidgets.QGroupBox("Beschreibung")
        desc_layout = QtWidgets.QVBoxLayout()
        
        desc_label = QtWidgets.QLabel(activity_data.get('beschreibung', ''))
        desc_label.setWordWrap(True)
        desc_layout.addWidget(desc_label)
        
        desc_group.setLayout(desc_layout)
        container_layout.addWidget(desc_group)
        
        # Durchgeführt
        if activity_data.get('durchgeführte_arbeiten'):
            work_group = QtWidgets.QGroupBox("Durchgeführte Arbeiten")
            work_layout = QtWidgets.QVBoxLayout()
            
            for arbeit in activity_data['durchgeführte_arbeiten']:
                work_text = f"• {arbeit.get('arbeit', '')} ({arbeit.get('dauer_minuten', 0)} Min)"
                if arbeit.get('bemerkungen'):
                    work_text += f" - {arbeit.get('bemerkungen', '')}"
                work_label = QtWidgets.QLabel(work_text)
                work_label.setWordWrap(True)
                work_layout.addWidget(work_label)
            
            work_group.setLayout(work_layout)
            container_layout.addWidget(work_group)
        
        # Verwendete Materialien
        if activity_data.get('verwendete_materialien'):
            material_group = QtWidgets.QGroupBox("Verwendete Materialien")
            material_layout = QtWidgets.QVBoxLayout()
            
            material_table = QtWidgets.QTableWidget()
            material_table.setColumnCount(4)
            material_table.setHorizontalHeaderLabels(["Teilenummer", "Bezeichnung", "Menge", "Einheit"])
            material_table.setRowCount(len(activity_data['verwendete_materialien']))
            material_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
            
            for row, material in enumerate(activity_data['verwendete_materialien']):
                material_table.setItem(row, 0, QtWidgets.QTableWidgetItem(material.get('teilenummer', '')))
                material_table.setItem(row, 1, QtWidgets.QTableWidgetItem(material.get('bezeichnung', '')))
                material_table.setItem(row, 2, QtWidgets.QTableWidgetItem(str(material.get('menge_verbraucht', ''))))
                material_table.setItem(row, 3, QtWidgets.QTableWidgetItem(material.get('einheit', '')))
            
            material_table.horizontalHeader().setStretchLastSection(True)
            material_layout.addWidget(material_table)
            material_group.setLayout(material_layout)
            container_layout.addWidget(material_group)
        
        container_layout.addStretch()
        scroll_area.setWidget(container)
        layout.addWidget(scroll_area)
        
        # Buttons
        button_layout = QtWidgets.QHBoxLayout()
        
        export_btn = QtWidgets.QPushButton("Einzel-Export")
        export_btn.clicked.connect(self.export_single_activity)
        
        button_layout.addWidget(export_btn)
        button_layout.addStretch()
        
        close_btn = QtWidgets.QPushButton("Schließen")
        close_btn.clicked.connect(self.reject)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
    
    def export_single_activity(self):
        html = self._generate_single_activity_html()
        
        fname, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Aktivität exportieren",
            str(Path.home() / f"{self.activity_data.get('activity_type', 'aktivität')}_{self.activity_data.get('datum', '')}.html"),
            "HTML Dateien (*.html)"
        )
        
        if fname:
            with open(fname, 'w', encoding='utf-8') as f:
                f.write(html)
            QtWidgets.QMessageBox.information(self, "Erfolg", f"Aktivität exportiert: {fname}")
    
    def _generate_single_activity_html(self):
        activity = self.activity_data
        
        html = f"""
        <!DOCTYPE html>
        <html lang="de">
        <head>
            <meta charset="utf-8">
            <title>{activity.get('activity_type', '')} - {activity.get('datum', '')}</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; background-color: #0f0f12; color: #e6eef8; }}
                .header {{ background: #2c3e50; color: white; padding: 20px; border-radius: 5px; margin-bottom: 20px; }}
                .section {{ background: #1b1b20; border: 1px solid #2a2a33; border-radius: 5px; padding: 15px; margin-bottom: 15px; }}
                table {{ width: 100%; border-collapse: collapse; margin: 10px 0; }}
                th, td {{ border: 1px solid #2a2a33; padding: 8px; text-align: left; }}
                th {{ background: #2c3e50; color: white; }}
            </style>
        </head>
        <body>
            <div class="header">
                <h1>{activity.get('activity_type', '')}</h1>
                <p>Fahrzeug: {activity.get('fahrzeug', '')} | Datum: {activity.get('datum', '')} | KM-Stand: {activity.get('km_stand', '')} km</p>
            </div>
            
            <div class="section">
                <h3>Beschreibung</h3>
                <p>{activity.get('beschreibung', '').replace(chr(10), '<br>')}</p>
            </div>
        """
        
        if activity.get('durchgeführte_arbeiten'):
            html += """
            <div class="section">
                <h3>Durchgeführte Arbeiten</h3>
                <table>
                    <tr><th>Aktivität</th><th>Dauer (Min)</th><th>Bemerkungen</th></tr>
            """
            for arbeit in activity['durchgeführte_arbeiten']:
                html += f"""
                <tr>
                    <td>{arbeit.get('arbeit', '')}</td>
                    <td>{arbeit.get('dauer_minuten', '')}</td>
                    <td>{arbeit.get('bemerkungen', '')}</td>
                </tr>
                """
            html += "</table></div>"
        
        if activity.get('verwendete_materialien'):
            html += """
            <div class="section">
                <h3>Verwendete Materialien</h3>
                <table>
                    <tr><th>Teilenummer</th><th>Bezeichnung</th><th>Menge</th><th>Einheit</th></tr>
            """
            for material in activity['verwendete_materialien']:
                html += f"""
                <tr>
                    <td>{material.get('teilenummer', '')}</td>
                    <td>{material.get('bezeichnung', '')}</td>
                    <td>{material.get('menge_verbraucht', '')}</td>
                    <td>{material.get('einheit', '')}</td>
                </tr>
                """
            html += "</table></div>"
        
        html += f"""
            <div style="margin-top: 30px; color: #666; border-top: 1px solid #2a2a33; padding-top: 15px;">
                <p>Durchgeführt durch: {activity.get('erstellt_durch', '')}<br><br>
                <p>Erstellt mit {o3NAME} {o3VERSION}<br>{o3COPYRIGHT}<br>
                Erstellt am: {activity.get('erstellt_am', '')}</p>
            </div>
        </body>
        </html>
        """
        
        return html


class MaterialSelectionDialog(QtWidgets.QDialog):
    def __init__(self, storage_manager: StorageManager, parent=None):
        super().__init__(parent)
        self.storage_manager = storage_manager
        self.selected_materials = []
        self.setWindowTitle("Material aus Lager auswählen")
        self.setMinimumSize(600, 400)
        
        layout = QtWidgets.QVBoxLayout(self)
        
        # Suchfeld
        search_layout = QtWidgets.QHBoxLayout()
        search_layout.addWidget(QtWidgets.QLabel("Suchen:"))
        self.search_edit = QtWidgets.QLineEdit()
        self.search_edit.setPlaceholderText("Teilenummer, Bezeichnung...")
        self.search_edit.textChanged.connect(self.filter_materials)
        search_layout.addWidget(self.search_edit)
        layout.addLayout(search_layout)
        
        # Material-Tabelle
        self.material_table = QtWidgets.QTableWidget()
        self.material_table.setColumnCount(4)
        self.material_table.setHorizontalHeaderLabels(["", "Teilenummer", "Bezeichnung", "Lagerbestand"])
        self.material_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self._load_materials()
        layout.addWidget(self.material_table)
        
        # Buttons
        button_layout = QtWidgets.QHBoxLayout()
        select_btn = QtWidgets.QPushButton("Ausgewählte hinzufügen")
        select_btn.clicked.connect(self.select_materials)
        cancel_btn = QtWidgets.QPushButton("Abbrechen")
        cancel_btn.clicked.connect(self.reject)
        
        button_layout.addWidget(select_btn)
        button_layout.addStretch()
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)
    
    def _load_materials(self):
        self.material_table.setRowCount(len(self.storage_manager.items))
        
        for row, item in enumerate(self.storage_manager.items):
            # Checkbox für Auswahl
            checkbox = QtWidgets.QCheckBox()
            checkbox_widget = QtWidgets.QWidget()
            checkbox_layout = QtWidgets.QHBoxLayout(checkbox_widget)
            checkbox_layout.addWidget(checkbox)
            checkbox_layout.setAlignment(QtCore.Qt.AlignCenter)
            checkbox_layout.setContentsMargins(0, 0, 0, 0)
            self.material_table.setCellWidget(row, 0, checkbox_widget)
            
            self.material_table.setItem(row, 1, QtWidgets.QTableWidgetItem(item.teilenummer))
            self.material_table.setItem(row, 2, QtWidgets.QTableWidgetItem(item.bezeichnung))
            
            # Lagerbestand mit Farbe bei niedrigem Bestand
            stock_item = QtWidgets.QTableWidgetItem(str(item.anzahl))
            if item.anzahl <= getattr(item, 'mindestbestand', 0):
                stock_item.setBackground(QtGui.QColor('#ffcccc'))
            self.material_table.setItem(row, 3, stock_item)
    
    def filter_materials(self):
        search_text = self.search_edit.text().lower()
        
        for row in range(self.material_table.rowCount()):
            teilenummer_item = self.material_table.item(row, 1)
            bezeichnung_item = self.material_table.item(row, 2)
            
            show_row = False
            if teilenummer_item and search_text in teilenummer_item.text().lower():
                show_row = True
            elif bezeichnung_item and search_text in bezeichnung_item.text().lower():
                show_row = True
            elif not search_text:
                show_row = True
            
            self.material_table.setRowHidden(row, not show_row)
    
    def select_materials(self):
        self.selected_materials = []
        
        for row in range(self.material_table.rowCount()):
            checkbox_widget = self.material_table.cellWidget(row, 0)
            checkbox = checkbox_widget.findChild(QtWidgets.QCheckBox)
            
            if checkbox and checkbox.isChecked():
                teilenummer_item = self.material_table.item(row, 1)
                bezeichnung_item = self.material_table.item(row, 2)
                einheit = getattr(self.storage_manager.items[row], 'einheit', 'Stück')
                
                if teilenummer_item:
                    material_data = {
                        'teilenummer': teilenummer_item.text(),
                        'bezeichnung': bezeichnung_item.text() if bezeichnung_item else "",
                        'einheit': einheit
                    }
                    self.selected_materials.append(material_data)
        
        if self.selected_materials:
            self.accept()
        else:
            QtWidgets.QMessageBox.warning(self, "Keine Auswahl", "Bitte wählen Sie mindestens ein Material aus.")
    
    def get_selected_materials(self):
        return self.selected_materials

# backups
class BackupManager:
    def __init__(self):
        self.backup_dir = Path.home() / ".o3measurement" / "backups"
        self.backup_dir.mkdir(parents=True, exist_ok=True)
    
    def create_full_backup(self, include_vehicles=True, include_templates=True, 
                          include_storage=True, include_data=True, backup_path=None):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if not backup_path:
                backup_path = self.backup_dir / f"o3measurement_backup_{timestamp}.zip"
            
            # Metadaten für Backup
            backup_info = {
                "app_name": o3NAME,
                "app_version": o3VERSION,
                "backup_date": datetime.now().isoformat(),
                "includes": {
                    "vehicles": include_vehicles,
                    "templates": include_templates,
                    "storage": include_storage,
                    "data": include_data
                },
                "total_files": 0,
                "total_size": 0
            }
            
            with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as backup_zip:
                files_added = 0
                total_size = 0
                
                # Backup-Info hinzufügen
                backup_info_file = "backup_info.json"
                backup_zip.writestr(backup_info_file, json.dumps(backup_info, indent=2))
                files_added += 1
                
                # 1. Fahrzeuge sichern
                if include_vehicles and VEHICLES_BASE_DIR.exists():
                    vehicle_files = list(VEHICLES_BASE_DIR.rglob("*"))
                    for file_path in vehicle_files:
                        if file_path.is_file():
                            arcname = f"vehicles/{file_path.relative_to(VEHICLES_BASE_DIR)}"
                            backup_zip.write(file_path, arcname)
                            files_added += 1
                            total_size += file_path.stat().st_size
                
                # 2. Vorlagen sichern
                if include_templates and TEMPLATES_DIR.exists():
                    template_files = list(TEMPLATES_DIR.rglob("*.json"))
                    for file_path in template_files:
                        arcname = f"templates/{file_path.name}"
                        backup_zip.write(file_path, arcname)
                        files_added += 1
                        total_size += file_path.stat().st_size
                
                # 3. Lagerbestand sichern
                if include_storage and STORAGE_DIR.exists():
                    storage_files = list(STORAGE_DIR.rglob("*"))
                    for file_path in storage_files:
                        if file_path.is_file():
                            arcname = f"storage/{file_path.relative_to(STORAGE_DIR)}"
                            backup_zip.write(file_path, arcname)
                            files_added += 1
                            total_size += file_path.stat().st_size
                
                # 4. DATA_DIR sichern (NEU)
                if include_data and DATA_DIR.exists():
                    data_files = list(DATA_DIR.rglob("*"))
                    for file_path in data_files:
                        if file_path.is_file():
                            arcname = f"data/{file_path.relative_to(DATA_DIR)}"
                            backup_zip.write(file_path, arcname)
                            files_added += 1
                            total_size += file_path.stat().st_size
                
                # Backup info aktualisieren
                backup_info["total_files"] = files_added
                backup_info["total_size"] = total_size
                backup_zip.writestr(backup_info_file, json.dumps(backup_info, indent=2))
            
            return {
                "success": True,
                "backup_path": backup_path,
                "file_count": files_added,
                "total_size": total_size,
                "backup_info": backup_info
            }
            
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def restore_backup(self, backup_path, restore_vehicles=True, restore_templates=True,
                      restore_storage=True, restore_data=True, overwrite_existing=True):
        try:
            if not Path(backup_path).exists():
                return {"success": False, "error": "Backup-Datei nicht gefunden"}
            
            restore_info = {
                "files_restored": 0,
                "files_skipped": 0,
                "files_overwritten": 0,
                "errors": []
            }
            
            with zipfile.ZipFile(backup_path, 'r') as backup_zip:
                # Backup-Info lesen
                backup_info = {}
                if 'backup_info.json' in backup_zip.namelist():
                    with backup_zip.open('backup_info.json') as f:
                        backup_info = json.load(f)
                
                # Dateien extrahieren
                for file_info in backup_zip.infolist():
                    try:
                        filename = file_info.filename
                        
                        # Verzeichnis struktur bestimmen
                        if filename.startswith('vehicles/') and restore_vehicles:
                            target_path = VEHICLES_BASE_DIR / Path(filename).relative_to('vehicles')
                        elif filename.startswith('templates/') and restore_templates:
                            target_path = TEMPLATES_DIR / Path(filename).name
                        elif filename.startswith('storage/') and restore_storage:
                            target_path = STORAGE_DIR / Path(filename).relative_to('storage')
                        elif filename.startswith('data/') and restore_data:  # NEU
                            target_path = DATA_DIR / Path(filename).relative_to('data')
                        elif filename == 'backup_info.json':
                            continue  # Backup-Info überspringen
                        else:
                            continue  # Unbekannte Dateien überspringen
                        
                        # Prüfen ob Datei bereits existiert
                        if target_path.exists():
                            if not overwrite_existing:
                                restore_info["files_skipped"] += 1
                                continue
                            else:
                                restore_info["files_overwritten"] += 1
                                target_path.unlink()
                        
                        # Verzeichnis erstellen falls nötig
                        target_path.parent.mkdir(parents=True, exist_ok=True)
                        
                        # Datei extrahieren
                        with backup_zip.open(filename) as source, open(target_path, 'wb') as target:
                            target.write(source.read())
                        
                        restore_info["files_restored"] += 1
                        
                    except Exception as e:
                        restore_info["errors"].append(f"Fehler bei {filename}: {str(e)}")
                
                restore_info["success"] = True
                restore_info["backup_info"] = backup_info
                return restore_info
                
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def list_backups(self):
        backups = []
        if self.backup_dir.exists():
            for backup_file in self.backup_dir.glob("o3measurement_backup_*.zip"):
                try:
                    with zipfile.ZipFile(backup_file, 'r') as zipf:
                        if 'backup_info.json' in zipf.namelist():
                            with zipf.open('backup_info.json') as f:
                                info = json.load(f)
                            backups.append({
                                'file': backup_file,
                                'date': info.get('backup_date', ''),
                                'version': info.get('app_version', ''),
                                'files': info.get('total_files', 0),
                                'size': backup_file.stat().st_size,
                                'includes': info.get('includes', {})
                            })
                except:
                    # Fallback für Backups ohne Metadaten
                    backups.append({
                        'file': backup_file,
                        'date': backup_file.stat().st_mtime,
                        'version': 'Unbekannt',
                        'files': 'Unbekannt',
                        'size': backup_file.stat().st_size,
                        'includes': {}
                    })
        
        # Nach Datum sortieren
        backups.sort(key=lambda x: x['date'], reverse=True)
        return backups
    
    def get_backup_info(self, backup_path):
        try:
            with zipfile.ZipFile(backup_path, 'r') as backup_zip:
                if 'backup_info.json' in backup_zip.namelist():
                    with backup_zip.open('backup_info.json') as f:
                        return json.load(f)
                else:
                    return {"error": "Keine Backup-Informationen gefunden"}
        except Exception as e:
            return {"error": str(e)}


class BackupDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.backup_manager = BackupManager()
        self.setWindowTitle("o3Measurement - Backup & Wiederherstellung")
        screen = QtWidgets.QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        self.setFixedSize(740, 500) 
        self.resize(int(screen_geometry.width() * 0.7), int(screen_geometry.height() * 0.65))
        
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        
        # Tabs
        tab_widget = QtWidgets.QTabWidget()
        tab_widget.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #2a2a33; background: #1b1b20; }
            QTabBar::tab { background: #2a2a33; color: #e6eef8; padding: 6px 12px; margin: 1px; border: 1px solid #444; }
            QTabBar::tab:selected { background: #0088ff; color: white; border: 1px solid #00aaff; }
            QTabBar::tab:hover:!selected { background: #34343e; }
        """)
        
        # Backup Tab
        backup_tab = QtWidgets.QWidget()
        backup_layout = QtWidgets.QHBoxLayout(backup_tab)
        backup_layout.setContentsMargins(5, 5, 5, 5)
        backup_layout.setSpacing(8)
        
        # Linke Spalt
        left_column = QtWidgets.QVBoxLayout()
        left_column.setSpacing(6)
        
        # Datenauswahl
        data_group = QtWidgets.QGroupBox("Zu sichernde Daten auswählen")
        data_layout = QtWidgets.QVBoxLayout(data_group)
        data_layout.setSpacing(4)
        
        self.backup_vehicles = QtWidgets.QCheckBox("Fahrzeugdaten (Fahrzeuge/Service-Historien/Tabellen)")
        self.backup_vehicles.setChecked(True)
        data_layout.addWidget(self.backup_vehicles)
        
        self.backup_templates = QtWidgets.QCheckBox("Messvorlagen (Alle Vorlagen)")
        self.backup_templates.setChecked(True)
        data_layout.addWidget(self.backup_templates)
        
        self.backup_storage = QtWidgets.QCheckBox("Lagerbestand (Kompletter Bestand)")
        self.backup_storage.setChecked(True)
        data_layout.addWidget(self.backup_storage)
        
        self.backup_data = QtWidgets.QCheckBox("Datasets (Alle Datasets)")
        self.backup_data.setChecked(True)
        data_layout.addWidget(self.backup_data)

        left_column.addWidget(data_group)
        
        # Speicherort
        location_group = QtWidgets.QGroupBox("Backup-Speicherort festlegen")
        location_layout = QtWidgets.QVBoxLayout(location_group)
        
        # Pfad mit Button
        path_layout = QtWidgets.QHBoxLayout()
        self.location_edit = QtWidgets.QLineEdit()
        self.location_edit.setReadOnly(True)
        path_layout.addWidget(self.location_edit)
        
        browse_btn = QtWidgets.QPushButton("Durchsuchen")
        browse_btn.setFixedWidth(80)
        browse_btn.clicked.connect(self._browse_backup_location)
        path_layout.addWidget(browse_btn)
        
        location_layout.addLayout(path_layout)
        
        # Info unter Pfad
        info_text = QtWidgets.QLabel("Ideal für externe Laufwerke und Serverfreigaben.")
        info_text.setStyleSheet("color: #9aa; font-size: 9px; padding: 2px;")
        location_layout.addWidget(info_text)
        
        # Namensgebung mit zeit
        auto_path = self.backup_manager.backup_dir / f"o3measurement_backup_{datetime.now().strftime('%d_%m_%Y_%H_%M_%S')}.zip"
        self.location_edit.setText(str(auto_path))
        
        left_column.addWidget(location_group)
        left_column.addStretch()
        
        # Rechte spalte
        right_column = QtWidgets.QVBoxLayout()
        right_column.setSpacing(6)
        
        # Info bereich
        info_group = QtWidgets.QGroupBox("Informationen")
        info_layout = QtWidgets.QVBoxLayout(info_group)
        
        info_text = QtWidgets.QLabel(
            "<p>Erstellen Sie ein Backup "
            "der o3Measurement-Daten,</p>"
            "<p>um auf eine neue Version umzusteigen "
            "oder ein neues Gerät einzurichten</p>"
        )
        info_text.setStyleSheet("font-size: 10px; line-height: 1.4;")
        info_text.setWordWrap(True)
        info_layout.addWidget(info_text)
        right_column.addWidget(info_group)
        
        # Backup button
        self.backup_btn = QtWidgets.QPushButton("Backup Erstellen")
        self.backup_btn.clicked.connect(self._create_backup)
        self.backup_btn.setStyleSheet("""
            QPushButton {
                background-color: #2e7d32; 
                color: white; 
                font-weight: bold; 
                padding: 10px;
                font-size: 12px;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #3d8b40; }
            QPushButton:pressed { background-color: #1b5e20; }
        """)
        right_column.addWidget(self.backup_btn)
        
        # log
        status_group = QtWidgets.QGroupBox("Backup-Fortschritt")
        status_layout = QtWidgets.QVBoxLayout(status_group)
        self.backup_status = QtWidgets.QLabel("Bereit für Backuperstellung\nWählen Sie Daten aus und erstellen Sie das Backup")
        self.backup_status.setWordWrap(True)
        self.backup_status.setStyleSheet("font-size: 10px; padding: 6px; min-height: 60px; background: #1b1b20; border-radius: 3px;")
        self.backup_status.setAlignment(QtCore.Qt.AlignTop | QtCore.Qt.AlignLeft)
        status_layout.addWidget(self.backup_status)
        right_column.addWidget(status_group)
        
        # Spalten zusammenfügen
        backup_layout.addLayout(left_column, 45)  
        backup_layout.addLayout(right_column, 55) 
        
        tab_widget.addTab(backup_tab, "Backup")
        
        # ZWEISPALTIG
        restore_tab = QtWidgets.QWidget()
        restore_layout = QtWidgets.QHBoxLayout(restore_tab)
        restore_layout.setContentsMargins(5, 5, 5, 5)
        restore_layout.setSpacing(8)
        
        # Linke Spalte
        restore_left = QtWidgets.QVBoxLayout()
        restore_left.setSpacing(6)
        
        # quelle
        source_group = QtWidgets.QGroupBox("Quelle auswählen")
        source_layout = QtWidgets.QVBoxLayout(source_group)
        
        # Auswahlmethode
        method_layout = QtWidgets.QHBoxLayout()
        self.use_existing_backups = QtWidgets.QRadioButton("Vorhandene Backups")
        self.use_existing_backups.setChecked(True)
        self.use_existing_backups.toggled.connect(self._on_select_method_changed)
        method_layout.addWidget(self.use_existing_backups)
        
        self.use_custom_backup = QtWidgets.QRadioButton("Externes Backup")
        self.use_custom_backup.toggled.connect(self._on_select_method_changed)
        method_layout.addWidget(self.use_custom_backup)
        source_layout.addLayout(method_layout)
        
        # Backup liste
        self.backups_list = QtWidgets.QListWidget()
        self.backups_list.setMaximumHeight(140)
        self._refresh_backups_list()
        source_layout.addWidget(self.backups_list)
        
        # Eigene Backups
        custom_layout = QtWidgets.QHBoxLayout()
        self.custom_backup_edit = QtWidgets.QLineEdit()
        self.custom_backup_edit.setReadOnly(True)
        self.custom_backup_edit.setEnabled(False)
        custom_layout.addWidget(self.custom_backup_edit)
        
        self.custom_browse_btn = QtWidgets.QPushButton("Durchsuchen")
        self.custom_browse_btn.setFixedWidth(80)
        self.custom_browse_btn.clicked.connect(self._browse_custom_backup)
        self.custom_browse_btn.setEnabled(False)
        custom_layout.addWidget(self.custom_browse_btn)
        source_layout.addLayout(custom_layout)
        
        restore_left.addWidget(source_group)
        
        # Backup info
        info_group2 = QtWidgets.QGroupBox("Ausgewähltes Backup")
        info_layout2 = QtWidgets.QVBoxLayout(info_group2)
        self.backup_info_label = QtWidgets.QLabel("Wählen Sie ein Backup aus\n\nBackup-Datum & Version\nEnthaltene Daten\nDateigröße & Anzahl")
        self.backup_info_label.setWordWrap(True)
        self.backup_info_label.setStyleSheet("font-size: 10px; padding: 6px; min-height: 80px; background: #1b1b20; border-radius: 3px;")
        self.backup_info_label.setAlignment(QtCore.Qt.AlignTop | QtCore.Qt.AlignLeft)
        info_layout2.addWidget(self.backup_info_label)
        restore_left.addWidget(info_group2)
        
        restore_left.addStretch()
        
        # Rechte Spalte
        restore_right = QtWidgets.QVBoxLayout()
        restore_right.setSpacing(6)
        
        options_group = QtWidgets.QGroupBox("Wiederherstellungsoptionen")
        options_layout = QtWidgets.QVBoxLayout(options_group)
        options_layout.setSpacing(4)
        
        self.restore_vehicles = QtWidgets.QCheckBox("Fahrzeugdaten wiederherstellen")
        self.restore_vehicles.setChecked(True)
        options_layout.addWidget(self.restore_vehicles)
        
        self.restore_templates = QtWidgets.QCheckBox("Messvorlagen wiederherstellen")
        self.restore_templates.setChecked(True)
        options_layout.addWidget(self.restore_templates)
        
        self.restore_storage = QtWidgets.QCheckBox("Lagerbestand wiederherstellen")
        self.restore_storage.setChecked(True)
        options_layout.addWidget(self.restore_storage)

        self.restore_data = QtWidgets.QCheckBox("Datasets wiederherstellen")
        self.restore_data.setChecked(True)
        options_layout.addWidget(self.restore_data)

        self.overwrite_existing = QtWidgets.QCheckBox("Vorhandene Daten überschreiben") # ROT
        self.overwrite_existing.setChecked(True)
        self.overwrite_existing.setStyleSheet("color: #ff6b6b; font-weight: bold;")
        options_layout.addWidget(self.overwrite_existing)
        
        restore_right.addWidget(options_group)
        
        # Restore button
        self.restore_btn = QtWidgets.QPushButton("Backup wiederherstellen")
        self.restore_btn.clicked.connect(self._restore_backup)
        self.restore_btn.setStyleSheet("""
            QPushButton {
                background-color: #d32f2f; 
                color: white; 
                font-weight: bold; 
                padding: 10px;
                font-size: 12px;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #e53935; }
            QPushButton:pressed { background-color: #b71c1c; }
            QPushButton:disabled { background-color: #666; color: #999; }
        """)
        self.restore_btn.setEnabled(False)
        restore_right.addWidget(self.restore_btn)
        
        # log
        status_group2 = QtWidgets.QGroupBox("Wiederherstellungs-Status")
        status_layout2 = QtWidgets.QVBoxLayout(status_group2)
        self.restore_status = QtWidgets.QLabel("Bereit für Wiederherstellung\nWählen Sie ein Backup und den Datensatz aus")
        self.restore_status.setWordWrap(True)
        self.restore_status.setStyleSheet("font-size: 10px; padding: 6px; min-height: 60px; background: #1b1b20; border-radius: 3px;")
        self.restore_status.setAlignment(QtCore.Qt.AlignTop | QtCore.Qt.AlignLeft)
        status_layout2.addWidget(self.restore_status)
        restore_right.addWidget(status_group2)
        
        # Spalten zusammenfügen
        restore_layout.addLayout(restore_left, 50)
        restore_layout.addLayout(restore_right, 50)
        
        tab_widget.addTab(restore_tab, "Wiederherstellen")
        
        layout.addWidget(tab_widget)
        
        # Schließen
        btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)
        
        self.backups_list.currentItemChanged.connect(self._on_backup_selected)
    
    def _on_select_method_changed(self):
        if self.use_existing_backups.isChecked():
            self.backups_list.setEnabled(True)
            self.custom_backup_edit.setEnabled(False)
            self.custom_browse_btn.setEnabled(False)
            self._on_backup_selected(self.backups_list.currentItem(), None)
        else:
            self.backups_list.setEnabled(False)
            self.custom_backup_edit.setEnabled(True)
            self.custom_browse_btn.setEnabled(True)
            self._on_custom_backup_selected()
    
    def _browse_backup_location(self):
        fname, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, 
            "Backup speichern unter", 
            self.location_edit.text(),
            "Zip-Backup (*.zip)"
        )
        if fname:
            self.location_edit.setText(fname)
    
    def _browse_custom_backup(self):
        fname, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Backup-Datei auswählen",
            str(Path.home()),
            "Zip-Backup (*.zip)"
        )
        if fname:
            self.custom_backup_edit.setText(fname)
            self._on_custom_backup_selected()
    
    def _on_custom_backup_selected(self):
        backup_path = self.custom_backup_edit.text().strip()
        if backup_path and Path(backup_path).exists():
            info = self.backup_manager.get_backup_info(backup_path)
            
            if "error" not in info:
                backup_date = datetime.fromisoformat(info['backup_date']).strftime("%d.%m.%Y %H:%M:%S")
                includes = []
                if info.get('includes', {}).get('vehicles'):
                    includes.append("Fahrzeuge")
                if info.get('includes', {}).get('templates'):
                    includes.append("Vorlagen")
                if info.get('includes', {}).get('datasets'):
                    includes.append("Daten")
                if info.get('includes', {}).get('storage'):
                    includes.append("Lager")
                if info.get('includes', {}).get('settings'):
                    includes.append("Einstellungen")
                
                info_text = (
                    f"Backup vom: {backup_date}\n"
                    f"Version: {info.get('app_version', 'Unbekannt')}\n"
                    f"Dateien: {info.get('total_files', 0)}\n"
                    f"Größe: {info.get('total_size', 0) / 1024 / 1024:.2f} MB\n"
                    f"Enthält: {', '.join(includes)}"
                )
            else:
                info_text = f"Backup-Informationen nicht verfügbar\nDatei: {Path(backup_path).name}"
            
            self.backup_info_label.setText(info_text)
            self.restore_btn.setEnabled(True)
        else:
            self.backup_info_label.setText("Bitte wählen Sie eine gültige Backup-Datei aus")
            self.restore_btn.setEnabled(False)
    

    def _create_backup(self):
        backup_path = self.location_edit.text().strip()
        if not backup_path:
            QtWidgets.QMessageBox.warning(self, "Speicherort fehlt", "Bitte wählen Sie einen Speicherort für das Backup.")
            return
        
        # Sicherheitsabfrage falls Datei bereits existiert
        if Path(backup_path).exists():
            reply = QtWidgets.QMessageBox.question(
                self,
                "Datei existiert bereits",
                f"Die Datei '{Path(backup_path).name}' existiert bereits.\nMöchten Sie sie überschreiben?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
            )
            if reply != QtWidgets.QMessageBox.Yes:
                return
        
        # Backup Button deaktivieren während des Backups
        self.backup_btn.setEnabled(False)
        self.backup_status.setText("Backup wird erstellt...")
        
        QtCore.QTimer.singleShot(100, lambda: self._execute_backup(backup_path))

    def _execute_backup(self, backup_path):
        result = self.backup_manager.create_full_backup(
            include_vehicles=self.backup_vehicles.isChecked(),
            include_templates=self.backup_templates.isChecked(),
            include_storage=self.backup_storage.isChecked(),
            include_data=self.backup_data.isChecked(),
            backup_path=backup_path
        )
        
        # UI wieder aktivieren bzw. antifreeze
        self.backup_btn.setEnabled(True)
        
        if result["success"]:
            self.backup_status.setText(
                f"Backup erfolgreich erstellt!\n"
                f"Datei: {result['backup_path']}\n"
                f"Dateien: {result['file_count']}\n"
                f"Größe: {result['total_size'] / 1024 / 1024:.2f} MB"
            )
            self.backup_status.setStyleSheet("color: #90ee90;")
            self._refresh_backups_list()
        else:
            self.backup_status.setText(f"Fehler beim Backup: {result['error']}")
            self.backup_status.setStyleSheet("color: #ff6b6b;")
    
    def _refresh_backups_list(self):
        self.backups_list.clear()
        backups = self.backup_manager.list_backups()
        
        for backup in backups:
            try:
                backup_date = datetime.fromisoformat(backup['date']).strftime("%d.%m.%Y %H:%M")
            except:
                backup_date = datetime.fromtimestamp(backup['date']).strftime("%d.%m.%Y %H:%M") if isinstance(backup['date'], (int, float)) else "Unbekannt"
            
            item_text = f"{backup_date} - v{backup['version']} - {backup['files']} Dateien - {backup['size'] / 1024 / 1024:.1f} MB"
            item = QtWidgets.QListWidgetItem(item_text)
            item.setData(QtCore.Qt.UserRole, backup['file'])
            self.backups_list.addItem(item)
    
    def _on_backup_selected(self, current, previous):
        if current:
            backup_path = current.data(QtCore.Qt.UserRole)
            info = self.backup_manager.get_backup_info(backup_path)
            
            if "error" not in info:
                backup_date = datetime.fromisoformat(info['backup_date']).strftime("%d.%m.%Y %H:%M:%S")
                includes = []
                if info.get('includes', {}).get('vehicles'):
                    includes.append("Fahrzeuge")
                if info.get('includes', {}).get('templates'):
                    includes.append("Vorlagen")
                if info.get('includes', {}).get('datasets'):
                    includes.append("Daten")
                if info.get('includes', {}).get('storage'):
                    includes.append("Lager")
                if info.get('includes', {}).get('settings'):
                    includes.append("Einstellungen")
                
                info_text = (
                    f"Backup vom: {backup_date}\n"
                    f"Version: {info.get('app_version', 'Unbekannt')}\n"
                    f"Dateien: {info.get('total_files', 0)}\n"
                    f"Größe: {info.get('total_size', 0) / 1024 / 1024:.2f} MB\n"
                    f"Enthält: {', '.join(includes)}"
                )
            else:
                info_text = f"Backup-Informationen nicht verfügbar\nDatei: {backup_path.name}"
            
            self.backup_info_label.setText(info_text)
            self.restore_btn.setEnabled(True)
        else:
            self.backup_info_label.setText("Wählen Sie ein Backup aus der Liste")
            self.restore_btn.setEnabled(False)
    
    def _restore_backup(self):
        # Pfad bestimmen
        if self.use_existing_backups.isChecked():
            current_item = self.backups_list.currentItem()
            if not current_item:
                QtWidgets.QMessageBox.warning(self, "Kein Backup", "Bitte wählen Sie ein Backup aus der Liste aus.")
                return
            backup_path = current_item.data(QtCore.Qt.UserRole)
        else:
            backup_path = self.custom_backup_edit.text().strip()
            if not backup_path or not Path(backup_path).exists():
                QtWidgets.QMessageBox.warning(self, "Ungültige Datei", "Bitte wählen Sie eine gültige Backup-Datei aus.")
                return
        
        # Sicherheitsabfrage
        reply = QtWidgets.QMessageBox.question(
            self,
            "Backup wiederherstellen",
            "Sind Sie sicher, dass Sie dieses Backup wiederherstellen möchten?\n"
            "Vorhandene Daten können überschrieben werden!",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
        )
        
        if reply != QtWidgets.QMessageBox.Yes:
            return
        
        # Restore btn deaktivieren
        self.restore_btn.setEnabled(False)
        self.restore_status.setText("Backup wird wiederhergestellt...")
        
        QtCore.QTimer.singleShot(100, lambda: self._execute_restore(backup_path))
    
    def _execute_restore(self, backup_path):
        result = self.backup_manager.restore_backup(
            backup_path=backup_path,
            restore_vehicles=self.restore_vehicles.isChecked(),
            restore_templates=self.restore_templates.isChecked(),
            restore_storage=self.restore_storage.isChecked(),
            overwrite_existing=self.overwrite_existing.isChecked()
        )
        
        # UI wieder aktivieren
        self.restore_btn.setEnabled(True)
        
        if result["success"]:
            status_text = (
                f"Wiederherstellung erfolgreich!\n"
                f"Wiederhergestellte Dateien: {result['files_restored']}\n"
            )
            
            if result.get('files_overwritten', 0) > 0:
                status_text += f"Überschriebene Dateien: {result['files_overwritten']}\n"
            
            if result.get('files_skipped', 0) > 0:
                status_text += f"Übersprungene Dateien: {result['files_skipped']}\n"
                
            status_text += "Programm neu starten um Änderungen zu übernehmen."
            
            self.restore_status.setText(status_text)
            self.restore_status.setStyleSheet("color: #90ee90;")
            
            if result["errors"]:
                self.restore_status.setText(
                    self.restore_status.text() + f"\n\nFehler/Info ({len(result['errors'])}):\n" + 
                    "\n".join(result['errors'][:3])
                )
        else:
            self.restore_status.setText(f"Fehler: {result['error']}")
            self.restore_status.setStyleSheet("color: #ff6b6b;")

class SaveTableDialog(QtWidgets.QDialog):
    def __init__(self, current_vehicle: Vehicle, main_window, parent=None):
        super().__init__(parent)
        self.current_vehicle = current_vehicle
        self.main_window = main_window
        self.setWindowTitle("Tabelle speichern")
        self.setMinimumSize(400, 200)
        
        layout = QtWidgets.QFormLayout(self)
        
        # Prüfen ob bereits eine Tabelle existiert
        existing_table_name = self.current_vehicle.get_existing_table_name()
        
        # Tabellenname
        self.name_edit = QtWidgets.QLineEdit()
        if existing_table_name:
            # Bestehenden Namen vorschlagen
            self.name_edit.setText(existing_table_name)
            self.name_edit.setEnabled(False)  # Namen nicht ändern lassen
            info_text = f"Wird gespeichert als: {existing_table_name}"
        else:
            # Neue Tabelle - Namen vorschlagen
            default_name = f"Messung_{datetime.now().strftime('%Y%m%d_%H%M')}"
            self.name_edit.setText(default_name)
            info_text = "Neue Tabelle wird erstellt"
        
        self.name_edit.setPlaceholderText("Name der Tabelle")
        layout.addRow("Tabellenname:", self.name_edit)
        
        # Info-Text
        info_label = QtWidgets.QLabel(info_text)
        info_label.setStyleSheet("color: #90caf9; font-style: italic;")
        layout.addRow(info_label)
        
        # Beschreibung
        self.description_edit = QtWidgets.QTextEdit()
        self.description_edit.setMaximumHeight(60)
        self.description_edit.setPlaceholderText("Beschreibung der Tabelle (optional)")
        layout.addRow("Beschreibung:", self.description_edit)
        
        # Buttons
        button_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Save | 
            QtWidgets.QDialogButtonBox.Cancel
        )
        button_box.accepted.connect(self.save_table)
        button_box.rejected.connect(self.reject)
        layout.addRow(button_box)
    
    def save_table(self):
        name = self.name_edit.text().strip()
        if not name:
            QtWidgets.QMessageBox.warning(self, "Fehler", "Bitte geben Sie einen Namen für die Tabelle ein.")
            return
        
        # Template und Tabellendaten sammeln
        template_data = self.main_window._gather_template_from_ui().to_dict()
        
        # Tabellendaten sammeln
        table_data = []
        for row in range(self.main_window.table.rowCount()):
            row_data = []
            for col in range(self.main_window.table.columnCount()):
                item = self.main_window.table.item(row, col)
                cell_data = {
                    'text': item.text() if item else '',
                    'background': item.background().color().name() if item and item.background().style() != QtCore.Qt.NoBrush else '',
                    'foreground': item.foreground().color().name() if item and item.foreground().style() != QtCore.Qt.NoBrush else ''
                }
                row_data.append(cell_data)
            table_data.append(row_data)
        
        # Prüfen ob Tabelle bereits existiert
        existing_table = None
        for table in self.current_vehicle.saved_tables:
            if table.name == name:
                existing_table = table
                break
        
        if existing_table:
            # Bestehende Tabelle aktualisieren
            existing_table.template_data = template_data
            existing_table.table_data = table_data
            existing_table.description = self.description_edit.toPlainText().strip()
            existing_table.modified_date = datetime.now().strftime("%Y-%m-%d %H:%M")
            saved_table = existing_table
        else:
            # Neue Tabelle erstellen
            saved_table = SavedTable(
                name=name,
                description=self.description_edit.toPlainText().strip(),
                template_data=template_data,
                table_data=table_data,
                vehicle_name=self.current_vehicle.name
            )
        
        # Speichern
        if self.current_vehicle.save_table(saved_table):
            action = "aktualisiert" if existing_table else "gespeichert"
            QtWidgets.QMessageBox.information(self, "Erfolg", f"Tabelle '{name}' wurde {action}.")
            self.accept()
        else:
            QtWidgets.QMessageBox.warning(self, "Fehler", "Tabelle konnte nicht gespeichert werden.")

class LoadTableDialog(QtWidgets.QDialog):
    def __init__(self, current_vehicle: Vehicle, parent=None):
        super().__init__(parent)
        self.current_vehicle = current_vehicle
        self.selected_table = None
        self.setWindowTitle("Gespeicherte Tabelle laden")
        self.setMinimumSize(500, 400)
        
        layout = QtWidgets.QVBoxLayout(self)
        
        # Suchfeld
        search_layout = QtWidgets.QHBoxLayout()
        search_layout.addWidget(QtWidgets.QLabel("Suchen:"))
        self.search_edit = QtWidgets.QLineEdit()
        self.search_edit.setPlaceholderText("Tabellenname suchen...")
        self.search_edit.textChanged.connect(self.filter_tables)
        search_layout.addWidget(self.search_edit)
        layout.addLayout(search_layout)
        
        # Tabellenliste
        self.tables_list = QtWidgets.QListWidget()
        self.tables_list.itemDoubleClicked.connect(self.load_selected_table)
        self.tables_list.currentItemChanged.connect(self.on_table_selected)
        layout.addWidget(self.tables_list)
        
        # Tabelleninfo
        self.table_info = QtWidgets.QTextEdit()
        self.table_info.setMaximumHeight(100)
        self.table_info.setReadOnly(True)
        layout.addWidget(self.table_info)
        
        # Buttons
        button_layout = QtWidgets.QHBoxLayout()
        
        load_btn = QtWidgets.QPushButton("Tabelle laden")
        load_btn.clicked.connect(self.load_selected_table)
        
        delete_btn = QtWidgets.QPushButton("Tabelle löschen")
        delete_btn.clicked.connect(self.delete_selected_table)
        delete_btn.setStyleSheet("background-color: #d32f2f; color: white;")
        
        button_layout.addWidget(load_btn)
        button_layout.addWidget(delete_btn)
        button_layout.addStretch()
        
        close_btn = QtWidgets.QPushButton("Schließen")
        close_btn.clicked.connect(self.reject)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        
        self.load_tables_list()
    
    def load_tables_list(self):
        self.tables_list.clear()
        for table in self.current_vehicle.saved_tables:
            item = QtWidgets.QListWidgetItem(f"{table.name} - {table.created_date}")
            item.setData(QtCore.Qt.UserRole, table.table_id)
            self.tables_list.addItem(item)
    
    def on_table_selected(self, current, previous):
        if current:
            table_id = current.data(QtCore.Qt.UserRole)
            table = self.current_vehicle.load_table(table_id)
            if table:
                info_text = (
                    f"Name: {table.name}\n"
                    f"Erstellt: {table.created_date}\n"
                    f"Geändert: {table.modified_date}\n"
                    f"Beschreibung: {table.description}\n"
                    f"Zeilen: {len(table.table_data)}\n"
                    f"Spalten: {len(table.template_data.get('columns', []))}"
                )
                self.table_info.setText(info_text)
    
    def filter_tables(self):
        search_text = self.search_edit.text().lower()
        for i in range(self.tables_list.count()):
            item = self.tables_list.item(i)
            item_text = item.text().lower()
            item.setHidden(search_text not in item_text)
    
    def load_selected_table(self):
        current_item = self.tables_list.currentItem()
        if not current_item:
            QtWidgets.QMessageBox.warning(self, "Keine Auswahl", "Bitte wählen Sie eine Tabelle aus.")
            return
        
        table_id = current_item.data(QtCore.Qt.UserRole)
        self.selected_table = self.current_vehicle.load_table(table_id)
        
        if self.selected_table:
            self.accept()
        else:
            QtWidgets.QMessageBox.warning(self, "Fehler", "Tabelle konnte nicht geladen werden.")
    
    def delete_selected_table(self):
        current_item = self.tables_list.currentItem()
        if not current_item:
            QtWidgets.QMessageBox.warning(self, "Keine Auswahl", "Bitte wählen Sie eine Tabelle aus.")
            return
        
        table_name = current_item.text().split(" - ")[0]
        table_id = current_item.data(QtCore.Qt.UserRole)
        
        reply = QtWidgets.QMessageBox.question(
            self, 
            "Tabelle löschen", 
            f"Soll die Tabelle '{table_name}' wirklich gelöscht werden?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
        )
        
        if reply == QtWidgets.QMessageBox.Yes:
            # Tabelle löschen
            self.current_vehicle.delete_table(table_id)
            
            # Liste aktualisieren
            self.load_tables_list()
            self.table_info.clear()
            
            QtWidgets.QMessageBox.information(self, "Erfolg", f"Tabelle '{table_name}' wurde gelöscht.")
    
    def get_selected_table(self) -> SavedTable:
        return self.selected_table

# erweiterte suche nach kfz
class AdvancedVehicleSelectionDialog(QtWidgets.QDialog):
    def __init__(self, vehicles, parent=None):
        super().__init__(parent)
        self.vehicles = vehicles
        self.selected_vehicle = None
        self.setWindowTitle(f"{o3NAME} - Fahrzeuge suchen und auswählen")
        self.setMinimumSize(1000, 600)
        
        layout = QtWidgets.QVBoxLayout(self)
        
        search_filter_group = QtWidgets.QGroupBox("Suchfilter")
        search_filter_layout = QtWidgets.QFormLayout(search_filter_group)
        
        # Suchfeld
        self.search_edit = QtWidgets.QLineEdit()
        self.search_edit.setPlaceholderText("Fahrzeugname, FIN, HSN, Modell, Farbe...")
        self.search_edit.textChanged.connect(self.filter_vehicles)
        search_filter_layout.addRow("Suche:", self.search_edit)
        
        # in einer Zeile
        filter_layout = QtWidgets.QHBoxLayout()
        
        filter_layout.addWidget(QtWidgets.QLabel("HSN:"))
        self.manufacturer_combo = QtWidgets.QComboBox()
        self.manufacturer_combo.addItem("Alle HSN-Kennungen", "")
        self.manufacturer_combo.currentTextChanged.connect(self.filter_vehicles)
        filter_layout.addWidget(self.manufacturer_combo)
        
        # Baujahr filter
        filter_layout.addWidget(QtWidgets.QLabel("Baujahr:"))
        self.year_combo = QtWidgets.QComboBox()
        self.year_combo.addItem("Alle Jahre", "")
        self.year_combo.currentTextChanged.connect(self.filter_vehicles)
        filter_layout.addWidget(self.year_combo)
        
        filter_layout.addWidget(QtWidgets.QLabel("Antrieb:"))
        self.drive_combo = QtWidgets.QComboBox()
        self.drive_combo.addItem("Alle Antriebe", "")
        self.drive_combo.currentTextChanged.connect(self.filter_vehicles)
        filter_layout.addWidget(self.drive_combo)
        
        filter_layout.addStretch()
        search_filter_layout.addRow(filter_layout)
        
        layout.addWidget(search_filter_group)
        
        self.vehicles_table = QtWidgets.QTableWidget()
        self.vehicles_table.setColumnCount(6)
        self.vehicles_table.setHorizontalHeaderLabels([
            "Fahrzeug", "HSN", "FIN", "Baujahr", "Antrieb", "Letzter Service"
        ])
        self.vehicles_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.vehicles_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.vehicles_table.doubleClicked.connect(self.accept_selection)
        self.vehicles_table.setSortingEnabled(True)
        
        # Spaltenbreiten
        self.vehicles_table.setColumnWidth(0, 230)  # Fahrzeug
        self.vehicles_table.setColumnWidth(1, 90)  # HSN
        self.vehicles_table.setColumnWidth(2, 170)  # FIN
        self.vehicles_table.setColumnWidth(3, 80)   # Baujahr
        self.vehicles_table.setColumnWidth(4, 100)  # Antrieb
        self.vehicles_table.setColumnWidth(5, 100)  # Service
        
        layout.addWidget(self.vehicles_table, 1)
        
        # Statistik
        self.stats_label = QtWidgets.QLabel("Lade Fahrzeuge...")
        self.stats_label.setStyleSheet("color: #888; font-weight: bold;")
        layout.addWidget(self.stats_label)
        
        # Buttons
        button_layout = QtWidgets.QHBoxLayout()
        
        select_btn = QtWidgets.QPushButton("Fahrzeug auswählen")
        select_btn.clicked.connect(self.accept_selection)
        
        view_btn = QtWidgets.QPushButton("Details anzeigen")
        view_btn.clicked.connect(self.view_vehicle_details)
        
        button_layout.addWidget(select_btn)
        button_layout.addWidget(view_btn)
        button_layout.addStretch()
        
        cancel_btn = QtWidgets.QPushButton("Abbrechen")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(button_layout)
        
        # Filter-Daten vorbereiten
        self._prepare_filters()
        self.load_vehicles_table()
    
    def _prepare_filters(self):
        manufacturers = set()
        years = set()
        drives = set()
        
        for vehicle in self.vehicles:
            hsn = vehicle.specifications.get('hsn', '')
            manufacturer = f"HSN {hsn}" if hsn else vehicle.name.split()[0]
            manufacturers.add(manufacturer)
            
            # Baujahr
            baujahr = vehicle.specifications.get('baujahr', '')
            if baujahr:
                years.add(baujahr)
            
            antrieb = vehicle.specifications.get('antrieb', '')
            if antrieb:
                drives.add(antrieb)
        
        for manufacturer in sorted(manufacturers):
            self.manufacturer_combo.addItem(manufacturer, manufacturer)
        
        for year in sorted(years, reverse=True):
            self.year_combo.addItem(year, year)

        for drive in sorted(drives):
            self.drive_combo.addItem(drive, drive)
    
    def _get_from_hsn(self, hsn):
        if not hsn or hsn.strip() == "":
            return "-"
        return f"HSN {hsn}"
        
    def load_vehicles_table(self, vehicles=None):
        if vehicles is None:
            vehicles = self.vehicles
        
        self.vehicles_table.setRowCount(len(vehicles))
        
        for row, vehicle in enumerate(vehicles):
            # Fahrzeugname
            self.vehicles_table.setItem(row, 0, QtWidgets.QTableWidgetItem(vehicle.name))
            
            # Hersteller
            hsn = vehicle.specifications.get('hsn', '')
            manufacturer = self._get_from_hsn(hsn) if hsn else vehicle.name.split()[0]
            self.vehicles_table.setItem(row, 1, QtWidgets.QTableWidgetItem(manufacturer))
            
            # FIN
            fin = vehicle.specifications.get('fin', '')
            self.vehicles_table.setItem(row, 2, QtWidgets.QTableWidgetItem(fin))
            
            # Baujahr
            baujahr = vehicle.specifications.get('baujahr', '')
            self.vehicles_table.setItem(row, 3, QtWidgets.QTableWidgetItem(baujahr))
            
            # Antrieb
            antrieb = vehicle.specifications.get('antrieb', '')
            self.vehicles_table.setItem(row, 4, QtWidgets.QTableWidgetItem(antrieb))
            
            # Letzter Service
            last_service = vehicle.last_service.get('ölwechsel_km', '')
            last_service_text = f"{last_service} km" if last_service else "Keine Daten"
            self.vehicles_table.setItem(row, 5, QtWidgets.QTableWidgetItem(last_service_text))
        
        self.update_stats(len(vehicles))
    
    def filter_vehicles(self):
        search_text = self.search_edit.text().lower()
        manufacturer_filter = self.manufacturer_combo.currentData()
        year_filter = self.year_combo.currentData()
        drive_filter = self.drive_combo.currentData()
        
        filtered_vehicles = []
        
        for vehicle in self.vehicles:
            # Suchtext in verschiedenen Feldern prüfen
            search_match = (
                not search_text or
                search_text in vehicle.name.lower() or
                search_text in vehicle.specifications.get('fin', '').lower() or
                search_text in vehicle.specifications.get('hsn', '').lower() or
                search_text in vehicle.specifications.get('motor', '').lower() or
                search_text in vehicle.specifications.get('farbe', '').lower() or
                search_text in vehicle.description.lower()
            )
            
            # Hersteller-Filter
            hsn = vehicle.specifications.get('hsn', '')
            manufacturer = self._get_from_hsn(hsn) if hsn else vehicle.name.split()[0]
            manufacturer_match = not manufacturer_filter or manufacturer == manufacturer_filter
            
            # Baujahr-Filter
            baujahr = vehicle.specifications.get('baujahr', '')
            year_match = not year_filter or baujahr == year_filter
            
            # Antriebs-Filter
            antrieb = vehicle.specifications.get('antrieb', '')
            drive_match = not drive_filter or antrieb == drive_filter
            
            if search_match and manufacturer_match and year_match and drive_match:
                filtered_vehicles.append(vehicle)
        
        self.load_vehicles_table(filtered_vehicles)
    
    def update_stats(self, count):
        total = len(self.vehicles)
        self.stats_label.setText(f"Zeige {count} von {total} Fahrzeugen")
    
    def accept_selection(self):
        current_row = self.vehicles_table.currentRow()
        if current_row >= 0:
            vehicle_name = self.vehicles_table.item(current_row, 0).text()
            for vehicle in self.vehicles:
                if vehicle.name == vehicle_name:
                    self.selected_vehicle = vehicle
                    self.accept()
                    return
        
        QtWidgets.QMessageBox.warning(self, "Keine Auswahl", "Bitte wählen Sie ein Fahrzeug aus der Liste aus.")
    
    def view_vehicle_details(self):
        current_row = self.vehicles_table.currentRow()
        if current_row >= 0:
            vehicle_name = self.vehicles_table.item(current_row, 0).text()
            for vehicle in self.vehicles:
                if vehicle.name == vehicle_name:
                    dlg = VehicleViewDialog(vehicle, self)
                    dlg.exec_()
                    return
        
        QtWidgets.QMessageBox.warning(self, "Keine Auswahl", "Bitte wählen Sie ein Fahrzeug aus der Liste aus.")
    
    def get_selected_vehicle(self):
        return self.selected_vehicle

class UnsavedChangesDialog(QtWidgets.QMessageBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ungesicherte Änderungen")
        self.setText("Es gibt ungesicherte Änderungen.")
        self.setInformativeText("Möchten Sie Ihre Änderungen speichern, bevor Sie fortfahren?")
        
        self.setStandardButtons(
            QtWidgets.QMessageBox.Save |
            QtWidgets.QMessageBox.Discard |
            QtWidgets.QMessageBox.Cancel
        )
        self.setDefaultButton(QtWidgets.QMessageBox.Save)
        self.setIcon(QtWidgets.QMessageBox.Warning)
        
        self.button(QtWidgets.QMessageBox.Save).setText("&Speichern")
        self.button(QtWidgets.QMessageBox.Discard).setText("Nicht &speichern")
        self.button(QtWidgets.QMessageBox.Cancel).setText("&Abbrechen")

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{o3NAME} - v{o3VERSION} | {o3COPYRIGHT}")
        self.resize(1000, 600)
        self.setMinimumSize(1000, 600)
        self.center_start()
        self.current_template = None
        self.setpoint_row_index = -1
        self.export_short_description = ""
        self.export_long_description = ""
        self.global_tolerance = 0.0
        self.current_vehicle = None
        self.vehicles = self._load_all_vehicles()
        self.unsaved_changes = False
        self.auto_save_in_progress = False
        self.create_actions()
        self.create_toolbar()
        self._create_main_ui()
        self._apply_dark_style()
        self.backup_manager = BackupManager()
        self.setup_auto_save_check()

# ANFANG DROPDOWNS/BUTTONS 

# speichererinnerung
    def setup_auto_save_check(self):
        # Überwache Änderungen in der Tabelle
        self.table.cellChanged.connect(self.mark_unsaved_changes)
        self.template_name_edit.textChanged.connect(self.mark_unsaved_changes)
        self.description_edit.textChanged.connect(self.mark_unsaved_changes)
        
        # Überwache Spalten-Änderungen
        self.columns_list.itemChanged.connect(self.mark_unsaved_changes)
        
        # Fahrzeug-Änderungen sind Auto-Saves und zählen nicht als ungespeichert
        # self.vehicle_combo.currentTextChanged.connect(self.mark_unsaved_changes) <- AUSKOMMENTIEREN

    def mark_unsaved_changes(self):
        # Wenn Auto-Save aktiv ist, keine ungespeicherten Änderungen markieren
        if hasattr(self, 'auto_save_in_progress') and self.auto_save_in_progress:
            return
            
        # Prüfen, ob die Änderung von einer relevanten Komponente kommt
        sender = self.sender()
        
        # Nur bestimmte Komponenten als "wichtige" Änderungen betrachten
        relevant_senders = [
            self.table,           # Tabellenänderungen
            self.template_name_edit,  # Vorlagenname
            self.description_edit,    # Vorlagenbeschreibung
            self.columns_list        # Spaltenänderungen
        ]
        
        if sender in relevant_senders and not self.unsaved_changes:
            self.unsaved_changes = True
            self.update_window_title()

    def clear_unsaved_changes(self):
        self.unsaved_changes = False
        self.update_window_title()

    def update_window_title(self):
        base_title = f"{o3NAME} - v{o3VERSION} | {o3COPYRIGHT}"
        if self.unsaved_changes:
            self.setWindowTitle(f"* {base_title} - Ungesicherte Änderungen")
        else:
            self.setWindowTitle(base_title)

    def closeEvent(self, event):
        if self.unsaved_changes:
            reply = self.show_unsaved_changes_dialog()
            if reply == QtWidgets.QMessageBox.Save:
                if self.save_all_changes():
                    event.accept()
                else:
                    # User hat Speichern abgebrochen -> Schließen verhindern
                    event.ignore()
            elif reply == QtWidgets.QMessageBox.Discard:
                event.accept()
            else:  # Cancel
                event.ignore()
        else:
            event.accept()

    def show_unsaved_changes_dialog(self):
        dialog = UnsavedChangesDialog(self)
        return dialog.exec_()

    def save_all_changes(self):
        try:
            # Wenn Fahrzeug ausgewählt ist -> Normales Tabellen-Speichern Menü öffnen
            if self.current_vehicle:
                # Öffne den normalen SaveTableDialog
                dlg = SaveTableDialog(self.current_vehicle, self, self)
                if dlg.exec_() == QtWidgets.QDialog.Accepted:
                    self._auto_save_vehicle(self.current_vehicle)
                    self.clear_unsaved_changes()
                    return True
                else:
                    # User hat Abbrechen geklickt
                    return False
            
            else:
                # Öffne den normalen Dataset-Speichern Dialog
                fname, _ = QtWidgets.QFileDialog.getSaveFileName(
                    self, 
                    "Tabellendaten speichern", 
                    str(DATA_DIR), 
                    "CSV Dateien (*.csv);;JSON Dateien (*.json)"
                )
                if fname:
                    if fname.lower().endswith('.csv'):
                        self._save_csv(fname)
                    else:
                        self._save_json(fname)
                    QtWidgets.QMessageBox.information(self, "Gespeichert", f"Tabellendaten gespeichert: {fname}")
                    self.clear_unsaved_changes()
                    return True
                else:
                    # User hat Abbrechen geklickt
                    return False
            
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Speicherfehler", 
                                        f"Änderungen konnten nicht gespeichert werden: {e}")
            return False


    def center_start(self): # hauptfenster mittig startem (spez. kleine monitore)
        frame_geom = self.frameGeometry()
        screen = QtWidgets.QApplication.primaryScreen().availableGeometry().center()
        frame_geom.moveCenter(screen)
        self.move(frame_geom.topLeft())

    def show_hu_au_manager(self):
        if not self.vehicles:
            QtWidgets.QMessageBox.warning(self, "Keine Fahrzeuge", "Bitte zuerst Fahrzeuge anlegen.")
            return
            
        dlg = hu_auManagerDialog(self.vehicles, self)
        dlg.exec_()

    def show_storage_manager(self):
        dlg = StorageManagerDialog(self)
        dlg.exec_()

    def show_defect_reports(self):
        if not self.current_vehicle:
            QtWidgets.QMessageBox.warning(self, "Kein Fahrzeug", "Bitte zuerst ein Fahrzeug auswählen.")
            return
            
        dlg = DefectReportsDialog(self.current_vehicle, self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            # Auto-Save-Flag setzen
            self.auto_save_in_progress = True
            try:
                self._auto_save_vehicle(self.current_vehicle)
            finally:
                self.auto_save_in_progress = False

    def show_units_dialog(self):
        dlg = UnitsViewDialog(self)
        dlg.exec_()

    def create_actions(self):
        self.new_template_act = QtWidgets.QAction("Neue Vorlage", self)
        self.new_template_act.triggered.connect(self.new_template)

        self.load_template_act = QtWidgets.QAction("Vorlage laden", self)
        self.load_template_act.triggered.connect(self.load_template)

        self.save_template_act = QtWidgets.QAction("Vorlage speichern", self)
        self.save_template_act.triggered.connect(self.save_template)

        self.new_dataset_act = QtWidgets.QAction("Neues Dataset", self)
        self.new_dataset_act.triggered.connect(self.new_dataset)

        self.load_dataset_act = QtWidgets.QAction("Dataset laden", self)
        self.load_dataset_act.triggered.connect(self.load_dataset)

        self.save_dataset_act = QtWidgets.QAction("Dataset speichern", self)
        self.save_dataset_act.triggered.connect(self.save_dataset)

        self.print_act = QtWidgets.QAction("Tabelle Drucken", self)
        self.print_act.triggered.connect(self.print_table)

        self.export_html_act = QtWidgets.QAction("HTML Tabellenexport", self)
        self.export_html_act.triggered.connect(self.export_html)

        self.add_description_act = QtWidgets.QAction("Tabellenbeschreibungen hinzufügen", self)
        self.add_description_act.triggered.connect(self.add_descriptions)

        self.check_values_act = QtWidgets.QAction("Sollwerte prüfen", self)
        self.check_values_act.triggered.connect(self.check_all_values)

        self.new_vehicle_act = QtWidgets.QAction("Neues Fahrzeug", self)
        self.new_vehicle_act.triggered.connect(self.new_vehicle)

        self.load_vehicle_act = QtWidgets.QAction("Fahrzeug laden", self)
        self.load_vehicle_act.triggered.connect(self.load_vehicle)

        self.save_vehicle_act = QtWidgets.QAction("Fahrzeug speichern", self)
        self.save_vehicle_act.triggered.connect(self.save_vehicle)

        self.view_vehicle_act = QtWidgets.QAction("Fahrzeug ansehen", self)
        self.view_vehicle_act.triggered.connect(self.view_vehicle)

        self.export_vehicle_act = QtWidgets.QAction("Fahrzeug exportieren", self)
        self.export_vehicle_act.triggered.connect(self.export_vehicle)

        self.manage_service_act = QtWidgets.QAction("Service verwalten", self)
        self.manage_service_act.triggered.connect(self.manage_service)

        self.pending_services_act = QtWidgets.QAction("Anstehende Services", self)
        self.pending_services_act.triggered.connect(self.show_pending_services)

        self.set_tolerance_act = QtWidgets.QAction("Toleranz einstellen", self)
        self.set_tolerance_act.triggered.connect(self.set_tolerance)

        self.ueber_act = QtWidgets.QAction("Über", self)
        self.ueber_act.triggered.connect(self.show_ueber)

        self.view_units_act = QtWidgets.QAction("Maßeinheitenkatalog", self)
        self.view_units_act.triggered.connect(self.show_units_dialog)

        self.defect_reports_act = QtWidgets.QAction("Beanstandungen erfassen", self)
        self.defect_reports_act.triggered.connect(self.show_defect_reports)

        self.storage_manager_act = QtWidgets.QAction("Lagerbestandsverwaltung", self)
        self.storage_manager_act.triggered.connect(self.show_storage_manager)

        self.hu_au_manager_act = QtWidgets.QAction("HU/AU Verwaltung", self)
        self.hu_au_manager_act.triggered.connect(self.show_hu_au_manager)

        self.document_activity_act = QtWidgets.QAction("Aktivität dokumentieren", self)
        self.document_activity_act.triggered.connect(self.document_activity)

        self.view_activities_act = QtWidgets.QAction("Aktivitäten Historie", self)
        self.view_activities_act.triggered.connect(self.view_activities_history)

        self.backup_act = QtWidgets.QAction("Backup/Wiederherstellung", self)
        self.backup_act.triggered.connect(self.show_backup_dialog)
        
        self.quick_backup_act = QtWidgets.QAction("Schnell Backup erstellen", self)
        self.quick_backup_act.triggered.connect(self.create_quick_backup)

        self.save_table_act = QtWidgets.QAction("Tabelle speichern", self)
        self.save_table_act.triggered.connect(self.save_current_table)
        
        self.load_table_act = QtWidgets.QAction("Tabelle laden", self)
        self.load_table_act.triggered.connect(self.load_saved_table)
        
        self.manage_tables_act = QtWidgets.QAction("Tabellen verwalten", self)
        self.manage_tables_act.triggered.connect(self.manage_saved_tables)

        self.advanced_vehicle_search_act = QtWidgets.QAction("Erweiterte Fahrzeugsuche", self)
        self.advanced_vehicle_search_act.triggered.connect(self.advanced_vehicle_search)

    def create_toolbar(self):
        tb = self.addToolBar("Main")
        tb.setMovable(False)
        # css
        toolbutton_style = """
        QToolButton {
            background-color: #1F1F1F; /* #2d2d2d ursprung */
            color: white;
            border: 1px solid #444;
            padding: 4px 11px;
            border-radius: 1px; /* Button rundungen */
            margin: 3px; /* abstand zu anderen elementen */
            font-weight: normal;
            min-width: 60px;
        }
        QToolButton:hover {
            background-color: #0088ff;
            border: 1px solid #00aaff;
        }
        QToolButton:pressed {
            background-color: #0066cc;
        }
        QToolButton::menu-indicator {
            image: none;
            width: 0px;
        }
        """
        
        # Style Menüs
        menu_style = """
        QMenu {
            background-color: #2d2d2d;
            color: white;
            border: 1px solid #444;
            margin: 0px;
        }
        QMenu::item {
            padding: 4px 20px 4px 12px;
            background-color: #2d2d2d;
        }
        QMenu::item:selected {
            background-color: #0088ff;
        }
        QMenu::item:hover {
            background-color: #0088ff;
        }
        QMenu::separator {
            height: 1px;
            background-color: #444;
            margin: 2px 6px;
        }
        """
        
        table_btn = QtWidgets.QToolButton()
        table_btn.setText("Tabellen ⮟")
        table_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        table_btn.setStyleSheet(toolbutton_style)
        table_menu = QtWidgets.QMenu(table_btn)
        table_menu.setStyleSheet(menu_style)
        table_menu.addAction(self.save_table_act)
        table_menu.addAction(self.load_table_act)
        table_menu.addAction(self.manage_tables_act)
        table_btn.setMenu(table_menu)
        tb.addWidget(table_btn)

        template_btn = QtWidgets.QToolButton()
        template_btn.setText("Vorlagen ⮟")
        template_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        template_btn.setStyleSheet(toolbutton_style)
        template_menu = QtWidgets.QMenu(template_btn)
        template_menu.setStyleSheet(menu_style)
        template_menu.addAction(self.new_template_act)
        template_menu.addAction(self.load_template_act)
        template_menu.addAction(self.save_template_act)
        template_btn.setMenu(template_menu)
        tb.addWidget(template_btn)

        dataset_btn = QtWidgets.QToolButton()
        dataset_btn.setText("Datasets ⮟")
        dataset_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        dataset_btn.setStyleSheet(toolbutton_style)
        dataset_menu = QtWidgets.QMenu(dataset_btn)
        dataset_menu.setStyleSheet(menu_style)
        dataset_menu.addAction(self.new_dataset_act)
        dataset_menu.addAction(self.load_dataset_act)
        dataset_menu.addAction(self.save_dataset_act)
        dataset_btn.setMenu(dataset_menu)
        tb.addWidget(dataset_btn)

        export_btn = QtWidgets.QToolButton()
        export_btn.setText("Export ⮟")
        export_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        export_btn.setStyleSheet(toolbutton_style)
        export_menu = QtWidgets.QMenu(export_btn)
        export_menu.setStyleSheet(menu_style)
        export_menu.addAction(self.print_act)
        export_menu.addAction(self.export_html_act)
        export_btn.setMenu(export_menu)
        tb.addWidget(export_btn)

        data_btn = QtWidgets.QToolButton()
        data_btn.setText("Allgemein ⮟")
        data_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        data_btn.setStyleSheet(toolbutton_style)
        data_menu = QtWidgets.QMenu(data_btn)
        data_menu.setStyleSheet(menu_style)
        data_menu.addAction(self.add_description_act)
        data_menu.addAction(self.check_values_act)
        data_menu.addAction(self.set_tolerance_act)
        data_menu.addAction(self.view_units_act)
        data_btn.setMenu(data_menu)
        tb.addWidget(data_btn)

        backup_btn = QtWidgets.QToolButton()
        backup_btn.setText("Backup ⮟")
        backup_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        backup_btn.setStyleSheet(toolbutton_style)
        backup_menu = QtWidgets.QMenu(backup_btn)
        backup_menu.setStyleSheet(menu_style)
        backup_menu.addAction(self.backup_act)
        backup_menu.addAction(self.quick_backup_act)
        backup_btn.setMenu(backup_menu)
        tb.addWidget(backup_btn)


# abstandshalter für unterschied tabelle zu weiterem
        sep_label = QtWidgets.QLabel("|") 
        sep_label.setAlignment(QtCore.Qt.AlignCenter)
        sep_label.setStyleSheet("""
            QLabel {
                background: transparent;
                border: none;
                color: #808080;           /* Grau – dezent */
                font-weight: bold;
                padding: 0px;
                margin-left: 6px;
                margin-right: 6px;
            }
        """)
        sep_label.setFixedHeight(data_btn.sizeHint().height() - 2)
        sep_label.setContentsMargins(0, 0, 0, 0)

        tb.addWidget(sep_label)


        vehicle_btn = QtWidgets.QToolButton()
        vehicle_btn.setText("Fahrzeug ⮟")
        vehicle_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        vehicle_btn.setStyleSheet(toolbutton_style)
        vehicle_menu = QtWidgets.QMenu(vehicle_btn)
        vehicle_menu.setStyleSheet(menu_style)
        vehicle_menu.addAction(self.advanced_vehicle_search_act)
        vehicle_menu.addAction(self.new_vehicle_act)
        vehicle_menu.addAction(self.load_vehicle_act)
        vehicle_menu.addAction(self.save_vehicle_act)
        vehicle_menu.addAction(self.view_vehicle_act)
        vehicle_menu.addAction(self.export_vehicle_act)
        vehicle_btn.setMenu(vehicle_menu)
        vehicle_menu.addAction(self.defect_reports_act)
        vehicle_menu.addAction(self.hu_au_manager_act)
        vehicle_menu.addAction(self.document_activity_act)
        vehicle_menu.addAction(self.view_activities_act)
        tb.addWidget(vehicle_btn)

        service_btn = QtWidgets.QToolButton()
        service_btn.setText("Service ⮟")
        service_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        service_btn.setStyleSheet(toolbutton_style)
        service_menu = QtWidgets.QMenu(service_btn)
        service_menu.setStyleSheet(menu_style)
        service_menu.addAction(self.manage_service_act)
        service_menu.addAction(self.pending_services_act)
        service_btn.setMenu(service_menu)
        tb.addWidget(service_btn)

        storage_btn = QtWidgets.QToolButton()
        storage_btn.setText("Lager ⮟")
        storage_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        storage_btn.setStyleSheet(toolbutton_style)
        storage_menu = QtWidgets.QMenu(storage_btn)
        storage_menu.setStyleSheet(menu_style)
        storage_menu.addAction(self.storage_manager_act)
        storage_btn.setMenu(storage_menu)
        tb.addWidget(storage_btn)

        # tb.addAction(self.ueber_act)  ------ Über button, ohne css. *AUSKOMMENTIERT*
        ueber_btn = QtWidgets.QToolButton()
        ueber_btn.setText("Über")
        ueber_btn.setStyleSheet(toolbutton_style) 
        ueber_btn.clicked.connect(self.show_ueber)
        tb.addWidget(ueber_btn)

# ENDE DROPDOWNS/BUTTONS

    def advanced_vehicle_search(self):
        if not self.vehicles:
            QtWidgets.QMessageBox.information(self, "Keine Fahrzeuge", 
                                            "Es sind keine Fahrzeuge zum Durchsuchen vorhanden.")
            return
            
        dlg = AdvancedVehicleSelectionDialog(self.vehicles, self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            selected_vehicle = dlg.get_selected_vehicle()
            if selected_vehicle:
                self.current_vehicle = selected_vehicle
                self._refresh_vehicle_combo()
                self.vehicle_combo.setCurrentText(selected_vehicle.name)
                self._update_vehicle_display()

    def save_current_table(self):
        if not self.current_vehicle:
            QtWidgets.QMessageBox.warning(self, "Kein Fahrzeug", 
                                        "Bitte wählen Sie zuerst ein Fahrzeug aus, um die Tabelle zu speichern.")
            return
        
        if not self.current_template and self.table.rowCount() == 0:
            QtWidgets.QMessageBox.warning(self, "Keine Daten", 
                                        "Keine Tabellendaten zum Speichern vorhanden.")
            return
        
        dlg = SaveTableDialog(self.current_vehicle, self, self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            self._auto_save_vehicle(self.current_vehicle)

    def load_saved_table(self):
        if not self.current_vehicle:
            QtWidgets.QMessageBox.warning(self, "Kein Fahrzeug", 
                                        "Bitte wählen Sie zuerst ein Fahrzeug aus.")
            return
        
        if not self.current_vehicle.saved_tables:
            QtWidgets.QMessageBox.information(self, "Keine Tabellen", 
                                            "Für dieses Fahrzeug sind keine Tabellen gespeichert.")
            return
        
        dlg = LoadTableDialog(self.current_vehicle, self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            saved_table = dlg.get_selected_table()
            if saved_table:
                self.load_saved_table_data(saved_table)

    def load_saved_table_data(self, saved_table: SavedTable):
        try:
            # Template laden
            template = Template.from_dict(saved_table.template_data)
            self._apply_template_to_ui(template)
            
            # Tabellendaten laden
            self.table.clear()
            self.table.setRowCount(len(saved_table.table_data))
            self.table.setColumnCount(len(template.columns))
            
            # Header setzen
            headers = []
            for c in template.columns:
                unit = c.get('unit', '')
                if unit:
                    headers.append(f"{c.get('name', '')} ({unit})")
                else:
                    headers.append(c.get('name', ''))
            self.table.setHorizontalHeaderLabels(headers)
            
            # Zellendaten laden
            for row, row_data in enumerate(saved_table.table_data):
                for col, cell_data in enumerate(row_data):
                    if col < len(template.columns):
                        item = QtWidgets.QTableWidgetItem(cell_data.get('text', ''))
                        
                        # Formatierung wiederherstellen
                        bg_color = cell_data.get('background')
                        fg_color = cell_data.get('foreground')
                        
                        if bg_color:
                            item.setBackground(QtGui.QColor(bg_color))
                        if fg_color:
                            item.setForeground(QtGui.QColor(fg_color))
                        
                        self.table.setItem(row, col, item)
            
            # Sollwert-Zeile wiederherstellen falls vorhanden
            self._add_setpoint_row(template.columns)
            
            self.current_template = template
            self.template_name_edit.setText(template.name)
            self.description_edit.setPlainText(template.description)
            
            QtWidgets.QMessageBox.information(self, "Erfolg", 
                                            f"Tabelle '{saved_table.name}' wurde geladen.")
            
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Fehler", 
                                        f"Fehler beim Laden der Tabelle: {str(e)}")

    def manage_saved_tables(self):
        if not self.current_vehicle:
            QtWidgets.QMessageBox.warning(self, "Kein Fahrzeug", 
                                        "Bitte wählen Sie zuerst ein Fahrzeug aus.")
            return
        
        print(f"DEBUG: Öffne Tabellenverwaltung für {self.current_vehicle.name}")
        print(f"DEBUG: Anzahl gespeicherter Tabellen: {len(self.current_vehicle.saved_tables)}")
        
        # Dialog öffnen
        dlg = LoadTableDialog(self.current_vehicle, self)
        result = dlg.exec_()
        
        if result == QtWidgets.QDialog.Accepted:
            # Tabelle wurde geladen
            saved_table = dlg.get_selected_table()
            if saved_table:
                self.load_saved_table_data(saved_table)
        else:
            self._refresh_vehicle_combo()
            print(f"DEBUG: Tabellen nach Dialog: {len(self.current_vehicle.saved_tables)}")

    def show_backup_dialog(self):
        dlg = BackupDialog(self)
        dlg.exec_()

    def create_quick_backup(self):
        # Pfad vorschlagen
        default_path = self.backup_manager.backup_dir / f"o3measurement_quick_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        
        fname, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Schnell-Backup speichern unter",
            str(default_path),
            "Zip-Dateien (*.zip)"
        )
        
        if not fname:
            return  # wenn abgebrochen
        
        # Backup erstellen
        result = self.backup_manager.create_full_backup(backup_path=fname)
        
        if result["success"]:
            QtWidgets.QMessageBox.information(
                self, 
                "Backup erfolgreich", 
                f"Schnell-Backup wurde erstellt:\n{result['backup_path']}\n\n"
                f"{result['file_count']} Dateien, {result['total_size'] / 1024 / 1024:.2f} MB"
            )
        else:
            QtWidgets.QMessageBox.warning(
                self,
                "Backup fehlgeschlagen",
                f"Backup konnte nicht erstellt werden:\n{result['error']}"
            )


    def view_activities_history(self):
        if not self.current_vehicle:
            QtWidgets.QMessageBox.warning(self, "Kein Fahrzeug", "Bitte zuerst ein Fahrzeug auswählen")
            return
        
        dlg = ActivitiesHistoryDialog(self.current_vehicle, self)
        dlg.exec_()

    def document_activity(self):
        if not self.current_vehicle:
            QtWidgets.QMessageBox.warning(self, "Kein Fahrzeug", "Bitte zuerst ein Fahrzeug auswählen")
            return
        
        dlg = ActivityDocumentationDialog(self.current_vehicle, self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            # Auto-Save-Flag setzen
            self.auto_save_in_progress = True
            try:
                self._auto_save_vehicle(self.current_vehicle)
            finally:
                self.auto_save_in_progress = False

    def _create_main_ui(self):
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        layout = QtWidgets.QHBoxLayout(central)

        left = QtWidgets.QFrame()
        left.setMinimumWidth(400)
        left_layout = QtWidgets.QVBoxLayout(left)
        left_layout.setContentsMargins(8, 8, 8, 8)

        # Fahrzeug bereich
        vehicle_group = QtWidgets.QGroupBox("Fahrzeug")
        vehicle_layout = QtWidgets.QVBoxLayout()
        
        # Fahrzeug dropdown
        vehicle_select_layout = QtWidgets.QHBoxLayout()
        vehicle_select_layout.addWidget(QtWidgets.QLabel("Fahrzeug:"))
        
        self.vehicle_combo = QtWidgets.QComboBox()
        self.vehicle_combo.setEditable(True)
        self.vehicle_combo.completer().setCompletionMode(QtWidgets.QCompleter.PopupCompletion)
        self._refresh_vehicle_combo()
        self.vehicle_combo.currentTextChanged.connect(self.on_vehicle_selected)
        vehicle_select_layout.addWidget(self.vehicle_combo)
        
        # Fahrzeugbuttons oben links
        vehicle_btn_layout = QtWidgets.QHBoxLayout()
        new_vehicle_btn = QtWidgets.QPushButton("Neu")
        new_vehicle_btn.clicked.connect(self.new_vehicle)
        edit_vehicle_btn = QtWidgets.QPushButton("Bearbeiten")
        edit_vehicle_btn.clicked.connect(self.edit_vehicle)
        view_vehicle_btn = QtWidgets.QPushButton("Ansehen")
        view_vehicle_btn.clicked.connect(self.view_vehicle)
        search_vehicle_btn = QtWidgets.QPushButton("Erweiterte Fahrzeugsuche")
        search_vehicle_btn.clicked.connect(self.advanced_vehicle_search)
        # btns            
        vehicle_btn_layout.addWidget(new_vehicle_btn)
        vehicle_btn_layout.addWidget(edit_vehicle_btn)
        vehicle_btn_layout.addWidget(view_vehicle_btn)
        vehicle_btn_layout.addWidget(search_vehicle_btn)
        vehicle_btn_layout.addStretch()
        
        vehicle_layout.addLayout(vehicle_select_layout)
        vehicle_layout.addLayout(vehicle_btn_layout)
        
        self.vehicle_name_label = QtWidgets.QLabel("Kein Fahrzeug ausgewählt")
        self.vehicle_name_label.setFont(QtGui.QFont(None, 10, QtGui.QFont.Bold))
        vehicle_layout.addWidget(self.vehicle_name_label)
        
        self.vehicle_desc_label = QtWidgets.QLabel("")
        self.vehicle_desc_label.setWordWrap(True)
        vehicle_layout.addWidget(self.vehicle_desc_label)
        
        vehicle_group.setLayout(vehicle_layout)
        left_layout.addWidget(vehicle_group)

        # Vorlagen
        template_group = QtWidgets.QGroupBox("Vorlagen / Spaltenkonfiguration")
        template_layout = QtWidgets.QVBoxLayout()

        self.template_name_edit = QtWidgets.QLineEdit()
        self.template_name_edit.setPlaceholderText("Vorlagenname")
        template_layout.addWidget(self.template_name_edit)

        self.description_edit = QtWidgets.QTextEdit()
        self.description_edit.setMaximumHeight(60)
        self.description_edit.setPlaceholderText("Vorlagenbeschreibung (optional)")
        template_layout.addWidget(self.description_edit)

        self.columns_list = QtWidgets.QListWidget()
        self.columns_list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        template_layout.addWidget(self.columns_list, 1)

        # Vorlagen cfg buttons
        col_controls = QtWidgets.QHBoxLayout()
        self.add_col_btn = QtWidgets.QPushButton("+ Spalte")
        self.add_col_btn.clicked.connect(self.add_column)
        self.edit_col_btn = QtWidgets.QPushButton("Eigenschaften")
        self.edit_col_btn.clicked.connect(self.edit_column)
        self.rem_col_btn = QtWidgets.QPushButton("- Spalte")
        self.rem_col_btn.clicked.connect(self.remove_column)
        col_controls.addWidget(self.add_col_btn)
        col_controls.addWidget(self.edit_col_btn)
        col_controls.addWidget(self.rem_col_btn)
        template_layout.addLayout(col_controls)

        # zeilen verschiebung
        move_controls = QtWidgets.QHBoxLayout()
        self.up_btn = QtWidgets.QPushButton("▲")
        self.up_btn.clicked.connect(self.move_column_up)
        self.down_btn = QtWidgets.QPushButton("▼")
        self.down_btn.clicked.connect(self.move_column_down)
        move_controls.addWidget(self.up_btn)
        move_controls.addWidget(self.down_btn)
        template_layout.addLayout(move_controls)

        template_layout.addWidget(QtWidgets.QLabel("Verfügbare Vorlagen:"))
        self.templates_combo = QtWidgets.QComboBox()
        self.templates_combo.setEditable(False)
        template_layout.addWidget(self.templates_combo)
        load_selected_btn = QtWidgets.QPushButton("Ausgewählte Vorlage laden")
        load_selected_btn.clicked.connect(self.load_selected_template)
        template_layout.addWidget(load_selected_btn)

        template_group.setLayout(template_layout)
        left_layout.addWidget(template_group, 1)

        layout.addWidget(left)

        right = QtWidgets.QFrame()
        right_layout = QtWidgets.QVBoxLayout(right)
        right_layout.setContentsMargins(8, 8, 8, 8)

        header = QtWidgets.QHBoxLayout()
        title = QtWidgets.QLabel(f"{o3NAME} - Workspace")
        title.setFont(QtGui.QFont(None, 12, QtGui.QFont.Bold))
        header.addWidget(title)
        header.addStretch()
        right_layout.addLayout(header)

        self.table = QtWidgets.QTableWidget()
        self.table.setAlternatingRowColors(False)
        self.table.verticalHeader().setVisible(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.cellChanged.connect(self.on_cell_changed)
        right_layout.addWidget(self.table, 1)

        row_ops = QtWidgets.QHBoxLayout()
        add_row = QtWidgets.QPushButton("+ Zeile")
        add_row.clicked.connect(self.add_row)
        rem_row = QtWidgets.QPushButton("- Zeile")
        rem_row.clicked.connect(self.remove_row)
        fill_defaults = QtWidgets.QPushButton("Sollwerte/Defaults einfügen")
        fill_defaults.clicked.connect(self.fill_setpoints)

        row_ops.addWidget(add_row)
        row_ops.addWidget(rem_row)
        row_ops.addWidget(fill_defaults)

        text_label = QtWidgets.QLabel(f"{o3NAME} {o3VERSION} | © openw3rk INVENT")
        row_ops.addWidget(text_label)

        row_ops.insertStretch(3)  # vor dem Text

        right_layout.addLayout(row_ops)
        layout.addWidget(right, 1)

        self.columns_list.currentItemChanged.connect(self.on_column_selected)
        self._refresh_templates_list()

    def new_template(self):
        if self.unsaved_changes:
            reply = self.show_unsaved_changes_dialog()
            if reply == QtWidgets.QMessageBox.Save:
                if not self.save_all_changes():
                    return  # Abbruch wenn Speichern fehlschlägt
            elif reply == QtWidgets.QMessageBox.Cancel:
                return  # Abbruch
        
        # Original Code fortsetzen
        self.template_name_edit.clear()
        self.description_edit.clear()
        self.columns_list.clear()
        self.current_template = None
        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.setpoint_row_index = -1
        self.export_short_description = ""
        self.export_long_description = ""
        self.global_tolerance = 0.0
        self.clear_unsaved_changes()

    def load_template(self):
        if self.unsaved_changes:
            reply = self.show_unsaved_changes_dialog()
            if reply == QtWidgets.QMessageBox.Save:
                if not self.save_all_changes():
                    return
            elif reply == QtWidgets.QMessageBox.Cancel:
                return
        
        fname, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Vorlage öffnen", str(TEMPLATES_DIR), "JSON Dateien (*.json)")
        if not fname:
            return
        with open(fname, "r", encoding="utf-8") as f:
            d = json.load(f)
        tmpl = Template.from_dict(d)
        self._apply_template_to_ui(tmpl)
        self.clear_unsaved_changes()

    def save_template(self):
        tmpl = self._gather_template_from_ui()
        name, ok = QtWidgets.QInputDialog.getText(self, 'Vorlage speichern', 'Vorlagenname:', text=tmpl.name)
        if not ok:
            return
        tmpl.name = name.strip() or tmpl.name
        fname = TEMPLATES_DIR / (tmpl.name + ".json")
        with open(fname, "w", encoding="utf-8") as f:
            json.dump(tmpl.to_dict(), f, indent=2, ensure_ascii=False)
        QtWidgets.QMessageBox.information(self, "Vorlage gespeichert", f"Vorlage gespeichert: {fname}")
        self._refresh_templates_list()
        self.clear_unsaved_changes()

    def new_dataset(self):
        if self.unsaved_changes:
            reply = self.show_unsaved_changes_dialog()
            if reply == QtWidgets.QMessageBox.Save:
                if not self.save_all_changes():
                    return
            elif reply == QtWidgets.QMessageBox.Cancel:
                return
        
        tmpl = self._gather_template_from_ui()
        if not tmpl.columns:
            QtWidgets.QMessageBox.warning(self, "Keine Spalten", "Bitte zuerst eine Vorlage mit Spalten anlegen")
            return
        self.current_template = tmpl
        self._apply_template_to_table(tmpl)
        self.clear_unsaved_changes()

    def load_dataset(self):
        if self.unsaved_changes:
            reply = self.show_unsaved_changes_dialog()
            if reply == QtWidgets.QMessageBox.Save:
                if not self.save_all_changes():
                    return
            elif reply == QtWidgets.QMessageBox.Cancel:
                return
        
        fname, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Dataset laden", str(DATA_DIR), "CSV Dateien (*.csv);;JSON Dateien (*.json)")
        if not fname:
            return
        if fname.lower().endswith('.csv'):
            self._load_csv(fname)
        else:
            self._load_json(fname)
        self.clear_unsaved_changes()

    def save_dataset(self):
        fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Dataset speichern", str(DATA_DIR), "CSV Dateien (*.csv);;JSON Dateien (*.json)")
        if not fname:
            return
        if fname.lower().endswith('.csv'):
            self._save_csv(fname)
        else:
            self._save_json(fname)
        QtWidgets.QMessageBox.information(self, "Gespeichert", f"Dataset gespeichert: {fname}")
        self.clear_unsaved_changes()

    def print_table(self):
        html = self._generate_html_for_print()
        printer = QPrinter()
        dlg = QPrintDialog(printer, self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            doc = QtGui.QTextDocument()
            doc.setHtml(html)
            doc.print_(printer)

    def export_html(self):
        html = self._generate_html_for_print()
        fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, "HTML speichern", str(DATA_DIR / "export.html"), "HTML Dateien (*.html)")
        if not fname:
            return
        with open(fname, 'w', encoding='utf-8') as f:
            f.write(html)
        QtWidgets.QMessageBox.information(self, "Export erfolgreich", f"HTML wurde gespeichert: {fname}")

    def add_descriptions(self):
        dlg = ExportDescriptionDialog(self.export_short_description, self.export_long_description, self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            self.export_short_description, self.export_long_description = dlg.get_descriptions()

    def check_all_values(self):
        if not self.current_template:
            QtWidgets.QMessageBox.warning(self, "Keine Vorlage", "Bitte zuerst eine Vorlage laden")
            return
        
        for row in range(self.table.rowCount()):
            if row == self.setpoint_row_index:  # Sollwert zeile überspringen
                continue
            
            for column in range(self.table.columnCount()):
                if column >= len(self.current_template.columns):
                    continue
                    
                col_info = self.current_template.columns[column]
                setpoint_value = col_info.get('setpoint')
            
                if setpoint_value is None:  # Nur Spalten mit Sollwerten prüfen
                    continue
                
                cell_item = self.table.item(row, column)
                if not cell_item:
                    continue
                
                actual_value = cell_item.text()
            
                tolerance = col_info.get('tolerance', self.global_tolerance)
            
                if actual_value.strip():  # Nur prüfen wenn Wert vorhanden
                    actual_clean = self._remove_unit(actual_value)
                    setpoint_clean = self._remove_unit(setpoint_value)
                
                    is_ok = self.check_value(actual_clean, setpoint_clean, tolerance)
                
                    unit = col_info.get('unit', '')
                    if unit and actual_clean:
                        formatted_value = self._format_value_with_unit(actual_clean, unit)
                        cell_item.setText(formatted_value)
                
                    if tolerance > 0 or col_info.get('always_highlight', False):
                    # Farbe setzen basierend auf Prüfung
                        if is_ok:
                            cell_item.setBackground(QtGui.QColor('#1e3a1e'))  
                            cell_item.setForeground(QtGui.QColor('#90ee90'))  
                        else:
                            cell_item.setBackground(QtGui.QColor('#3a1e1e'))  
                            cell_item.setForeground(QtGui.QColor('#ff6b6b'))  
                    else:
                        cell_item.setBackground(QtGui.QColor(0, 0, 0, 0))     
                        cell_item.setForeground(QtGui.QColor(255, 255, 255))  

    def set_tolerance(self):
        dlg = ToleranceDialog(self.global_tolerance, self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            self.global_tolerance = dlg.get_tolerance()
            QtWidgets.QMessageBox.information(self, "Toleranz gesetzt", 
                                            f"Globale Toleranz wurde auf {self.global_tolerance}% eingestellt.\n"
                                            f"Klicken Sie auf 'Sollwerte prüfen' um alle Werte neu zu bewerten.")

    def show_ueber(self):
        dlg = UeberWindow(self)
        dlg.exec_()


    def _load_all_vehicles(self):
        vehicles = []
        if VEHICLES_BASE_DIR.exists():
            for vehicle_dir in VEHICLES_BASE_DIR.iterdir():
                if vehicle_dir.is_dir():
                    vehicle_file = vehicle_dir / "vehicle.json"
                    if vehicle_file.exists():
                        try:
                            with open(vehicle_file, "r", encoding="utf-8") as f:
                                data = json.load(f)
                            vehicle = Vehicle.from_dict(data)
                            vehicles.append(vehicle)
                        except Exception as e:
                            print(f"Fehler beim Laden von {vehicle_file}: {e}")
        return vehicles

    def _refresh_vehicle_combo(self):
        self.vehicle_combo.clear()
        
        # Prüfen ob vehicles existiert
        if not hasattr(self, 'vehicles') or not self.vehicles:
            self.vehicles = self._load_all_vehicles()
            
        if not self.vehicles:
            return
        
        # ComboBox konfigurieren für bessere Suche
        self.vehicle_combo.setEditable(True)
        self.vehicle_combo.setInsertPolicy(QtWidgets.QComboBox.NoInsert)
        
        # Auto-Vervollständigung
        completer = QtWidgets.QCompleter([v.name for v in self.vehicles])
        completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        completer.setFilterMode(QtCore.Qt.MatchContains)
        completer.setCompletionMode(QtWidgets.QCompleter.PopupCompletion)
        self.vehicle_combo.setCompleter(completer)
        
        for vehicle in self.vehicles:
            self.vehicle_combo.addItem(vehicle.name, vehicle)


    def on_vehicle_selected(self, vehicle_name):
        if not vehicle_name:
            return
            
        # Finde das entsprechende Vehicle-Objekt
        for vehicle in self.vehicles:
            if vehicle.name == vehicle_name:
                self.current_vehicle = vehicle
                self._update_vehicle_display()
                break

    # Fahrzeugfunktionen
    def new_vehicle(self):
        dlg = VehicleDialog()
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            # Auto-Save-Flag setzen
            self.auto_save_in_progress = True
            try:
                new_vehicle = dlg.get_vehicle()
                if not new_vehicle.name:
                    QtWidgets.QMessageBox.warning(self, "Fehler", "Bitte geben Sie einen Fahrzeugnamen ein.")
                    return
                    
                self.vehicles.append(new_vehicle)
                self.current_vehicle = new_vehicle
                self._refresh_vehicle_combo()
                self.vehicle_combo.setCurrentText(new_vehicle.name)
                self._update_vehicle_display()
                # Automatisch speichern
                self._auto_save_vehicle(new_vehicle)
            finally:
                self.auto_save_in_progress = False

    def edit_vehicle(self):
        if not self.current_vehicle:
            QtWidgets.QMessageBox.warning(self, "Kein Fahrzeug", "Bitte zuerst ein Fahrzeug auswählen")
            return
            
        old_name = self.current_vehicle.name
        dlg = VehicleDialog(self.current_vehicle)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            # Save-Flag setzen
            self.auto_save_in_progress = True
            try:
                updated_vehicle = dlg.get_vehicle()
                
                if old_name != updated_vehicle.name:
                    old_dir = VEHICLES_BASE_DIR / old_name
                    if old_dir.exists():
                        temp_dir = VEHICLES_BASE_DIR / f"temp_{old_name}"
                        old_dir.rename(temp_dir)
                        
                        new_dir = VEHICLES_BASE_DIR / updated_vehicle.name
                        if new_dir.exists():
                            QtWidgets.QMessageBox.warning(self, "Fehler", f"Ein Fahrzeug mit dem Namen '{updated_vehicle.name}' existiert bereits.")
                            temp_dir.rename(old_dir)  # Zurück benennen
                            return
                        
                        # Endgültig umbenennen
                        temp_dir.rename(new_dir)
                
                for i, vehicle in enumerate(self.vehicles):
                    if vehicle.name == old_name:
                        self.vehicles[i] = updated_vehicle
                        self.current_vehicle = updated_vehicle
                        break
                        
                self._refresh_vehicle_combo()
                self.vehicle_combo.setCurrentText(updated_vehicle.name)
                self._update_vehicle_display()
                # Automatisch speichern
                self._auto_save_vehicle(updated_vehicle)
            finally:
                self.auto_save_in_progress = False

    def view_vehicle(self):
        if not self.current_vehicle:
            QtWidgets.QMessageBox.warning(self, "Kein Fahrzeug", "Bitte zuerst ein Fahrzeug auswählen")
            return
            
        dlg = VehicleViewDialog(self.current_vehicle, self)
        dlg.exec_()

    def _auto_save_vehicle(self, vehicle):
        if not vehicle.name:
            return
            
        try:
            # Auto-Save-Flag setzen
            self.auto_save_in_progress = True
            
            # Fahrzeugverzeichnis
            vehicle.ensure_directories()
            
            # Sicherung Fahrzeugdaten
            vehicle_file = vehicle.get_vehicle_dir() / "vehicle.json"
            with open(vehicle_file, "w", encoding="utf-8") as f:
                json.dump(vehicle.to_dict(), f, indent=2, ensure_ascii=False)
            print(f"Fahrzeug automatisch gespeichert: {vehicle_file}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Auto-Save Fehler", f"Fehler beim automatischen Speichern: {e}")
        finally:
            # Save-Flag zurücksetzen
            self.auto_save_in_progress = False


    def load_vehicle(self):
        vehicle_names = [v.name for v in self.vehicles]
        if not vehicle_names:
            QtWidgets.QMessageBox.information(self, "Info", "Keine Fahrzeuge zum Laden verfügbar.")
            return
            
        name, ok = QtWidgets.QInputDialog.getItem(self, "Fahrzeug laden", "Fahrzeug auswählen:", vehicle_names, 0, False)
        if ok and name:
            for vehicle in self.vehicles:
                if vehicle.name == name:
                    self.current_vehicle = vehicle
                    self._refresh_vehicle_combo()
                    self.vehicle_combo.setCurrentText(vehicle.name)
                    self._update_vehicle_display()
                    break

    def save_vehicle(self):
        if not self.current_vehicle:
            QtWidgets.QMessageBox.warning(self, "Kein Fahrzeug", "Bitte zuerst ein Fahrzeug erstellen oder laden")
            return
        
        self._auto_save_vehicle(self.current_vehicle)
        QtWidgets.QMessageBox.information(self, "Fahrzeug gespeichert", f"Fahrzeug '{self.current_vehicle.name}' wurde gespeichert.")
        self.clear_unsaved_changes()
    def export_vehicle(self):
        if not self.current_vehicle:
            QtWidgets.QMessageBox.warning(self, "Kein Fahrzeug", "Bitte zuerst ein Fahrzeug erstellen oder laden")
            return
        
        include_images = self.current_vehicle.include_attachments_in_export
        
        if include_images:
            folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Export-Ordner auswählen", str(VEHICLES_BASE_DIR))
            if not folder:
                return
            export_dir = Path(folder) / f"{self.current_vehicle.name}_export"
            export_dir.mkdir(exist_ok=True)
            
            # Bilder kopieren, wenn für Export freigegeben
            if include_images:
                for attachment in self.current_vehicle.attachments:
                    if 'scan_path' in attachment and Path(attachment['scan_path']).exists():
                        shutil.copy2(attachment['scan_path'], export_dir / attachment['filename'])
            
            html_path = export_dir / f"{self.current_vehicle.name}.html"
            html = self._generate_vehicle_html(include_images=include_images, export_dir=export_dir)
        else:
            fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Fahrzeug exportieren", str(VEHICLES_BASE_DIR / f"{self.current_vehicle.name}.html"), "HTML Dateien (*.html)")
            if not fname:
                return
            html_path = Path(fname)
            html = self._generate_vehicle_html(include_images=False)
        
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html)
        QtWidgets.QMessageBox.information(self, "Export erfolgreich", f"Fahrzeugbericht wurde exportiert: {html_path}")

# html generieren
    def _generate_vehicle_html(self, include_images=False, export_dir=None):
        if not self.current_vehicle:
            return ""

        html = [f"""
            <!--
            
            ***********************************************************************
            * o3Measurement Version {o3VERSION} by openw3rk INVENT - VEHICLE SOLUTIONS *
            *             https://openw3rk.de | https://vs.openw3rk.de            *
            *                  - Copyright (c) openw3rk INVENT -                  *
            ***********************************************************************
            
            -->
            <!--
            
            INFORMATION:                                                                  
            o3Measurement EXPORT | VEHICLE: {self.current_vehicle.name}                    
            o3Measurement EXPORT | VEHICLE DESCRIPTION: {self.current_vehicle.description} 
            
            -->
            
        <html>
        <head>
            <meta charset='utf-8'>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                .header {{ background: #2c3e50; color: white; padding: 20px; border-radius: 5px; }}
                .section {{ margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px; }}
                .specs-table {{ width: 100%; border-collapse: collapse; }}
                .specs-table th, .specs-table td {{ padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }}
                .specs-table th {{ background: #f2f2f2; }}
                .fluids-table, .parts-table {{ width: 100%; border-collapse: collapse; }}
                .fluids-table th, .fluids-table td, .parts-table th, .parts-table td {{ padding: 8px; text-align: left; border: 1px solid #ddd; }}
                .attachments {{ display: flex; flex-wrap: wrap; gap: 10px; }}
                .attachment {{ border: 1px solid #ddd; padding: 10px; border-radius: 5px; }}
                .attachment img {{
                    max-width: 200px;
                    max-height: 200px;
                    cursor: pointer;
                    transition: transform 0.2s;
                }}
                /* Modal (großes Bild) */
                .modal {{
                    display: none;
                    position: fixed;
                    z-index: 9999;
                    padding-top: 50px;
                    left: 0;
                    top: 0;
                    width: 100%;
                    height: 100%;
                    background-color: rgba(0,0,0,0.8);
                }}
                .modal img {{
                    margin: auto;
                    display: block;
                    max-width: 90%;
                    max-height: 90%;
                    border-radius: 5px;
                }}
            </style>
        </head>
        <body>
            <div class='header'>
                <h1>{self.current_vehicle.name}</h1>
                <p>{self.current_vehicle.description}</p>
            </div>

            
            <div id="imgModal" class="modal" onclick="this.style.display='none'">
                <img id="modalImg">
            </div>

            <script>
            function showModal(src) {{
                var modal = document.getElementById("imgModal");
                var modalImg = document.getElementById("modalImg");
                modal.style.display = "block";
                modalImg.src = src;
            }}
            </script>
        """]
    
        # Spezifikationen
        html.append("<div class='section'>")
        html.append("<h2>Fahrzeugspezifikationen</h2>")
        html.append("<table class='specs-table'>")
        specs = self.current_vehicle.specifications
        for key, value in specs.items():
            if key != 'fluessigkeiten' and value:
                display_key = key.capitalize()
                html.append(f"<tr><th>{display_key}</th><td>{value}</td></tr>")
        html.append("</table>")
        html.append("</div>")
    
        # Flüssigkeiten
        fluids = specs.get('fluessigkeiten', {})
        if fluids:
            html.append("<div class='section'>")
            html.append("<h2>Flüssigkeiten</h2>")
            html.append("<table class='fluids-table'>")
            html.append("<tr><th>Flüssigkeit</th><th>Spezifikation</th><th>Menge</th></tr>")
            for fluid_name, fluid_spec in fluids.items():
                html.append(f"<tr><td>{fluid_name}</td><td>{fluid_spec.get('spezifikation', '')}</td><td>{fluid_spec.get('menge', '')}</td></tr>")
            html.append("</table>")
            html.append("</div>")
        
        if self.current_vehicle.parts:
            html.append("<div class='section'>")
            html.append("<h2>Bauteile und Teilenummern</h2>")
            html.append("<table class='parts-table'>")
            html.append("<tr><th>Bauteil</th><th>Teilenummer</th><th>Motor Code</th></tr>")
            for part in self.current_vehicle.parts:
                html.append(f"<tr><td>{part.get('bauteil', '')}</td><td>{part.get('teilenummer', '')}</td><td>{part.get('motor_code', '')}</td></tr>")
            html.append("</table>")
            html.append("</div>")
    
        if self.current_vehicle.service_history or self.current_vehicle.service_intervals:
            html.append("<div class='section'>")
            html.append("<h2>Service-Informationen</h2>")
        
            if self.current_vehicle.last_service.get('ölwechsel_km'):
                html.append(f"<p><strong>Letzter Ölwechsel:</strong> {self.current_vehicle.last_service.get('ölwechsel_km')} km am {self.current_vehicle.last_service.get('ölwechsel_datum', '')}</p>")
        
            if self.current_vehicle.service_intervals:
                html.append("<h3>Service-Intervalle</h3>")
                html.append("<table class='fluids-table'>")
                html.append("<tr><th>Service</th><th>Intervall</th></tr>")
                for service, interval in self.current_vehicle.service_intervals.items():
                    unit = SERVICE_TYPES.get(service, {}).get('unit', 'km')
                    html.append(f"<tr><td>{service}</td><td>{interval} {unit}</td></tr>")
                html.append("</table>")
        
            html.append("</div>")
    
        # Anhänge/scans
        if self.current_vehicle.attachments and include_images:
            html.append("<div class='section'>")
            html.append("<h2>Dokumente und Scans</h2>")
            html.append("<div class='attachments'>")
            for attachment in self.current_vehicle.attachments:
                if attachment['mimetype'] == 'image':
                    if include_images and export_dir:
                        img_src = attachment['filename']
                    else:
                        img_src = f"data:image/jpeg;base64,{attachment['data']}"
                
                    html.append(f"""
                    <div class='attachment'>
                        <img src='{img_src}' alt='{attachment['filename']}' onclick="showModal(this.src)"><br>
                        {attachment['filename']}
                    </div>
                    """)
                else:
                    html.append(f"""
                    <div class='attachment'>
                        <strong>{attachment['filename']}</strong><br>
                        ({attachment['mimetype']})
                    </div>
                    """)
            html.append("</div>")
            html.append("</div>")
        html.append(f"""
        <div class='footer'>
            <h5>Erstellt mit:<br>{o3NAME} Version {o3VERSION}<br>openw3rk INVENT - Vehicle Solutions</h5>
            <p>Zeitpunkt: <span id="now"></span><script>document.getElementById('now').textContent=new Date().toLocaleString()</script>
        </div>
        """)
        html.append("</body></html>")
        return ''.join(html)

    def _update_vehicle_display(self):
        if self.current_vehicle:
            self.vehicle_name_label.setText(self.current_vehicle.name)
            self.vehicle_desc_label.setText(self.current_vehicle.description)
        else:
            self.vehicle_name_label.setText("Kein Fahrzeug ausgewählt")
            self.vehicle_desc_label.setText("")

    def manage_service(self):
        if not self.current_vehicle:
            QtWidgets.QMessageBox.warning(self, "Kein Fahrzeug", "Bitte zuerst ein Fahrzeug auswählen")
            return
            
        dlg = ServiceDialog(self.current_vehicle, self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            # Auto-Save-Flag setzen bevor Vehicle aktualisiert wird
            self.auto_save_in_progress = True
            try:
                updated_vehicle = dlg.get_service_data()
                self.current_vehicle = updated_vehicle
                for i, vehicle in enumerate(self.vehicles):
                    if vehicle.name == self.current_vehicle.name:
                        self.vehicles[i] = updated_vehicle
                        break
                self._auto_save_vehicle(updated_vehicle)
            finally:
                self.auto_save_in_progress = False


    def show_pending_services(self):
        dlg = PendingServicesDialog(self.vehicles, self)
        dlg.exec_()

    def _refresh_templates_list(self):
        self.templates_combo.clear()
        for p in sorted(TEMPLATES_DIR.glob("*.json")):
            self.templates_combo.addItem(p.stem)

    def check_value(self, actual_value, setpoint_value, tolerance_percent):
        try:
            # Entferne Einheiten für den Vergleich
            actual_clean = self._remove_unit(actual_value)
            setpoint_clean = self._remove_unit(setpoint_value)
            
            actual = float(actual_clean)
            setpoint = float(setpoint_clean)
            
            if tolerance_percent == 0:
                # Exakter Vergleich
                return actual == setpoint
            else:
                # Toleranz-basierter Vergleich
                tolerance_abs = abs(setpoint) * (tolerance_percent / 100.0)
                lower_bound = setpoint - tolerance_abs
                upper_bound = setpoint + tolerance_abs
                return lower_bound <= actual <= upper_bound
                
        except (ValueError, TypeError):
            # Wenn Werte nicht in Zahlen umwandelbar sind
            return False

    def _remove_unit(self, value):
        if not value:
            return value
        return value.split(' ')[0]

    def _format_value_with_unit(self, value, unit):
        if not value or not unit:
            return value
        return f"{value} {unit}"

    def on_cell_changed(self, row, column):
        if row == self.setpoint_row_index:  
            return
        
        if not self.current_template or column >= len(self.current_template.columns):
            return
        
        col_info = self.current_template.columns[column]
        setpoint_value = col_info.get('setpoint')
    
        if setpoint_value is None:  # Nur Spalten mit Sollwerten prüfen
            return
        
        cell_item = self.table.item(row, column)
        if not cell_item:
            return
        
        actual_value = cell_item.text()
    
        tolerance = col_info.get('tolerance', self.global_tolerance)
    
        if actual_value.strip():  # Nur prüfen wenn Wert vorhanden
            actual_clean = self._remove_unit(actual_value)
            setpoint_clean = self._remove_unit(setpoint_value)
        
            is_ok = self.check_value(actual_clean, setpoint_clean, tolerance)
        
            unit = col_info.get('unit', '')
            if unit and actual_clean:
                formatted_value = self._format_value_with_unit(actual_clean, unit)
                cell_item.setText(formatted_value)
        
            if tolerance > 0 or col_info.get('always_highlight', False):
                if is_ok:
                    cell_item.setBackground(QtGui.QColor('#1e3a1e'))  
                    cell_item.setForeground(QtGui.QColor('#90ee90'))  
                else:
                    cell_item.setBackground(QtGui.QColor('#3a1e1e'))  
                    cell_item.setForeground(QtGui.QColor('#ff6b6b'))  
            else:
                cell_item.setBackground(QtGui.QColor(0, 0, 0, 0))     
                cell_item.setForeground(QtGui.QColor(255, 255, 255))  

    def add_column(self):
        name, ok = QtWidgets.QInputDialog.getText(self, "Spaltenname", "Name der neuen Spalte:")
        if not ok or not name.strip():
            return
        col = {"name": name.strip(), "type": "Text", "setpoint": None, "default": None, "readonly": True, "width": 120, "tolerance": 0.0, "unit": ""}
        item = QtWidgets.QListWidgetItem(col["name"])
        item.setData(QtCore.Qt.UserRole, col)
        self.columns_list.addItem(item)

    def edit_column(self):
        row = self.columns_list.currentRow()
        if row < 0:
            QtWidgets.QMessageBox.warning(self, "Keine Spalte", "Bitte zuerst eine Spalte auswählen")
            return
        
        item = self.columns_list.item(row)
        col = item.data(QtCore.Qt.UserRole)
        dlg = ColumnPropertiesDialog(col, self)
        
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            newcol = dlg.col
            item.setText(newcol.get('name', ''))
            item.setData(QtCore.Qt.UserRole, newcol)
            
            if self.current_template:
                # Nur diese eine Spalte aktualisieren statt komplette Tabelle
                self._update_single_column(row, newcol)

    def _update_single_column(self, column_index, new_col_info):        
        # Header aktualisieren
        unit = new_col_info.get('unit', '')
        if unit:
            header_text = f"{new_col_info.get('name', '')} ({unit})"
        else:
            header_text = new_col_info.get('name', '')
        
        self.table.setHorizontalHeaderItem(column_index, QtWidgets.QTableWidgetItem(header_text))
        
        # Spaltenbreite aktualisieren
        width = new_col_info.get('width', 120)
        self.table.setColumnWidth(column_index, width)
        
        # Sollwert-Zeile aktualisieren falls vorhanden
        if self.setpoint_row_index >= 0:
            setpoint_item = self.table.item(self.setpoint_row_index, column_index)
            if setpoint_item:
                setpoint_value = new_col_info.get('setpoint', '')
                unit = new_col_info.get('unit', '')
                
                if unit and setpoint_value:
                    display_value = f"{setpoint_value} {unit}"
                else:
                    display_value = str(setpoint_value) if setpoint_value else ''
                
                setpoint_item.setText(display_value)
                
                # Schreibschutz entsprechend der Spalteneinstellung
                if new_col_info.get('readonly', True):
                    setpoint_item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                else:
                    setpoint_item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsEditable)
        
        # Template aktualisieren
        if self.current_template and column_index < len(self.current_template.columns):
            self.current_template.columns[column_index] = new_col_info

    def remove_column(self):
        row = self.columns_list.currentRow()
        if row >= 0:
            # Bestätigung einholen
            reply = QtWidgets.QMessageBox.question(
                self, 
                "Spalte löschen", 
                "Soll diese Spalte wirklich gelöscht werden?\nAlle Werte in dieser Spalte gehen verloren.",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
            )
            
            if reply == QtWidgets.QMessageBox.Yes:
                # Spalte aus Liste entfernen
                self.columns_list.takeItem(row)
                
                # Template aktualisieren
                if self.current_template and row < len(self.current_template.columns):
                    self.current_template.columns.pop(row)
                    
                    # Tabelle neu aufbauen mit verbleibenden Spalten
                    self._apply_template_to_table(self.current_template, preserve_data=True)

    def move_column_up(self):
        row = self.columns_list.currentRow()
        if row > 0:
            # Daten vor Bewegung sichern
            old_data = self._get_current_table_data()
            
            # Spalte in der Liste bewegen
            item = self.columns_list.takeItem(row)
            self.columns_list.insertItem(row - 1, item)
            self.columns_list.setCurrentRow(row - 1)
            
            # Template aktualisieren
            if self.current_template:
                col = self.current_template.columns.pop(row)
                self.current_template.columns.insert(row - 1, col)
                
                # Tabelle mit neuen Spaltenreihenfolge aber gleichen Daten
                self._reorder_table_columns(old_data, row, row - 1)

    def move_column_down(self):
        row = self.columns_list.currentRow()
        if row >= 0 and row < self.columns_list.count() - 1:
            # Daten vor Bewegung sichern
            old_data = self._get_current_table_data()
            
            # Spalte in der Liste bewegen
            item = self.columns_list.takeItem(row)
            self.columns_list.insertItem(row + 1, item)
            self.columns_list.setCurrentRow(row + 1)
            
            # Template aktualisieren
            if self.current_template:
                col = self.current_template.columns.pop(row)
                self.current_template.columns.insert(row + 1, col)
                
                # Tabelle mit neuen Spaltenreihenfolge aber gleichen Daten
                self._reorder_table_columns(old_data, row, row + 1)

    def _get_current_table_data(self):
        data = []
        for row in range(self.table.rowCount()):
            row_data = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                row_data.append({
                    'text': item.text() if item else "",
                    'background': item.background().color().name() if item and item.background() else "",
                    'foreground': item.foreground().color().name() if item and item.foreground() else "",
                    'flags': item.flags() if item else QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsEditable
                })
            data.append(row_data)
        return data

    def _reorder_table_columns(self, old_data, from_index, to_index):
        if not self.current_template:
            return
            
        # Temporär Daten sichern
        temp_data = old_data
        
        # Neue Tabellenstruktur erstellen
        self.table.clear()
        self.table.setColumnCount(len(self.current_template.columns))
        
        # Header setzen
        headers = []
        for c in self.current_template.columns:
            unit = c.get('unit', '')
            if unit:
                headers.append(f"{c.get('name', '')} ({unit})")
            else:
                headers.append(c.get('name', ''))
        
        self.table.setHorizontalHeaderLabels(headers)
        
        # Spaltenbreiten setzen
        for idx, c in enumerate(self.current_template.columns):
            w = int(c.get('width', 120))
            self.table.setColumnWidth(idx, w)
        
        # Daten wiederherstellen mit neuer Spaltenreihenfolge
        if temp_data:
            self.table.setRowCount(len(temp_data))
            for row, row_data in enumerate(temp_data):
                for col in range(len(self.current_template.columns)):
                    if col < len(row_data):
                        item_data = row_data[col]
                        item = QtWidgets.QTableWidgetItem(item_data['text'])
                        
                        # Formatierung wiederherstellen
                        if item_data.get('background'):
                            item.setBackground(QtGui.QColor(item_data['background']))
                        if item_data.get('foreground'):
                            item.setForeground(QtGui.QColor(item_data['foreground']))
                        
                        item.setFlags(item_data.get('flags', QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsEditable))
                        self.table.setItem(row, col, item)
        
        # Sollwert-Zeile wiederherstellen
        self._add_setpoint_row(self.current_template.columns)

    def on_column_selected(self, current, previous):
        if current is None:
            return
        col = current.data(QtCore.Qt.UserRole)
        self.template_name_edit.setPlaceholderText(f"Vorlagenname (ausgewählte Spalte: {col.get('name')})")

    def apply_column_properties(self):
        pass

    def _gather_template_from_ui(self):
        name = self.template_name_edit.text().strip() or "unnamed"
        description = self.description_edit.toPlainText().strip()
        cols = []
        for i in range(self.columns_list.count()):
            col = self.columns_list.item(i).data(QtCore.Qt.UserRole)
            cols.append(col)
        return Template(name, cols, description)

    def load_selected_template(self):
        txt = self.templates_combo.currentText()
        if not txt:
            return
        fname = TEMPLATES_DIR / (txt + ".json")
        if not fname.exists():
            QtWidgets.QMessageBox.warning(self, "Nicht gefunden", "Die Datei existiert nicht")
            self._refresh_templates_list()
            return
        with open(fname, "r", encoding="utf-8") as f:
            d = json.load(f)
        tmpl = Template.from_dict(d)
        self._apply_template_to_ui(tmpl)

    def _apply_template_to_ui(self, tmpl: Template):
        self.template_name_edit.setText(tmpl.name)
        self.description_edit.setPlainText(tmpl.description)
        self.columns_list.clear()
        
        for c in tmpl.columns:
            item = QtWidgets.QListWidgetItem(c.get("name", ""))
            item.setData(QtCore.Qt.UserRole, c)
            self.columns_list.addItem(item)
        
        self.current_template = tmpl
        
        # Tabelle mit Daten-Erhalt aktualisieren
        self._apply_template_to_table(tmpl, preserve_data=True)

    def _apply_template_to_table(self, tmpl: Template, preserve_data=True):        
        # Alte Daten sichern, falls gewünscht
        old_data = []
        if preserve_data and self.table.rowCount() > 0 and self.table.columnCount() > 0:
            for row in range(self.table.rowCount()):
                row_data = []
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    row_data.append(item.text() if item else "")
                old_data.append(row_data)
        
        cols = tmpl.columns
        self.table.clear()
        self.table.setColumnCount(len(cols))
        
        headers = []
        for c in cols:
            unit = c.get('unit', '')
            if unit:
                headers.append(f"{c.get('name', '')} ({unit})")
            else:
                headers.append(c.get('name', ''))
        
        self.table.setHorizontalHeaderLabels(headers)
        
        # Spaltenbreiten setzen
        for idx, c in enumerate(cols):
            w = int(c.get('width', 120))
            self.table.setColumnWidth(idx, w)
        
        # Zeilen wiederherstellen
        if preserve_data and old_data:
            self.table.setRowCount(len(old_data))
            for row, row_data in enumerate(old_data):
                for col in range(min(len(row_data), len(cols))):
                    if col < len(row_data):
                        item = QtWidgets.QTableWidgetItem(row_data[col])
                        flags = QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsEditable
                        item.setFlags(flags)
                        self.table.setItem(row, col, item)
        else:
            self.table.setRowCount(0)
        
        # Sollwert-Zeile nur hinzufügen, wenn nicht bereits Daten vorhanden
        if not old_data or self.setpoint_row_index == -1:
            self._add_setpoint_row(cols)
        
        self.current_template = tmpl

    def _add_setpoint_row(self, columns):
        has_setpoints = any(col.get('setpoint') is not None for col in columns)
        if has_setpoints:
            self.setpoint_row_index = 0
            self.table.insertRow(0)
            
            for col_idx, col_info in enumerate(columns):
                setpoint_value = col_info.get('setpoint', '')
                unit = col_info.get('unit', '')
                
                if unit and setpoint_value:
                    display_value = f"{setpoint_value} {unit}"
                else:
                    display_value = str(setpoint_value)
                    
                item = QtWidgets.QTableWidgetItem(display_value)
                
                if col_info.get('readonly', True):
                    item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                else:
                    item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsEditable)
                
                item.setBackground(QtGui.QColor('#2a2a33'))  
                item.setForeground(QtGui.QColor('#9aa'))     
                
                self.table.setItem(0, col_idx, item)
            
            self.table.setVerticalHeaderItem(0, QtWidgets.QTableWidgetItem("Sollwerte"))
        else:
            self.setpoint_row_index = -1

    def add_row(self):
        c = self.table.columnCount()
        if c == 0:
            QtWidgets.QMessageBox.warning(self, "Keine Spalten", "Bitte zuerst eine Vorlage/Dataset erstellen")
            return
    
        # Neue Zeile immer am Ende einfügen
        insert_position = self.table.rowCount()
        self.table.insertRow(insert_position)
    
        cols = self.current_template.columns if self.current_template else [self.columns_list.item(i).data(QtCore.Qt.UserRole) for i in range(self.columns_list.count())]
        for col_idx in range(c):
            colinfo = cols[col_idx]
            default = colinfo.get('default')
            unit = colinfo.get('unit', '')
        
            text = str(default) if default is not None else ''
            if unit and text:
                text = f"{text} {unit}"
            
            it = QtWidgets.QTableWidgetItem(text)
        
            flags = QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsEditable
            it.setFlags(flags)
        
            self.table.setItem(insert_position, col_idx, it)

    def remove_row(self):
        r = self.table.currentRow()
        if r >= 0:
            if r == self.setpoint_row_index:
                QtWidgets.QMessageBox.warning(self, "Nicht erlaubt", "Die Sollwert-Zeile kann nicht gelöscht werden.")
                return
            self.table.removeRow(r)

    def fill_setpoints(self):
        cols = []
        if self.current_template:
            cols = self.current_template.columns
        else:
            tmpl = self._gather_template_from_ui()
            cols = tmpl.columns
    
        if not cols:
            QtWidgets.QMessageBox.warning(self, "Keine Spalten", "Keine Sollwerte vorhanden")
            return
    
        if self.setpoint_row_index < 0:
            self._add_setpoint_row(cols)
    
    # Sollwerte in die Zeile eintragen
        for cidx, col in enumerate(cols):
            sp = col.get("setpoint")
            if sp is not None:
                it = self.table.item(self.setpoint_row_index, cidx)
                if it is not None:
                    unit = col.get('unit', '')
                    if unit:
                        display_value = f"{sp} {unit}"
                    else:
                        display_value = str(sp)
                    it.setText(display_value)
                # Schreibschutz entsprechend der Spalteneinstellung
                    if col.get('readonly', True):
                        it.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                    else:
                        it.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsEditable)

    def _save_csv(self, path):
        headers = [self.table.horizontalHeaderItem(i).text() if self.table.horizontalHeaderItem(i) else f"Column{i}" for i in range(self.table.columnCount())]
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            for r in range(self.table.rowCount()):
                row = [self.table.item(r, c).text() if self.table.item(r, c) else '' for c in range(self.table.columnCount())]
                writer.writerow(row)

    def _save_json(self, path):
        data = {"headers": [self.table.horizontalHeaderItem(i).text() if self.table.horizontalHeaderItem(i) else f"Column{i}" for i in range(self.table.columnCount())],
                "rows": [[self.table.item(r, c).text() if self.table.item(r, c) else '' for c in range(self.table.columnCount())] for r in range(self.table.rowCount())]}
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def _load_csv(self, path):
        with open(path, newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
        if not rows:
            return
        headers = rows[0]
        data = rows[1:]
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(0)
        for rdata in data:
            r = self.table.rowCount()
            self.table.insertRow(r)
            for c, val in enumerate(rdata):
                it = QtWidgets.QTableWidgetItem(val)
                self.table.setItem(r, c, it)

    def _load_json(self, path):
        with open(path, encoding='utf-8') as f:
            d = json.load(f)
        headers = d.get('headers', [])
        rows = d.get('rows', [])
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(0)
        for rdata in rows:
            r = self.table.rowCount()
            self.table.insertRow(r)
            for c, val in enumerate(rdata):
                self.table.setItem(r, c, QtWidgets.QTableWidgetItem(val))

    def _generate_html_for_print(self):
        title = f"{o3NAME} - Export"
        html = [f"<html><head><meta charset='utf-8'><style>body{{font-family:Arial, Helvetica, sans-serif;font-size:12px;background:#121217;color:#e6eef8}}table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #444;padding:6px;text-align:left}}th{{background:#222;color:#fff}}.meta{{margin-bottom:10px}}.setpoint{{font-style:italic;color:#9aa;font-weight:bold}}.short-description{{margin-bottom:15px;padding:10px;border:1px solid #444;background:#1b1b20}}.long-description{{margin-top:20px;padding:10px;border:1px solid #444;background:#1b1b20}}.value-ok{{background:#1e3a1e;color:#90ee90}}.value-error{{background:#3a1e1e;color:#ff6b6b}}img.banner{{max-height:80px}}</style></head><body>"]
        if LOGO_PATH.exists():
            html.append(f"<div style='display:flex;align-items:center;margin-bottom:10px'><img class='banner' src='file://{LOGO_PATH}' alt='logo' style='height:60px;margin-right:12px'><div><h2>{title}</h2><div class='meta'>Vorlage: {self.template_name_edit.text() or '-'} &nbsp;&nbsp; Version: {o3VERSION}</div></div></div>")
        else:
            html.append(f"<h2>{title}</h2>")
            html.append(f"<div class='meta'>Vorlage: {self.template_name_edit.text() or '-'} &nbsp;&nbsp; Version: {o3VERSION}</div>")

        if self.export_short_description:
            html.append(f"<div class='short-description'><strong>Kurzbeschreibung:</strong><br>{self.export_short_description.replace(chr(10), '<br>')}</div>")

        
        html.append("<table>")
        html.append("<tr>")
        headers = [self.table.horizontalHeaderItem(i).text() if self.table.horizontalHeaderItem(i) else f"Column{i+1}" for i in range(self.table.columnCount())]
        for h in headers:
            html.append(f"<th>{h}</th>")
        html.append("</tr>")
        cols = []
        if self.current_template and self.current_template.columns:
            cols = self.current_template.columns
        else:
            cols = [self.columns_list.item(i).data(QtCore.Qt.UserRole) for i in range(self.columns_list.count())]
        if cols:
            any_sp = any(c.get('setpoint') for c in cols)
            if any_sp:
                html.append("<tr class='setpoint'>")
                for c in cols:
                    sp = c.get('setpoint')
                    unit = c.get('unit', '')
                    if unit and sp:
                        display_value = f"{sp} {unit}"
                    else:
                        display_value = sp if sp is not None else ''
                    html.append(f"<td>{display_value}</td>")
                html.append("</tr>")
        for r in range(self.table.rowCount()):
            if r == self.setpoint_row_index:
                continue
            html.append("<tr>")
            for c in range(self.table.columnCount()):
                val = self.table.item(r, c).text() if self.table.item(r, c) else ""
                cell_item = self.table.item(r, c)
                if cell_item and cell_item.background().color().name() == "#1e3a1e":
                    html.append(f"<td class='value-ok'>{val}</td>")
                elif cell_item and cell_item.background().color().name() == "#3a1e1e":
                    html.append(f"<td class='value-error'>{val}</td>")
                else:
                    html.append(f"<td>{val}</td>")
            html.append("</tr>")
        html.append("</table>")
        
        # Langbeschreibung unter der Tabelle
        if self.export_long_description:
            html.append(f"<div class='long-description'><strong>Beschreibung:</strong><br>{self.export_long_description.replace(chr(10), '<br>')}</div>")
        
        html.append("</body></html>")
        return ''.join(html)

    def _apply_dark_style(self):
        style = """
        QWidget{ background-color: #121217; color: #e6eef8; }
        QLineEdit, QPlainTextEdit, QTextEdit, QSpinBox, QComboBox, QListWidget{ background-color: #1b1b20; border: 1px solid #2a2a33; padding:4px }
        QPushButton{ background-color: #2b2b34; border: 1px solid #3a3a45; padding:6px; border-radius:6px }
        QPushButton:hover{ background-color: #34343e }
        QTableWidget{ background-color: #0f0f12; gridline-color: #222; selection-background-color:#2b6ea3 }
        QHeaderView::section{ background-color: #16161a; padding:6px; border: none }
        QGroupBox{ border: 1px solid #24242b; border-radius:8px; margin-top:6px }
        QGroupBox:title{ subcontrol-origin: margin; left:8px; padding:0 3px }
        QToolBar{ background: transparent; spacing:6px }
        QMenuBar{ background: transparent }
        QMenu{ background-color: #161618 }
        QMessageBox{ background-color: #1b1b1f }
        QDialog{ background-color: #121217; color: #e6eef8; }
        QLabel{ color: #e6eef8; }
        """
        self.setStyleSheet(style)


def show_splash(app):
    w, h = 640, 300
    pix = QtGui.QPixmap(w, h)
    pix.fill(QtGui.QColor('#121217'))
    painter = QtGui.QPainter(pix)
    if LOGO_PATH.exists():
        logo = QtGui.QPixmap(str(LOGO_PATH))
        logo_scaled = logo.scaledToHeight(120, QtCore.Qt.SmoothTransformation)
        painter.drawPixmap(24, 24, logo_scaled)
        x = 24 + logo_scaled.width() + 12
    else:
        x = 24
    painter.setPen(QtGui.QColor('#e6eef8'))
    font = QtGui.QFont('Arial', 20, QtGui.QFont.Bold)
    painter.setFont(font)
    painter.drawText(x, 60, f"{o3NAME}")
    font2 = QtGui.QFont('Arial', 10)
    painter.setFont(font2)
    painter.drawText(x, 90, f"Version {o3VERSION}")
    painter.drawText(x, 120, f"{o3COPYRIGHT}")
    painter.drawText(x, 140, "Copyright © 2025 openw3rk INVENT")
    painter.drawText(x, 260, "https://openw3rk.de")
    painter.drawText(x, 280, "https://o3measurement.openw3rk.de")
    painter.end()

    splash = QtWidgets.QSplashScreen(pix)
    splash.show()
# 3 sekunden splash
    loop = QtCore.QEventLoop()
    QtCore.QTimer.singleShot(3000, loop.quit)
    loop.exec_()
    splash.close()


def main():
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
    app = QtWidgets.QApplication(sys.argv)

    app.setWindowIcon(QtGui.QIcon(str(ICON_PATH)))
    # splash
    show_splash(app)

    win = MainWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()

# --------------------------------------------------------------------------------
# o3Measurement 12.5.1 | Copyright (c) 2025 openw3rk INVENT, All rights reserved.
# Licensed under GNU General Public License v3.0 (GPLv3), see "LICENSE.txt".
# o3Measurement comes with ABSOLUTELY NO WARRANTY.
# --------------------------------------------------------------------------------
# https://openw3rk.de | https://o3measurement.openw3rk.de | develop@openw3rk.de
# --------------------------------------------------------------------------------
