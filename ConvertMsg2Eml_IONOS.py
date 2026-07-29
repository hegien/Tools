#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
msg_to_eml.py
==============

Rekursives Durchsuchen eines Ordners nach *.msg*-Dateien,
Umwandeln in *.eml* (unter Nutzung der Outlook-COM-Schnittstelle)
und Schreiben eines einfachen Text-Logfiles.

Features
--------
* Rekursive Suche in allen Unterordnern
* Keine Ueberschreibung vorhandener *.eml*-Dateien --> Namenssuffix "_2"
* Logfile im Startordner (msg-to-eml.log)
* Ausfuehrung ueber die Kommandozeile mit Parametern
* Hintergrund-/Silent-Modus (keine Konsolenausgabe)

Author: Dora Developer (IONOS GPT)
"""

import argparse
import logging
import os
import sys
from pathlib import Path

# --------------------------------------------------------------
# Outlook-COM-Helper
# --------------------------------------------------------------
def msg_to_eml_via_outlook(msg_path: Path) -> bytes:
    """
    Oeffnet eine *.msg*-Datei mit Outlook und gibt den Inhalt
    als MIME-Message (bytes) zurueck, die anschliessend als *.eml*
    gespeichert werden kann.

    Parameters
    ----------
    msg_path: pathlib.Path
        Vollstaendiger Pfad zur *.msg*-Datei.

    Returns
    -------
    bytes
        EML-Daten im RFC-822-Format.
    """
    import win32com.client  # Teil von pywin32

    # Outlook-Instanz holen (oder neue starten)
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Die .msg-Datei als MailItem laden
    mail_item = namespace.OpenSharedItem(str(msg_path))

    # Die MailItem-Methode `SaveAs` mit Format 5 (= olRFC822) erzeugt EML-Daten.
    # Wir speichern zunaechst in einen temporaeren Pfad und lesen die Bytes ein,
    # weil Outlook selbst nur in Dateien schreibt.
    temp_eml_path = msg_path.with_suffix(".tmp_eml")
    mail_item.SaveAs(str(temp_eml_path), 5)  # 5 == olRFC822

    # Dateiinhalt einlesen und temporaere Datei loeschen
    eml_bytes = temp_eml_path.read_bytes()
    temp_eml_path.unlink(missing_ok=True)

    return eml_bytes


# --------------------------------------------------------------
# Hilfsfunktionen fuer Dateinamen
# --------------------------------------------------------------
def build_target_path(src_path: Path) -> Path:
    """
    Ermittelt den Zielpfad fuer die *.eml*-Datei.
    Wenn die Datei bereits existiert, wird ein Suffix ``_2`` angehaengt,
    bevor die Endung ``.eml`` gesetzt wird.

    Parameters
    ----------
    src_path: pathlib.Path
        Pfad zur Original-*.msg*-Datei.

    Returns
    -------
    pathlib.Path
        Pfad, unter dem die *.eml* geschrieben werden soll.
    """
    target = src_path.with_suffix(".eml")
    if target.exists():
        # Zaehler erhoehen, bis ein freier Name gefunden ist
        counter = 2
        while True:
            new_name = f"{src_path.stem}_{counter}.eml"
            target = src_path.parent / new_name
            if not target.exists():
                break
            counter += 1
    return target


# --------------------------------------------------------------
# Hauptlogik
# --------------------------------------------------------------
def convert_folder(start_dir: Path, silent: bool = False) -> None:
    """
    Durchsucht `start_dir` rekursiv, konvertiert *.msg* nach *.eml*
    und protokolliert das Ergebnis.

    Parameters
    ----------
    start_dir: pathlib.Path
        Startverzeichnis, in dem gesucht wird.
    silent: bool
        Wenn True, werden keine Meldungen auf die Konsole geschrieben.
    """
    # Logfile anlegen (im Startordner)
    log_file = start_dir / "msg-to-eml.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout) if not silent else logging.NullHandler(),
        ],
    )

    logging.info(f"Startverzeichnis: {start_dir}")
    logging.info("Suche nach *.msg*-Dateien …")

    msg_files = list(start_dir.rglob("*.msg"))
    logging.info(f"{len(msg_files)} *.msg*-Datei(en) gefunden.")

    for msg_path in msg_files:
        try:
            logging.info(f"Verarbeite: {msg_path}")

            # Zielpfad ermitteln (inkl. _2-Logik)
            eml_path = build_target_path(msg_path)

            # Konvertierung via Outlook
            eml_bytes = msg_to_eml_via_outlook(msg_path)

            # Schreiben
            eml_path.write_bytes(eml_bytes)

            logging.info(f"Gespeichert als: {eml_path}")
        except Exception as exc:  # noqa: BLE001
            logging.error(f"Fehler bei {msg_path}: {exc}")

    logging.info("Fertig.")


# --------------------------------------------------------------
# CLI-Parser
# --------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Rekursives Konvertieren von Outlook .msg-Dateien nach .eml."
    )
    parser.add_argument(
        "folder",
        type=Path,
        help="Pfad zum Startordner, in dem nach *.msg* gesucht wird.",
    )
    parser.add_argument(
        "-s",
        "--silent",
        action="store_true",
        help="Kein Konsolenausgabe - nur Logfile (fuer Hintergrund-Ausfuehrung).",
    )
    args = parser.parse_args()

    if not args.folder.is_dir():
        print(f"Der Pfad {args.folder} ist kein gueltiges Verzeichnis.", file=sys.stderr)
        sys.exit(1)

    convert_folder(args.folder.resolve(), silent=args.silent)


if __name__ == "__main__":
    main()