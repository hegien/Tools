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


# Aus Optimierung für Formatierte Mail
import win32com.client
import email
import email.policy
import email.generator
# from pathlib import Path
import datetime
import uuid

# import win32com.client
# import email
from email.message import EmailMessage
from email.generator import BytesGenerator
# import email.policy
# from pathlib import Path
# import uuid
# import datetime

# --------------------------------------------------------------
# Outlook-COM-Helper
# --------------------------------------------------------------
# def msg_to_eml_via_outlook(msg_path: Path) -> bytes:
#     """
#     Oeffnet eine *.msg*-Datei mit Outlook und gibt den Inhalt
#     als MIME-Message (bytes) zurueck, die anschliessend als *.eml*
#     gespeichert werden kann.

#     Parameters
#     ----------
#     msg_path: pathlib.Path
#         Vollstaendiger Pfad zur *.msg*-Datei.

#     Returns
#     -------
#     bytes
#         EML-Daten im RFC-822-Format.
#     """
#     import win32com.client  # Teil von pywin32

#     # Outlook-Instanz holen (oder neue starten)
#     outlook = win32com.client.Dispatch("Outlook.Application")
#     namespace = outlook.GetNamespace("MAPI")

#     # Die .msg-Datei als MailItem laden
#     mail_item = namespace.OpenSharedItem(str(msg_path))

#     # Die MailItem-Methode `SaveAs` mit Format 5 (= olRFC822) erzeugt EML-Daten.
#     # Wir speichern zunaechst in einen temporaeren Pfad und lesen die Bytes ein,
#     # weil Outlook selbst nur in Dateien schreibt.
#     temp_eml_path = msg_path.with_suffix(".tmp_eml")
#     mail_item.SaveAs(str(temp_eml_path), 5)  # 5 == olRFC822

#     # Dateiinhalt einlesen und temporaere Datei loeschen
#     eml_bytes = temp_eml_path.read_bytes()
#     temp_eml_path.unlink(missing_ok=True)

#     return eml_bytes

# Aus Optimierung für Formatierte Mail
def msg_to_eml_via_outlook(msg_path: Path) -> bytes:
    """Exportiert eine .msg‑Datei zu einer voll‑wertigen RFC‑822‑Nachricht."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    mail = ns.OpenSharedItem(str(msg_path))

    # ---------- 1. Daten aus dem MailItem holen ----------
    # Header‑Infos
    headers = {
        "From": mail.SenderEmailAddress,
        "To": ", ".join(mail.Recipients[i].Address for i in range(mail.Recipients.Count)),
        "Subject": mail.Subject,
        "Date": mail.SentOn.strftime("%a, %d %b %Y %H:%M:%S %z"),
        "Message-ID": f"<{uuid.uuid4()}@outlook>",
    }

    # Body – Outlook liefert HTML über .HTMLBody
    html_body = mail.HTMLBody or ""
    # Optional: Plain‑Text‑Fallback via .Body (Outlook liefert bereits reinen Text)
    plain_body = mail.Body or ""

    # ---------- 2. MIME‑Nachricht zusammenbauen ----------
    # multipart/alternative (Plain‑Text + HTML)
    outer = email.message.EmailMessage(policy=email.policy.SMTP)
    for k, v in headers.items():
        outer[k] = v

    # Wenn Anhänge existieren, benutzen wir multipart/mixed → inner = multipart/alternative
    has_attachments = mail.Attachments.Count > 0
    if has_attachments:
        outer.set_type("multipart/mixed")
        alternative = email.message.EmailMessage()
        alternative.set_type("multipart/alternative")
    else:
        alternative = outer   # kein extra Wrapper nötig

    # –‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑-
    # Plain‑Text‑Teil
    alternative.set_content(plain_body, subtype="plain", charset="utf-8")

    # HTML‑Teil (hier wird charset explizit auf utf‑8 gesetzt)
    alternative.add_alternative(html_body, subtype="html", charset="utf-8")

    # Wenn wir einen multipart/mixed‑Wrapper benötigen, hängen wir das alternative‑Objekt an:
    if has_attachments:
        outer.attach(alternative)

        # ---------- 3. Anhänge einbetten ----------
        for i in range(1, mail.Attachments.Count + 1):
            att = mail.Attachments.Item(i)

            # Temporär sichern, weil Outlook nur in Dateien schreiben kann
            tmp_path = Path.cwd() / f"__tmp_att_{uuid.uuid4().hex}"
            att.SaveAsFile(str(tmp_path))

            # Erkennen des MIME‑Typs (einfacher Ansatz: anhand der Extension)
            maintype, subtype = ("application", "octet-stream")
            if tmp_path.suffix.lower() in {".jpg", ".jpeg"}:
                maintype, subtype = ("image", "jpeg")
            elif tmp_path.suffix.lower() == ".png":
                maintype, subtype = ("image", "png")
            elif tmp_path.suffix.lower() == ".pdf":
                maintype, subtype = ("application", "pdf")

            with open(tmp_path, "rb") as fh:
                outer.add_attachment(
                    fh.read(),
                    maintype=maintype,
                    subtype=subtype,
                    filename=att.FileName,          # Original‑Dateiname
                )
            tmp_path.unlink(missing_ok=True)

    # ---------- 4. RFC‑822‑Bytes erzeugen ----------
    # Der EmailMessage‑Generator gibt korrekt formatierte Zeilenumbrüche (\r\n) zurück.
    gen = email.generator.BytesGenerator(policy=email.policy.SMTP)
    from io import BytesIO
    buf = BytesIO()
    gen.flatten(outer, buf)
    return buf.getvalue()

# Korrektur wegen Fehler
def msg_to_eml_via_outlook(msg_path: Path) -> bytes:
    """
    Exportiert eine .msg‑Datei zu einer vollwertigen RFC‑822‑Nachricht.
    Der erzeugte Byte‑String kann unverändert mit Path.write_bytes()
    als *.eml* abgespeichert werden.
    """
    # ---------- Outlook‑Objekt ----------
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    mail = ns.OpenSharedItem(str(msg_path))

    # ---------- 1. Header‑Informationen ----------
    headers = {
        "From": mail.SenderEmailAddress,
        "To": ", ".join(rec.Address for rec in mail.Recipients),
        "Subject": mail.Subject,
        "Date": mail.SentOn.strftime("%a, %d %b %Y %H:%M:%S %z"),
        "Message-ID": f"<{uuid.uuid4()}@outlook>",
    }

    # ---------- 2. Body ----------
    html_body = mail.HTMLBody or ""
    plain_body = mail.Body or ""

    # ---------- 3. MIME‑Nachricht bauen ----------
    # Wenn Anhänge vorhanden sind, verwenden wir multipart/mixed,
    # sonst reicht multipart/alternative.
    has_attachments = mail.Attachments.Count > 0

    if has_attachments:
        outer = EmailMessage(policy=email.policy.SMTP)
        outer.set_type("multipart/mixed")
        # Der eigentliche textuelle Teil wird als multipart/alternative angehängt
        alternative = EmailMessage()
        alternative.set_type("multipart/alternative")
        outer.attach(alternative)
    else:
        outer = EmailMessage(policy=email.policy.SMTP)
        alternative = outer  # kein extra Wrapper nötig

    # Header eintragen
    for k, v in headers.items():
        outer[k] = v

    # ----- Text‑ und HTML‑Teil -----
    alternative.set_content(plain_body, subtype="plain", charset="utf-8")
    alternative.add_alternative(html_body, subtype="html", charset="utf-8")

    # ---------- 4. Anhänge einbetten (falls vorhanden) ----------
    if has_attachments:
        for i in range(1, mail.Attachments.Count + 1):
            att = mail.Attachments.Item(i)

            # Temporär in eine Datei schreiben – Outlook kann nur in Dateien exportieren
            tmp_path = Path.cwd() / f"__tmp_att_{uuid.uuid4().hex}{Path(att.FileName).suffix}"
            att.SaveAsFile(str(tmp_path))

            # MIME‑Typ rudimentär bestimmen (oder einfach application/octet-stream verwenden)
            maintype, subtype = ("application", "octet-stream")
            ext = tmp_path.suffix.lower()
            if ext in {".jpg", ".jpeg"}:
                maintype, subtype = ("image", "jpeg")
            elif ext == ".png":
                maintype, subtype = ("image", "png")
            elif ext == ".pdf":
                maintype, subtype = ("application", "pdf")

            with open(tmp_path, "rb") as fh:
                outer.add_attachment(
                    fh.read(),
                    maintype=maintype,
                    subtype=subtype,
                    filename=att.FileName,           # Original‑Dateiname im Header
                )
            tmp_path.unlink(missing_ok=True)   # Aufräumen

    # ---------- 5. RFC‑822‑Bytes erzeugen ----------
    buf = BytesIO()
    gen = BytesGenerator(buf, policy=email.policy.SMTP)   # **Wichtig: outfp übergeben**
    gen.flatten(outer)
    return buf.getvalue()



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