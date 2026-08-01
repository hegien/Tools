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
import win32com.client
import email
import email.policy
import email.generator
import datetime
import uuid
from io import BytesIO
from email.message import EmailMessage
from email.generator import BytesGenerator
import base64



# Finale Version, mit der die Konvertierung von 350 .msg Files geklappt hat
# ------------------------------------------------------------
def _guess_mime_type(suffix: str):
    suffix = suffix.lower()
    if suffix in {".jpg", ".jpeg"}:   return ("image", "jpeg")
    if suffix == ".png":              return ("image", "png")
    if suffix == ".gif":              return ("image", "gif")
    if suffix == ".bmp":              return ("image", "bmp")
    if suffix == ".svg":              return ("image", "svg+xml")
    if suffix == ".pdf":              return ("application", "pdf")
    if suffix == ".doc":              return ("application", "msword")
    if suffix == ".docx":             return ("application",
                                            "vnd.openxmlformats-officedocument.wordprocessingml.document")
    if suffix == ".xls":              return ("application", "vnd.ms-excel")
    if suffix == ".xlsx":             return ("application",
                                            "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return ("application", "octet-stream")

def _to_str(value):
    """Listen → kommagetrennte Strings, sonst unverändert."""
    if isinstance(value, (list, tuple)):
        return ", ".join(str(v) for v in value if v is not None)
    if value is None:
        return ""
    return str(value)

def _to_body(value):
    """Listen → Zeilen‑String, sonst unverändert."""
    if isinstance(value, (list, tuple)):
        return "\n".join(str(v) for v in value if v is not None)
    if value is None:
        return ""
    return str(value)

# ------------------------------------------------------------
def msg_to_eml_via_outlook(msg_path: Path) -> bytes:
    """Exportiert .msg → .eml mit garantiertem multipart/mixed‑Root."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    mail = ns.OpenSharedItem(str(msg_path))

    # ---------- Header ----------
    hdr = {
        "From":    _to_str(mail.SenderEmailAddress),
        "To":      _to_str([rec.Address for rec in mail.Recipients]),
        "Subject": _to_str(mail.Subject),
        "Date":    mail.SentOn.strftime("%a, %d %b %Y %H:%M:%S %z"),
        "Message-ID": f"<{uuid.uuid4()}@outlook>",
    }
    if hasattr(mail, "CC"):
        hdr["Cc"] = _to_str(getattr(mail, "CC"))
    if hasattr(mail, "BCC"):
        hdr["Bcc"] = _to_str(getattr(mail, "BCC"))

    # ---------- Body ----------
    plain_body = _to_body(mail.Body)
    html_body  = _to_body(mail.HTMLBody)

    have_rtf_body = False
    rtf_body_bytes = b""
    if getattr(mail, "BodyFormat", 1) == 3:      # Rich‑Text
        rtf_body_bytes = _to_body(mail.Body).encode("utf-8")
        have_rtf_body = True

    # ---------- Attachments ----------
    inline_parts = []      # (cid, maintype, subtype, data, filename)
    normal_parts = []      # (maintype, subtype, data, filename)
    rtf_inline = None      # (maintype, subtype, data, filename)

    ATTACH_OLE          = 0x00000008
    ATTACH_EMBEDDED_OLE = 0x00000010

    def _read_att(att):
        tmp = Path.cwd() / f"__tmp_{uuid.uuid4().hex}{Path(att.FileName).suffix}"
        att.SaveAsFile(str(tmp))
        data = tmp.read_bytes()
        tmp.unlink(missing_ok=True)
        return data

    for i in range(1, mail.Attachments.Count + 1):
        att = mail.Attachments.Item(i)

        try:
            flags = att.PropertyAccessor.GetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x37120003"
            )
        except Exception:
            flags = 0

        data = _read_att(att)
        suffix = Path(att.FileName).suffix
        maintype, subtype = _guess_mime_type(suffix)

        # ----- Inline‑Bilder (nach Endung) -----
        if suffix.lower() in {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".svg"}:
            cid = f"<{uuid.uuid4()}@outlook>"
            inline_parts.append((cid, maintype, subtype, data, att.FileName))
            html_body = html_body.replace(
                f'src="cid:{i}"', f'src="cid:{cid[1:-1]}"'
            ).replace(
                f"src='cid:{i}'", f"src='cid:{cid[1:-1]}'"
            )
            continue

        # ----- RTF‑Attachment -----
        if suffix.lower() == ".rtf":
            rtf_inline = (maintype, subtype, data, att.FileName)
            have_rtf_body = True
            continue

        # ----- OLE / Embedded‑Message -----
        if flags & ATTACH_OLE or flags & ATTACH_EMBEDDED_OLE:
            normal_parts.append((maintype, subtype, data, att.FileName))
            continue

        # ----- Restliche Anhänge -----
        normal_parts.append((maintype, subtype, data, att.FileName))

    # ---------- 4. MIME‑Baum bauen ----------
    # 4.1 Root ist **immer** multipart/mixed
    root = EmailMessage(policy=email.policy.SMTP)
    root.set_type("multipart/mixed")
    for k, v in hdr.items():
        root[k] = _to_str(v)

    # 4.2 Wenn Inline‑Bilder oder ein RTF‑Inline‑Teil existieren, bauen wir multipart/related
    need_related = bool(inline_parts) or rtf_inline is not None
    if need_related:
        related = EmailMessage()
        related.set_type("multipart/related")
        root.attach(related)          # related ist Kind von root
        body_container = related      # alle nachfolgenden Teile gehen hier rein
    else:
        body_container = root         # kein related → body kommt direkt unter root

    # 4.3 multipart/alternative (plain + html + optional rtf)
    alt = EmailMessage()               # kein set_type → wird automatisch zu multipart/alternative
    alt.set_content(plain_body, subtype="plain", charset="utf-8")
    alt.add_alternative(html_body, subtype="html", charset="utf-8")
    if have_rtf_body and rtf_body_bytes:
        alt.add_alternative(
            rtf_body_bytes.decode("utf-8", errors="ignore"),
            subtype="rtf",
            charset="utf-8",
        )
    body_container.attach(alt)          # alt wird in related (falls vorhanden) oder root angehängt

    # 4.4 Inline‑Bilder + evtl. RTF‑Attachment in multipart/related einbetten
    if need_related:
        for cid, mtype, stype, data, fname in inline_parts:
            related.add_related(
                data,
                maintype=mtype,
                subtype=stype,
                cid=cid,
                filename=fname,
                disposition="inline",
            )
        if rtf_inline:
            mtype, stype, data, fname = rtf_inline
            related.add_related(
                data,
                maintype=mtype,
                subtype=stype,
                filename=fname,
                disposition="inline",
            )

    # 4.5 Normale Anhänge (immer unter root, weil root bereits multipart/mixed ist)
    for mtype, stype, data, fname in normal_parts:
        root.add_attachment(
            data,
            maintype=mtype,
            subtype=stype,
            filename=fname,
            disposition="attachment",
        )

    # ---------- Debug‑Ausgabe (optional) ----------
    # Entfernen Sie die folgenden Zeilen, wenn Sie das Logging nicht mehr benötigen.
    print("\n--- MIME‑Baum (Debug) ---")
    print("Root Content-Type :", root.get_content_type())
    for idx, part in enumerate(root.walk()):
        disp = part.get('Content-Disposition')
        print(f" part {idx}: {part.get_content_type()}, disposition={disp}")

    # ---------- 5. RFC‑822‑Bytes erzeugen ----------
    buf = BytesIO()
    gen = BytesGenerator(buf, policy=email.policy.SMTP)   # outfp = buf
    gen.flatten(root)
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