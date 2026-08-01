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
from io import BytesIO

# import win32com.client
# import email
from email.message import EmailMessage
from email.generator import BytesGenerator
# import email.policy
# from pathlib import Path
# import uuid
# import datetime

# import win32com.client
# import email
# from email.message import EmailMessage
# from email.generator import BytesGenerator
# import email.policy
# from pathlib import Path
# import uuid
# import datetime
import base64

# Initiale Version
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
# def msg_to_eml_via_outlook(msg_path: Path) -> bytes:
#     """Exportiert eine .msg‑Datei zu einer voll‑wertigen RFC‑822‑Nachricht."""
#     outlook = win32com.client.Dispatch("Outlook.Application")
#     ns = outlook.GetNamespace("MAPI")
#     mail = ns.OpenSharedItem(str(msg_path))

#     # ---------- 1. Daten aus dem MailItem holen ----------
#     # Header‑Infos
#     headers = {
#         "From": mail.SenderEmailAddress,
#         "To": ", ".join(mail.Recipients[i].Address for i in range(mail.Recipients.Count)),
#         "Subject": mail.Subject,
#         "Date": mail.SentOn.strftime("%a, %d %b %Y %H:%M:%S %z"),
#         "Message-ID": f"<{uuid.uuid4()}@outlook>",
#     }

#     # Body – Outlook liefert HTML über .HTMLBody
#     html_body = mail.HTMLBody or ""
#     # Optional: Plain‑Text‑Fallback via .Body (Outlook liefert bereits reinen Text)
#     plain_body = mail.Body or ""

#     # ---------- 2. MIME‑Nachricht zusammenbauen ----------
#     # multipart/alternative (Plain‑Text + HTML)
#     outer = email.message.EmailMessage(policy=email.policy.SMTP)
#     for k, v in headers.items():
#         outer[k] = v

#     # Wenn Anhänge existieren, benutzen wir multipart/mixed → inner = multipart/alternative
#     has_attachments = mail.Attachments.Count > 0
#     if has_attachments:
#         outer.set_type("multipart/mixed")
#         alternative = email.message.EmailMessage()
#         alternative.set_type("multipart/alternative")
#     else:
#         alternative = outer   # kein extra Wrapper nötig

#     # –‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑-
#     # Plain‑Text‑Teil
#     alternative.set_content(plain_body, subtype="plain", charset="utf-8")

#     # HTML‑Teil (hier wird charset explizit auf utf‑8 gesetzt)
#     alternative.add_alternative(html_body, subtype="html", charset="utf-8")

#     # Wenn wir einen multipart/mixed‑Wrapper benötigen, hängen wir das alternative‑Objekt an:
#     if has_attachments:
#         outer.attach(alternative)

#         # ---------- 3. Anhänge einbetten ----------
#         for i in range(1, mail.Attachments.Count + 1):
#             att = mail.Attachments.Item(i)

#             # Temporär sichern, weil Outlook nur in Dateien schreiben kann
#             tmp_path = Path.cwd() / f"__tmp_att_{uuid.uuid4().hex}"
#             att.SaveAsFile(str(tmp_path))

#             # Erkennen des MIME‑Typs (einfacher Ansatz: anhand der Extension)
#             maintype, subtype = ("application", "octet-stream")
#             if tmp_path.suffix.lower() in {".jpg", ".jpeg"}:
#                 maintype, subtype = ("image", "jpeg")
#             elif tmp_path.suffix.lower() == ".png":
#                 maintype, subtype = ("image", "png")
#             elif tmp_path.suffix.lower() == ".pdf":
#                 maintype, subtype = ("application", "pdf")

#             with open(tmp_path, "rb") as fh:
#                 outer.add_attachment(
#                     fh.read(),
#                     maintype=maintype,
#                     subtype=subtype,
#                     filename=att.FileName,          # Original‑Dateiname
#                 )
#             tmp_path.unlink(missing_ok=True)

#     # ---------- 4. RFC‑822‑Bytes erzeugen ----------
#     # Der EmailMessage‑Generator gibt korrekt formatierte Zeilenumbrüche (\r\n) zurück.
#     gen = email.generator.BytesGenerator(policy=email.policy.SMTP)
#     from io import BytesIO
#     buf = BytesIO()
#     gen.flatten(outer, buf)
#     return buf.getvalue()







# Korrektur wegen Fehler
# def msg_to_eml_via_outlook(msg_path: Path) -> bytes:
#     """
#     Exportiert eine .msg‑Datei zu einer vollwertigen RFC‑822‑Nachricht.
#     Der erzeugte Byte‑String kann unverändert mit Path.write_bytes()
#     als *.eml* abgespeichert werden.
#     """
#     # ---------- Outlook‑Objekt ----------
#     outlook = win32com.client.Dispatch("Outlook.Application")
#     ns = outlook.GetNamespace("MAPI")
#     mail = ns.OpenSharedItem(str(msg_path))

#     # ---------- 1. Header‑Informationen ----------
#     headers = {
#         "From": mail.SenderEmailAddress,
#         "To": ", ".join(rec.Address for rec in mail.Recipients),
#         "Subject": mail.Subject,
#         "Date": mail.SentOn.strftime("%a, %d %b %Y %H:%M:%S %z"),
#         "Message-ID": f"<{uuid.uuid4()}@outlook>",
#     }

#     # ---------- 2. Body ----------
#     html_body = mail.HTMLBody or ""
#     plain_body = mail.Body or ""

#     # ---------- 3. MIME‑Nachricht bauen ----------
#     # Wenn Anhänge vorhanden sind, verwenden wir multipart/mixed,
#     # sonst reicht multipart/alternative.
#     has_attachments = mail.Attachments.Count > 0

#     if has_attachments:
#         outer = EmailMessage(policy=email.policy.SMTP)
#         outer.set_type("multipart/mixed")
#         # Der eigentliche textuelle Teil wird als multipart/alternative angehängt
#         alternative = EmailMessage()
#         alternative.set_type("multipart/alternative")
#         outer.attach(alternative)
#     else:
#         outer = EmailMessage(policy=email.policy.SMTP)
#         alternative = outer  # kein extra Wrapper nötig

#     # Header eintragen
#     for k, v in headers.items():
#         outer[k] = v

#     # ----- Text‑ und HTML‑Teil -----
#     alternative.set_content(plain_body, subtype="plain", charset="utf-8")
#     alternative.add_alternative(html_body, subtype="html", charset="utf-8")

#     # ---------- 4. Anhänge einbetten (falls vorhanden) ----------
#     if has_attachments:
#         for i in range(1, mail.Attachments.Count + 1):
#             att = mail.Attachments.Item(i)

#             # Temporär in eine Datei schreiben – Outlook kann nur in Dateien exportieren
#             tmp_path = Path.cwd() / f"__tmp_att_{uuid.uuid4().hex}{Path(att.FileName).suffix}"
#             att.SaveAsFile(str(tmp_path))

#             # MIME‑Typ rudimentär bestimmen (oder einfach application/octet-stream verwenden)
#             maintype, subtype = ("application", "octet-stream")
#             ext = tmp_path.suffix.lower()
#             if ext in {".jpg", ".jpeg"}:
#                 maintype, subtype = ("image", "jpeg")
#             elif ext == ".png":
#                 maintype, subtype = ("image", "png")
#             elif ext == ".pdf":
#                 maintype, subtype = ("application", "pdf")

#             with open(tmp_path, "rb") as fh:
#                 outer.add_attachment(
#                     fh.read(),
#                     maintype=maintype,
#                     subtype=subtype,
#                     filename=att.FileName,           # Original‑Dateiname im Header
#                 )
#             tmp_path.unlink(missing_ok=True)   # Aufräumen

#     # ---------- 5. RFC‑822‑Bytes erzeugen ----------
#     buf = BytesIO()
#     gen = BytesGenerator(buf, policy=email.policy.SMTP)   # **Wichtig: outfp übergeben**
#     gen.flatten(outer)
#     return buf.getvalue()








# Erweiterung, um auch alle Einbettungen (Bilder, RTF‑Teile, Anlagen) korrekt zu übernehmen
# ------------------------------------------------------------
# Hilfsfunktion: MIME‑Typ anhand der Dateiendung bestimmen
# ------------------------------------------------------------
# def _guess_mime_type(suffix: str):
#     suffix = suffix.lower()
#     if suffix in {".jpg", ".jpeg"}:
#         return ("image", "jpeg")
#     if suffix == ".png":
#         return ("image", "png")
#     if suffix == ".gif":
#         return ("image", "gif")
#     if suffix == ".bmp":
#         return ("image", "bmp")
#     if suffix == ".svg":
#         return ("image", "svg+xml")
#     if suffix == ".pdf":
#         return ("application", "pdf")
#     if suffix == ".doc":
#         return ("application", "msword")
#     if suffix == ".docx":
#         return ("application", "vnd.openxmlformats-officedocument.wordprocessingml.document")
#     if suffix == ".xls":
#         return ("application", "vnd.ms-excel")
#     if suffix == ".xlsx":
#         return ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
#     # fallback
#     return ("application", "octet-stream")

# # ------------------------------------------------------------
# def msg_to_eml_via_outlook(msg_path: Path) -> bytes:
#     """
#     Exportiert eine .msg‑Datei zu einer vollständig MIME‑konformen
#     RFC‑822‑Nachricht, inkl.:
#         • plain‑text + html (multipart/alternative)
#         • inline‑Bilder (Content‑ID, multipart/related)
#         • RTF‑Body (optional, text/rtf)
#         • normale Anlagen (multipart/mixed)
#     """
#     # ---------- Outlook ----------
#     outlook = win32com.client.Dispatch("Outlook.Application")
#     ns = outlook.GetNamespace("MAPI")
#     mail = ns.OpenSharedItem(str(msg_path))

#     # ---------- Header ----------
#     hdr = {
#         "From": mail.SenderEmailAddress,
#         "To": ", ".join(rec.Address for rec in mail.Recipients),
#         "Subject": mail.Subject,
#         "Date": mail.SentOn.strftime("%a, %d %b %Y %H:%M:%S %z"),
#         "Message-ID": f"<{uuid.uuid4()}@outlook>",
#     }

#     # ---------- Body ----------
#     html_body = mail.HTMLBody or ""
#     plain_body = mail.Body or ""

#     # RTF‑Body (falls vorhanden)
#     have_rtf = False
#     rtf_bytes = b""
#     if getattr(mail, "BodyFormat", 1) == 3:          # 3 = olFormatRichText
#         # Outlook liefert den RTF‑String als Unicode‑Text → wir codieren ihn in UTF‑8
#         # (alternativ: mail.RTFBody gibt binäre Daten, hier nicht verwendet)
#         rtf_bytes = mail.Body.encode("utf-8")
#         have_rtf = True

#     # ---------- Attachments ----------
#     # Wir bauen drei Listen:
#     inline_parts = []      # (cid, maintype, subtype, data, filename)
#     normal_parts = []      # (maintype, subtype, data, filename)
#     rtf_inline = None      # (maintype, subtype, data) falls als inline‑RTF eingebettet

#     # MAPI‑Konstanten (aus Outlook‑Objekt‑Modell)
#     ATTACH_EMBEDDED_MSG = 0x00000004
#     ATTACH_OLE          = 0x00000008
#     ATTACH_EMBEDDED_OLE = 0x00000010

#     # Hilfsfunktion zum Lesen eines Attachments in ein bytes‑Objekt
#     def _read_attachment(att):
#         tmp = Path.cwd() / f"__tmp_{uuid.uuid4().hex}{Path(att.FileName).suffix}"
#         att.SaveAsFile(str(tmp))
#         data = tmp.read_bytes()
#         tmp.unlink(missing_ok=True)
#         return data

#     # Durch alle Attachments iterieren
#     for i in range(1, mail.Attachments.Count + 1):
#         att = mail.Attachments.Item(i)

#         # Prüfen, ob es ein “embedded” Item ist (inline‑Bild o.ä.)
#         # Der PropertyAccessor liefert das Flag PR_ATTACHMENT_FLAGS (0x3712)
#         prop = att.PropertyAccessor
#         try:
#             flags = prop.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37120003")
#         except Exception:
#             flags = 0

#         data = _read_attachment(att)
#         suffix = Path(att.FileName).suffix
#         maintype, subtype = _guess_mime_type(suffix)

#         # ------------------------------------------------------------
#         # 1️⃣ Inline‑Bilder (PR_ATTACHMENT_FLAGS & ATTACH_OLE == 0)
#         #    – Outlook kennzeichnet sie nicht mit speziellen Flags,
#         #      aber sie haben typischerweise eine Bild‑Extension.
#         # ------------------------------------------------------------
#         if suffix.lower() in {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".svg"}:
#             cid = f"<{uuid.uuid4()}@outlook>"
#             inline_parts.append((cid, maintype, subtype, data, att.FileName))
#             # Ersetze ggf. Referenzen im HTML‑Body:
#             # Outlook speichert oft src="cid:AttachmentX"
#             # Wir ersetzten ANY "src="cid:…"" durch unseren cid.
#             # Dafür bauen wir ein simples Replace‑Muster:
#             html_body = html_body.replace(
#                 f"src=\"cid:{i}\"", f"src=\"cid:{cid[1:-1]}\""
#             ).replace(
#                 f"src='cid:{i}'", f"src='cid:{cid[1:-1]}'"
#             )
#             continue

#         # ------------------------------------------------------------
#         # 2️⃣ RTF‑Teil (wenn der Attachment‑Typ "rtf" ist)
#         # ------------------------------------------------------------
#         if suffix.lower() == ".rtf":
#             # RTF‑Datei kann als eigener Teil (text/rtf) eingefügt werden.
#             # Wir behandeln sie nicht als normalen Anhang.
#             rtf_inline = (maintype, subtype, data, att.FileName)
#             have_rtf = True
#             continue

#         # ------------------------------------------------------------
#         # 3️⃣ OLE/Embedded‑Message (z. B. ein eingebettetes Word‑Doc)
#         # ------------------------------------------------------------
#         if flags & ATTACH_OLE or flags & ATTACH_EMBEDDED_OLE:
#             # Wir behandeln es als normalen Anhang – Clients können es öffnen.
#             normal_parts.append((maintype, subtype, data, att.FileName))
#             continue

#         # ------------------------------------------------------------
#         # 4️⃣ Alles andere = "normale" Anlage
#         # ------------------------------------------------------------
#         normal_parts.append((maintype, subtype, data, att.FileName))

#     # ---------- 5. MIME‑Struktur bauen ----------
#     # outer = multipart/mixed (wenn wir normale Anhänge haben)
#     # inner = multipart/related (wenn inline‑Bilder oder RTF‑inline existieren)
#     # base  = multipart/alternative (plain + html [+ rtf])

#     # 1️⃣ Basis‑Message (outer)
#     outer = EmailMessage(policy=email.policy.SMTP)

#     # Header eintragen
#     for k, v in hdr.items():
#         outer[k] = v

#     # Wir benötigen multipart/mixed nur, wenn es **normale** Anhänge gibt
#     need_mixed = len(normal_parts) > 0
#     need_related = len(inline_parts) > 0 or rtf_inline is not None

#     if need_mixed:
#         outer.set_type("multipart/mixed")
#     else:
#         # ohne mixed verwenden wir gleich das "root"-Objekt für das nächste Level
#         outer = outer  # nichts zu ändern

#     # 2️⃣ optional: multipart/related (für inline‑Bilder / rtf)
#     if need_related:
#         related = EmailMessage()
#         related.set_type("multipart/related")
#         # Das Relation-Objekt wird an outer (oder an outer’s Teil) angehängt
#         if need_mixed:
#             outer.attach(related)
#         else:
#             outer = related   # es wird zum neuen Root
#     else:
#         related = None

#     # 3️⃣ multipart/alternative (plain + html [+ rtf])
#     alternative = EmailMessage()
#     alternative.set_type("multipart/alternative")

#     if related:
#         related.attach(alternative)
#     else:
#         if need_mixed:
#             outer.attach(alternative)
#         else:
#             outer = alternative   # kein mixed, kein related → root ist alternative

#     # ---- plain text ----
#     alternative.set_content(plain_body, subtype="plain", charset="utf-8")

#     # ---- html ----
#     alternative.add_alternative(html_body, subtype="html", charset="utf-8")

#     # ---- optional rtf (als eigener Teil, nicht inline) ----
#     if have_rtf and rtf_bytes:
#         # RTF kann als eigenständiger MIME‑Teil (text/rtf) an das
#         # multipart/alternative hängen – das ist laut RFC erlaubt.
#         alternative.add_alternative(
#             rtf_bytes.decode("utf-8", errors="ignore"),
#             subtype="rtf",
#             charset="utf-8",
#         )

#     # ---- inline‑Bilder (wenn vorhanden) ----
#     if inline_parts:
#         for cid, mtype, stype, data, filename in inline_parts:
#             # Bild‑Part hinzufügen
#             related.add_related(
#                 data,
#                 maintype=mtype,
#                 subtype=stype,
#                 cid=cid,                     # <…> already enthalten
#                 filename=filename,
#                 disposition="inline",        # wichtig für "inline"
#             )

#     # ---- RTF als inline (falls wir ein RTF‑Attachment hatten) ----
#     if rtf_inline:
#         mtype, stype, data, filename = rtf_inline
#         # RFC‑empfohlen: `text/rtf` als inline‑Teil, Content‑ID optional
#         related.add_related(
#             data,
#             maintype=mtype,
#             subtype=stype,
#             filename=filename,
#             disposition="inline",
#         )

#     # ---- normale Anhänge (wenn vorhanden) ----
#     if normal_parts:
#         for mtype, stype, data, filename in normal_parts:
#             outer.add_attachment(
#                 data,
#                 maintype=mtype,
#                 subtype=stype,
#                 filename=filename,
#                 disposition="attachment",
#             )

#     # ---------- 6. Bytes erzeugen ----------
#     buf = BytesIO()
#     gen = BytesGenerator(buf, policy=email.policy.SMTP)   # outfp = buf
#     gen.flatten(outer)
#     return buf.getvalue()









# Korrigierte Version wegen Fehler "set_content not valid on multipart"
# ------------------------------------------------------------
# Hilfsfunktion: MIME‑Typ anhand Dateiendung bestimmen
# ------------------------------------------------------------
# def _guess_mime_type(suffix: str):
#     suffix = suffix.lower()
#     if suffix in {".jpg", ".jpeg"}:
#         return ("image", "jpeg")
#     if suffix == ".png":
#         return ("image", "png")
#     if suffix == ".gif":
#         return ("image", "gif")
#     if suffix == ".bmp":
#         return ("image", "bmp")
#     if suffix == ".svg":
#         return ("image", "svg+xml")
#     if suffix == ".pdf":
#         return ("application", "pdf")
#     if suffix == ".doc":
#         return ("application", "msword")
#     if suffix == ".docx":
#         return ("application",
#                 "vnd.openxmlformats-officedocument.wordprocessingml.document")
#     if suffix == ".xls":
#         return ("application", "vnd.ms-excel")
#     if suffix == ".xlsx":
#         return ("application",
#                 "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
#     # Fallback
#     return ("application", "octet-stream")

# # ------------------------------------------------------------
# def msg_to_eml_via_outlook(msg_path: Path) -> bytes:
#     """
#     Exportiert eine .msg‑Datei zu einer voll‑wertigen RFC‑822‑Nachricht,
#     inkl.:
#         • plain‑text + html (multipart/alternative)
#         • inline‑Bilder (Content‑ID, multipart/related)
#         • RTF‑Body (optional, text/rtf)
#         • normale Anlagen (multipart/mixed)
#     """
#     # ---------- Outlook ----------
#     outlook = win32com.client.Dispatch("Outlook.Application")
#     ns = outlook.GetNamespace("MAPI")
#     mail = ns.OpenSharedItem(str(msg_path))

#     # ---------- Header ----------
#     hdr = {
#         "From": mail.SenderEmailAddress,
#         "To": ", ".join(rec.Address for rec in mail.Recipients),
#         "Subject": mail.Subject,
#         "Date": mail.SentOn.strftime("%a, %d %b %Y %H:%M:%S %z"),
#         "Message-ID": f"<{uuid.uuid4()}@outlook>",
#     }

#     # ---------- Body ----------
#     html_body = mail.HTMLBody or ""
#     plain_body = mail.Body or ""

#     # RTF‑Body (wenn die Nachricht im Rich‑Text‑Format vorliegt)
#     have_rtf_body = False
#     rtf_body_bytes = b""
#     if getattr(mail, "BodyFormat", 1) == 3:           # 3 == olFormatRichText
#         # Outlook liefert den RTF‑String als Unicode‑Text → UTF‑8‑Kodierung
#         rtf_body_bytes = mail.Body.encode("utf-8")
#         have_rtf_body = True

#     # ---------- Attachments ----------
#     inline_parts = []   # (cid, maintype, subtype, data, filename)
#     normal_parts = []   # (maintype, subtype, data, filename)
#     rtf_inline = None   # optional (maintype, subtype, data, filename)

#     # MAPI‑Flag‑Konstanten (aus Microsoft‑Docs)
#     ATTACH_OLE          = 0x00000008
#     ATTACH_EMBEDDED_OLE = 0x00000010

#     def _read_attachment(att):
#         """Speichert das Attachment kurzzeitig in eine Datei und liest die Bytes."""
#         tmp = Path.cwd() / f"__tmp_{uuid.uuid4().hex}{Path(att.FileName).suffix}"
#         att.SaveAsFile(str(tmp))
#         data = tmp.read_bytes()
#         tmp.unlink(missing_ok=True)
#         return data

#     for i in range(1, mail.Attachments.Count + 1):
#         att = mail.Attachments.Item(i)

#         # MAPI‑Flags des Attachments (PR_ATTACHMENT_FLAGS, 0x37120003)
#         try:
#             flags = att.PropertyAccessor.GetProperty(
#                 "http://schemas.microsoft.com/mapi/proptag/0x37120003"
#             )
#         except Exception:
#             flags = 0

#         data = _read_attachment(att)
#         suffix = Path(att.FileName).suffix
#         maintype, subtype = _guess_mime_type(suffix)

#         # ---- Inline‑Bilder (nach Endung erkennen) ----
#         if suffix.lower() in {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".svg"}:
#             cid = f"<{uuid.uuid4()}@outlook>"
#             inline_parts.append((cid, maintype, subtype, data, att.FileName))

#             # Ersetze Outlook‑typische Referenz "cid:attachmentX" durch unser CID
#             html_body = html_body.replace(
#                 f'src="cid:{i}"', f'src="cid:{cid[1:-1]}"'
#             ).replace(
#                 f"src='cid:{i}'", f"src='cid:{cid[1:-1]}'"
#             )
#             continue

#         # ---- RTF‑Attachment (falls vorhanden) ----
#         if suffix.lower() == ".rtf":
#             # Wir behandeln das RTF als eigenen Inline‑Teil, nicht als normalen Anhang
#             rtf_inline = (maintype, subtype, data, att.FileName)
#             have_rtf_body = True          # signalisiere, dass RTF vorhanden ist
#             continue

#         # ---- OLE / Embedded‑Message ----
#         if flags & ATTACH_OLE or flags & ATTACH_EMBEDDED_OLE:
#             # Wir legen diese einfach als normalen Anhang ab – Clients können sie öffnen.
#             normal_parts.append((maintype, subtype, data, att.FileName))
#             continue

#         # ---- alles andere = regulärer Anhang ----
#         normal_parts.append((maintype, subtype, data, att.FileName))

#     # ---------- 5. MIME‑Struktur bauen ----------
#     # 1️⃣ Root‑Message (outer)
#     outer = EmailMessage(policy=email.policy.SMTP)

#     # Header eintragen
#     for k, v in hdr.items():
#         outer[k] = v

#     need_mixed   = len(normal_parts) > 0
#     need_related = len(inline_parts) > 0 or rtf_inline is not None

#     # -----------------------------------------------------------------
#     #   multipart/mixed  (nur wenn es "normale" Anhänge gibt)
#     # -----------------------------------------------------------------
#     if need_mixed:
#         outer.set_type("multipart/mixed")

#     # -----------------------------------------------------------------
#     #   multipart/related  (für inline‑Bilder + evtl. RTF‑Inline)
#     # -----------------------------------------------------------------
#     if need_related:
#         related = EmailMessage()
#         related.set_type("multipart/related")
#         if need_mixed:
#             outer.attach(related)
#         else:
#             outer = related          # kein mixed → related wird das Root‑Objekt
#     else:
#         related = None

#     # -----------------------------------------------------------------
#     #   multipart/alternative  (plain + html + optional rtf)
#     # -----------------------------------------------------------------
#     #   Wichtig: Wir **erzeugen** ein *normales* EmailMessage‑Objekt,
#     #   rufen dann set_content() auf → es wird automatisch zu multipart/alternative.
#     alternative = EmailMessage()                # kein set_type hier!
#     # Plain‑Text‑Teil (erstes)
#     alternative.set_content(plain_body, subtype="plain", charset="utf-8")
#     # HTML‑Teil (zweites)
#     alternative.add_alternative(html_body, subtype="html", charset="utf-8")
#     # Optionaler RTF‑Teil (Dritte) – nur wenn wir ein echtes RTF‑Body haben
#     if have_rtf_body and rtf_body_bytes:
#         # RFC‑konform: text/rtf kann als eigener Teil im alternative‑Block stehen
#         alternative.add_alternative(
#             rtf_body_bytes.decode("utf-8", errors="ignore"),
#             subtype="rtf",
#             charset="utf-8",
#         )

#     # Jetzt das fertige alternative-Objekt in den übergeordneten Container hängen:
#     if need_related:
#         related.attach(alternative)
#     elif need_mixed:
#         outer.attach(alternative)
#     else:
#         # weder mixed noch related → alternative ist das Root‑Objekt
#         outer = alternative

#     # -----------------------------------------------------------------
#     #   Inline‑Bilder (und ggf. RTF‑Attachment) in multipart/related einbetten
#     # -----------------------------------------------------------------
#     if need_related:
#         for cid, mtype, stype, data, filename in inline_parts:
#             related.add_related(
#                 data,
#                 maintype=mtype,
#                 subtype=stype,
#                 cid=cid,               # <…>
#                 filename=filename,
#                 disposition="inline",
#             )

#         if rtf_inline:
#             mtype, stype, data, filename = rtf_inline
#             # RTF‑Attachment wird ebenfalls als inline‑Teil eingefügt
#             related.add_related(
#                 data,
#                 maintype=mtype,
#                 subtype=stype,
#                 filename=filename,
#                 disposition="inline",
#             )

#     # -----------------------------------------------------------------
#     #   Normale Anhänge (multipart/mixed) anhängen
#     # -----------------------------------------------------------------
#     if need_mixed:
#         for mtype, stype, data, filename in normal_parts:
#             outer.add_attachment(
#                 data,
#                 maintype=mtype,
#                 subtype=stype,
#                 filename=filename,
#                 disposition="attachment",
#             )

#     # ---------- 6. Bytes erzeugen ----------
#     buf = BytesIO()
#     gen = BytesGenerator(buf, policy=email.policy.SMTP)   # outfp ist buf
#     gen.flatten(outer)

#     return buf.getvalue()








# Korrigierte Version, um die Header Zeolen immer im Kopf des Fensters zu bekommen
# ------------------------------------------------------------
# def _guess_mime_type(suffix: str):
#     """Einfache Zuordnung von Dateiendungen zu MIME‑Typen."""
#     suffix = suffix.lower()
#     if suffix in {".jpg", ".jpeg"}:
#         return ("image", "jpeg")
#     if suffix == ".png":
#         return ("image", "png")
#     if suffix == ".gif":
#         return ("image", "gif")
#     if suffix == ".bmp":
#         return ("image", "bmp")
#     if suffix == ".svg":
#         return ("image", "svg+xml")
#     if suffix == ".pdf":
#         return ("application", "pdf")
#     if suffix == ".doc":
#         return ("application", "msword")
#     if suffix == ".docx":
#         return ("application",
#                 "vnd.openxmlformats-officedocument.wordprocessingml.document")
#     if suffix == ".xls":
#         return ("application", "vnd.ms-excel")
#     if suffix == ".xlsx":
#         return ("application",
#                 "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
#     return ("application", "octet-stream")
# # ------------------------------------------------------------

# def msg_to_eml_via_outlook(msg_path: Path) -> bytes:
#     """Exportiert .msg → .eml inkl. Inline‑Bilder, RTF und Anhänge."""
#     # ---------- 1. Outlook ----------
#     outlook = win32com.client.Dispatch("Outlook.Application")
#     ns = outlook.GetNamespace("MAPI")
#     mail = ns.OpenSharedItem(str(msg_path))

#     # ---------- 2. Header ----------
#     hdr = {
#         "From": mail.SenderEmailAddress,
#         "To": ", ".join(rec.Address for rec in mail.Recipients),
#         "Subject": mail.Subject,
#         "Date": mail.SentOn.strftime("%a, %d %b %Y %H:%M:%S %z"),
#         "Message-ID": f"<{uuid.uuid4()}@outlook>",
#     }

#     # ---------- 3. Body ----------
#     html_body = mail.HTMLBody or ""
#     plain_body = mail.Body or ""

#     # RTF‑Body, wenn die Originalnachricht im Rich‑Text‑Format ist
#     have_rtf_body = False
#     rtf_body_bytes = b""
#     if getattr(mail, "BodyFormat", 1) == 3:            # 3 = olFormatRichText
#         rtf_body_bytes = mail.Body.encode("utf-8")
#         have_rtf_body = True

#     # ---------- 4. Attachments sammeln ----------
#     inline_parts = []   # (cid, maintype, subtype, data, filename)
#     normal_parts = []   # (maintype, subtype, data, filename)
#     rtf_inline = None   # (maintype, subtype, data, filename)

#     ATTACH_OLE          = 0x00000008
#     ATTACH_EMBEDDED_OLE = 0x00000010

#     def _read_att(att):
#         tmp = Path.cwd() / f"__tmp_{uuid.uuid4().hex}{Path(att.FileName).suffix}"
#         att.SaveAsFile(str(tmp))
#         data = tmp.read_bytes()
#         tmp.unlink(missing_ok=True)
#         return data

#     for i in range(1, mail.Attachments.Count + 1):
#         att = mail.Attachments.Item(i)

#         # PR_ATTACHMENT_FLAGS (0x3712) – liefert Flags zu Inline‑ / OLE‑Attachments
#         try:
#             flags = att.PropertyAccessor.GetProperty(
#                 "http://schemas.microsoft.com/mapi/proptag/0x37120003"
#             )
#         except Exception:
#             flags = 0

#         data = _read_att(att)
#         suffix = Path(att.FileName).suffix
#         maintype, subtype = _guess_mime_type(suffix)

#         # ---- Inline‑Bilder (nach Endung) ----
#         if suffix.lower() in {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".svg"}:
#             cid = f"<{uuid.uuid4()}@outlook>"
#             inline_parts.append((cid, maintype, subtype, data, att.FileName))

#             # Ersetze Outlook‑interne cid‑Referenzen (cid:1, cid:2, …)
#             html_body = html_body.replace(
#                 f'src="cid:{i}"', f'src="cid:{cid[1:-1]}"'
#             ).replace(
#                 f"src='cid:{i}'", f"src='cid:{cid[1:-1]}'"
#             )
#             continue

#         # ---- RTF‑Attachment (falls vorhanden) ----
#         if suffix.lower() == ".rtf":
#             rtf_inline = (maintype, subtype, data, att.FileName)
#             have_rtf_body = True
#             continue

#         # ---- OLE / Embedded‑Message ----
#         if flags & ATTACH_OLE or flags & ATTACH_EMBEDDED_OLE:
#             normal_parts.append((maintype, subtype, data, att.FileName))
#             continue

#         # ---- Alles andere = normaler Anhang ----
#         normal_parts.append((maintype, subtype, data, att.FileName))

#     # ---------- 5. MIME‑Struktur bauen ----------
#     # 5.1 Root‑Message (immer existent) – hier kommen die Header hin
#     root = EmailMessage(policy=email.policy.SMTP)
#     for k, v in hdr.items():
#         root[k] = v

#     need_mixed   = len(normal_parts) > 0
#     need_related = len(inline_parts) > 0 or rtf_inline is not None

#     # 5.2 optionaler multipart/mixed‑Container (nur bei normalen Anhängen)
#     if need_mixed:
#         root.set_type("multipart/mixed")
#         container = root                     # wir hängen alles an dieses Objekt
#     else:
#         container = root                     # kein mixed → root ist unser Container

#     # 5.3 optionaler multipart/related‑Container (für Inline‑Bilder/RTF)
#     if need_related:
#         related = EmailMessage()
#         related.set_type("multipart/related")
#         # Wenn wir bereits einen multipart/mixed‑Container haben,
#         # hängen wir `related` dort an, sonst wird er das neue Container‑Objekt.
#         container.attach(related)
#         container = related                  # ab jetzt hängen wir an `related` an
#         # (für den Fall, dass kein mixed existiert, ist `container` jetzt `related`)

#     # 5.4 multipart/alternative – enthält plain, html und ggf. rtf
#     alternative = EmailMessage()             # kein set_type, wird durch set_content()
#                                             # automatisch zu multipart/alternative
#     alternative.set_content(plain_body, subtype="plain", charset="utf-8")
#     alternative.add_alternative(html_body, subtype="html", charset="utf-8")
#     if have_rtf_body and rtf_body_bytes:
#         alternative.add_alternative(
#             rtf_body_bytes.decode("utf-8", errors="ignore"),
#             subtype="rtf",
#             charset="utf-8",
#         )
#     # Alternative in den derzeitigen Container (mixed / related / root) einhängen
#     container.attach(alternative)

#     # 5.5 Inline‑Bilder (nur wenn we have_related == True)
#     if need_related:
#         for cid, mtype, stype, data, filename in inline_parts:
#             related.add_related(
#                 data,
#                 maintype=mtype,
#                 subtype=stype,
#                 cid=cid,               # <…>
#                 filename=filename,
#                 disposition="inline",
#             )
#         if rtf_inline:
#             mtype, stype, data, filename = rtf_inline
#             related.add_related(
#                 data,
#                 maintype=mtype,
#                 subtype=stype,
#                 filename=filename,
#                 disposition="inline",
#             )

#     # 5.6 Normale Anhänge (nur wenn need_mixed == True)
#     if need_mixed:
#         for mtype, stype, data, filename in normal_parts:
#             root.add_attachment(
#                 data,
#                 maintype=mtype,
#                 subtype=stype,
#                 filename=filename,
#                 disposition="attachment",
#             )

#     # ---------- 6. Bytes erzeugen ----------
#     buf = BytesIO()
#     gen = BytesGenerator(buf, policy=email.policy.SMTP)   # outfp = buf
#     gen.flatten(root)
#     return buf.getvalue()






# Korrigierte Version wegen Fehler "'list' object has no attribute 'encode'"
# ------------------------------------------------------------
# def _guess_mime_type(suffix: str):
#     suffix = suffix.lower()
#     if suffix in {".jpg", ".jpeg"}:
#         return ("image", "jpeg")
#     if suffix == ".png":
#         return ("image", "png")
#     if suffix == ".gif":
#         return ("image", "gif")
#     if suffix == ".bmp":
#         return ("image", "bmp")
#     if suffix == ".svg":
#         return ("image", "svg+xml")
#     if suffix == ".pdf":
#         return ("application", "pdf")
#     if suffix == ".doc":
#         return ("application", "msword")
#     if suffix == ".docx":
#         return ("application",
#                 "vnd.openxmlformats-officedocument.wordprocessingml.document")
#     if suffix == ".xls":
#         return ("application", "vnd.ms-excel")
#     if suffix == ".xlsx":
#         return ("application",
#                 "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
#     return ("application", "octet-stream")

# def _ensure_str(value):
#     """Konvertiert Listen/Tuples zu kommagetrennten Strings."""
#     if isinstance(value, (list, tuple)):
#         return ", ".join(str(v) for v in value)
#     return str(value) if value is not None else ""

# def _ensure_body(value):
#     """Stellt sicher, dass der Body ein einfacher String ist."""
#     if isinstance(value, (list, tuple)):
#         return "\n".join(str(v) for v in value)
#     return str(value) if value is not None else ""

# # ------------------------------------------------------------
# def msg_to_eml_via_outlook(msg_path: Path) -> bytes:
#     """Exportiert .msg → .eml inkl. Inline‑Bilder, RTF und Anhänge."""
#     outlook = win32com.client.Dispatch("Outlook.Application")
#     ns = outlook.GetNamespace("MAPI")
#     mail = ns.OpenSharedItem(str(msg_path))

#     # ----- 1. Header (immer Strings) -----
#     hdr = {
#         "From": _ensure_str(mail.SenderEmailAddress),
#         "To":   _ensure_str([rec.Address for rec in mail.Recipients]),
#         "Subject": _ensure_str(mail.Subject),
#         "Date": mail.SentOn.strftime("%a, %d %b %Y %H:%M:%S %z"),
#         "Message-ID": f"<{uuid.uuid4()}@outlook>",
#     }
#     # Optional Cc / Bcc (falls vorhanden)
#     if hasattr(mail, "CC"):
#         hdr["Cc"] = _ensure_str(mail.CC)
#     if hasattr(mail, "BCC"):
#         hdr["Bcc"] = _ensure_str(mail.BCC)

#     # ----- 2. Body (plain, html, ggf. rtf) -----
#     plain_body = _ensure_body(mail.Body)
#     html_body  = _ensure_body(mail.HTMLBody)

#     have_rtf_body = False
#     rtf_body_bytes = b""
#     if getattr(mail, "BodyFormat", 1) == 3:        # 3 = Rich‑Text‑Format
#         rtf_body_bytes = _ensure_body(mail.Body).encode("utf-8")
#         have_rtf_body = True

#     # ----- 3. Attachments sammeln -----
#     inline_parts = []   # (cid, maintype, subtype, data, filename)
#     normal_parts = []   # (maintype, subtype, data, filename)
#     rtf_inline = None   # (maintype, subtype, data, filename)

#     ATTACH_OLE          = 0x00000008
#     ATTACH_EMBEDDED_OLE = 0x00000010

#     def _read_att(att):
#         tmp = Path.cwd() / f"__tmp_{uuid.uuid4().hex}{Path(att.FileName).suffix}"
#         att.SaveAsFile(str(tmp))
#         data = tmp.read_bytes()
#         tmp.unlink(missing_ok=True)
#         return data

#     for i in range(1, mail.Attachments.Count + 1):
#         att = mail.Attachments.Item(i)

#         try:
#             flags = att.PropertyAccessor.GetProperty(
#                 "http://schemas.microsoft.com/mapi/proptag/0x37120003"
#             )
#         except Exception:
#             flags = 0

#         data = _read_att(att)
#         suffix = Path(att.FileName).suffix
#         maintype, subtype = _guess_mime_type(suffix)

#         # ----- Inline‑Bilder -----
#         if suffix.lower() in {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".svg"}:
#             cid = f"<{uuid.uuid4()}@outlook>"
#             inline_parts.append((cid, maintype, subtype, data, att.FileName))
#             # Outlook benutzt manchmal src="cid:1", src="cid:2" …
#             html_body = html_body.replace(
#                 f'src="cid:{i}"', f'src="cid:{cid[1:-1]}"'
#             ).replace(
#                 f"src='cid:{i}'", f"src='cid:{cid[1:-1]}'"
#             )
#             continue

#         # ----- RTF‑Attachment (wenn vorhanden) -----
#         if suffix.lower() == ".rtf":
#             rtf_inline = (maintype, subtype, data, att.FileName)
#             have_rtf_body = True
#             continue

#         # ----- OLE / Embedded‑Message -----
#         if flags & ATTACH_OLE or flags & ATTACH_EMBEDDED_OLE:
#             normal_parts.append((maintype, subtype, data, att.FileName))
#             continue

#         # ----- Alle übrigen Anhänge -----
#         normal_parts.append((maintype, subtype, data, att.FileName))

#     # ----- 4. MIME‑Struktur bauen -----
#     root = EmailMessage(policy=email.policy.SMTP)
#     for k, v in hdr.items():
#         root[k] = v

#     need_mixed   = len(normal_parts) > 0
#     need_related = len(inline_parts) > 0 or rtf_inline is not None

#     # 4.1 optional multipart/mixed
#     if need_mixed:
#         root.set_type("multipart/mixed")
#         container = root
#     else:
#         container = root

#     # 4.2 optional multipart/related
#     if need_related:
#         related = EmailMessage()
#         related.set_type("multipart/related")
#         container.attach(related)
#         container = related           # ab jetzt hängen wir an `related` an

#     # 4.3 multipart/alternative (plain, html, optional rtf)
#     alternative = EmailMessage()      # kein set_type – wird automatisch multipart/alternative
#     alternative.set_content(plain_body, subtype="plain", charset="utf-8")
#     alternative.add_alternative(html_body, subtype="html", charset="utf-8")
#     if have_rtf_body and rtf_body_bytes:
#         alternative.add_alternative(
#             rtf_body_bytes.decode("utf-8", errors="ignore"),
#             subtype="rtf",
#             charset="utf-8",
#         )
#     container.attach(alternative)

#     # 4.4 Inline‑Bilder & evtl. RTF‑Attachment in multipart/related einbetten
#     if need_related:
#         for cid, mtype, stype, data, filename in inline_parts:
#             related.add_related(
#                 data,
#                 maintype=mtype,
#                 subtype=stype,
#                 cid=cid,
#                 filename=filename,
#                 disposition="inline",
#             )
#         if rtf_inline:
#             mtype, stype, data, filename = rtf_inline
#             related.add_related(
#                 data,
#                 maintype=mtype,
#                 subtype=stype,
#                 filename=filename,
#                 disposition="inline",
#             )

#     # 4.5 Normale Anhänge (unter multipart/mixed)
#     if need_mixed:
#         for mtype, stype, data, filename in normal_parts:
#             root.add_attachment(
#                 data,
#                 maintype=mtype,
#                 subtype=stype,
#                 filename=filename,
#                 disposition="attachment",
#             )

#     # ----- 5. Bytes erzeugen -----
#     from io import BytesIO
#     buf = BytesIO()
#     gen = BytesGenerator(buf, policy=email.policy.SMTP)   # outfp = buf
#     gen.flatten(root)
#     return buf.getvalue()




# Korrigierte Version, da immer noch Fehler "'list' object has no attribute 'encode'"
import win32com.client
import email
from email.message import EmailMessage
from email.generator import BytesGenerator
import email.policy
from pathlib import Path
import uuid
import datetime

# ------------------------------------------------------------
def _guess_mime_type(suffix: str):
    """Einfacher MIME‑Typ‑Mapper."""
    suffix = suffix.lower()
    if suffix in {".jpg", ".jpeg"}:
        return ("image", "jpeg")
    if suffix == ".png":
        return ("image", "png")
    if suffix == ".gif":
        return ("image", "gif")
    if suffix == ".bmp":
        return ("image", "bmp")
    if suffix == ".svg":
        return ("image", "svg+xml")
    if suffix == ".pdf":
        return ("application", "pdf")
    if suffix == ".doc":
        return ("application", "msword")
    if suffix == ".docx":
        return ("application",
                "vnd.openxmlformats-officedocument.wordprocessingml.document")
    if suffix == ".xls":
        return ("application", "vnd.ms-excel")
    if suffix == ".xlsx":
        return ("application",
                "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return ("application", "octet-stream")

# ------------------------------------------------------------
def _to_str(value):
    """Konvertiert beliebige Eingaben (Liste, Tuple, None, etc.) in einen String."""
    if isinstance(value, (list, tuple)):
        # Kommagetrennte Darstellung
        return ", ".join(str(v) for v in value if v is not None)
    if value is None:
        return ""
    return str(value)

def _to_body(value):
    """Stellt sicher, dass der Body ein einfacher String ist."""
    if isinstance(value, (list, tuple)):
        return "\n".join(str(v) for v in value if v is not None)
    if value is None:
        return ""
    return str(value)

# ------------------------------------------------------------
def msg_to_eml_via_outlook_fixed(msg_path: Path) -> bytes:
    """Exportiert .msg → .eml – komplett robust gegen List‑Erscheinungen."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    mail = ns.OpenSharedItem(str(msg_path))

    # -------------------- 1. Header --------------------
    hdr = {
        "From": _to_str(mail.SenderEmailAddress),
        "To":   _to_str([rec.Address for rec in mail.Recipients]),
        "Subject": _to_str(mail.Subject),
        "Date": mail.SentOn.strftime("%a, %d %b %Y %H:%M:%S %z"),
        "Message-ID": f"<{uuid.uuid4()}@outlook>",
    }

    # Optional Cc / Bcc – Outlook liefert hier häufig Listen
    if hasattr(mail, "CC"):
        hdr["Cc"] = _to_str(getattr(mail, "CC"))
    if hasattr(mail, "BCC"):
        hdr["Bcc"] = _to_str(getattr(mail, "BCC"))

    # -------------------- 2. Body --------------------
    plain_body = _to_body(mail.Body)
    html_body  = _to_body(mail.HTMLBody)

    # RTF‑Body (falls das Original im RTF‑Format vorliegt)
    have_rtf_body = False
    rtf_body_bytes = b""
    if getattr(mail, "BodyFormat", 1) == 3:   # 3 == olFormatRichText
        rtf_body_bytes = _to_body(mail.Body).encode("utf-8")
        have_rtf_body = True

    # -------------------- 3. Attachments --------------------
    inline_parts = []   # (cid, maintype, subtype, data, filename)
    normal_parts = []   # (maintype, subtype, data, filename)
    rtf_inline = None   # (maintype, subtype, data, filename)

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

        # MAPI‑Flag‑Abfrage – gibt ein Integer zurück, kann aber fehlen
        try:
            flags = att.PropertyAccessor.GetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x37120003"
            )
        except Exception:
            flags = 0

        data = _read_att(att)
        suffix = Path(att.FileName).suffix
        maintype, subtype = _guess_mime_type(suffix)

        # --------------------------------------------------------
        # Inline‑Bilder (Erkennung über Dateiendung)
        # --------------------------------------------------------
        if suffix.lower() in {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".svg"}:
            cid = f"<{uuid.uuid4()}@outlook>"
            inline_parts.append((cid, maintype, subtype, data, att.FileName))

            # Outlook referenziert Inline‑Bilder oft mit src="cid:1", …
            html_body = html_body.replace(
                f'src="cid:{i}"', f'src="cid:{cid[1:-1]}"'
            ).replace(
                f"src='cid:{i}'", f"src='cid:{cid[1:-1]}'"
            )
            continue

        # --------------------------------------------------------
        # RTF‑Attachment (wenn vorhanden)
        # --------------------------------------------------------
        if suffix.lower() == ".rtf":
            rtf_inline = (maintype, subtype, data, att.FileName)
            have_rtf_body = True
            continue

        # --------------------------------------------------------
        # OLE / Embedded‑Message (z. B. Word‑Dokument als OLE‑Objekt)
        # --------------------------------------------------------
        if flags & ATTACH_OLE or flags & ATTACH_EMBEDDED_OLE:
            normal_parts.append((maintype, subtype, data, att.FileName))
            continue

        # --------------------------------------------------------
        # Alles andere = normaler Anhang
        # --------------------------------------------------------
        normal_parts.append((maintype, subtype, data, att.FileName))

    # -------------------- 4. MIME‑Struktur --------------------
    root = EmailMessage(policy=email.policy.SMTP)

    # Header immer als String eintragen
    for k, v in hdr.items():
        root[k] = _to_str(v)

    need_mixed   = len(normal_parts) > 0
    need_related = len(inline_parts) > 0 or rtf_inline is not None

    # ------ multipart/mixed (falls nötig) ------
    if need_mixed:
        root.set_type("multipart/mixed")
        container = root
    else:
        container = root

    # ------ multipart/related (falls nötig) ------
    if need_related:
        related = EmailMessage()
        related.set_type("multipart/related")
        container.attach(related)
        container = related   # ab jetzt hängen wir an `related` an

    # ------ multipart/alternative (plain + html + optional rtf) ------
    alternative = EmailMessage()        # kein set_type – wird durch set_content()
    alternative.set_content(_to_body(plain_body), subtype="plain", charset="utf-8")
    alternative.add_alternative(_to_body(html_body), subtype="html", charset="utf-8")
    if have_rtf_body and rtf_body_bytes:
        alternative.add_alternative(
            rtf_body_bytes.decode("utf-8", errors="ignore"),
            subtype="rtf",
            charset="utf-8",
        )
    container.attach(alternative)

    # ------ Inline‑Bilder & ggf. RTF‑Attachment in multipart/related ------
    if need_related:
        for cid, mtype, stype, data, filename in inline_parts:
            related.add_related(
                data,
                maintype=mtype,
                subtype=stype,
                cid=cid,
                filename=filename,
                disposition="inline",
            )
        if rtf_inline:
            mtype, stype, data, filename = rtf_inline
            related.add_related(
                data,
                maintype=mtype,
                subtype=stype,
                filename=filename,
                disposition="inline",
            )

    # ------ Normale Anhänge (unter multipart/mixed) ------
    if need_mixed:
        for mtype, stype, data, filename in normal_parts:
            root.add_attachment(
                data,
                maintype=mtype,
                subtype=stype,
                filename=filename,
                disposition="attachment",
            )

    # -------------------- 5. Bytes erzeugen --------------------
    from io import BytesIO
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