#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
eurlex_pdf.py - Telecharge un acte EUR-Lex en PDF dans la langue voulue,
sans se faire bloquer par le WAF.

Pourquoi un detour ?
    eur-lex.europa.eu est derriere un AWS WAF : toutes les URL
    /legal-content/.../PDF/ repondent
        202 Accepted
        x-amzn-waf-action: challenge
    Un simple curl / requests ne peut pas franchir ce challenge (il faudrait
    executer du JavaScript). Idem pour /eli/... et pour data.europa.eu, qui
    redirige vers eur-lex.

    L'infrastructure CELLAR / OP Portal n'est pas filtree :
        1. on resout l'UUID CELLAR depuis le CELEX / OJ (HTTP 303 -> Location)
        2. on telecharge via le handler officiel de l'Office des publications

Usage :
    python eurlex_pdf.py 32024R2982
    python eurlex_pdf.py https://eur-lex.europa.eu/eli/reg_impl/2024/2982/oj
    python eurlex_pdf.py 32024R2982 32024R2979 32016R0679 -l fr -o pdf
    python eurlex_pdf.py --file liste.txt -l de

Aucune dependance externe (stdlib uniquement).
"""

import argparse
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
)

RESOURCE = "https://publications.europa.eu/resource"
HANDLER = "https://op.europa.eu/o/opportal-service/download-handler"

# type ELI -> lettre CELEX
ELI_LETTER = {
    "reg": "R", "reg_impl": "R", "reg_del": "R", "reg_fin": "R",
    "dir": "L", "dir_impl": "L", "dir_del": "L",
    "dec": "D", "dec_impl": "D", "dec_del": "D",
    "rec": "H",
}

UUID_RE = re.compile(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}")
# 32024R2982 (acte) ou 02024R2982-20260811 (version consolidee)
CELEX_RE = re.compile(r"[03]\d{4}[A-Z]\d{4}(?:-\d{8})?")


class NoRedirect(urllib.request.HTTPRedirectHandler):
    """Empeche urllib de suivre les 3xx : on veut lire le header Location."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


_opener = urllib.request.build_opener(NoRedirect)


def _get(url, accept="*/*", timeout=60):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": accept})
    return urllib.request.urlopen(req, timeout=timeout)


def normalize(raw):
    """Renvoie ('celex'|'oj'|'uuid', valeur)."""
    s = raw.strip()

    if UUID_RE.fullmatch(s):
        return "uuid", s

    # URL eur-lex avec ?uri=CELEX:... / ?uri=OJ:L_...
    if "uri=" in s:
        m = re.search(r"[?&]uri=([^&]+)", s)
        if m:
            s = urllib.parse.unquote(m.group(1))

    if CELEX_RE.fullmatch(s):
        return "celex", s

    m = re.fullmatch(r"OJ:?(L|C)_(\d{4})(\d{6})", s, re.I)
    if m:
        return "oj", "%s_%s%s" % (m.group(1).upper(), m.group(2), m.group(3))

    # ELI : .../eli/reg_impl/2024/2982/oj
    #       .../eli/reg_impl/2024/2982/2026-08-11   (version consolidee)
    m = re.search(r"/?eli/([a-z_]+)/(\d{4})/(\d+)(?:/(\d{4}-\d{2}-\d{2}))?(?:/|$)", s)
    if m:
        kind, year, num, date = m.group(1), m.group(2), m.group(3), m.group(4)
        letter = ELI_LETTER.get(kind)
        if not letter:
            raise ValueError("type ELI inconnu : %s" % kind)
        celex = "%s%s%s%04d" % ("0" if date else "3", year, letter, int(num))
        if date:
            celex += "-" + date.replace("-", "")
        return "celex", celex

    raise ValueError("identifiant non reconnu : %r" % raw)


def resource_uri(kind, value):
    if kind == "uuid":
        return None
    return "%s/%s/%s" % (RESOURCE, kind, value)


def resolve_uuid(uri):
    """Suit le 303 de publications.europa.eu et extrait l'UUID CELLAR."""
    try:
        with _opener.open(
            urllib.request.Request(uri, headers={"User-Agent": UA}),
            timeout=60,
        ) as r:
            loc = r.headers.get("Location") or r.geturl()
    except urllib.error.HTTPError as e:
        if e.code in (301, 302, 303, 307, 308):
            loc = e.headers.get("Location", "")
        else:
            raise RuntimeError("resolution %s -> HTTP %s" % (uri, e.code))

    m = re.search(r"cellar/([0-9a-f-]{36})", loc or "")
    if not m:
        raise RuntimeError("UUID introuvable dans Location: %r" % loc)
    return m.group(1)


def download(uuid, lang, outpath, attempts=4):
    url = HANDLER + "?" + urllib.parse.urlencode(
        {
            "identifier": uuid,
            "format": "PDF",
            "language": lang,
            "productionSystem": "cellar",
        }
    )
    last = None
    for i in range(attempts):
        try:
            req = urllib.request.Request(
                url,
                headers={
                    "User-Agent": UA,
                    "Accept": "application/pdf",
                    "Accept-Language": lang,
                },
            )
            with urllib.request.urlopen(req, timeout=180) as r:
                data = r.read()
                ctype = r.headers.get("Content-Type", "")
            if not data.startswith(b"%PDF"):
                raise RuntimeError("reponse non-PDF (Content-Type: %s)" % ctype)
            with open(outpath, "wb") as f:
                f.write(data)
            return len(data), ctype
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(2 * (i + 1))
    raise RuntimeError("echec telechargement : %s" % last)


def main(argv=None):
    p = argparse.ArgumentParser(
        description="Telecharge des actes EUR-Lex en PDF (contourne le WAF)."
    )
    p.add_argument("ids", nargs="*", help="CELEX, ELI, URL eur-lex ou UUID CELLAR")
    p.add_argument("-l", "--lang", default="fr", help="langue (fr, en, de, ...)")
    p.add_argument("-o", "--outdir", default=".", help="dossier de sortie")
    p.add_argument("--file", help="fichier texte : un identifiant par ligne")
    p.add_argument("-f", "--force", action="store_true", help="ecraser l'existant")
    a = p.parse_args(argv)

    ids = list(a.ids)
    if a.file:
        raw = open(a.file, "rb").read()
        for enc in ("utf-8-sig", "utf-16", "latin-1"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        ids += [
            ln.strip()
            for ln in text.splitlines()
            if ln.strip() and not ln.lstrip().startswith("#")
        ]

    if not ids:
        p.error("aucun identifiant fourni")

    os.makedirs(a.outdir, exist_ok=True)
    errors = 0

    for raw in ids:
        try:
            kind, val = normalize(raw)
            uuid = val if kind == "uuid" else resolve_uuid(resource_uri(kind, val))
            name = "%s_%s.pdf" % (val if kind != "uuid" else uuid, a.lang)
            path = os.path.join(a.outdir, name)

            if os.path.exists(path) and not a.force:
                print("[=] %-22s deja present -> %s" % (val, path))
                continue

            size, ctype = download(uuid, a.lang, path)
            print("[OK] %-22s %7d octets  %s" % (val, size, path))
        except Exception as e:  # noqa: BLE001
            errors += 1
            print("[!!] %-22s %s" % (raw, e))

    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
