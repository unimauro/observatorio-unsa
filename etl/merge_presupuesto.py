#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fusiona la serie histórica 2012-2024 (ca_unsa_raw.json) en data/presupuesto-unsa.json
SIN tocar 2025/2026 ni detalle_ultimo_anio. Ordena ascendente, recalcula ejec_pct."""
import json, os, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.join(HERE, "ca_unsa_raw.json")
TARGET = os.path.join(HERE, "..", "data", "presupuesto-unsa.json")

raw = json.load(open(RAW, encoding="utf-8"))
doc = json.load(open(TARGET, encoding="utf-8"))

serie = doc.get("serie", [])
existing = {r["year"] for r in serie}
cur_year = datetime.date.today().year

for y in sorted(int(k) for k in raw):
    if y in existing:
        continue  # nunca sobrescribir años ya presentes (2025/2026)
    r = raw[str(y)]
    pim = r["pim"]; dev = r["dev"]
    rec = {
        "year": y,
        "pia": r["pia"],
        "pim": pim,
        "cert": r["cert"],
        "dev": dev,
        "gir": r["gir"],
        "ejec_pct": round(100 * dev / pim, 1) if pim else 0,
    }
    if y == cur_year and rec["ejec_pct"] < 70:
        rec["parcial"] = True
    serie.append(rec)

serie.sort(key=lambda x: x["year"])
doc["serie"] = serie
doc.setdefault("_meta", {})["nota"] = (
    "Serie histórica completa 2012-2026 (pliego 513) vía Consulta Amigable MEF. "
    "2025 cerrado; 2026 en ejecución (parcial)."
)

json.dump(doc, open(TARGET, "w", encoding="utf-8"),
          ensure_ascii=False, separators=(",", ":"))

print("Serie final:", [(r["year"], round(r["pim"]/1e6, 1)) for r in serie])
print("2025 PIM =", next(r["pim"] for r in serie if r["year"] == 2025))
