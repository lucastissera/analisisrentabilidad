"""Descarga el IPC oficial (nivel general nacional, base dic 2016) y genera data/ipc_argentina.json."""
import csv
import io
import json
import os
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "data", "ipc_argentina.json")

CSV_URL = (
    "https://infra.datos.gob.ar/catalog/sspm/dataset/145/distribution/145.3/download/"
    "indice-precios-al-consumidor-nivel-general-base-diciembre-2016-mensual.csv"
)


def main():
    req = urllib.request.Request(CSV_URL, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=120) as resp:
        raw = resp.read().decode("utf-8")
    reader = csv.DictReader(io.StringIO(raw))
    indices = {}
    for row in reader:
        fecha = (row.get("indice_tiempo") or "").strip()[:10]
        if len(fecha) < 10:
            continue
        y_m = f"{fecha[:4]}-{fecha[5:7]}"
        val = row.get("ipc_ng_nacional")
        if val is None or str(val).strip() == "":
            continue
        try:
            indices[y_m] = round(float(str(val).replace(",", ".")), 4)
        except ValueError:
            continue
    meta = {
        "fuente": "INDEC vía SSPM / Datos Argentina — Serie 145.3, ipc_ng_nacional (IPC nivel general nacional, base dic 2016 = 100).",
        "csv_oficial": CSV_URL,
        "comando_actualizar": "python scripts/gen_ipc_json.py",
    }
    obj = {"meta": meta, "indices": indices}
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    keys = sorted(indices.keys())
    print("Wrote", OUT, "meses:", len(indices), "desde", keys[0], "hasta", keys[-1])


if __name__ == "__main__":
    main()
