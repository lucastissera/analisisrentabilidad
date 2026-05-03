"""
Coeficientes IPC (Argentina) para restatement a moneda de cierre de ejercicio.
Serie en data/ipc_argentina.json (claves YYYY-MM, nivel acumulado coherente).
"""

from __future__ import annotations

import csv
import io
import json
import os
import time
import urllib.error
import urllib.request
from copy import deepcopy
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple

_BASE = os.path.dirname(os.path.abspath(__file__))
IPC_PATH = os.path.join(_BASE, "data", "ipc_argentina.json")

# Serie 145.3 — IPC nivel general nacional, base dic 2016 = 100 (INDEC / Datos Argentina)
OFFICIAL_IPC_CSV_URL = (
    "https://infra.datos.gob.ar/catalog/sspm/dataset/145/distribution/145.3/download/"
    "indice-precios-al-consumidor-nivel-general-base-diciembre-2016-mensual.csv"
)

# Magnitudes monetarias cargadas por sucursal; no se deflactan empleados, m² ni transacciones
MONETARY_BRANCH_KEYS = frozenset(
    {
        "ventasBrutas",
        "devoluciones",
        "otrosIngresos",
        "cmv",
        "moDirecta",
        "materiales",
        "alquiler",
        "sueldos",
        "servicios",
        "mantenimiento",
        "marketing",
        "admGeneral",
        "inventario",
        "equipamiento",
        "mobiliario",
    }
)


def _load_ipc_indices() -> Dict[str, float]:
    if not os.path.isfile(IPC_PATH):
        return {}
    try:
        with open(IPC_PATH, encoding="utf-8") as f:
            raw = json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}
    if isinstance(raw, dict) and "indices" in raw:
        return {str(k): float(v) for k, v in raw["indices"].items()}
    if isinstance(raw, dict):
        return {str(k): float(v) for k, v in raw.items() if isinstance(v, (int, float))}
    return {}


_IPC_CACHE: Optional[Dict[str, float]] = None
_IPC_MTIME: Optional[float] = None


def load_ipc_series() -> Dict[str, float]:
    """Recarga si cambió el archivo en disco (p. ej. tras gen_ipc_json.py)."""
    global _IPC_CACHE, _IPC_MTIME
    try:
        mtime = os.path.getmtime(IPC_PATH)
    except OSError:
        mtime = None
    if _IPC_CACHE is None or _IPC_MTIME != mtime:
        _IPC_CACHE = _load_ipc_indices()
        _IPC_MTIME = mtime
    return _IPC_CACHE


def _invalidate_ipc_cache() -> None:
    global _IPC_CACHE, _IPC_MTIME
    _IPC_CACHE = None
    _IPC_MTIME = None


def _parse_official_ipc_csv(text: str) -> Dict[str, float]:
    out: Dict[str, float] = {}
    reader = csv.DictReader(io.StringIO(text))
    for row in reader:
        fecha = (row.get("indice_tiempo") or "").strip()[:10]
        if len(fecha) < 10:
            continue
        y_m = f"{fecha[:4]}-{fecha[5:7]}"
        val = row.get("ipc_ng_nacional")
        if val is None or str(val).strip() == "":
            continue
        try:
            out[y_m] = round(float(str(val).replace(",", ".")), 4)
        except ValueError:
            continue
    return out


def sync_ipc_from_official(
    force: bool = False,
    min_interval_seconds: int = 3600,
    timeout: int = 25,
) -> Dict[str, Any]:
    """
    Descarga la serie oficial 145.3 y fusiona en data/ipc_argentina.json.
    - Cada YYYY-MM publicado en el CSV reemplaza el valor en disco (fuente INDEC).
    - Los meses que solo existen en el archivo (carga manual) se conservan.
    """
    now = time.time()
    existing: Dict[str, float] = {}
    meta: Dict[str, Any] = {}
    if os.path.isfile(IPC_PATH):
        try:
            with open(IPC_PATH, encoding="utf-8") as f:
                blob = json.load(f)
            if isinstance(blob, dict):
                raw_i = blob.get("indices")
                if isinstance(raw_i, dict):
                    existing = {str(k): float(v) for k, v in raw_i.items()}
                raw_m = blob.get("meta")
                if isinstance(raw_m, dict):
                    meta = dict(raw_m)
        except (json.JSONDecodeError, OSError, TypeError, ValueError):
            existing, meta = {}, {}

    last_epoch = meta.get("ultima_sincronizacion_official_epoch")
    if not force and last_epoch is not None:
        try:
            if now - float(last_epoch) < float(min_interval_seconds):
                return {
                    "ok": True,
                    "skipped": True,
                    "reason": "interval",
                    "min_interval_seconds": int(min_interval_seconds),
                    "ultimo_mes_archivo": max(existing.keys()) if existing else None,
                }
        except (TypeError, ValueError):
            pass

    req = urllib.request.Request(
        OFFICIAL_IPC_CSV_URL,
        headers={"User-Agent": "AnalisisRentabilidad/1.0 (IPC sync)"},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            text = resp.read().decode("utf-8")
    except (urllib.error.URLError, OSError, TimeoutError) as e:
        return {
            "ok": False,
            "skipped": False,
            "error": str(e),
            "ultimo_mes_archivo": max(existing.keys()) if existing else None,
        }

    official = _parse_official_ipc_csv(text)
    if not official:
        return {
            "ok": False,
            "skipped": False,
            "error": "CSV oficial sin valores ipc_ng_nacional",
            "ultimo_mes_archivo": max(existing.keys()) if existing else None,
        }

    merged = dict(existing)
    for k, v in official.items():
        merged[k] = v

    meta["ultima_sincronizacion_official_iso"] = datetime.now(timezone.utc).strftime(
        "%Y-%m-%dT%H:%M:%SZ"
    )
    meta["ultima_sincronizacion_official_epoch"] = now
    meta["ultima_sincronizacion_official_ok"] = True
    meta.pop("ultima_sincronizacion_official_error", None)
    meta["csv_oficial_url"] = OFFICIAL_IPC_CSV_URL

    out_obj = {"meta": meta, "indices": merged}
    os.makedirs(os.path.dirname(IPC_PATH), exist_ok=True)
    tmp_path = IPC_PATH + ".tmp"
    try:
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(out_obj, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, IPC_PATH)
    except OSError as e:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        return {
            "ok": False,
            "skipped": False,
            "error": str(e),
            "ultimo_mes_archivo": max(existing.keys()) if existing else None,
        }

    _invalidate_ipc_cache()
    return {
        "ok": True,
        "skipped": False,
        "meses_en_csv": len(official),
        "meses_totales_archivo": len(merged),
        "ultimo_mes_oficial": max(official.keys()),
        "ultimo_mes_archivo": max(merged.keys()) if merged else None,
    }


def ipc_series_meta() -> Dict[str, Any]:
    """Último mes en archivo y metadatos."""
    series = load_ipc_series()
    keys = sorted(series.keys())
    last = keys[-1] if keys else None
    meta = {}
    if os.path.isfile(IPC_PATH):
        try:
            with open(IPC_PATH, encoding="utf-8") as f:
                raw = json.load(f)
            if isinstance(raw, dict) and "meta" in raw:
                meta = raw["meta"]
        except (json.JSONDecodeError, OSError):
            pass
    return {"lastMonth": last, "meta": meta, "count": len(series)}


def resolve_index(series: Dict[str, float], year: int, month: int, max_back: int = 36) -> Optional[float]:
    """
    Índice del mes (year, month); si no está publicado, retrocede hasta encontrar
    el último mes anterior con dato (proxy 'último oficial disponible').
    """
    y, m = int(year), int(month)
    for _ in range(max_back):
        key = f"{y}-{m:02d}"
        if key in series:
            return float(series[key])
        m -= 1
        if m < 1:
            m = 12
            y -= 1
    return None


def fiscal_year_end_for_data_month(data_y: int, data_m: int, fy_end_month: int) -> Tuple[int, int]:
    """
    Mes de cierre del ejercicio fiscal que contiene (data_y, data_m).
    fy_end_month: 1..12 último mes del ejercicio.
    """
    fe = int(fy_end_month)
    if fe < 1 or fe > 12:
        fe = 12
    dy, dm = int(data_y), int(data_m)
    if dm <= fe:
        return dy, fe
    return dy + 1, fe


def parse_period_to_month(period_str: str, period_type: str, fy_end_month: int) -> Optional[Tuple[int, int]]:
    """Devuelve (año, mes) del dato para enlazar con la serie IPC."""
    fy = int(fy_end_month) if fy_end_month else 12
    if fy < 1 or fy > 12:
        fy = 12
    s = str(period_str).strip()
    if period_type == "annual" and len(s) == 4 and s.isdigit():
        return int(s), fy
    if len(s) >= 7 and s[4] == "-":
        parts = s.split("-")
        try:
            return int(parts[0]), int(parts[1])
        except (ValueError, IndexError):
            return None
    return None


def ipc_restatement_factor(period_str: str, cfg: Dict[str, Any]) -> Optional[float]:
    """
    Coeficiente: IPC(cierre de ejercicio resuelto) / IPC(mes del dato resuelto).
    Redondeado a 2 decimales (half-up en magnitudes típicas).
    No aplica a montos corporativos anuales (se dejan en la convención actual).
    """
    if not cfg.get("ipcAdjust"):
        return None
    series = load_ipc_series()
    if not series:
        return None
    fy_end = int(cfg.get("fiscalYearEndMonth", 12) or 12)
    ptype = cfg.get("periodType", "monthly")
    parsed = parse_period_to_month(period_str, ptype, fy_end)
    if not parsed:
        return None
    dy, dm = parsed
    cy, cm = fiscal_year_end_for_data_month(dy, dm, fy_end)
    idx_target = resolve_index(series, cy, cm)
    idx_data = resolve_index(series, dy, dm)
    if idx_target is None or idx_data is None or idx_data == 0:
        return None
    return round(idx_target / idx_data, 2)


def adjust_branches_for_ipc(branches: List[Dict[str, Any]], period_key: str, cfg: Dict[str, Any]) -> List[Dict[str, Any]]:
    factor = ipc_restatement_factor(period_key, cfg)
    if factor is None:
        return branches
    out = deepcopy(branches)
    for b in out:
        for k in MONETARY_BRANCH_KEYS:
            if k not in b:
                continue
            v = b[k]
            if isinstance(v, (int, float)):
                b[k] = v * factor
    return out


def adjust_monthly_store(monthly: Dict[str, List[dict]], cfg: Dict[str, Any]) -> Dict[str, List[dict]]:
    """Copia de monthly_data con sucursales ajustadas por IPC (no muta el original)."""
    return {p: adjust_branches_for_ipc(list(branches), p, cfg) for p, branches in monthly.items()}
