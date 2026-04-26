"""
Sistema de Análisis de Rentabilidad
Backend Flask — app.py
"""

from flask import Flask, request, jsonify, render_template, send_file, session
import json, os, io
from datetime import datetime, timedelta
from excel_engine import ExcelExporter, ExcelImporter, ExcelTemplateGenerator
from pdf_engine import build_pdf_report
from auth_users import init_db, create_user, verify_user, ensure_bootstrap_user

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "cambiar-SECRET_KEY-en-produccion")
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=30)
init_db()
ensure_bootstrap_user()

# ─── In-memory store por usuario (swap for DB + persistencia en disco cuando haga falta) ─
_store = {}  # user_id (int) → config + monthly data


def get_state():
    uid = session.get("user_id")
    if not uid:
        return None
    if uid not in _store:
        _store[uid] = {"config": None, "monthly_data": {}}
    return _store[uid]


@app.before_request
def _require_login():
    if not request.path.startswith("/api"):
        return
    if request.path in ("/api/auth/login", "/api/auth/me") or request.path == "/api/auth/register":
        return
    if session.get("user_id"):
        return
    return jsonify({"error": "unauthorized", "message": "Iniciá sesión para continuar."}), 401

# ─── Routes ───────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/auth/login", methods=["POST"])
def auth_login():
    data = request.json or {}
    u = (data.get("username") or "").strip()
    p = data.get("password") or ""
    uid = verify_user(u, p)
    if not uid:
        return jsonify({"ok": False, "message": "Usuario/contraseña incorrectos"}), 401
    session["user_id"] = uid
    session["username"] = u
    session.permanent = True
    return jsonify({"ok": True, "username": u})


@app.route("/api/auth/logout", methods=["POST"])
def auth_logout():
    session.clear()
    return jsonify({"ok": True})


@app.route("/api/auth/me", methods=["GET"])
def auth_me():
    if not session.get("user_id"):
        return jsonify({"loggedIn": False})
    return jsonify({"loggedIn": True, "username": session.get("username", "")})


@app.route("/api/auth/register", methods=["POST"])
def auth_register():
    """Alta vía API (p. ej. admin); requiere REGISTER_SECRET. El alta por defecto: create_user.py o WhatsApp."""
    secret = os.environ.get("REGISTER_SECRET", "")
    if not secret or request.headers.get("X-Register-Secret", "") != secret:
        return jsonify({"ok": False, "message": "No autorizado."}), 403
    data = request.json or {}
    u = (data.get("username") or "").strip()
    p = data.get("password") or ""
    ok, err = create_user(u, p)
    if not ok:
        return jsonify({"ok": False, "message": err}), 400
    return jsonify({"ok": True, "username": u})

@app.route("/api/config", methods=["POST"])
def save_config():
    state = get_state()
    state["config"] = request.json
    return jsonify({"ok": True})

@app.route("/api/config", methods=["GET"])
def load_config():
    state = get_state()
    return jsonify(state.get("config") or {})

@app.route("/api/data", methods=["GET"])
def get_data():
    state = get_state()
    return jsonify(state.get("monthly_data", {}))

@app.route("/api/data/month", methods=["POST"])
def save_month():
    """Save or update one month of branch data.
    Body: { period: "2024-01", branches: [{...}, ...] }
    """
    state = get_state()
    body = request.json
    period = body.get("period")          # e.g. "2024-01" or "2024" for annual
    if not period:
        return jsonify({"error": "period required"}), 400
    state["monthly_data"][period] = body.get("branches", [])
    return jsonify({"ok": True, "periods": sorted(state["monthly_data"].keys())})

@app.route("/api/data/month/<period>", methods=["DELETE"])
def delete_month(period):
    state = get_state()
    state["monthly_data"].pop(period, None)
    return jsonify({"ok": True})

def _compute_all_metrics():
    """Métricas por período + acumulado; misma lógica que /api/metrics."""
    state = get_state()
    cfg = state.get("config", {}) or {}
    monthly = state.get("monthly_data", {})
    corp_costs = cfg.get("corpCosts", {})
    alloc_method = cfg.get("allocMethod", "ventas")
    manual_pct = cfg.get("manualPct", [])

    result = {}
    for period, branches in monthly.items():
        result[period] = _compute_period(branches, corp_costs, alloc_method, manual_pct)

    if len(monthly) > 1:
        result["__cumulative__"] = _compute_cumulative(
            monthly, corp_costs, alloc_method, manual_pct
        )
    return cfg, monthly, result


@app.route("/api/metrics", methods=["GET"])
def get_metrics():
    """Return computed metrics for all stored periods."""
    _, _, result = _compute_all_metrics()
    return jsonify(result)

@app.route("/api/export/excel", methods=["GET"])
def export_excel():
    state = get_state()
    cfg = state.get("config", {}) or {}
    monthly = state.get("monthly_data", {})
    exporter = ExcelExporter(cfg, monthly)
    buf = exporter.build()
    empresa = cfg.get("empresa", "Empresa")
    fname = f"Rentabilidad_{empresa.replace(' ','_')}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(buf, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True)

@app.route("/api/export/pdf", methods=["GET"])
def export_pdf():
    state = get_state()
    cfg, monthly, metrics = _compute_all_metrics()
    buf = build_pdf_report(cfg, monthly, metrics)
    empresa = cfg.get("empresa", "Empresa")
    fname = f"Rentabilidad_{empresa.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.pdf"
    return send_file(
        buf,
        download_name=fname,
        mimetype="application/pdf",
        as_attachment=True,
    )

@app.route("/api/template/excel", methods=["GET"])
def download_template():
    """Download a blank import template."""
    state = get_state()
    cfg = state.get("config", {}) or {}
    gen = ExcelTemplateGenerator(cfg)
    buf = gen.build()
    return send_file(buf, download_name="plantilla_importacion.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True)

@app.route("/api/import/excel", methods=["POST"])
def import_excel():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    try:
        importer = ExcelImporter(f.stream)
        result = importer.parse()
        state = get_state()
        # Merge imported periods into existing store
        for period, branches in result["monthly_data"].items():
            state["monthly_data"][period] = branches
        if result.get("config"):
            if not state.get("config"):
                state["config"] = result["config"]
        return jsonify({
            "ok": True,
            "imported_periods": sorted(result["monthly_data"].keys()),
            "config": result.get("config"),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 400

# ─── Calc engine ──────────────────────────────────────────────────────────────
def _alloc_weights(branches, method, manual_pct):
    n = len(branches)
    if method == "ventas":
        w = [b.get("ventasBrutas", 0) - b.get("devoluciones", 0) + b.get("otrosIngresos", 0) for b in branches]
    elif method == "empleados":
        w = [b.get("empleados", 1) for b in branches]
    elif method == "m2":
        w = [b.get("m2", 1) for b in branches]
    elif method == "transacciones":
        w = [b.get("transacciones", 1) for b in branches]
    elif method == "igualitario":
        w = [1] * n
    elif method == "manual":
        total = sum(manual_pct[:n]) or 1
        return [p / total for p in manual_pct[:n]]
    else:
        w = [1] * n
    total = sum(w) or 1
    return [x / total for x in w]

def _calc_branch(b, corp_share):
    vn = b.get("ventasBrutas", 0) - b.get("devoluciones", 0) + b.get("otrosIngresos", 0)
    cmv = b.get("cmv", 0) + b.get("moDirecta", 0) + b.get("materiales", 0)
    mb = vn - cmv
    go = sum(b.get(k, 0) for k in ["alquiler", "sueldos", "servicios", "mantenimiento", "marketing", "admGeneral"])
    ebitda_d = mb - go
    ebitda_n = ebitda_d - corp_share
    activos = b.get("inventario", 0) + b.get("equipamiento", 0) + b.get("mobiliario", 0)
    cf = sum(b.get(k, 0) for k in ["alquiler", "sueldos", "servicios", "mantenimiento", "admGeneral"])
    cv_pct = (cmv + b.get("marketing", 0)) / vn if vn else 0
    pe = cf / (1 - cv_pct) if (1 - cv_pct) > 0 else 0
    return {
        "vn": vn, "cmv": cmv, "mb": mb, "pctMb": mb / vn if vn else 0,
        "go": go, "ebitdaDirecto": ebitda_d, "pctEbitdaDirecto": ebitda_d / vn if vn else 0,
        "corpShare": corp_share, "ebitdaNeto": ebitda_n, "pctEbitdaNeto": ebitda_n / vn if vn else 0,
        "activos": activos, "roa": ebitda_n / activos if activos else 0,
        "pe": pe,
        "ventaXemp": vn / b["empleados"] if b.get("empleados") else 0,
        "ventaXm2": vn / b["m2"] if b.get("m2") else 0,
        "ticket": vn / b["transacciones"] if b.get("transacciones") else 0,
    }

def _compute_period(branches, corp_costs, alloc_method, manual_pct):
    corp_total = sum(corp_costs.values())
    weights = _alloc_weights(branches, alloc_method, manual_pct)
    results = [_calc_branch(b, corp_total * weights[i]) for i, b in enumerate(branches)]
    vnt = sum(r["vn"] for r in results)
    return {
        "branches": results,
        "totals": {
            "vn": vnt,
            "mb": sum(r["mb"] for r in results),
            "pctMb": sum(r["mb"] for r in results) / vnt if vnt else 0,
            "ebitdaDirecto": sum(r["ebitdaDirecto"] for r in results),
            "ebitdaNeto": sum(r["ebitdaNeto"] for r in results),
            "pctEbitdaNeto": sum(r["ebitdaNeto"] for r in results) / vnt if vnt else 0,
            "activos": sum(r["activos"] for r in results),
            "corpTotal": corp_total,
        }
    }

def _compute_cumulative(monthly, corp_costs, alloc_method, manual_pct):
    """Sum all numeric fields across periods, then recompute metrics."""
    from collections import defaultdict
    n = None
    acc = None
    for period in sorted(monthly.keys()):
        branches = monthly[period]
        if n is None:
            n = len(branches)
            acc = [{} for _ in range(n)]
        for i, b in enumerate(branches):
            for k, v in b.items():
                if isinstance(v, (int, float)):
                    acc[i][k] = acc[i].get(k, 0) + v
                else:
                    acc[i][k] = v   # strings: keep last
    return _compute_period(acc, corp_costs, alloc_method, manual_pct)

if __name__ == "__main__":
    app.run(debug=True, port=5050)
