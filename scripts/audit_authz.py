#!/usr/bin/env python3
"""T-01 / FASE 2.1 — inventaris endpoint TULIS tanpa gerbang peran (AST atas backend/routes/**).

    python scripts/audit_authz.py                 # ringkasan + CSV ke memory/AUDIT_AUTHZ.csv
    python scripts/audit_authz.py --gate          # exit 1 bila jumlah tanpa gerbang > baseline
    python scripts/audit_authz.py --set-baseline  # simpan angka sekarang sebagai baseline

Definisi "gerbang" (heuristik, sengaja longgar supaya tidak ada positif palsu yang memaksa refactor):
  - di badan fungsi: menyebut role/permission (`role`, `_permissions`, `assert_can_act`, `check_role`,
    `require_roles`, `deny_`, `_require_fin*`, `APPROVER_ROLES`, `ADMIN_ROLES`, `FINANCE_ROLES`, `has_perm`,
    `require_perm`, `require_write_actor`, `is_vendor`, `vendor_identity`, `scope`, `assert_*`, `_require_*`)
  - atau Depends(...) pada parameter selain require_auth/get_db/_paginate_params
  - atau router berkas punya `dependencies=[...]` (gerbang level router di berkas)
  - atau router dipasang di server.py dengan `dependencies=` (gerbang level server, FASE 2.2)
"""
import ast
import csv
import os
import re
import sys

ROOT = os.path.join(os.path.dirname(__file__), "..")
ROUTES = os.path.join(ROOT, "backend", "routes")
SERVER = os.path.join(ROOT, "backend", "server.py")
CSV_OUT = os.path.join(ROOT, "memory", "AUDIT_AUTHZ.csv")
BASELINE = os.path.join(ROOT, "scripts", "authz_baseline.txt")
WRITE = {"post", "put", "patch", "delete"}
GATE_RX = re.compile(
    r"\brole\b|_permissions|assert_can_act|check_role|require_roles|deny_|_require_fin|_require_finance"
    r"|APPROVER_ROLES|ADMIN_ROLES|FINANCE_ROLES|has_perm|require_perm|require_write_actor|is_vendor"
    r"|vendor_identity|\bscope|assert_\w+\(|_require_\w+\(|require_admin|require_role|ensure_role"
    r"|PROD_ADMIN_ROLES|PROD_VENDOR_ROLES|forbid_|only_roles|allowed_roles|\.get\(['\"]role['\"]\)"
    # kepemilikan: dokumen difilter user/karyawan yang login (self-service)
    r"|user_id['\"]\s*:\s*user(\[|\.get)|_get_linked_employee|employee_id['\"]\s*:\s*emp\[")
PLAIN_DEPS = {"require_auth", "get_db", "_paginate_params", "_sort_params", "get_current_user"}


def _server_gated_modules() -> set:
    src = open(SERVER).read()
    mod_of = {}
    for m, v in re.findall(r"from routes\.([\w\.]+) import (?:\w+ as )?(\w+_router)\b", src):
        mod_of[v] = m
    for m, v in re.findall(r"from routes\.([\w\.]+) import \(\s*(?:[\w, \n]*?)router as (\w+)", src, re.S):
        mod_of.setdefault(v, m)
    gated = set()
    for v in re.findall(r"app\.include_router\((\w+),\s*dependencies=", src):
        if v in mod_of:
            gated.add(mod_of[v].replace(".", "/"))
    return gated


def _dep_names(fn: ast.AsyncFunctionDef | ast.FunctionDef) -> list:
    names = []
    for d in fn.args.defaults + fn.args.kw_defaults:
        if isinstance(d, ast.Call) and getattr(d.func, "id", "") == "Depends" and d.args:
            a = d.args[0]
            names.append(getattr(a, "id", None) or getattr(a, "attr", None) or ast.unparse(a))
    return names


def _route_of(fn) -> tuple[str, str] | None:
    for dec in fn.decorator_list:
        if isinstance(dec, ast.Call) and isinstance(dec.func, ast.Attribute):
            method = dec.func.attr.lower()
            if method in WRITE | {"get", "api_route", "websocket"} and dec.args:
                path = dec.args[0].value if isinstance(dec.args[0], ast.Constant) else "?"
                return method, path
    return None


def scan() -> list[dict]:
    server_gated = _server_gated_modules()
    rows = []
    for dirpath, _, files in os.walk(ROUTES):
        for f in sorted(files):
            if not f.endswith(".py"):
                continue
            p = os.path.join(dirpath, f)
            rel = os.path.relpath(p, os.path.join(ROOT, "backend", "routes"))[:-3]
            try:
                tree = ast.parse(open(p).read())
            except SyntaxError:
                continue
            src_all = open(p).read()
            file_router_gated = bool(re.search(r"APIRouter\([^)]*dependencies\s*=", src_all, re.S))
            for node in ast.walk(tree):
                if not isinstance(node, (ast.AsyncFunctionDef, ast.FunctionDef)):
                    continue
                r = _route_of(node)
                if not r or r[0] not in WRITE:
                    continue
                body = ast.unparse(node)
                deps = _dep_names(node)
                fn_gate = bool(GATE_RX.search(body)) or any(d not in PLAIN_DEPS for d in deps)
                rows.append({
                    "router": rel, "method": r[0].upper(), "path": r[1], "function": node.name,
                    "line": node.lineno, "fn_gate": int(fn_gate), "file_router_gate": int(file_router_gated),
                    "server_gate": int(rel in server_gated),
                    "no_gate": int(not fn_gate and not file_router_gated and not (rel in server_gated)),
                })
    return rows


def main():
    rows = scan()
    total = len(rows)
    no_gate = [r for r in rows if r["no_gate"]]
    no_fn_gate = [r for r in rows if not r["fn_gate"]]
    deletes = [r for r in no_gate if r["method"] == "DELETE"]
    os.makedirs(os.path.dirname(CSV_OUT), exist_ok=True)
    with open(CSV_OUT, "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(sorted(rows, key=lambda r: (-r["no_gate"], r["router"], r["line"])))
    per_router: dict = {}
    for r in no_gate:
        per_router[r["router"]] = per_router.get(r["router"], 0) + 1
    print(f"endpoint tulis: {total} · tanpa gerbang fungsi: {len(no_fn_gate)} · "
          f"TANPA GERBANG SAMA SEKALI (fungsi/berkas/server): {len(no_gate)} · DELETE tanpa gerbang: {len(deletes)}")
    print("router terbanyak tanpa gerbang:", sorted(per_router.items(), key=lambda kv: -kv[1])[:10])
    print("CSV:", CSV_OUT)
    if "--set-baseline" in sys.argv:
        open(BASELINE, "w").write(f"{len(no_gate)}\n")
        print("baseline disimpan:", len(no_gate))
    if "--gate" in sys.argv:
        base = int(open(BASELINE).read().strip()) if os.path.exists(BASELINE) else len(no_gate)
        if len(no_gate) > base:
            print(f"GAGAL: endpoint tulis tanpa gerbang NAIK {base} → {len(no_gate)}")
            return 1
        print(f"OK: {len(no_gate)} ≤ baseline {base}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
