from dotenv import load_dotenv
import os
import requests
import csv
import json
import base64
from datetime import datetime
from typing import List, Dict, Optional

load_dotenv()

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────
GRAPH_BASE = "https://graph.microsoft.com/beta"


# ─────────────────────────────────────────────────────────────────────────────
# Auth
# ─────────────────────────────────────────────────────────────────────────────

def get_token(scope: str) -> str:
    tenant_id     = os.getenv("TENANT_ID")
    client_id     = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")
    if not all([tenant_id, client_id, client_secret]):
        raise ValueError("Missing TENANT_ID / CLIENT_ID / CLIENT_SECRET in .env")
    r = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={"client_id": client_id, "client_secret": client_secret,
              "scope": scope, "grant_type": "client_credentials"},
        timeout=60,
    )
    r.raise_for_status()
    return r.json()["access_token"]

def get_graph_token() -> str: return get_token("https://graph.microsoft.com/.default")


# ─────────────────────────────────────────────────────────────────────────────
# Microsoft Graph — users + manager
# ─────────────────────────────────────────────────────────────────────────────

def get_all_users(token: str) -> List[Dict]:
    url     = f"{GRAPH_BASE}/users"
    headers = {"Authorization": f"Bearer {token}", "ConsistencyLevel": "eventual"}
    params  = {
        "$select": (
            "id,displayName,userPrincipalName,mail,accountEnabled,"
            "userType,employeeType,customSecurityAttributes"
        ),
        "$expand": "manager($select=displayName)",
        "$top":    "999",
    }
    users: List[Dict] = []
    while url:
        r = requests.get(url, headers=headers, params=params, timeout=120)
        r.raise_for_status()
        data = r.json()
        users.extend(data.get("value", []))
        url    = data.get("@odata.nextLink")
        params = None
    return users

def get_account_type(user: Dict) -> str:
    csa = (user.get("customSecurityAttributes") or {})
    return (((csa.get("AccountClassification") or {}).get("AccountType")) or "").strip().lower()

def get_manager_display_name(user: Dict) -> Optional[str]:
    return (user.get("manager") or {}).get("displayName")

def is_active_employee_non_service(user: Dict, require_employeeType: bool = False) -> bool:
    if user.get("accountEnabled") is not True:                              return False
    if (user.get("userType") or "").strip().lower() != "member":           return False
    if require_employeeType:
        if (user.get("employeeType") or "").strip().lower() != "employee": return False
    if get_account_type(user) == "service":                                return False
    return True

def flatten_user_row(user: Dict) -> Dict:
    mgr = get_manager_display_name(user)
    return {
        "id":                 user.get("id"),
        "displayName":        user.get("displayName"),
        "userPrincipalName":  user.get("userPrincipalName"),
        "mail":               user.get("mail"),
        "accountEnabled":     user.get("accountEnabled"),
        "employeeType":       user.get("employeeType"),
        "managerDisplayName": mgr if mgr else "CEO",
    }


# ─────────────────────────────────────────────────────────────────────────────
# CSV export
# ─────────────────────────────────────────────────────────────────────────────

def export_csv(rows: List[Dict], filename: str = "active_employees.csv") -> str:
    fieldnames = ["id","displayName","userPrincipalName","mail",
                  "accountEnabled","employeeType","managerDisplayName"]
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)
    return filename


# ─────────────────────────────────────────────────────────────────────────────
# Dashboard HTML
# ─────────────────────────────────────────────────────────────────────────────

def export_morse_employee_dashboard(
    rows: List[Dict],
    filename: str = "employee_dashboard.html",
) -> str:
    """
    Export the Morse Employee Dashboard as a fully self-contained HTML file.

    Renders a polished report with:
      - Summary stat cards (total / employee / contractor / intern counts)
      - Donut chart — workforce type breakdown
      - Horizontal bar chart — top managers by headcount
      - Filterable, searchable, sortable employee table
      - Category dropdown filter (All / Employee / Contractor / Intern)
      - Download button that exports the currently selected category as CSV

    Args:
        rows:     Flat employee dicts, each with keys:
                  id, displayName, userPrincipalName, mail,
                  accountEnabled, employeeType, managerDisplayName
        filename: Output file path (default: employee_dashboard.html)

    Returns:
        The resolved path of the written file.
    """
    now           = datetime.now()
    friendly_date = now.strftime(f"%A, %B {now.day}, %Y")   # e.g. "Tuesday, May 13, 2026"

    # ── Counts by employee type ────────────────────────────────────────────────
    type_counts: Dict[str, int] = {}
    for row in rows:
        t = (row.get("employeeType") or "Unknown").strip()
        type_counts[t] = type_counts.get(t, 0) + 1

    total        = len(rows)
    n_employee   = type_counts.get("Employee", 0)
    n_contractor = type_counts.get("Contractor", 0)
    n_intern     = type_counts.get("Intern", 0)

    pct = lambda n: f"{round(n / total * 100)}%" if total else "0%"

    # ── Top managers ───────────────────────────────────────────────────────────
    mgr_counts: Dict[str, int] = {}
    for row in rows:
        mgr = (row.get("managerDisplayName") or "Unknown").strip()
        if mgr.upper() != "CEO":
            mgr_counts[mgr] = mgr_counts.get(mgr, 0) + 1
    top_mgrs = sorted(mgr_counts.items(), key=lambda x: x[1], reverse=True)[:10]

    # ── Serialise rows for embedding in JS ────────────────────────────────────
    js_data = json.dumps(rows, ensure_ascii=False)

    # ── Bar-chart data ─────────────────────────────────────────────────────────
    bar_labels = json.dumps([m for m, _ in top_mgrs])
    bar_values = json.dumps([c for _, c in top_mgrs])

    html = (
        "<!DOCTYPE html>\n"
        '<html lang="en">\n'
        "<head>\n"
        '<meta charset="UTF-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
        f"<title>Morse Data — Active Employee Dashboard</title>\n"
        '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.5.0/dist/chart.umd.js" '
        'integrity="sha384-iU8HYtnGQ8Cy4zl7gbNMOhsDTTKX02BTXptVP/vqAWIaTfM7isw76iyZCsjL2eVi" '
        'crossorigin="anonymous"></script>\n'
        "<style>\n"
        "  :root {\n"
        "    color-scheme: light;\n"
        "    --bg: #f4f6fb; --surface: #ffffff; --border: #e2e8f0;\n"
        "    --primary: #1e40af; --primary-light: #dbeafe; --accent: #0ea5e9;\n"
        "    --employee: #22c55e; --employee-light: #dcfce7;\n"
        "    --contractor: #f59e0b; --contractor-light: #fef3c7;\n"
        "    --intern: #8b5cf6; --intern-light: #ede9fe;\n"
        "    --text: #1e293b; --text-muted: #64748b;\n"
        "    --shadow: 0 1px 3px rgba(0,0,0,.08), 0 4px 16px rgba(0,0,0,.06);\n"
        "    --radius: 14px;\n"
        "  }\n"
        "  * { box-sizing: border-box; margin: 0; padding: 0; }\n"
        "  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;\n"
        "         background: var(--bg); color: var(--text); min-height: 100vh; padding: 0 0 40px; }\n"
        "  .header { background: linear-gradient(135deg,#1e3a8a 0%,#1e40af 50%,#1d4ed8 100%);\n"
        "            color:#fff; padding:28px 32px 24px; display:flex; align-items:center;\n"
        "            justify-content:space-between; flex-wrap:wrap; gap:12px; }\n"
        "  .header-left { display:flex; align-items:center; gap:14px; }\n"
        "  .logo-mark { width:46px; height:46px; background:rgba(255,255,255,.15);\n"
        "               border:2px solid rgba(255,255,255,.3); border-radius:12px;\n"
        "               display:flex; align-items:center; justify-content:center;\n"
        "               font-size:22px; font-weight:800; color:#fff; letter-spacing:-1px; }\n"
        "  .header-title { font-size:22px; font-weight:700; line-height:1.2; }\n"
        "  .header-sub { font-size:13px; color:rgba(255,255,255,.7); margin-top:2px; }\n"
        "  .header-badge { background:rgba(255,255,255,.15); border:1px solid rgba(255,255,255,.3);\n"
        "                  border-radius:20px; padding:6px 16px; font-size:13px;\n"
        "                  color:rgba(255,255,255,.9); font-weight:500; }\n"
        "  .page { max-width:1200px; margin:0 auto; padding:28px 24px 0; }\n"
        "  .stats-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:16px; margin-bottom:24px; }\n"
        "  @media(max-width:800px){ .stats-grid{ grid-template-columns:repeat(2,1fr); } }\n"
        "  .stat-card { background:var(--surface); border-radius:var(--radius); padding:20px 22px;\n"
        "               box-shadow:var(--shadow); border:1px solid var(--border);\n"
        "               display:flex; align-items:center; gap:14px;\n"
        "               transition:transform .15s,box-shadow .15s; cursor:pointer; }\n"
        "  .stat-card:hover { transform:translateY(-2px); box-shadow:0 6px 24px rgba(0,0,0,.1); }\n"
        "  .stat-card.active { border-width:2px; }\n"
        "  .stat-card.active-all { border-color:var(--primary); background:var(--primary-light); }\n"
        "  .stat-card.active-emp { border-color:var(--employee); background:var(--employee-light); }\n"
        "  .stat-card.active-con { border-color:var(--contractor); background:var(--contractor-light); }\n"
        "  .stat-card.active-int { border-color:var(--intern); background:var(--intern-light); }\n"
        "  .stat-icon { width:48px; height:48px; border-radius:12px;\n"
        "               display:flex; align-items:center; justify-content:center; font-size:22px; flex-shrink:0; }\n"
        "  .stat-icon.all        { background:var(--primary-light); }\n"
        "  .stat-icon.employee   { background:var(--employee-light); }\n"
        "  .stat-icon.contractor { background:var(--contractor-light); }\n"
        "  .stat-icon.intern     { background:var(--intern-light); }\n"
        "  .stat-info { flex:1; min-width:0; }\n"
        "  .stat-value { font-size:30px; font-weight:800; line-height:1; color:var(--text); }\n"
        "  .stat-label { font-size:12px; color:var(--text-muted); font-weight:500; margin-top:3px;\n"
        "                text-transform:uppercase; letter-spacing:.5px; }\n"
        "  .stat-pct { font-size:12px; font-weight:600; padding:2px 8px; border-radius:20px;\n"
        "              margin-top:6px; display:inline-block; }\n"
        "  .charts-row { display:grid; grid-template-columns:320px 1fr; gap:16px; margin-bottom:24px; }\n"
        "  @media(max-width:800px){ .charts-row{ grid-template-columns:1fr; } }\n"
        "  .chart-card { background:var(--surface); border-radius:var(--radius); padding:22px;\n"
        "                box-shadow:var(--shadow); border:1px solid var(--border); }\n"
        "  .chart-title { font-size:14px; font-weight:700; color:var(--text); margin-bottom:16px;\n"
        "                 display:flex; align-items:center; gap:8px; }\n"
        "  .donut-wrap, .bar-wrap { position:relative; height:220px; }\n"
        "  .table-card { background:var(--surface); border-radius:var(--radius);\n"
        "                box-shadow:var(--shadow); border:1px solid var(--border); overflow:hidden; }\n"
        "  .table-toolbar { padding:18px 22px; display:flex; align-items:center;\n"
        "                   justify-content:space-between; flex-wrap:wrap; gap:12px;\n"
        "                   border-bottom:1px solid var(--border); }\n"
        "  .toolbar-left { display:flex; align-items:center; gap:12px; flex-wrap:wrap; }\n"
        "  .search-box { position:relative; display:flex; align-items:center; }\n"
        "  .search-icon { position:absolute; left:10px; color:var(--text-muted); font-size:14px; pointer-events:none; }\n"
        "  .search-input { border:1px solid var(--border); border-radius:8px; padding:8px 12px 8px 32px;\n"
        "                  font-size:13px; color:var(--text); background:var(--bg); outline:none; width:220px;\n"
        "                  transition:border-color .15s,box-shadow .15s; }\n"
        "  .search-input:focus { border-color:var(--accent); box-shadow:0 0 0 3px rgba(14,165,233,.12); }\n"
        "  .filter-group { display:flex; gap:8px; align-items:center; }\n"
        "  .filter-label { font-size:12px; font-weight:600; color:var(--text-muted);\n"
        "                  text-transform:uppercase; letter-spacing:.4px; white-space:nowrap; }\n"
        "  .filter-select-wrap { position:relative; display:flex; align-items:center; }\n"
        "  .filter-select-wrap::after { content:'\\25BE'; position:absolute; right:10px;\n"
        "                               pointer-events:none; font-size:11px; color:var(--text-muted); }\n"
        "  .filter-select { appearance:none; -webkit-appearance:none; border:1px solid var(--border);\n"
        "                   border-radius:8px; padding:8px 28px 8px 12px; font-size:13px; font-weight:600;\n"
        "                   color:var(--text); background:#fff; cursor:pointer; outline:none;\n"
        "                   transition:border-color .15s,box-shadow .15s; min-width:150px; }\n"
        "  .filter-select:focus { border-color:var(--accent); box-shadow:0 0 0 3px rgba(14,165,233,.12); }\n"
        "  .filter-select.sel-all        { border-color:var(--primary); color:var(--primary); }\n"
        "  .filter-select.sel-employee   { border-color:var(--employee); color:#15803d; }\n"
        "  .filter-select.sel-contractor { border-color:var(--contractor); color:#b45309; }\n"
        "  .filter-select.sel-intern     { border-color:var(--intern); color:#6d28d9; }\n"
        "  .toolbar-right { display:flex; gap:8px; }\n"
        "  .dl-btn { display:flex; align-items:center; gap:6px; padding:8px 16px; border-radius:8px;\n"
        "            font-size:13px; font-weight:600; cursor:pointer; transition:all .15s; border:none;\n"
        "            background:var(--primary); color:#fff; }\n"
        "  .dl-btn:hover { background:#1d4ed8; }\n"
        "  .table-wrap { overflow-x:auto; }\n"
        "  table { width:100%; border-collapse:collapse; font-size:13px; }\n"
        "  thead th { padding:12px 16px; text-align:left; font-size:11px; font-weight:700;\n"
        "             color:var(--text-muted); text-transform:uppercase; letter-spacing:.5px;\n"
        "             background:var(--bg); border-bottom:1px solid var(--border);\n"
        "             white-space:nowrap; cursor:pointer; user-select:none; }\n"
        "  thead th:hover { color:var(--text); }\n"
        "  thead th .sort-icon { margin-left:4px; opacity:.4; }\n"
        "  thead th.sorted .sort-icon { opacity:1; color:var(--primary); }\n"
        "  tbody tr { border-bottom:1px solid var(--border); transition:background .1s; }\n"
        "  tbody tr:last-child { border-bottom:none; }\n"
        "  tbody tr:hover { background:#f8fafc; }\n"
        "  tbody td { padding:11px 16px; vertical-align:middle; }\n"
        "  .avatar { width:32px; height:32px; border-radius:50%; display:flex; align-items:center;\n"
        "            justify-content:center; font-size:12px; font-weight:700; color:#fff; flex-shrink:0; }\n"
        "  .name-cell { display:flex; align-items:center; gap:10px; }\n"
        "  .name-text { font-weight:600; color:var(--text); }\n"
        "  .email-text { font-size:11px; color:var(--text-muted); margin-top:1px; }\n"
        "  .badge { display:inline-flex; align-items:center; gap:4px;\n"
        "           padding:3px 10px; border-radius:20px; font-size:11px; font-weight:600; }\n"
        "  .badge-employee   { background:var(--employee-light);   color:#15803d; }\n"
        "  .badge-contractor { background:var(--contractor-light); color:#b45309; }\n"
        "  .badge-intern     { background:var(--intern-light);     color:#6d28d9; }\n"
        "  .badge-unknown    { background:#f1f5f9; color:#64748b; }\n"
        "  .badge-dot { width:6px; height:6px; border-radius:50%; }\n"
        "  .badge-employee .badge-dot   { background:var(--employee); }\n"
        "  .badge-contractor .badge-dot { background:var(--contractor); }\n"
        "  .badge-intern .badge-dot     { background:var(--intern); }\n"
        "  .badge-unknown .badge-dot    { background:#94a3b8; }\n"
        "  .enabled-dot { display:inline-block; width:8px; height:8px; border-radius:50%; background:var(--employee); }\n"
        "  .table-footer { padding:12px 22px; display:flex; align-items:center; justify-content:space-between;\n"
        "                  border-top:1px solid var(--border); font-size:12px; color:var(--text-muted); background:var(--bg); }\n"
        "  .no-results { text-align:center; padding:60px 20px; color:var(--text-muted); font-size:14px; }\n"
        "  .no-results .emoji { font-size:36px; display:block; margin-bottom:10px; }\n"
        "</style>\n"
        "</head>\n"
        "<body>\n"
        "\n"
        '<div class="header">\n'
        '  <div class="header-left">\n'
        '    <div class="logo-mark">M</div>\n'
        "    <div>\n"
        '      <div class="header-title">Morse Data Enterprises</div>\n'
        '      <div class="header-sub">Active Employee Directory &amp; Report Dashboard</div>\n'
        "    </div>\n"
        "  </div>\n"
        f'  <div class="header-badge">📅 {friendly_date}</div>\n'
        "</div>\n"
        "\n"
        '<div class="page">\n'
        "\n"
        '  <!-- Stat Cards -->\n'
        '  <div class="stats-grid">\n'
        f'    <div class="stat-card" id="card-all" onclick="setFilter(\'All\')">\n'
        '      <div class="stat-icon all">\U0001f465</div>\n'
        "      <div class=\"stat-info\">\n"
        f'        <div class="stat-value">{total}</div>\n'
        '        <div class="stat-label">Total Users</div>\n'
        '        <div class="stat-pct" style="background:#dbeafe;color:#1e40af;">All Active</div>\n'
        "      </div>\n"
        "    </div>\n"
        f'    <div class="stat-card" id="card-emp" onclick="setFilter(\'Employee\')">\n'
        '      <div class="stat-icon employee">\U0001f3e2</div>\n'
        "      <div class=\"stat-info\">\n"
        f'        <div class="stat-value">{n_employee}</div>\n'
        '        <div class="stat-label">Employees</div>\n'
        f'        <div class="stat-pct" style="background:var(--employee-light);color:#15803d;">{pct(n_employee)}</div>\n'
        "      </div>\n"
        "    </div>\n"
        f'    <div class="stat-card" id="card-con" onclick="setFilter(\'Contractor\')">\n'
        '      <div class="stat-icon contractor">\U0001f527</div>\n'
        "      <div class=\"stat-info\">\n"
        f'        <div class="stat-value">{n_contractor}</div>\n'
        '        <div class="stat-label">Contractors</div>\n'
        f'        <div class="stat-pct" style="background:var(--contractor-light);color:#b45309;">{pct(n_contractor)}</div>\n'
        "      </div>\n"
        "    </div>\n"
        f'    <div class="stat-card" id="card-int" onclick="setFilter(\'Intern\')">\n'
        '      <div class="stat-icon intern">\U0001f393</div>\n'
        "      <div class=\"stat-info\">\n"
        f'        <div class="stat-value">{n_intern}</div>\n'
        '        <div class="stat-label">Interns</div>\n'
        f'        <div class="stat-pct" style="background:var(--intern-light);color:#6d28d9;">{pct(n_intern)}</div>\n'
        "      </div>\n"
        "    </div>\n"
        "  </div>\n"
        "\n"
        '  <!-- Charts -->\n'
        '  <div class="charts-row">\n'
        '    <div class="chart-card">\n'
        '      <div class="chart-title"><span>\U0001f369</span> Workforce Breakdown</div>\n'
        '      <div class="donut-wrap"><canvas id="donutChart"></canvas></div>\n'
        "    </div>\n"
        '    <div class="chart-card">\n'
        '      <div class="chart-title"><span>\U0001f4ca</span> Top Managers by Headcount</div>\n'
        '      <div class="bar-wrap"><canvas id="barChart"></canvas></div>\n'
        "    </div>\n"
        "  </div>\n"
        "\n"
        '  <!-- Table -->\n'
        '  <div class="table-card">\n'
        '    <div class="table-toolbar">\n'
        '      <div class="toolbar-left">\n'
        '        <div class="search-box">\n'
        '          <span class="search-icon">\U0001f50d</span>\n'
        '          <input class="search-input" id="searchInput" type="text" placeholder="Search name or email…" oninput="applyFilters()">\n'
        "        </div>\n"
        '        <div class="filter-group">\n'
        '          <span class="filter-label">Category:</span>\n'
        '          <div class="filter-select-wrap">\n'
        '            <select class="filter-select sel-all" id="filterSelect" onchange="setFilter(this.value)">\n'
        '              <option value="All">\U0001f465 All Users</option>\n'
        '              <option value="Employee">\U0001f3e2 Employee</option>\n'
        '              <option value="Contractor">\U0001f527 Contractor</option>\n'
        '              <option value="Intern">\U0001f393 Intern</option>\n'
        "            </select>\n"
        "          </div>\n"
        "        </div>\n"
        "      </div>\n"
        '      <div class="toolbar-right">\n'
        '        <button class="dl-btn" id="dlBtn" onclick="downloadCSV()">⬇ Download List</button>\n'
        "      </div>\n"
        "    </div>\n"
        "\n"
        '    <div class="table-wrap">\n'
        '      <table id="empTable">\n'
        "        <thead>\n"
        "          <tr>\n"
        '            <th onclick="sortBy(\'displayName\')" id="th-displayName">Name <span class="sort-icon">↕</span></th>\n'
        '            <th onclick="sortBy(\'employeeType\')" id="th-employeeType">Type <span class="sort-icon">↕</span></th>\n'
        '            <th onclick="sortBy(\'mail\')" id="th-mail">Email <span class="sort-icon">↕</span></th>\n'
        '            <th onclick="sortBy(\'managerDisplayName\')" id="th-managerDisplayName">Manager <span class="sort-icon">↕</span></th>\n'
        '            <th onclick="sortBy(\'accountEnabled\')" id="th-accountEnabled">Status <span class="sort-icon">↕</span></th>\n'
        "          </tr>\n"
        "        </thead>\n"
        '        <tbody id="tableBody"></tbody>\n'
        "      </table>\n"
        '      <div class="no-results" id="noResults" style="display:none;">\n'
        '        <span class="emoji">\U0001f50d</span>No users match your current filter.\n'
        "      </div>\n"
        "    </div>\n"
        "\n"
        '    <div class="table-footer">\n'
        f'      <span id="rowCount">Showing {total} of {total} users</span>\n'
        "      <span>Morse Data Enterprises — Active Directory Export</span>\n"
        "    </div>\n"
        "  </div>\n"
        "\n"
        "</div>\n"
        "\n"
        "<script>\n"
        f"const ALL_USERS = {js_data};\n"
        "\n"
        "let currentFilter = 'All';\n"
        "let currentSearch = '';\n"
        "let sortKey = 'displayName';\n"
        "let sortAsc = true;\n"
        "let filteredData = [...ALL_USERS];\n"
        "\n"
        "const COLORS = ['#1e40af','#0891b2','#0d9488','#15803d','#a16207','#b45309','#9333ea','#be185d','#dc2626','#1d4ed8'];\n"
        "function avatarColor(name) {\n"
        "  let h = 0; for (let c of name) h = (h * 31 + c.charCodeAt(0)) & 0xffff;\n"
        "  return COLORS[h % COLORS.length];\n"
        "}\n"
        "function initials(name) {\n"
        "  const p = name.trim().split(' ');\n"
        "  return (p[0][0] + p[p.length > 1 ? p.length-1 : 0][0]).toUpperCase();\n"
        "}\n"
        "\n"
        "function buildCharts() {\n"
        "  const dCtx = document.getElementById('donutChart').getContext('2d');\n"
        f"  const typeCounts = {json.dumps(type_counts)};\n"
        "  const typeLabels = Object.keys(typeCounts);\n"
        "  const typeValues = Object.values(typeCounts);\n"
        "  const palette = { Employee:'#22c55e', Contractor:'#f59e0b', Intern:'#8b5cf6' };\n"
        "  const bgColors = typeLabels.map(l => palette[l] || '#94a3b8');\n"
        "  new Chart(dCtx, {\n"
        "    type: 'doughnut',\n"
        "    data: { labels: typeLabels, datasets: [{ data: typeValues,\n"
        "      backgroundColor: bgColors, borderColor: '#fff', borderWidth: 3, hoverOffset: 8 }] },\n"
        "    options: { responsive:true, maintainAspectRatio:false, cutout:'68%',\n"
        "      plugins: { legend: { position:'bottom', labels:{ padding:14, font:{size:12,weight:'600'}, usePointStyle:true, pointStyleWidth:8 }},\n"
        "        tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${ctx.parsed} (${Math.round(ctx.parsed/ALL_USERS.length*100)}%)` }} } }\n"
        "  });\n"
        "\n"
        "  const bCtx = document.getElementById('barChart').getContext('2d');\n"
        f"  const barLabels = {bar_labels};\n"
        f"  const barValues = {bar_values};\n"
        "  new Chart(bCtx, {\n"
        "    type: 'bar',\n"
        "    data: { labels: barLabels, datasets: [{ label:'Direct Reports', data: barValues,\n"
        "      backgroundColor: barValues.map((_,i)=> i===0?'#1e40af':i===1?'#0ea5e9':i===2?'#22c55e':'#94a3b8'),\n"
        "      borderRadius:6, borderSkipped:false }] },\n"
        "    options: { responsive:true, maintainAspectRatio:false, indexAxis:'y',\n"
        "      plugins: { legend:{display:false}, tooltip:{callbacks:{label:ctx=>` ${ctx.parsed.x} direct reports`}} },\n"
        "      scales: { x:{grid:{color:'#f1f5f9'},ticks:{stepSize:1,font:{size:11}}},\n"
        "                y:{grid:{display:false},ticks:{font:{size:11},color:'#374151'}} } }\n"
        "  });\n"
        "}\n"
        "\n"
        "function setFilter(type) {\n"
        "  currentFilter = type;\n"
        "  const sel = document.getElementById('filterSelect');\n"
        "  sel.value = type; sel.className = 'filter-select sel-' + type.toLowerCase();\n"
        "  ['card-all','card-emp','card-con','card-int'].forEach(id => document.getElementById(id).className = 'stat-card');\n"
        "  const cardMap = { All:'card-all active-all', Employee:'card-emp active-emp', Contractor:'card-con active-con', Intern:'card-int active-int' };\n"
        "  if (cardMap[type]) document.getElementById(cardMap[type].split(' ')[0]).className = 'stat-card active';\n"
        "  const dlLabel = type === 'All' ? 'All Users' : type + 's';\n"
        "  document.getElementById('dlBtn').textContent = '\\u2b07 Download ' + dlLabel;\n"
        "  applyFilters();\n"
        "}\n"
        "\n"
        "function applyFilters() {\n"
        "  currentSearch = document.getElementById('searchInput').value.toLowerCase();\n"
        "  filteredData = ALL_USERS.filter(u => {\n"
        "    const typeMatch = currentFilter === 'All' || u.employeeType === currentFilter;\n"
        "    const searchMatch = !currentSearch ||\n"
        "      (u.displayName||'').toLowerCase().includes(currentSearch) ||\n"
        "      (u.mail||'').toLowerCase().includes(currentSearch) ||\n"
        "      (u.managerDisplayName||'').toLowerCase().includes(currentSearch);\n"
        "    return typeMatch && searchMatch;\n"
        "  });\n"
        "  filteredData.sort((a,b) => {\n"
        "    let va = a[sortKey]||'', vb = b[sortKey]||'';\n"
        "    return sortAsc ? va.localeCompare(vb) : vb.localeCompare(va);\n"
        "  });\n"
        "  renderTable();\n"
        "}\n"
        "\n"
        "function sortBy(key) {\n"
        "  if (sortKey === key) sortAsc = !sortAsc; else { sortKey = key; sortAsc = true; }\n"
        "  document.querySelectorAll('thead th').forEach(th => th.classList.remove('sorted'));\n"
        "  document.getElementById('th-'+key).classList.add('sorted');\n"
        "  applyFilters();\n"
        "}\n"
        "\n"
        "function renderTable() {\n"
        "  const tbody = document.getElementById('tableBody');\n"
        "  const noRes = document.getElementById('noResults');\n"
        "  const count = document.getElementById('rowCount');\n"
        "  if (filteredData.length === 0) {\n"
        "    tbody.innerHTML = ''; noRes.style.display = 'block';\n"
        "    count.textContent = 'No results'; return;\n"
        "  }\n"
        "  noRes.style.display = 'none';\n"
        "  count.textContent = `Showing ${filteredData.length} of ${ALL_USERS.length} users`;\n"
        "  tbody.innerHTML = filteredData.map(u => {\n"
        "    const tc = u.employeeType === 'Employee' ? 'badge-employee'\n"
        "             : u.employeeType === 'Contractor' ? 'badge-contractor'\n"
        "             : u.employeeType === 'Intern' ? 'badge-intern' : 'badge-unknown';\n"
        "    const col = avatarColor(u.displayName||'?');\n"
        "    const ini = initials(u.displayName||'?');\n"
        "    return `<tr>\n"
        "      <td><div class=\"name-cell\"><div class=\"avatar\" style=\"background:${col}\">${ini}</div>\n"
        "        <div><div class=\"name-text\">${esc(u.displayName)}</div>\n"
        "        <div class=\"email-text\">${esc(u.userPrincipalName)}</div></div></div></td>\n"
        "      <td><span class=\"badge ${tc}\"><span class=\"badge-dot\"></span>${u.employeeType||'&mdash;'}</span></td>\n"
        "      <td style=\"color:var(--text-muted);font-size:12px\">${esc(u.mail)}</td>\n"
        "      <td>${esc(u.managerDisplayName)}</td>\n"
        "      <td><span class=\"enabled-dot\" title=\"Active\"></span></td>\n"
        "    </tr>`;\n"
        "  }).join('');\n"
        "}\n"
        "\n"
        "function esc(s) {\n"
        "  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'&quot;');\n"
        "}\n"
        "\n"
        "function downloadCSV() {\n"
        "  const header = ['id','displayName','userPrincipalName','mail','accountEnabled','employeeType','managerDisplayName'];\n"
        "  const rows = [header, ...filteredData.map(u => header.map(k => `\"${String(u[k]||'').replace(/\"/g,'\"\"')}\"`))];\n"
        "  const csv = rows.map(r => r.join(',')).join('\\n');\n"
        "  const blob = new Blob([csv], {type:'text/csv'});\n"
        "  const url = URL.createObjectURL(blob);\n"
        "  const a = document.createElement('a');\n"
        "  a.href = url;\n"
        "  a.download = currentFilter === 'All'\n"
        "    ? 'morse_active_employees_all.csv'\n"
        "    : `morse_active_employees_${currentFilter.toLowerCase()}s.csv`;\n"
        "  a.click(); URL.revokeObjectURL(url);\n"
        "}\n"
        "\n"
        "buildCharts();\n"
        "setFilter('All');\n"
        "</script>\n"
        "</body>\n"
        "</html>\n"
    )

    with open(filename, "w", encoding="utf-8") as f:
        f.write(html)
    return os.path.abspath(filename)


# ─────────────────────────────────────────────────────────────────────────────
# Email — send dashboard as HTML attachment via Microsoft Graph
# ─────────────────────────────────────────────────────────────────────────────

def send_dashboard_email(token: str, html_path: str) -> None:
    """
    Send the employee dashboard HTML as an email attachment using Graph API.

    Required .env variables:
        EMAIL_SENDER      — UPN of the mailbox to send from (e.g. jislam@itsmorse.com)
                            The app registration needs Mail.Send application permission.
        EMAIL_RECIPIENTS  — Comma-separated recipient addresses
                            (e.g. ceo@itsmorse.com,cfo@itsmorse.com)

    Optional .env variables:
        EMAIL_SUBJECT     — Subject line (default: "Active Employee Report — <friendly date>")
    """
    sender         = (os.getenv("EMAIL_SENDER") or "").strip()
    recipients_raw = (os.getenv("EMAIL_RECIPIENTS") or "").strip()

    if not sender:
        raise ValueError("Missing EMAIL_SENDER in .env")
    if not recipients_raw:
        raise ValueError("Missing EMAIL_RECIPIENTS in .env")

    recipients    = [r.strip() for r in recipients_raw.split(",") if r.strip()]
    now           = datetime.now()
    friendly_date = now.strftime(f"%A, %B {now.day}, %Y")
    subject       = (os.getenv("EMAIL_SUBJECT") or "").strip() or f"Active Employee Report — {friendly_date}"
    attachment_name = os.path.basename(html_path)

    body_html = f"""
<div style="font-family:'Segoe UI',Arial,sans-serif;max-width:600px;margin:0 auto;background:#f4f6fb;">

  <!-- Header -->
  <div style="background-color:#1e3a8a;padding:32px 40px;border-radius:12px 12px 0 0;">
    <div style="display:inline-block;width:44px;height:44px;background-color:#2d4fa8;
                border:2px solid #4a6abf;border-radius:10px;text-align:center;
                line-height:44px;font-size:20px;font-weight:800;color:#ffffff;margin-bottom:12px;">M</div>
    <h1 style="margin:0;font-size:22px;color:#ffffff;font-weight:700;">Active Employee Report</h1>
    <p style="margin:4px 0 0;color:#a8c4f0;font-size:13px;">
      Morse Data Enterprises &mdash; {friendly_date}
    </p>
  </div>

  <!-- Body -->
  <div style="background:#ffffff;padding:32px 40px;border:1px solid #e2e8f0;border-top:none;">
    <p style="margin:0 0 20px;font-size:15px;color:#1e293b;">Hi,</p>
    <p style="margin:0 0 24px;font-size:15px;color:#374151;line-height:1.6;">
      Please find attached the latest <strong>Morse Data Enterprises Active Employee Report</strong>,
      generated on <strong>{friendly_date}</strong> directly from Azure Active Directory.
    </p>

    <!-- How to open -->
    <div style="background:#eff6ff;border-left:4px solid #1e40af;border-radius:4px;padding:18px 22px;margin:0 0 28px;">
      <p style="margin:0 0 12px;font-size:14px;font-weight:700;color:#1e40af;">
        How to open the dashboard
      </p>
      <ol style="margin:0;padding-left:20px;font-size:14px;color:#374151;line-height:2;">
        <li>Find the attachment <strong>{attachment_name}</strong> of this email</li>
        <li>Click it and choose <strong>Open</strong> &mdash; or save it to your desktop first</li>
        <li>Double click the downloaded file to open directly in your browser</li>
        <li>No login required &mdash; everything is self-contained in the file</li>
      </ol>
    </div>

    <!-- What's inside -->
    <p style="margin:0 0 10px;font-size:14px;font-weight:600;color:#1e293b;">What&rsquo;s inside</p>
    <table style="width:100%;border-collapse:collapse;font-size:14px;color:#374151;margin-bottom:28px;">
      <tr>
        <td style="padding:8px 0;border-bottom:1px solid #f1f5f9;">
          <span style="color:#1e40af;font-weight:700;margin-right:8px;">&#9654;</span>
          Headcount summary cards
        </td>
        <td style="padding:8px 0;border-bottom:1px solid #f1f5f9;color:#64748b;font-size:13px;">
          Total, Employees, Contractors, Interns
        </td>
      </tr>
      <tr>
        <td style="padding:8px 0;border-bottom:1px solid #f1f5f9;">
          <span style="color:#1e40af;font-weight:700;margin-right:8px;">&#9654;</span>
          Workforce breakdown chart
        </td>
        <td style="padding:8px 0;border-bottom:1px solid #f1f5f9;color:#64748b;font-size:13px;">
          Interactive donut chart by employee type
        </td>
      </tr>
      <tr>
        <td style="padding:8px 0;border-bottom:1px solid #f1f5f9;">
          <span style="color:#1e40af;font-weight:700;margin-right:8px;">&#9654;</span>
          Top managers by headcount
        </td>
        <td style="padding:8px 0;border-bottom:1px solid #f1f5f9;color:#64748b;font-size:13px;">
          Horizontal bar chart, top 10
        </td>
      </tr>
      <tr>
        <td style="padding:8px 0;">
          <span style="color:#1e40af;font-weight:700;margin-right:8px;">&#9654;</span>
          Employee directory
        </td>
        <td style="padding:8px 0;color:#64748b;font-size:13px;">
          Searchable, filterable &amp; exportable to CSV
        </td>
      </tr>
    </table>

    <p style="margin:0;font-size:13px;color:#94a3b8;border-top:1px solid #e2e8f0;padding-top:20px;line-height:1.6;">
      This report is generated automatically from Azure Active Directory.<br>
      For questions or access issues, contact your IT administrator.
    </p>
  </div>

  <!-- Footer -->
  <div style="background:#f8fafc;padding:16px 40px;border:1px solid #e2e8f0;border-top:none;
              border-radius:0 0 12px 12px;text-align:center;">
    <p style="margin:0;font-size:12px;color:#94a3b8;">
      Workforce Dashboard Pipeline &mdash; Morse Data Enterprises
    </p>
  </div>

</div>
"""

    # Read and base64-encode the HTML attachment
    with open(html_path, "rb") as f:
        attachment_bytes = base64.b64encode(f.read()).decode("utf-8")

    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": body_html,
            },
            "toRecipients": [
                {"emailAddress": {"address": addr}} for addr in recipients
            ],
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": attachment_name,
                    "contentType": "text/html",
                    "contentBytes": attachment_bytes,
                }
            ],
        },
        "saveToSentItems": "true",
    }

    url = f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail"
    r = requests.post(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=payload,
        timeout=60,
    )
    r.raise_for_status()
    print(f"  Email sent to: {', '.join(recipients)} ✓")


# ─────────────────────────────────────────────────────────────────────────────
# Main pipeline
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":

    now = datetime.now()
    print("=" * 60)
    print("  Workforce Dashboard Pipeline")
    print(f"  {now.strftime(f'%A, %B {now.day}, %Y')}")
    print("=" * 60)

    # ── 1. Fetch from Azure AD ────────────────────────────────────
    print("\n[1/4] Fetching users from Microsoft Graph...")
    graph_token = get_graph_token()
    users       = get_all_users(graph_token)
    print(f"  Total users fetched: {len(users)}")

    # ── 2. Filter + flatten ───────────────────────────────────────
    print("\n[2/4] Filtering active non-service accounts...")
    rows = [flatten_user_row(u) for u in users if is_active_employee_non_service(u)]
    print(f"  Filtered rows: {len(rows)}")

    # ── 3. Export CSV + Dashboard HTML ───────────────────────────
    print("\n[3/4] Exporting CSV and generating dashboard...")
    export_csv(rows)
    print("  CSV exported ✓")
    dashboard_path = export_morse_employee_dashboard(rows)
    print(f"  Dashboard saved → {dashboard_path} ✓")

    # ── 4. Email dashboard to executives ─────────────────────────
    print("\n[4/4] Sending dashboard email...")
    try:
        send_dashboard_email(graph_token, dashboard_path)
    except Exception as e:
        print(f"  Email skipped: {e}")

    print("\n" + "=" * 60)
    print(f"  Pipeline complete ✓  |  Rows: {len(rows)}")
    print("=" * 60)