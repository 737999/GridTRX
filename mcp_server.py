"""
GridTRX MCP Server — structured AI agent interface to the accounting engine.

Wraps models.py functions as MCP tools. Every tool takes db_path as its first
parameter. The GRIDTRX_WORKSPACE environment variable must be set to the
directory containing client books.db files — the server will refuse to start
without it, and rejects any db_path outside that directory.

Usage:
    pip install mcp
    GRIDTRX_WORKSPACE=~/clients python mcp_server.py
"""
import sys, os, csv
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from datetime import datetime, timedelta
from mcp.server.fastmcp import FastMCP
import models

mcp = FastMCP("GridTRX", instructions="""Double-entry accounting engine.

Key account naming conventions:
  BANK.CHQ, BANK.SAV  — bank accounts (debit-normal)
  AR                   — accounts receivable
  AP                   — accounts payable
  SH.LOAN              — shareholder loan (credit-normal = owed TO shareholder)
  CAPITAL              — share capital
  RE                   — retained earnings on the BS (perpetual TOTAL — never post here)
  RE.OB                — opening retained earnings (brought forward, credit-normal)
  REV.SVC, REV.DIV     — revenue accounts (credit-normal)
  EX.SC                — bank charges / service charges (ALWAYS use this for bank fees)
  EX.FEES              — professional fees (accounting, legal)
  EX.OFFICE            — office & general
  EX.SAL               — salaries & wages
  EX.RENT              — rent
  EX.SUSP              — suspense (unclassified — reclassify with reclassify_suspense)
  CLR.TSF              — clearing account for inter-account transfers

Bank charges and service charges: ALWAYS map to EX.SC. This includes monthly fees,
NSF charges, overdraft fees, interest charges, wire fees, and any other bank-imposed costs.

New client setup — THREE QUESTIONS FIRST, then five steps:
  Q1 What KIND of business? 'services' (no cost of sales) / 'retail' (purchases,
     freight, inventory) / 'construction' (materials, subcontract, labour,
     equipment, WIP, holdbacks). It changes the shape of the statements. ASK.
  Q2 What YEAR-END are you working on? The job being completed for the client.
  Q3 Are there opening balances? If the client has a prior year, yes.

  1. create_books(db_path, company_name, business_type, fiscal_year_end)
     — the ONLY way to make a set of books. Every other tool opens EXISTING
       books and refuses a path that does not exist.
  2. Customize: add_account for client-specific accounts (report_name required)
  3. Opening balances: post the PRIOR YEAR CLOSING figures with
     post_opening_balances DATED AT THE PRIOR YEAR-END. For a 2026-05-31 year
     end that is 2025-05-31 — NOT 2025-06-01. A day later puts them in the
     current year and the CY/PY comparative column is wrong. Everything enters
     through the TRX ledger and retained earnings is the computed residual.
  4. Import bank data via import_csv or import_ofx. An import that posts
     NOTHING raises — read the reason, it is usually the posting ceiling.
  5. Reclassify suspense items via reclassify_suspense (auto-learns rules).
     get_ledger('EX.SUSP') gives you the txn_id it wants.
  Run check_books() when you pick up a file and again before reporting done.

Opening balances (existing business onboarding):
  The goal is to post the PRIOR YEAR trial balance so that:
  (a) the balance sheet has correct opening balances, and
  (b) the income statement has PY comparative columns.

  Source data (in priority order):
    1. Formal prior-year trial balance (best — complete, balanced)
    2. PY balance sheet + income statement (from financial statements)
    3. Tax return schedules (CRA S100/S125, US Form 1120, etc.)
    4. Accountant's working papers

  IMPORTANT: Post BOTH balance sheet AND income statement accounts. If you only
  post the BS, your PY comparative on the IS will be blank. Always post the full TB.

  Do NOT use import_tb for this, and do NOT hand-post it with post_transaction.
  There is ONE tool: post_opening_balances(db_path, conversion_date, balances,
  expected_retained_earnings). It creates TRX.OPEN, posts every account against it,
  computes retained earnings and posts it to RE.OB, and leaves TRX.OPEN at zero —
  atomically, so a bad row posts nothing.

    post_opening_balances(db_path, "2024-09-30", [
        {"account": "BANK.CHQ", "description": "Chequing",    "amount": "33268.00"},
        {"account": "AP",       "description": "Payables",    "amount": "(1000.00)"},
        {"account": "REV",      "description": "Sales",       "amount": "(84000.00)"},
    ], expected_retained_earnings="35065.00")

  amount is a DOLLAR string signed the way Grid displays: plain = DEBIT,
  "(1000.00)" or "-1000.00" = CREDIT. One row per line of the PY trial balance.

  NEVER pass retained earnings in the list — it is COMPUTED as the residual and
  posted for you. RE, RE.OB, RE.OPEN, RE.CLOSE and EX.SUSP are all refused, as is
  any account missing from the chart (add it first). Never park a conversion in
  EX.SUSP and clear it down — suspense is for transactions you cannot identify.

  Pass expected_retained_earnings whenever you can read RE off the prior-year
  statements: the post then REFUSES unless the residual ties, which is the check
  that catches a transposed or missing number.

  Which RE figure comes out depends on what you post:
    balance sheet only        -> retained earnings BROUGHT FORWARD (closing RE).
    full TB incl. IS accounts -> retained earnings at the START of that year
                                 (T2 line 3660, NOT 3600). Post the full TB when
                                 you want the IS to have a prior-year comparative;
                                 PY net income then reaches closing RE by itself.
    Post DIVPAID (line 3701) too if dividends were paid in that year.

  Check with opening_balances_status(db_path). Wrong? delete_opening_balances(db_path)
  and post it again — Grid does not adjust a bad conversion, it redoes it.
  TRX.OPEN lives on the TRX report; open it to see every opening balance loaded.

import_csv date handling: the importer auto-detects date format but may misread DD-MM-YYYY
as MM-DD-YYYY when the day is <= 12. Verify dates after import if the bank uses DD-MM format.

Reports: BS and IS are created by create_starter_books. Always use list_reports to get
report names (they are short codes like 'BS', 'IS', not full names like 'Balance Sheet').
""")

_initialized_db = None
_workspace = None

def _get_workspace():
    """Return the resolved workspace path. Raises if GRIDTRX_WORKSPACE is not set."""
    global _workspace
    if _workspace is None:
        ws = os.environ.get("GRIDTRX_WORKSPACE", "")
        if not ws:
            raise RuntimeError(
                "GRIDTRX_WORKSPACE environment variable is not set. "
                "Set it to the directory containing client books.db files. "
                "Example: GRIDTRX_WORKSPACE=~/clients python mcp_server.py"
            )
        _workspace = os.path.realpath(os.path.expanduser(ws))
    return _workspace

def _check_path(file_path: str, label: str = "path"):
    """Enforce workspace boundary on any file path (imports, exports, etc.)."""
    resolved = os.path.realpath(os.path.expanduser(file_path))
    ws = _get_workspace()
    if not resolved.startswith(ws + os.sep) and resolved != ws:
        raise ValueError(
            f"Access denied: {label} '{file_path}' is outside the workspace ({ws})."
        )
    return resolved

def _init(db_path: str, create: bool = False):
    """Open a set of books. Enforces the workspace boundary.

    create=False (everything except create_books) REFUSES a path that does not
    exist. Opening and creating are different acts: asking a question about a
    file must never bring one into being. That used to happen — get_info on a
    mistyped path silently produced a half-built set of books with no chart of
    accounts, and every call after it failed for reasons that had nothing to do
    with the typo."""
    global _initialized_db
    resolved = os.path.realpath(os.path.expanduser(db_path))
    ws = _get_workspace()
    if not resolved.startswith(ws + os.sep) and resolved != ws:
        raise ValueError(
            f"Access denied: '{db_path}' is outside the workspace ({ws}). "
            f"Set GRIDTRX_WORKSPACE to change the allowed directory."
        )
    if not create and not os.path.exists(resolved):
        raise ValueError(
            f"No books at '{db_path}'. This tool opens EXISTING books and will "
            f"not create them. If this client is new, call "
            f"create_books(db_path, company_name, business_type, fiscal_year_end) "
            f"— it builds the chart of accounts and the statements. If the client "
            f"already exists, check the path."
        )
    if _initialized_db != resolved:
        try:
            models.init_db(resolved)
            _repair = models.re_repair_note()
            if _repair:
                # Not an exception — the books ARE open and correct now. But an
                # agent that never mentions this to the operator has hidden a
                # change to retained earnings, so it is put where every tool
                # call will trip over it.
                sys.stderr.write('GridTRX: ' + ' '.join(_repair) + '\n')
        except models.BooksLocked as e:
            # v130 — an agent is refused for the same reason a person is, and
            # gets the same door: the tool that clears it, named right here.
            raise ValueError(
                f"{e} "
                f"Do not guess: ask the operator whether anyone is in these books. "
                f"If they confirm nobody is, call clear_lock(db_path, "
                f"confirm='nobody is in these books') and open again."
            )
        _initialized_db = resolved


def _row_to_dict(row):
    """Convert a sqlite3.Row to a plain dict."""
    if row is None:
        return None
    return dict(row)


def _rows_to_dicts(rows):
    """Convert a list of sqlite3.Row to a list of plain dicts."""
    return [dict(r) for r in rows]


# ═══════════════════════════════════════════════════════════════════
# READ-ONLY TOOLS
# ═══════════════════════════════════════════════════════════════════

@mcp.tool()
def list_procedures(db_path: str) -> list[dict]:
    """The whole jobs this program knows how to do end to end.

    A procedure is an assembly the firm does often enough to have written down —
    the same reports, the same order, every time. Prefer one over composing the
    steps yourself: it is what the operator expects to receive, and it cannot
    drift from run to run. Judgement is NOT in here; deciding what a suspense
    row is remains a person's work."""
    _init(db_path)
    import app as _app
    return [{"name": p["key"], "title": p["title"], "note": p["note"],
             "asks_for": p["prompt"],
             "report_choice": bool(p.get("report_choice"))} for p in _app.PROCEDURES]


@mcp.tool()
def run_procedure(db_path: str, name: str, as_of: str, output_path: str,
                  report_name: str = "IS") -> dict:
    """Run one procedure and write its PDF to output_path.

    as_of        REQUIRED. The year-end (or last month) to report — never
                 assumed. Ask the operator which year they want if it is not
                 already clear; do not silently print the current one.
    report_name  only for procedures whose `report_choice` is true.
    """
    _init(db_path)
    import app as _app
    proc = _app._procedure(name)
    if not proc:
        raise ValueError(f"No procedure '{name}'. Use list_procedures to see them.")
    as_of = _normalize_date(as_of)
    if not as_of:
        raise ValueError("as_of must be a date, e.g. 2026-12-31.")
    resolved = _check_path(output_path, "output_path")
    if proc.get("report_choice"):
        pdf, filename, note = proc["run"](as_of, report_name)
    else:
        pdf, filename, note = proc["run"](as_of)
    with open(resolved, "wb") as fh:
        fh.write(pdf)
    return {"procedure": proc["title"], "as_of": as_of, "path": resolved,
            "bytes": len(pdf), "suggested_name": filename, "contains": note}


def _import_landed(result, what):
    """An import that posted NOTHING is a failure, not a result.

    It used to return quietly — posted: 0, skipped: 5, reasons in a list nobody
    was obliged to read — and every later call failed for reasons that made no
    sense. Worse, check_books then pronounced the empty books SOUND, because
    empty books have nothing wrong with them. The whole job could finish and
    print a confident, entirely blank year-end package.

    So: nothing landed = raise, and say WHY, in the reason's own words, with the
    fix named."""
    if result.get('rows_processed', 0) and not result.get('posted', 0):
        reasons, seen = [], set()
        for e in result.get('errors', []) or []:
            r = (e.get('reason') if isinstance(e, dict) else str(e)) or ''
            if r and r not in seen:
                seen.add(r); reasons.append(r)
        why = '; '.join(reasons[:4]) or 'no reason recorded'
        fix = ''
        if any('after fiscal year end' in r.lower() for r in reasons):
            fix = (" These rows are dated AFTER the year-end this file is working on. "
                   "Check get_info()['working_year_end'] — if you are meant to be doing a "
                   "later year, move it with set_fiscal_settings(working_year_end=...) "
                   "and import again. Do not silently drop the rows.")
        raise ValueError(
            f"{what}: {result['rows_processed']} row(s) read and NOTHING was posted. "
            f"Reason: {why}.{fix}")
    return result


@mcp.tool()
def create_books(db_path: str, company_name: str, business_type: str,
                 fiscal_year_end: str) -> dict:
    """Create a NEW set of books: chart of accounts, balance sheet, income
    statement, adjusting entries. This is the FIRST call for a new client —
    nothing else will create a file.

    Three questions, and they are the three an accountant asks:

    business_type    what the business DOES, which decides the top of the
                     income statement:
                       'services'      no cost of sales at all
                       'retail'        purchases / freight-in / inventory
                                       adjustment, and Inventory on the BS
                       'construction'  materials / subcontract / labour /
                                       equipment, with Work in Progress and
                                       holdbacks on the BS
                     Ask the operator if you do not know. Do NOT guess: it
                     changes the shape of the statements.

    fiscal_year_end  the year-end being WORKED ON, YYYY-MM-DD — the job you are
                     completing for this client. Posting stops there.

    Then opening balances. If the client has a prior year, post the PRIOR YEAR
    CLOSING figures dated at the PRIOR YEAR-END — for a 2026-05-31 year-end
    that is 2026-05-31 minus one year = 2025-05-31, NOT 2025-06-01. Dating them
    a day later puts them in the current year and the CY/PY comparative breaks.
    Use post_opening_balances: it enters everything through the TRX ledger and
    populates retained earnings as the computed residual. Never hand-post a
    conversion.
    """
    resolved = _check_path(db_path, "books")
    if os.path.exists(resolved):
        raise ValueError(
            f"'{db_path}' already exists — refusing to build over a set of books. "
            f"Open it instead, or choose a different folder.")
    fye = _normalize_date(fiscal_year_end)
    if not fye:
        raise ValueError("fiscal_year_end must be the year-end being worked on, "
                         "as YYYY-MM-DD (e.g. 2026-05-31).")
    bt = (business_type or "").strip().lower()
    if bt not in models.BUSINESS_TYPES:
        raise ValueError(
            "business_type must be one of: " +
            "; ".join(f"'{k}' — {v}" for k, v in models.BUSINESS_TYPE_NOTE.items()) +
            ". Ask the operator which one this client is.")
    os.makedirs(os.path.dirname(resolved), exist_ok=True)
    models.create_starter_books(resolved, company_name, fye[5:], business_type=bt)
    models.set_fiscal_settings(working_ye=fye)
    global _initialized_db
    _initialized_db = resolved
    anchor = models.fiscal_anchor()
    return {
        "created": resolved,
        "company_name": company_name,
        "business_type": bt,
        "business_type_note": models.BUSINESS_TYPE_NOTE[bt],
        "working_year_end": anchor["cy_end"],
        "comparative_year_end": anchor["py_end"],
        "accounts": len(models.get_accounts()),
        "reports": [r["name"] for r in models.get_reports()],
        "next_step": (
            f"If this client has a prior year, post its CLOSING trial balance with "
            f"post_opening_balances dated {anchor['py_end']} (the PRIOR year-end — "
            f"not the day after, or the comparative column will be wrong). "
            f"If it is a brand-new business with nothing brought forward, skip it "
            f"and import the bank data."),
    }


@mcp.tool()
def export_writeup(db_path: str, out_path: str = "") -> dict:
    """Export the write-up handoff for the working-paper program (Willy).

    Writes ONE JSON file: current + prior fiscal year fully itemized (so the
    write-up agent can compare this year's postings against last year's),
    FIVE fiscal years of per-account comparatives (balance at each year-end
    + activity within each year), the chart with leadsheet codes, the
    statement layouts, and the check_books verdict the write-up program's
    front-door gate reads. Requires the fiscal year to be set.

    out_path: where to write; default <Company>_<FY>_writeup.json beside the
    books. Returns the meta block + where it was written.
    """
    import json as _json
    _init(db_path)
    payload = models.export_writeup()
    if not out_path:
        safe = payload['meta']['company_name'].replace(' ', '_')[:40]
        out_path = os.path.join(
            os.path.dirname(models.get_db_path()),
            f"{safe}_{payload['meta']['fiscal_year']}_writeup.json")
    with open(out_path, 'w') as f:
        _json.dump(payload, f, indent=1)
    return {"written": out_path, "meta": payload['meta'],
            "accounts": len(payload['accounts']),
            "comparative_years": [c['fiscal_year']
                                  for c in payload['comparatives']]}


@mcp.tool()
def check_books(db_path: str) -> dict:
    """Is this set of books sound, and what is outstanding in it? Read-only.

    Run this FIRST when you pick up a file, and again before you report work as
    finished. One call replaces guessing: file integrity, the trial balance,
    whether every transaction balances, whether the report chain holds and the
    balance sheet balances, what is parked in Suspense, rows dated past the
    year-end, opening balances, reconciliation continuity, accounts that appear
    on no statement, and today's snapshot.

    Each check is 'ok', 'attention' (real work outstanding — not a fault) or
    'error' (the books are WRONG; fix before doing anything else). Report
    errors to the operator rather than working around them.
    """
    _init(db_path)
    return models.check_books()


@mcp.tool()
def lock_status(db_path: str) -> dict:
    """Who currently holds these books, if anyone. Read-only — never opens them.
    Use this when an open was refused, to report the holder to the operator."""
    resolved = _check_path(db_path, "books")
    info = models.books_lock_info(resolved)
    if not info:
        return {"locked": False}
    return {"locked": True, "pid": info.get("pid", "?"), "host": info.get("host", "?"),
            "prog": info.get("prog", "?"), "since": info.get("started", "?"),
            "lock_file": models._lock_path_for(resolved)}


@mcp.tool()
def clear_lock(db_path: str, confirm: str) -> dict:
    """Clear a leftover lock on a set of books so they can be opened.

    THE OPERATOR'S CALL, NEVER YOURS. A lock means someone may be in the file;
    clearing it while they are gives two sessions the same books. Only call this
    after a human has told you nobody is in them, and pass their confirmation as
    `confirm` (must say 'nobody'). Check lock_status first and report the holder.
    """
    if 'nobody' not in (confirm or '').lower():
        raise ValueError(
            "clear_lock needs the operator's confirmation that nobody is in these "
            "books — pass confirm='nobody is in these books'. Ask them first; report "
            "what lock_status says about who holds the file."
        )
    resolved = _check_path(db_path, "books")
    info = models.books_lock_info(resolved)
    if not info:
        return {"cleared": False, "note": "These books were not locked."}
    models.clear_books_lock(resolved)
    return {"cleared": True, "was_held_by": {
        "pid": info.get("pid", "?"), "host": info.get("host", "?"),
        "prog": info.get("prog", "?"), "since": info.get("started", "?")}}


@mcp.tool()
def list_accounts(db_path: str, query: str = "") -> list[dict]:
    """List all accounts in the chart of accounts. Optionally filter by name/description with query."""
    _init(db_path)
    if query:
        rows = models.search_accounts(query)
    else:
        rows = models.get_accounts()
    return [
        {
            "id": r["id"],
            "name": r["name"],
            "description": r["description"],
            "normal_balance": r["normal_balance"],
            "type": r["account_type"],
        }
        for r in rows
    ]


@mcp.tool()
def set_gifi(db_path: str, account_name: str, gifi_code: str) -> dict:
    """Set the GIFI code on an account. Code must be a valid CRA GIFI code.
    Use list_gifi_codes to see available codes. Pass empty string to clear."""
    _init(db_path)
    models.set_gifi(account_name, gifi_code)
    desc = models.GIFI_CODES.get(gifi_code, '') if gifi_code else ''
    return {"account": account_name, "gifi_code": gifi_code, "gifi_description": desc}


@mcp.tool()
def list_gifi_codes(db_path: str, query: str = "") -> list[dict]:
    """List available GIFI codes. Optionally filter by code or description."""
    _init(db_path)
    results = []
    for code, desc in sorted(models.GIFI_CODES.items()):
        if query:
            q = query.lower()
            if q not in code and q not in desc.lower():
                continue
        results.append({"code": code, "description": desc})
    return results


@mcp.tool()
def get_gifi_map(db_path: str) -> list[dict]:
    """Show all accounts with GIFI codes assigned."""
    _init(db_path)
    rows = models.get_gifi_map()
    return [{"account": r["name"], "description": r["description"],
             "gifi_code": r["gifi_code"],
             "gifi_description": models.GIFI_CODES.get(r["gifi_code"], '')}
            for r in rows]


@mcp.tool()
def gifi_export(db_path: str, date_from: str = "", date_to: str = "",
                output_path: str = "") -> dict:
    """Export GL balances rolled up by GIFI code for T2 S100/S125.

    Returns schedule_100 (balance sheet), schedule_125 (income statement),
    and t2engine_input (ready for T2Engine gifi.json).

    If output_path is provided, writes the t2engine_input as a JSON file
    that can be passed directly to: python3 t2engine.py oneshot <gifi.json> ...
    """
    _init(db_path)
    result = models.gifi_export(date_from or None, date_to or None)

    if output_path:
        import json
        output_path = _check_path(output_path, "output_path")
        company = models.get_meta('company_name', '')
        t2_data = dict(result['t2engine_input'])
        t2_data['_comment'] = f'{company} - {result["period"]}'
        with open(output_path, 'w') as f:
            json.dump(t2_data, f, indent=2)
        result['file_written'] = output_path

    return result


@mcp.tool()
def get_balance(
    db_path: str,
    account_name: str,
    date_from: str = "",
    date_to: str = "",
) -> dict:
    """Get the balance of a single account, optionally within a date range (YYYY-MM-DD)."""
    _init(db_path)
    acct = models.get_account_by_name(account_name)
    if not acct:
        raise ValueError(f"Account not found: {account_name}")
    raw = models.get_account_balance(
        acct["id"],
        date_from=date_from or None,
        date_to=date_to or None,
    )
    sign = 1 if acct["normal_balance"] == "D" else -1
    balance = raw * sign
    return {
        "account": acct["name"],
        "balance_cents": balance,
        "formatted": models.fmt_amount(balance),
    }


@mcp.tool()
def get_ledger(
    db_path: str,
    account_name: str,
    date_from: str = "",
    date_to: str = "",
) -> list[dict]:
    """Get the full ledger for an account with running balance. Optionally filter by date range."""
    _init(db_path)
    acct = models.get_account_by_name(account_name)
    if not acct:
        raise ValueError(f"Account not found: {account_name}")
    entries = models.get_ledger(
        acct["id"],
        date_from=date_from or None,
        date_to=date_to or None,
    )
    return [
        {
            "txn_id": e["txn_id"],
            "date": e["date"],
            "reference": e["reference"],
            "description": e["description"],
            "amount_cents": e["amount"],
            "amount_formatted": models.fmt_amount(e["amount"]),
            "running_balance_cents": e["running_balance"],
            "running_balance_formatted": models.fmt_amount(e["running_balance"]),
            "cross_accounts": e["cross_accounts"],
            "reconciled": bool(e["reconciled"]),
        }
        for e in entries
    ]


@mcp.tool()
def trial_balance(db_path: str, as_of_date: str = "") -> dict:
    """Get the trial balance — all posting accounts with non-zero balances, split into Dr/Cr columns."""
    _init(db_path)
    accounts, total_dr, total_cr = models.get_trial_balance(
        as_of_date=as_of_date or None,
    )
    return {
        "accounts": [
            {
                "name": a["name"],
                "description": a["description"],
                "normal_balance": a["normal_balance"],
                "debit_cents": a["debit"],
                "debit_formatted": models.fmt_amount(a["debit"]) if a["debit"] else "",
                "credit_cents": a["credit"],
                "credit_formatted": models.fmt_amount(a["credit"]) if a["credit"] else "",
            }
            for a in accounts
        ],
        "total_debit_cents": total_dr,
        "total_debit_formatted": models.fmt_amount(total_dr),
        "total_credit_cents": total_cr,
        "total_credit_formatted": models.fmt_amount(total_cr),
    }


@mcp.tool()
def generate_report(
    db_path: str,
    report_name: str,
    date_from: str = "",
    date_to: str = "",
    side: str = "",
) -> list[dict]:
    """Generate a financial report (BS, IS, AJE, etc.) with computed balances. Returns line items.

    side  ''  the net balance — an ordinary statement column.
          'D' DEBITS ONLY, 'C' CREDITS ONLY — the FLOWS through each account
              rather than the net. This is how you answer questions a netted
              report cannot: the debit side of a detailed AR is SALES BY
              CUSTOMER, the credit side of a detailed AP is PURCHASES BY
              SUPPLIER, the debit side of inventory is purchases and the credit
              side is cost of sales. Do not sum ledgers by hand to get these.
              Partial-payment boomerangs are excluded so flows are not inflated.
    """
    _init(db_path)
    report = models.find_report_by_name(report_name)
    if not report:
        raise ValueError(f"Report not found: {report_name}")
    side = (side or "").upper() or None
    if side not in (None, "D", "C"):
        raise ValueError(f"side must be '', 'D' or 'C' — got {side!r}")
    items = models.compute_report_column(
        report["id"],
        date_from=date_from or None,
        date_to=date_to or None,
        side=side,
    )
    result = []
    for item_dict, amount in items:
        entry = {
            "description": item_dict.get("description") or item_dict.get("acct_name") or "",
            "item_type": item_dict.get("item_type", ""),
            "indent": item_dict.get("indent", 0),
            "amount_cents": amount,
            "amount_formatted": models.fmt_amount(amount),
        }
        if item_dict.get("acct_name"):
            entry["account_name"] = item_dict["acct_name"]
        if item_dict.get("sep_style"):
            entry["separator_style"] = item_dict["sep_style"]
        result.append(entry)
    return result


@mcp.tool()
def get_transaction(db_path: str, txn_id: int) -> dict:
    """Get a single transaction by ID, including all its journal lines."""
    _init(db_path)
    txn, lines = models.get_transaction(txn_id)
    if not txn:
        raise ValueError(f"Transaction not found: {txn_id}")
    return {
        "id": txn["id"],
        "date": txn["date"],
        "description": txn["description"],
        "reference": txn["reference"],
        "lines": [
            {
                "account_name": l["account_name"],
                "amount_cents": l["amount"],
                "amount_formatted": models.fmt_amount(l["amount"]),
                "description": l["description"],
                "reconciled": bool(l["reconciled"]),
            }
            for l in lines
        ],
    }


@mcp.tool()
def search_transactions(
    db_path: str, query: str, limit: int = 100
) -> list[dict]:
    """Search transactions by description, reference, or account name."""
    _init(db_path)
    rows = models.search_transactions(query, limit=limit)
    return [
        {
            "txn_id": r["txn_id"],
            "date": r["date"],
            "reference": r["reference"],
            "description": r["description"],
            "accounts": r["accounts"],
            "total_amount_cents": r["total_amount"] or 0,
            "total_amount_formatted": models.fmt_amount(r["total_amount"] or 0),
        }
        for r in rows
    ]


@mcp.tool()
def list_reports(db_path: str) -> list[dict]:
    """List all available reports (BS, IS, AJE, etc.)."""
    _init(db_path)
    rows = models.get_reports()
    return [
        {
            "id": r["id"],
            "name": r["name"],
            "description": r["description"],
        }
        for r in rows
    ]


@mcp.tool()
def update_report(db_path: str, report_name: str, description: str) -> dict:
    """Update a report's description."""
    _init(db_path)
    rpt = models.find_report_by_name(report_name)
    if not rpt:
        raise ValueError(f"Report not found: {report_name}")
    return models.update_report(rpt['id'], description=description)


@mcp.tool()
def list_rules(db_path: str) -> list[dict]:
    """List all import rules for CSV auto-categorization."""
    _init(db_path)
    rows = models.get_import_rules()
    return [
        {
            "id": r["id"],
            "keyword": r["keyword"],
            "account": r["account_name"],
            "tax_code": r["tax_code"],
            "priority": r["priority"],
        }
        for r in rows
    ]


@mcp.tool()
def get_info(db_path: str) -> dict:
    """Get company metadata: name, the fiscal year end being worked on, the
    posting ceiling and the lock date."""
    _init(db_path)
    anchor = models.fiscal_anchor()
    return {
        "company_name": models.get_meta("company_name", ""),
        "fiscal_year_end": models.get_meta("fiscal_year_end", ""),
        "working_year_end": anchor["cy_end"] if anchor else "",
        "period_start": anchor["cy_start"] if anchor else "",
        "fy_ceiling": models.fiscal_ceiling(),
        "fy_ceiling_mode": models.get_meta("fy_ceiling_mode", "cy") or "cy",
        "lock_date": models.get_meta("lock_date", ""),
        "latest_transaction": models.latest_transaction_date(),
    }


@mcp.tool()
def set_fiscal_settings(db_path: str, working_year_end: str = "", ceiling_mode: str = "",
                        lock_date: str = "", company_name: str = "") -> dict:
    """Set the book's fiscal dates. Omitted/blank arguments keep their current value.

    working_year_end  YYYY-MM-DD — the year-end being WORKED ON (what the
                      statements report). Its year and month-day drive the
                      current and comparative periods.
    ceiling_mode      'cy'   posting stops at that year-end
                      'next' posting stays open one year past it, for a file
                             that needs following-year transactions while this
                             year-end is still being worked on.
    lock_date         YYYY-MM-DD, or the string 'none' to clear it.

    The ceiling is DERIVED from these — it is never keyed directly. It will not
    be set behind postings that already exist; the call is refused instead."""
    _init(db_path)
    lock = None if not lock_date else ("" if lock_date.strip().lower() == "none" else lock_date)
    return models.set_fiscal_settings(
        company_name=company_name or None,
        working_ye=working_year_end or None,
        ceiling_mode=ceiling_mode or None,
        lock_date=lock)


@mcp.tool()
def new_aje_year(db_path: str, account_name: str = "", description: str = "") -> dict:
    """Set up a year of adjusting entries — the journal account this year's AJEs
    post against, placed on the AJE report.

    Both arguments are optional: left blank, Grid suggests them from the fiscal
    year being worked on (26AJE / "2026 Adjusting Entries"). Pass either to
    override. Call this ONCE per year; posting an AJE needs the journal to exist."""
    _init(db_path)
    sug = models.suggest_aje_batch()
    acct = models.create_aje_batch(account_name or sug['account'],
                                   description or sug['description'])
    return {"account": acct["name"], "description": acct["description"],
            "account_id": acct["id"], "next_ref": models.next_aje_ref(acct["id"])}


@mcp.tool()
def list_aje_years(db_path: str) -> list:
    """The years of adjusting entries on file, newest reference first."""
    _init(db_path)
    out = []
    for j in models.aje_journals():
        out.append({"account": j["name"], "description": j["description"],
                    "entries": len(models.aje_groups(j["id"])),
                    "next_ref": models.next_aje_ref(j["id"]),
                    "journal_balance_cents": models.get_account_balance(j["id"])})
    return out


@mcp.tool()
def list_aje(db_path: str, journal: str) -> list:
    """Every adjustment in one year, grouped as an accountant reads them:
    one block per reference, carrying its legs. `journal` is the account name
    (e.g. '26AJE')."""
    _init(db_path)
    acct = models.get_account_by_name(journal)
    if not acct or not models.is_aje_journal(acct["id"]):
        raise ValueError(f"'{journal}' is not a year of adjusting entries. "
                         f"Use list_aje_years to see what is on file.")
    out = []
    for g in models.aje_groups(acct["id"]):
        out.append({"ref": g["ref"], "date": g["date"], "description": g["description"],
                    "balanced": g["balanced"],
                    "lines": [{"account": l["account"],
                               "amount": models.fmt_amount_plain(l["amount"]),
                               "amount_cents": l["amount"]} for l in g["lines"]]})
    return out


@mcp.tool()
def post_aje(db_path: str, journal: str, description: str, lines: list[dict],
             entry_date: str = "", ref: str = "", replace_ref: str = "") -> dict:
    """Post ONE adjusting entry — the only supported way to enter an AJE.
    Do not hand-build one with post_transaction.

    Each leg becomes its own 2-line transaction: the account being adjusted, and
    the year journal on the other side. All legs share the reference, the date and
    the description, so the journal nets to zero and the ledger and reports read
    it as ordinary postings — which is what it is.

    Args:
        journal: the year's journal account (e.g. '26AJE'). new_aje_year makes it.
        description: what the adjustment IS. It prints on EVERY leg, so it is
            capped at 50 characters — "Rcd 2026 tax provision", not a paragraph.
        lines: [{"account": "EX.TAX", "amount": "100.00"},
                {"account": "GST.PAY", "amount": "(80.00)"}, ...]
            amount is a DOLLAR string signed the way Grid displays: a plain number
            is a DEBIT, "(80.00)" or "-80.00" is a CREDIT. Must balance.
            NEVER include the journal account itself — Grid posts that side.
        entry_date: defaults to the fiscal year end being worked on.
        ref: defaults to the next in this journal's sequence (26AJE01, 26AJE02, …).
        replace_ref: an existing reference to re-post over, atomically. Use this
            to EDIT an entry — the old legs and the new ones swap in one step.
    """
    _init(db_path)
    acct = models.get_account_by_name(journal)
    if not acct or not models.is_aje_journal(acct["id"]):
        raise ValueError(f"'{journal}' is not a year of adjusting entries. "
                         f"Create it with new_aje_year, or list_aje_years to see what exists.")
    rows = []
    for i, l in enumerate(lines, start=1):
        if not isinstance(l, dict):
            raise ValueError(f"lines[{i}] must be an object with account/amount")
        amt = str(l.get("amount", "")).strip()
        try:
            cents = models.parse_amount(amt) if amt else 0
        except ValueError:
            raise ValueError(f"lines[{i}] ({l.get('account')}): '{amt}' is not an amount")
        rows.append({"account": str(l.get("account", "")).strip(), "amount": cents})

    anchor = models.fiscal_anchor()
    d = entry_date or (anchor["cy_end"] if anchor else "")
    v = models.post_aje(acct["id"], ref or models.next_aje_ref(acct["id"]), d,
                        description, rows, replace_ref=replace_ref or None)
    return {"ref": v["ref"], "date": d, "description": v["description"],
            "legs_posted": len(v["txn_ids"]), "batch": v["batch"],
            "replaced_lines": v["replaced"],
            "debits": models.fmt_amount_plain(v["debit_cents"]),
            "credits": models.fmt_amount_plain(v["credit_cents"]),
            "journal_balance_cents": models.get_account_balance(acct["id"]),
            "next_ref": models.next_aje_ref(acct["id"])}


@mcp.tool()
def delete_aje(db_path: str, journal: str, ref: str) -> dict:
    """Remove one adjustment entirely — every leg of it, or none."""
    _init(db_path)
    acct = models.get_account_by_name(journal)
    if not acct or not models.is_aje_journal(acct["id"]):
        raise ValueError(f"'{journal}' is not a year of adjusting entries.")
    n = models.delete_aje(acct["id"], ref)
    return {"ref": ref, "lines_removed": n,
            "journal_balance_cents": models.get_account_balance(acct["id"])}


@mcp.tool()
def set_ref_mark(db_path: str, report_name: str, account_name: str, mark: str) -> dict:
    """Set/clear the working-paper index mark (the red pencil: E-1, B-2.1) shown
    beside a line on a report. Empty mark removes it. The line is located by
    report name (e.g. 'BS') + account name."""
    _init(db_path)
    rpt = models.find_report_by_name(report_name)
    if not rpt:
        raise ValueError(f"Report not found: {report_name}")
    acct = models.get_account_by_name(account_name)
    if not acct:
        raise ValueError(f"Account not found: {account_name}")
    with models.get_db() as db:
        row = db.execute("SELECT id FROM report_items WHERE report_id=? AND account_id=?",
                         (rpt["id"], acct["id"])).fetchone()
    if not row:
        raise ValueError(f"{account_name} is not a line on {report_name}")
    models.set_ref_mark(row["id"], mark)
    return {"report": report_name, "account": account_name, "mark": (mark or "").strip()}


@mcp.tool()
def set_leadsheet(db_path: str, account_name: str, code: str) -> dict:
    """Assign an account to a working-paper lead sheet (codes like A, B-1, E-1;
    empty clears). Mirrors GIFI mapping — a practice-convention grouping."""
    _init(db_path)
    models.set_leadsheet(account_name, code)
    return {"account": account_name, "leadsheet": (code or "").strip()}


@mcp.tool()
def leadsheet_report(db_path: str, code: str = "") -> dict:
    """Lead-sheet data. With code: that sheet's accounts with CY/PY/$chg/%chg
    (BS accounts = closing balances, IS accounts = fiscal-year activity; dollars).
    Without code: the index of all lead-sheet codes with account counts."""
    _init(db_path)
    if not code.strip():
        return {"sheets": [{"code": r["code"], "accounts": r["n"]}
                           for r in models.leadsheet_index()]}
    d = models.leadsheet_data(code.strip())
    return {"code": d["code"],
            "rows": [{"account": r["name"], "description": r["description"],
                      "statement": r["type"], "cy": r["cy"] / 100, "py": r["py"] / 100,
                      "chg": r["chg"] / 100,
                      "pct": round(r["pct"], 1) if r["pct"] is not None else None}
                     for r in d["rows"]],
            "total_cy": d["cy"] / 100, "total_py": d["py"] / 100,
            "total_chg": d["chg"] / 100}


@mcp.tool()
def add_workpaper(db_path: str, ref: str, description: str = "",
                  path: str = "", fy: str = "", folder: str = "") -> dict:
    """Add a working paper to the index: ref (E-1, B-2.1), description, and an
    optional path RELATIVE to the client folder (e.g. 'WPFdocs/E-1.0.pdf').
    Blank path = placeholder to be linked later. `folder` = folder NAME in the
    Engagement File tree (Assets / Liabilities and Equity / Income Statement /
    custom); blank or unknown → the Engagement File root. Grid stores links
    only — the document itself stays in the client folder."""
    _init(db_path)
    fid = None
    if folder.strip():
        for f in models.list_wp_folders():
            if f["name"].strip().lower() == folder.strip().lower():
                fid = f["id"]
                break
    wp_id = models.add_workpaper(ref, description, path, fy or None, fid)
    return {"id": wp_id, "ref": ref.strip(), "fy": models.wp_fy(fy or None),
            "folder": folder.strip() or "Engagement File"}


@mcp.tool()
def sign_workpaper(db_path: str, ref: str, role: str, initials: str,
                   fy: str = "") -> dict:
    """Stamp a working paper: role 'prep' (preparer) or 'rev' (reviewer),
    with the signer's initials — e.g. an agent signs prep as its own name.
    Stamps '<INITIALS> <today>'. Empty initials raises."""
    _init(db_path)
    if role not in ("prep", "rev"):
        raise ValueError("role must be 'prep' or 'rev'")
    if not initials.strip():
        raise ValueError("initials required")
    fy = models.wp_fy(fy or None)
    with models.get_db() as db:
        row = db.execute("SELECT id FROM workpapers WHERE fy=? AND ref=? COLLATE NOCASE",
                         (fy, ref.strip())).fetchone()
    if not row:
        raise ValueError(f"Working paper {ref} not found for FY{fy}")
    stamp = f"{initials.strip().upper()} {datetime.now().strftime('%Y-%m-%d')}"
    models.update_workpaper(row["id"], "prep_by" if role == "prep" else "rev_by", stamp)
    return {"ref": ref.strip(), "role": role, "stamp": stamp}


@mcp.tool()
def workpaper_index(db_path: str, fy: str = "") -> dict:
    """The working-paper file: all papers for the year (ref, description, path,
    file_exists, prep/rev stamps, to_print) PLUS the completeness check —
    statement marks without a paper, papers without a present file, unindexed
    document files in the client folder, unprepared papers. An empty check =
    the file is tied together, old-school style."""
    _init(db_path)
    papers = models.list_workpapers(fy or None)
    return {"fy": models.wp_fy(fy or None),
            "papers": [{k: p[k] for k in ("ref", "description", "path",
                        "file_exists", "prep_by", "rev_by", "to_print")} for p in papers],
            "check": models.verify_workpapers(fy or None)}


@mcp.tool()
def backup_books(db_path: str) -> dict:
    """Snapshot the books to backups/<name>.<stamp>.db — a checkpointed,
    integrity-verified, single-file copy (restore = open or copy back a
    snapshot). A daily snapshot also happens automatically at every open;
    call this before risky bulk operations for an extra restore point."""
    _init(db_path)
    path = models.backup_books(force=True)
    snaps = models.list_backups()
    return {"backup": path, "snapshots_on_file": len(snaps),
            "keep_newest": models.BACKUP_KEEP}


# ═══════════════════════════════════════════════════════════════════
# WRITE TOOLS
# ═══════════════════════════════════════════════════════════════════

@mcp.tool()
def post_transaction(
    db_path: str,
    date: str,
    description: str,
    amount: str,
    debit_account: str,
    credit_account: str,
    reference: str = "",
) -> dict:
    """Post a simple 2-line transaction. Amount is in dollars (e.g. '1500.00'). Date is YYYY-MM-DD.
    Optional reference (e.g. 'OPEN' for opening balances). Auto-generated if blank."""
    _init(db_path)
    dr_acct = models.get_account_by_name(debit_account)
    if not dr_acct:
        raise ValueError(f"Debit account not found: {debit_account}")
    cr_acct = models.get_account_by_name(credit_account)
    if not cr_acct:
        raise ValueError(f"Credit account not found: {credit_account}")
    amount_cents = models.parse_amount(amount)
    if amount_cents <= 0:
        raise ValueError(f"Amount must be positive: {amount}")
    ref = reference if reference else models.generate_ref()
    txn_id = models.add_simple_transaction(
        date, ref, description, dr_acct["id"], cr_acct["id"], amount_cents
    )
    return {"txn_id": txn_id, "reference": ref}


@mcp.tool()
def delete_transaction(db_path: str, txn_id: int) -> dict:
    """Delete a transaction by ID. Respects lock date."""
    _init(db_path)
    models.delete_transaction(txn_id)
    return {"deleted": True, "txn_id": txn_id}


@mcp.tool()
def add_account(
    db_path: str,
    name: str,
    normal_balance: str,
    description: str = "",
    report_name: str = "",
    after_account: str = "",
    total_to: str = "",
) -> dict:
    """Add a new posting account. normal_balance is 'D' (debit-normal) or 'C' (credit-normal).

    Optionally place it on a report in the same call by providing report_name.
    Use after_account and/or total_to to control positioning (see add_to_report).
    """
    _init(db_path)
    if normal_balance not in ("D", "C"):
        raise ValueError("normal_balance must be 'D' or 'C'")
    if not report_name:
        raise ValueError(
            "Posting accounts must be placed on a report. "
            "Specify report_name (e.g. 'BS' or 'IS')."
        )
    account_id = models.add_account(name, normal_balance, description)
    result = {"account_id": account_id, "name": name}

    if report_name:
        placement = _place_on_report(report_name, name, after_account, total_to,
                                     indent=2, description="")
        result["report_item_id"] = placement["item_id"]
        result["report"] = placement["report"]
        result["position"] = placement["position"]

    return result


@mcp.tool()
def add_to_report(
    db_path: str,
    report_name: str,
    account_name: str,
    after_account: str = "",
    total_to: str = "",
    indent: int = 2,
    description: str = "",
) -> dict:
    """Place an existing account on a report (BS, IS, etc.).

    Positioning (in priority order):
      1. after_account — insert after this account on the report
      2. total_to — insert after the last item that rolls up to this total
      3. Neither — append to end of report

    Args:
        report_name: Report name ('BS', 'IS', etc.)
        account_name: Account to place on the report
        after_account: Insert after this account on the report
        total_to: Account name this item rolls up into (sets total_to_1)
        indent: Indentation level (default 2)
        description: Override display text (default: uses account description)
    """
    _init(db_path)
    return _place_on_report(report_name, account_name, after_account, total_to,
                            indent, description)


def _place_on_report(report_name, account_name, after_account, total_to,
                     indent, description):
    """Shared logic for placing an account on a report."""
    report = models.find_report_by_name(report_name)
    if not report:
        raise ValueError(f"Report not found: {report_name}")

    acct = models.get_account_by_name(account_name)
    if not acct:
        raise ValueError(f"Account not found: {account_name}")

    report_id = report["id"]
    items = models.get_report_items(report_id)
    position = None
    total_to_1 = total_to

    if after_account:
        # Find the after_account's position on this report
        found = False
        for item in items:
            if item["acct_name"] == after_account:
                position = item["position"] + 5
                found = True
                break
        if not found:
            raise ValueError(
                f"after_account '{after_account}' not found on report {report_name}")
    elif total_to:
        # Verify total_to account exists
        tt_acct = models.get_account_by_name(total_to)
        if not tt_acct:
            raise ValueError(f"total_to account not found: {total_to}")
        # Find the last item whose total_to_1 matches
        last_pos = None
        for item in items:
            if item["total_to_1"] == total_to:
                last_pos = item["position"]
        if last_pos is not None:
            position = last_pos + 5
        # else: fall through to append (position=None → add_report_item appends)

    item_id = models.add_report_item(
        report_id, "account", description, acct["id"],
        indent=indent, position=position, total_to_1=total_to_1,
    )

    return {
        "item_id": item_id,
        "report": report_name,
        "account": account_name,
        "position": position or "appended",
    }


@mcp.tool()
def bulk_report_layout(
    db_path: str,
    report_name: str,
    items: list[dict],
    after_account: str = "",
    mode: str = "append",
) -> dict:
    """Place a batch of items on a report in one call, with correct ordering guaranteed.

    Use this after importing a trial balance to lay out 20-60 accounts on the BS/IS
    in the correct sections with proper positioning, indentation, and total_to wiring.

    Args:
        report_name: Target report (BS, IS, etc.)
        items: Ordered list of items. Each dict can have:
            - account_name (required for account/total items)
            - item_type: "account" (default), "total", "label", "separator"
            - indent: indentation level (default 2)
            - total_to: account name this item rolls up into (sets total_to_1)
            - description: override display text
            - sep_style: for separators — "single", "double", "blank"
        after_account: Anchor — insert the batch after this account on the report.
            If empty, appends to end (or starts at 10 for replace mode).
        mode: "append" (default) adds to existing items;
            "replace" clears all existing items first then inserts.
    """
    _init(db_path)
    if mode not in ("append", "replace"):
        raise ValueError(f"Invalid mode: '{mode}'. Must be 'append' or 'replace'.")

    report = models.find_report_by_name(report_name)
    if not report:
        raise ValueError(f"Report not found: {report_name}")
    report_id = report["id"]

    # Replace mode: save cross-report chain links before wiping
    saved_chain = {}  # {account_name: total_to_1}
    if mode == "replace":
        existing = models.get_report_items(report_id)
        for item in existing:
            if item["acct_name"] and item["total_to_1"]:
                saved_chain[item["acct_name"]] = str(item["total_to_1"])
        models.clear_report_items(report_id)

    # Determine starting position
    if after_account:
        existing = models.get_report_items(report_id)
        anchor_pos = None
        for item in existing:
            if item["acct_name"] == after_account:
                anchor_pos = item["position"]
                break
        if anchor_pos is None:
            raise ValueError(
                f"after_account '{after_account}' not found on report {report_name}")
        position = anchor_pos + 5
    elif mode == "replace":
        position = 10
    else:
        # Append: start after current max
        existing = models.get_report_items(report_id)
        max_pos = max((item["position"] for item in existing), default=0)
        position = max_pos + 10

    placed = 0
    skipped = 0
    errors = []
    # Track which accounts the caller explicitly provided total_to for
    caller_specified_tt = set()

    def _resolve_total_to(val):
        """Ensure total_to is always an account name string, never an integer ID."""
        if not val:
            return ""
        val = str(val)
        # If it looks like a bare integer, resolve it to an account name
        if val.isdigit():
            with models.get_db() as lookup_db:
                row = lookup_db.execute(
                    "SELECT name FROM accounts WHERE id=?", (int(val),)
                ).fetchone()
                if row:
                    return row["name"]
        return val

    with models.get_db() as db:
        for idx, spec in enumerate(items):
            item_type = spec.get("item_type", "account")
            indent = spec.get("indent", 2)
            total_to_1 = _resolve_total_to(spec.get("total_to", ""))
            description = spec.get("description", "")
            sep_style = spec.get("sep_style", "")
            account_name = spec.get("account_name", "")

            # Validate item_type
            if item_type not in ("account", "total", "label", "separator"):
                errors.append({"index": idx, "reason": f"Invalid item_type: '{item_type}'"})
                skipped += 1
                continue

            # Labels and separators don't need an account
            if item_type in ("label", "separator"):
                db.execute(
                    "INSERT INTO report_items(report_id, position, item_type, description, "
                    "account_id, indent, total_to_1, sep_style) VALUES(?,?,?,?,?,?,?,?)",
                    (report_id, position, item_type, description, None, indent,
                     total_to_1, sep_style))
                position += 10
                placed += 1
                continue

            # Account and total types require account_name
            if not account_name:
                errors.append({"index": idx, "reason": f"Missing account_name for {item_type} item"})
                skipped += 1
                continue

            acct = models.get_account_by_name(account_name)
            if not acct:
                errors.append({"index": idx, "reason": f"Account not found: {account_name}"})
                skipped += 1
                continue

            if total_to_1:
                caller_specified_tt.add(account_name)

            db.execute(
                "INSERT INTO report_items(report_id, position, item_type, description, "
                "account_id, indent, total_to_1, sep_style) VALUES(?,?,?,?,?,?,?,?)",
                (report_id, position, item_type, description, acct["id"], indent,
                 total_to_1, sep_style))
            position += 10
            placed += 1

        # Restore cross-report chain links that the caller didn't explicitly provide.
        # This prevents replace mode from silently breaking links like
        # NETINC→NI or NI→RE that wire the IS net income to the BS.
        if mode == "replace" and saved_chain:
            for acct_name, old_tt in saved_chain.items():
                if acct_name in caller_specified_tt:
                    continue  # caller explicitly set this one, don't override
                old_tt = _resolve_total_to(old_tt)
                if old_tt:
                    db.execute(
                        "UPDATE report_items SET total_to_1=? "
                        "WHERE report_id=? AND account_id=(SELECT id FROM accounts WHERE name=?) "
                        "AND (total_to_1 IS NULL OR total_to_1='')",
                        (old_tt, report_id, acct_name))

        # Resequence once at the end
        models._resequence(db, report_id)

    return {
        "report": report_name,
        "placed": placed,
        "skipped": skipped,
        "errors": errors,
        "mode": mode,
    }


@mcp.tool()
def add_rule(
    db_path: str,
    keyword: str,
    account_name: str,
    tax_code: str = "",
    priority: int = 0,
) -> dict:
    """Add a CSV import rule. Transactions matching keyword are auto-posted to account_name."""
    _init(db_path)
    # Verify the target account exists
    acct = models.get_account_by_name(account_name)
    if not acct:
        raise ValueError(f"Account not found: {account_name}")
    models.save_import_rule(None, keyword, account_name, tax_code, priority)
    # Retrieve the newly created rule to get its ID
    rules = models.get_import_rules()
    rule_id = None
    for r in rules:
        if r["keyword"] == keyword and r["account_name"] == account_name:
            rule_id = r["id"]
            break
    return {"rule_id": rule_id}


@mcp.tool()
def delete_rule(db_path: str, rule_id: int) -> dict:
    """Delete an import rule by ID."""
    _init(db_path)
    models.delete_import_rule(rule_id)
    return {"deleted": True, "rule_id": rule_id}



@mcp.tool()
def post_opening_balances(
    db_path: str,
    conversion_date: str,
    balances: list[dict],
    expected_retained_earnings: str = "",
) -> dict:
    """Post a client's opening balances — the ONLY supported way to convert a set
    of books onto Grid. Do not hand-post a conversion with post_transaction.

    Give it the trial balance and it does the rest: creates TRX.OPEN on the TRX
    report, posts one 2-line entry per account against it (same date, reference
    OPEN), works out retained earnings as the residual and posts it to RE.OB, and
    leaves TRX.OPEN at zero. The whole conversion lands as ONE atomic batch — any
    bad row and nothing posts at all.

    NEVER put retained earnings in `balances`. It is COMPUTED. RE, RE.OB, RE.OPEN,
    RE.CLOSE and EX.SUSP are all refused, and so is any account not in the chart.

    Args:
        conversion_date: the balances are as at this date — normally the prior
            fiscal year end (YYYY-MM-DD).
        balances: [{"account": "BANK.CHQ", "description": "Chequing",
                    "amount": "33268.00"}, ...]
            amount is a DOLLAR string, signed the way Grid displays: a plain
            number is a DEBIT, "(1000.00)" or "-1000.00" is a CREDIT.
        expected_retained_earnings: optional check figure in dollars, as a credit
            (the normal case). Given it, the post REFUSES unless the computed
            residual ties to it — use it whenever the prior-year statements are
            in front of you.

    Balance sheet only  → the residual is retained earnings BROUGHT FORWARD.
    Full trial balance (income-statement accounts included, which is what gives
    the IS its prior-year comparative) → the residual is retained earnings at the
    START of that year (T2 line 3660, not 3600). Either is correct; the return
    value says which one you just posted.
    """
    _init(db_path)
    rows = []
    for i, b in enumerate(balances, start=1):
        if not isinstance(b, dict):
            raise ValueError(f"balances[{i}] must be an object with account/description/amount")
        amt = str(b.get("amount", "")).strip()
        try:
            cents = models.parse_amount(amt) if amt else 0
        except ValueError:
            raise ValueError(f"balances[{i}] ({b.get('account')}): '{amt}' is not an amount")
        rows.append({"account": str(b.get("account", "")).strip(),
                     "description": str(b.get("description", "")).strip(),
                     "amount": cents})
    expected = None
    if str(expected_retained_earnings).strip():
        expected = models.parse_amount(expected_retained_earnings)

    res = models.post_opening_balances(conversion_date, rows, expected)
    return {
        "batch": res["batch"],
        "entries_posted": len(res["txn_ids"]),
        "conversion_date": conversion_date,
        "debits": models.fmt_amount_plain(res["debit_cents"]),
        "credits": models.fmt_amount_plain(res["credit_cents"]),
        "retained_earnings_credit": models.fmt_amount_plain(res["re_credit_cents"]),
        "what_that_figure_is": res["re_note"],
        "conversion_account_balance": models.fmt_amount_plain(
            models.get_account_balance(models.get_account_by_name(models.CONVERSION_ACCT)["id"])),
    }


@mcp.tool()
def opening_balances_status(db_path: str) -> dict:
    """Where a set of books stands on opening balances: needed (nothing posted at
    all), posted, declined (brand-new client starting at zero), or later. Check
    this before converting — Grid refuses to layer a second conversion on a first."""
    _init(db_path)
    st = models.openings_state()
    if st["status"] == "posted":
        ob = models.get_account_by_name(models.OPENING_RE_ACCT)
        st["retained_earnings_credit"] = models.fmt_amount_plain(
            -models.get_account_balance(ob["id"]) if ob else 0)
    return st


@mcp.tool()
def delete_opening_balances(db_path: str) -> dict:
    """Delete an entire conversion so it can be re-entered. Grid does not adjust a
    wrong conversion — it deletes and redoes it."""
    _init(db_path)
    n = models.delete_opening_balances()
    return {"deleted": n, "status": models.openings_state()["status"]}

@mcp.tool()
def import_tb(db_path: str, csv_path: str, date: str = "2025-01-01") -> dict:
    """Import a trial balance CSV and post as one compound journal entry.

    CSV format: Account, Description, Debit, Credit.
    For each row, creates the account if it doesn't already exist
    (debit-balance accounts get normal_balance='D', credit get 'C').
    Posts one balanced compound journal entry dated as of the date parameter.
    The entry must balance — raises an error if total debits != total credits.

    NOTE: For opening balances (PY closing TB), do NOT use this tool.
    Use post_opening_balances — it creates TRX.OPEN, computes retained earnings,
    and posts the whole conversion atomically.
    """
    _init(db_path)
    csv_path = _check_path(csv_path, "csv_path")
    if not os.path.exists(csv_path):
        raise ValueError(f"File not found: {csv_path}")

    date = _normalize_date(date)
    if not date:
        raise ValueError("Invalid date. Use YYYY-MM-DD format.")

    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        rows_raw = list(reader)

    if not rows_raw:
        raise ValueError("Empty CSV file")

    # Detect header row
    header = [h.strip().lower() for h in rows_raw[0]]
    has_header = "account" in header or "description" in header
    start = 1 if has_header else 0

    accounts_created = []
    lines = []

    for row in rows_raw[start:]:
        if len(row) < 4:
            continue

        acct_name = row[0].strip()
        description = row[1].strip()
        dr_str = row[2].strip()
        cr_str = row[3].strip()

        if not acct_name or not description:
            continue
        # Skip total rows
        if acct_name.upper() in ("TOTALS", "TOTAL"):
            continue
        if description.upper() in ("TOTALS", "TOTAL"):
            continue

        dr_cents = models.parse_amount(dr_str) if dr_str else 0
        cr_cents = models.parse_amount(cr_str) if cr_str else 0

        if dr_cents == 0 and cr_cents == 0:
            continue

        # Determine normal balance and line amount
        if dr_cents > 0:
            normal = "D"
            amount_cents = dr_cents       # positive = debit
        else:
            normal = "C"
            amount_cents = -cr_cents      # negative = credit

        # Find or create account
        acct = models.get_account_by_name(acct_name)
        if not acct:
            acct_id = models.add_account(acct_name, normal, description)
            accounts_created.append(acct_name)
        else:
            acct_id = acct["id"]

        lines.append((acct_id, amount_cents, description))

    if not lines:
        raise ValueError("No valid data rows found in CSV")

    # Verify the entry balances
    total = sum(l[1] for l in lines)
    if total != 0:
        raise ValueError(
            f"Trial balance does not balance: off by {total / 100:.2f}"
        )

    ref = models.generate_ref()
    txn_id = models.add_transaction(date, ref, "Opening balances", lines)

    return {
        "txn_id": txn_id,
        "reference": ref,
        "date": date,
        "lines_posted": len(lines),
        "accounts_created": len(accounts_created),
        "new_accounts": accounts_created,
    }


@mcp.tool()
def import_csv(db_path: str, csv_path: str, bank_account: str) -> dict:
    """Import a bank CSV file into the books. Applies import rules to auto-categorize.

    CSV format: Date, Description, Amount (or Date, Description, Debit, Credit).
    Also handles multi-column bank exports (auto-detects columns).
    Positive amounts = deposits, negative = payments.
    Unmatched rows go to EX.SUSP (suspense). Review with get_ledger('EX.SUSP').
    """
    _init(db_path)

    csv_path = _check_path(csv_path, "csv_path")
    if not os.path.exists(csv_path):
        raise ValueError(f"File not found: {csv_path}")

    bank_acct = models.get_account_by_name(bank_account)
    if not bank_acct:
        raise ValueError(f"Bank account not found: {bank_account}")

    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        rows_raw = list(reader)

    if not rows_raw:
        raise ValueError("Empty CSV file")

    has_header, data_rows, csv_repairs = _normalize_csv(rows_raw)

    if not data_rows:
        raise ValueError("No data rows in CSV (only a header)")

    # Build row dicts for import_rows: parse amounts from CSV columns
    import_data = []
    parse_errors = []
    for row_num, row in enumerate(data_rows, start=2 if has_header else 1):
        if len(row) < 3:
            parse_errors.append({"row": row_num, "reason": "Too few columns"})
            continue

        row_date = row[0].strip()
        row_desc = row[1].strip()

        if not row_desc:
            parse_errors.append({"row": row_num, "reason": "Missing description"})
            continue

        if len(row) >= 4 and (row[2].strip() or row[3].strip()):
            try:
                dr = models.parse_amount(row[2]) if row[2].strip() else 0
                cr = models.parse_amount(row[3]) if row[3].strip() else 0
                amount_cents = dr - cr
            except Exception:
                parse_errors.append({"row": row_num, "reason": f"Bad amount '{row[2].strip()}/{row[3].strip()}'"})
                continue
        else:
            try:
                amount_cents = models.parse_amount(row[2])
            except Exception:
                parse_errors.append({"row": row_num, "reason": f"Bad amount '{row[2].strip()}'"})
                continue

        import_data.append({
            'date': row_date,
            'description': row_desc,
            'amount_cents': amount_cents,
        })

    result = models.import_rows(bank_acct['id'], import_data)
    result['rows_processed'] = len(data_rows)
    result['skipped'] = result['skipped'] + len(parse_errors)
    if csv_repairs:
        result["rows_repaired"] = len(csv_repairs)
    if parse_errors:
        all_errors = parse_errors + result.get('errors', [])
        result['errors'] = all_errors[:20]
    return _import_landed(result, 'import_csv')


@mcp.tool()
def import_ofx(db_path: str, ofx_path: str, bank_account: str) -> dict:
    """Import a bank OFX/QBO file. Applies import rules to auto-categorize.

    OFX/QBO files are standard bank download formats. Supports both XML-based
    and SGML-based OFX files. Positive amounts = deposits, negative = payments.
    Unmatched rows go to EX.SUSP (suspense). Review with get_ledger('EX.SUSP').
    """
    _init(db_path)

    ofx_path = _check_path(ofx_path, "ofx_path")
    if not os.path.exists(ofx_path):
        raise ValueError(f"File not found: {ofx_path}")

    bank_acct = models.get_account_by_name(bank_account)
    if not bank_acct:
        raise ValueError(f"Bank account not found: {bank_account}")

    rows = models.parse_ofx(ofx_path)
    return _import_landed(models.import_rows(bank_acct['id'], rows), 'import_ofx')


@mcp.tool()
def import_gl(db_path: str, csv_path: str, bank_account: str) -> dict:
    """Import a pre-categorized general ledger CSV. Cross-accounts are specified per row.

    Use this when converting from another accounting system (a legacy GL, QuickBooks GL,
    Sage, etc.) where every transaction already has its cross-account known.
    No import rules are applied — the cross-account column is used directly.

    CSV format: Date, Description, Amount, CrossAccount.
    Positive amounts = debits to the primary account, negative = credits.
    All cross-accounts must already exist in the chart of accounts.
    """
    _init(db_path)

    csv_path = _check_path(csv_path, "csv_path")
    if not os.path.exists(csv_path):
        raise ValueError(f"File not found: {csv_path}")

    bank_acct = models.get_account_by_name(bank_account)
    if not bank_acct:
        raise ValueError(f"Account not found: {bank_account}")

    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        rows_raw = list(reader)

    if not rows_raw:
        raise ValueError("Empty CSV file")

    # Detect header row
    first = rows_raw[0]
    has_header = any(h.strip().lower() in ('date', 'description', 'amount', 'crossaccount', 'cross_account')
                     for h in first)
    data_rows = rows_raw[1:] if has_header else rows_raw

    if not data_rows:
        raise ValueError("No data rows in CSV (only a header)")

    import_data = []
    parse_errors = []
    for row_num, row in enumerate(data_rows, start=2 if has_header else 1):
        if len(row) < 4:
            parse_errors.append({"row": row_num, "reason": f"Need 4 columns, got {len(row)}"})
            continue

        row_date = row[0].strip()
        row_desc = row[1].strip()
        cross_acct = row[3].strip()

        if not row_desc:
            parse_errors.append({"row": row_num, "reason": "Missing description"})
            continue

        if not cross_acct:
            parse_errors.append({"row": row_num, "reason": "Missing cross-account"})
            continue

        try:
            amount_cents = models.parse_amount(row[2])
        except Exception:
            parse_errors.append({"row": row_num, "reason": f"Bad amount '{row[2].strip()}'"})
            continue

        import_data.append({
            'date': row_date,
            'description': row_desc,
            'amount_cents': amount_cents,
            'cross_account': cross_acct,
        })

    result = models.import_gl_rows(bank_acct['id'], import_data)
    result['rows_processed'] = len(data_rows)
    result['skipped'] = result['skipped'] + len(parse_errors)
    if parse_errors:
        all_errors = parse_errors + result.get('errors', [])
        result['errors'] = all_errors[:20]
    return _import_landed(result, 'import_gl')


@mcp.tool()
def reclassify_suspense(db_path: str, txn_id: int, target_account: str, tax_code: str = '') -> dict:
    """Reclassify a suspense (EX.SUSP) transaction to the correct expense/revenue account.

    Swaps the EX.SUSP line for the target account in-place. If a tax_code is
    provided (e.g. 'G5'), splits the amount into net + tax components.
    Automatically learns an import rule from the transaction description
    if the vendor name is specific enough and no matching rule exists.

    txn_id: the transaction to reclassify (must have an EX.SUSP line).
    target_account: account name to reclassify to (e.g. 'EX.OFFICE').
    tax_code: optional tax code to apply (e.g. 'G5', 'H13'). Empty = no tax.
    """
    _init(db_path)
    return models.reclassify_suspense(txn_id, target_account, tax_code)




@mcp.tool()
def trace_account(db_path: str, account_name: str,
                  date_from: str = "", date_to: str = "") -> dict:
    """Trace the full accumulation tree for a report account.

    Shows what feeds into the account total, with amounts and sources.
    Useful for understanding why a total shows a particular number.

    Args:
        account_name: Account name to trace (e.g. RE, NETEARN, TOTREV)
        date_from: Optional start date (YYYY-MM-DD) for date range
        date_to: Optional end date (YYYY-MM-DD) for as-of or range

    Returns dict with: name, display, contributors list, feeds_into list.
    """
    _init(db_path)
    df = _normalize_date(date_from) if date_from else None
    dt = _normalize_date(date_to) if date_to else None
    return models.trace_account(account_name, df, dt)


@mcp.tool()
def set_lock_date(db_path: str, lock_date: str = "") -> dict:
    """Show or set the lock date. Transactions on or before the lock date cannot be posted, edited, or deleted.

    If lock_date is provided (YYYY-MM-DD), sets it. If omitted, returns the current lock date.
    """
    _init(db_path)

    if lock_date:
        normalized = _normalize_date(lock_date)
        if not normalized:
            raise ValueError(f"Invalid date: '{lock_date}'. Use YYYY-MM-DD format.")
        models.set_fiscal_settings(lock_date=normalized)   # guarded: refuses a lock past the ceiling
        return {"lock_date": normalized, "message": f"Lock date set to {normalized}"}
    else:
        current = models.get_meta("lock_date", "")
        return {"lock_date": current or None}


@mcp.tool()
def setup_detailed_ar(db_path: str) -> dict:
    """Setup a Detailed Accounts Receivable subledger report.

    Creates an AR report on the home screen with 3 sample client accounts
    (Gretzky, Lemieux, Orr), a total account (ARDET), and links it to the
    Balance Sheet via AR.DET. The cross-report total chain flows:
    R. client accounts → ARDET → AR.DET (on BS) → AR.TOT → CA → TA.

    Run once per set of books. Returns an error if the AR report already exists.
    After setup, add real clients with add_account (e.g. R.SMIJOH, D, "Smith, John")
    and place them on the AR report with total_to_1 set to ARDET.
    """
    _init(db_path)
    result = models.setup_detailed_ar()
    return {"success": True, "message": result}


@mcp.tool()
def setup_detailed_ap(db_path: str) -> dict:
    """Setup a Detailed Accounts Payable subledger report.

    Creates an AP.SUB report on the home screen with 3 sample vendor accounts
    (Bauer, CCM, Warrior), a total account (APDET), and links it to the
    Balance Sheet via AP.DET → AP.TOT. Also restructures the BS to add an
    AP.TOT subtotal for all payable accounts. The cross-report total chain flows:
    P. vendor accounts → APDET → AP.DET (on BS) → AP.TOT → CL → TL.

    Run once per set of books. Returns an error if the AP.SUB report already exists.
    After setup, add real vendors with add_account (e.g. P.SMISUP, C, "Smith, Supply")
    and place them on the AP.SUB report with total_to_1 set to APDET.
    """
    _init(db_path)
    result = models.setup_detailed_ap()
    return {"success": True, "message": result}


# ═══════════════════════════════════════════════════════════════════
# HELPERS (CSV import)
# ═══════════════════════════════════════════════════════════════════

def _normalize_date(s):
    """Try to normalize a date string to YYYY-MM-DD."""
    return models.normalize_date(s)


def _normalize_csv(rows_raw):
    """Delegates to models.normalize_csv — ONE csv normalizer for all interfaces
    (content-based header detection, extra-field repair, signed Debit/Credit
    netting on multi-column exports)."""
    return models.normalize_csv(rows_raw)

# ═══════════════════════════════════════════════════════════════════
# EXPORT / PDF TOOLS
# ═══════════════════════════════════════════════════════════════════

@mcp.tool()
def export_report_pdf(db_path: str, report_name: str, output_path: str,
                      date_from: str = "", date_to: str = "",
                      compare_date_from: str = "", compare_date_to: str = "",
                      hide_zero: bool = False) -> dict:
    """Export a report (BS, IS, etc.) to PDF. Optionally comparative with $chg and %chg."""
    import pdf_reports

    _init(db_path)
    resolved = _check_path(output_path, "output_path")
    report = models.find_report_by_name(report_name)
    if not report:
        raise ValueError(f"Report not found: {report_name}")
    company = models.get_meta("company_name", "My Books")

    items = models.get_report_items(report["id"])
    all_items = models.get_all_report_items()

    df = _normalize_date(date_from) if date_from else None
    dt = _normalize_date(date_to) if date_to else None

    # Current period column
    cur_data = models.compute_report_column(
        report["id"], date_from=df, date_to=dt,
        _display_items=items, _all_items=all_items)

    columns = [{"type": "actual", "label": dt or "Current", "data": cur_data}]

    # Comparative column (if requested)
    if compare_date_to:
        cdf = _normalize_date(compare_date_from) if compare_date_from else None
        cdt = _normalize_date(compare_date_to)
        prior_data = models.compute_report_column(
            report["id"], date_from=cdf, date_to=cdt,
            _display_items=items, _all_items=all_items)
        # Current year furthest LEFT, prior beside it — the order the screen
        # uses and the order the year-end package prints. This used to insert
        # the prior year in front, so the same books read one way on screen and
        # the other way in a PDF.
        columns.append({"type": "actual", "label": cdt, "data": prior_data})

        # $chg column
        change_data = []
        for j in range(len(prior_data)):
            item_p, bal_p = prior_data[j]
            _, bal_c = cur_data[j] if j < len(cur_data) else (None, 0)
            change_data.append((item_p, bal_c - bal_p))
        columns.append({"type": "change", "label": "$ chg", "data": change_data})

        # %chg column
        pct_data = []
        for j in range(len(prior_data)):
            item_p, bal_p = prior_data[j]
            _, bal_c = cur_data[j] if j < len(cur_data) else (None, 0)
            if bal_p != 0:
                pct = round((bal_c - bal_p) * 10000 / abs(bal_p))
            else:
                pct = 0
            pct_data.append((item_p, pct))
        columns.append({"type": "pct_change", "label": "% chg", "data": pct_data})

    # Build rows
    base_items = columns[0]["data"]
    rows = []
    for idx, (item, _) in enumerate(base_items):
        bals = []
        for col in columns:
            if col.get("data") and idx < len(col["data"]):
                bals.append(col["data"][idx][1])
            else:
                bals.append(0)
        if hide_zero and item.get("item_type") in ("account",) and all(b == 0 for b in bals):
            continue
        rows.append((item, bals))

    col_labels = [c["label"] for c in columns]
    col_types = [c["type"] for c in columns]

    pdf_bytes = pdf_reports.report_pdf(company, report_name, col_labels, col_types, rows, hide_zero)
    os.makedirs(os.path.dirname(resolved) or ".", exist_ok=True)
    with open(resolved, "wb") as f:
        f.write(pdf_bytes)

    return {
        "output_path": resolved,
        "bytes": len(pdf_bytes),
        "rows": len(rows),
        "columns": col_labels,
    }


@mcp.tool()
def export_gl_pdf(db_path: str, output_path: str,
                  date_from: str = "", date_to: str = "",
                  dr_cr_filter: str = "all") -> dict:
    """Export the full General Ledger to PDF."""
    import pdf_reports

    _init(db_path)
    resolved = _check_path(output_path, "output_path")
    company = models.get_meta("company_name", "My Books")

    accounts = (pdf_reports._get_report_account_order("BS")
                + pdf_reports._get_report_account_order("IS"))
    bs_ids = pdf_reports._get_bs_account_ids()

    df = _normalize_date(date_from) if date_from else ""
    dt = _normalize_date(date_to) if date_to else ""

    pdf_bytes = pdf_reports.gl_pdf(company, accounts, bs_ids, df, dt, dr_cr_filter)
    os.makedirs(os.path.dirname(resolved) or ".", exist_ok=True)
    with open(resolved, "wb") as f:
        f.write(pdf_bytes)

    return {
        "output_path": resolved,
        "bytes": len(pdf_bytes),
        "accounts": len(accounts),
    }


@mcp.tool()
def export_aje_pdf(db_path: str, account_name: str, output_path: str,
                   date_from: str = "", date_to: str = "") -> dict:
    """Export an Adjusting Journal Entry report for one account to PDF."""
    import pdf_reports

    _init(db_path)
    resolved = _check_path(output_path, "output_path")
    company = models.get_meta("company_name", "My Books")

    acct = models.get_account_by_name(account_name)
    if not acct:
        raise ValueError(f"Account not found: {account_name}")

    df = _normalize_date(date_from) if date_from else ""
    dt = _normalize_date(date_to) if date_to else ""

    pdf_bytes = pdf_reports.aje_pdf(company, acct["id"], acct["name"],
                                    acct["description"] or "", df, dt)
    os.makedirs(os.path.dirname(resolved) or ".", exist_ok=True)
    with open(resolved, "wb") as f:
        f.write(pdf_bytes)

    return {
        "output_path": resolved,
        "bytes": len(pdf_bytes),
        "account": acct["name"],
    }






@mcp.tool()
def import_aje(db_path: str, file_path: str, ref_prefix: str,
               journal_account: str = "", ye_date: str = "") -> dict:
    """Import adjusting journal entries from a file.

    Supports native formats (IIF, TXT/Venice) via parse_csw_aje(), and
    PDF/image/Excel via aje_extract.py (Anthropic API extraction).

    Auto-matches CaseWare account names to GridTRX accounts. (Retained earnings is
    computed off-book now — there is no roll to undo/redo; AJEs just post.)

    file_path: path to the AJE file (IIF, TXT, PDF, image, Excel).
    ref_prefix: reference prefix for posted transactions (e.g. 'AJE').
    journal_account: account name for journal routing (defaults to ref_prefix).
    ye_date: fiscal year end date if re-roll is needed (YYYY-MM-DD, optional).
    """
    _init(db_path)

    resolved = _check_path(file_path, "file_path")

    jrnl = journal_account.strip() if journal_account.strip() else None
    yed = _normalize_date(ye_date) if ye_date.strip() else None

    result = models.process_aje_file(resolved, ref_prefix, jrnl, yed)
    return result


if __name__ == "__main__":
    mcp.run()
