---
name: gridtrx
description: Double-entry, full-cycle accounting suite built for AI agents. Converts bank CSVs, OFX, and QBO files into balanced, auditable books — balance sheet, income statement, general ledger, trial balance. All data stays in a single local SQLite file.
homepage: https://github.com/737999/GridTRX
requires_tools:
  - exec
  - read
  - write
metadata:
  clawdbot:
    requires:
      env:
        - GRIDTRX_WORKSPACE
      bins:
        - python3
    primaryEnv: GRIDTRX_WORKSPACE
    files:
      - "*.py"
---

# Skill: GridTRX Accounting

## What it does

Use this skill when the user explicitly asks for bookkeeping work — "do the books," "categorize expenses," "import bank transactions," "run a balance sheet." Do not activate on casual mentions of money or finances, and confirm with the user before creating books or writing to existing ones. GridTRX is a full-cycle double-entry accounting engine. You prompt in plain English, and the agent completes the books correctly.

**Destructive operations — confirm first.** Before deleting transactions, re-importing, changing the lock date, or rewriting report layouts, tell the user exactly what you are about to do and get their OK. The engine has its own guardrails — deletions respect the lock date, every import is tagged as one batch, and each book is snapshotted daily on first open — but the agent's habit is still: state the action, confirm, then act. Every transaction balances. Every amount is deterministic. All data is local — no cloud services, no external APIs.

GridTRX produces a full set of auditable books: balance sheet, income statement, general ledger, trial balance, adjusting journal entries, and a perpetual retained-earnings statement. Reports are exportable to CSV and PDF for any time period.

## Architecture

GridTRX has three interfaces to the same engine (`models.py` → `books.db`):

1. **MCP Server (preferred for agents)** — Structured JSON tools. 57 tools wrapping `models.py` directly. No text parsing, typed parameters, deterministic output.
2. **CLI (fallback for agents, power users)** — One-shot shell commands via `python cli.py`. Zero dependencies beyond Python 3.7+ standard library. Any terminal-based agent can drive it via subprocess.
3. **Browser UI (for humans)** — Flask web interface at `localhost:5000` via `python run.py`. Ledger browsing, report viewer with drill-down, comparative reports up to 13 columns, bank import with rule preview, reconciliation marking, dark mode.

All three hit the same `models.py` data layer. Nothing is out of sync. Use MCP when available. Fall back to CLI otherwise. The browser UI is for human review.

### Prerequisites

Install dependencies before first use (one-time setup):

```bash
pip install -r requirements.txt
```

Or install individually: `pip install mcp` (MCP server), `pip install flask` (browser UI). The CLI has no dependencies beyond the Python 3.7+ standard library.

**No packages are installed at runtime.** All dependencies must be pre-installed.

### MCP Setup

Add to the agent's MCP config with `GRIDTRX_WORKSPACE` set to the user's client folder:
```json
{
  "command": "python",
  "args": ["/path/to/mcp_server.py"],
  "env": {"GRIDTRX_WORKSPACE": "/path/to/clients"}
}
```

`GRIDTRX_WORKSPACE` is mandatory — the MCP server will refuse to start without it. Any `db_path` outside the workspace is rejected at runtime. Every MCP tool takes `db_path` as its first parameter, which must resolve to a `books.db` file inside the workspace.

### CLI Usage

```
GRIDTRX_WORKSPACE=/path/to/clients python cli.py /path/to/clients/acme/books.db <command>
```

Runs one command, prints plain text to stdout, exits. When `GRIDTRX_WORKSPACE` is set, the CLI enforces the same workspace boundary as the MCP server — paths outside the workspace are rejected.

## Inputs needed

- The absolute path to the client's books (`books.db` file or its parent folder).
- The absolute path to the bank file (`.csv`, `.ofx`, or `.qbo`).
- The bank account name to post against (typically `BANK.CHQ` for chequing).

## Core concepts

- **Double-entry:** Every transaction is a balanced zero-sum entry. Debits = Credits. Always.
- **Sign convention:** Positive = Debit. Parentheses `(1,500.00)` = Credit. `—` = Zero.
- **Amounts:** Stored as integer cents internally. Displayed as dollars with two decimals.
- **Account names:** Case-insensitive, UPPER by convention. Common prefixes: `BANK.` `EX.` `REV.` `AR.` `AP.` `GST.` `RE.` `SHL.` — When importing a trial balance or creating accounts, ALWAYS use GridTRX naming. Never use numeric account codes. If the source data has numeric codes (1010, 5800, etc.), ignore the codes and map by description to the nearest GridTRX account name. If no match exists, create the account using the `EX.` or `REV.` prefix convention. Always call `list_accounts` first before creating anything.
- **EX.SUSP (Suspense):** Where unrecognized transactions land. This is the triage queue. Tell the AI what the suspense items are and it will clear them. Or clear them yourself through the GUI.
- **Import rules:** Keyword → account mappings. Case-insensitive match, highest priority wins. Optional tax code splits the amount into net + tax automatically.
- **Lock date:** Prevents changes to closed periods. Check before importing historical data.
- **Architecture:** Each client is one SQLite file. Copy it, back it up. It is confidential financial data — do not transmit it without the user's explicit approval. One data layer (`models.py`) — CLI, MCP server, and browser UI all call the same functions.

## Workflow

### Step 1: Initialize (if no books exist)

**MCP:** No direct tool — use exec to run CLI.
**CLI:** `python cli.py` then `new /path/to/folder "Company Name"`

This creates `books.db` with a full chart of accounts (~60 posting accounts), four reports (BS, IS, AJE, TRX), sample import rules, and four tax codes. **Always use starter books as the base** — they include the critical perpetual retained earnings chain (IS → NI → RE on the BS, with `RE.OB` holding opening retained earnings; `RE.OPEN`/`RE.CLOSE` are computed Statement-of-RE display lines). Never build reports from scratch without this chain. The fiscal year end is set to the most recent one that has PASSED — the year-end being worked on — and posting stops there until you open the next (`set_fiscal_settings(ceiling_mode='next')`).

After setup, run `validate` to confirm the report chain is intact:
**CLI:** `python cli.py /path/to/books.db validate`

### Step 1b: Import Prior-Year Balances (Existing Business)

For existing businesses, you need the prior-year trial balance to set up opening balances AND prior-year comparatives. Source data (in priority order):

1. Formal prior-year trial balance (best — complete, balanced)
2. PY balance sheet + income statement (from financial statements)
3. Tax return schedules (CRA S100/S125, US Form 1120, etc.)
4. Accountant's working papers

**IMPORTANT:** Post **both** balance sheet AND income statement accounts. If you only post BS accounts, your PY comparative on the IS will be blank.

**Process:**

1. Create starter books (Step 1 above).
2. Post the whole conversion with **one call** — `post_opening_balances`. It creates
   `TRX.OPEN` on the TRX report, posts a 2-line entry per account against it (same date,
   reference `OPEN`), computes retained earnings as the residual and posts it to `RE.OB`,
   and leaves `TRX.OPEN` at 0.00. Atomic: a bad row posts nothing.

   ```
   post_opening_balances(db_path, "2024-12-31", [
       {"account": "BANK.CHQ", "description": "Chequing", "amount": "33268.00"},
       {"account": "AP",       "description": "Payables", "amount": "(1000.00)"},
       {"account": "REV",      "description": "Sales",    "amount": "(84000.00)"},
   ], expected_retained_earnings="35065.00")
   ```

   `amount` is a dollar string signed the way Grid displays: plain = **debit**,
   `(1000.00)` or `-1000.00` = **credit**. One row per line of the PY trial balance.

3. **Never pass retained earnings.** It is computed and posted for you. `RE`, `RE.OB`,
   `RE.OPEN`, `RE.CLOSE` and `EX.SUSP` are refused, as is any account not in the chart
   (add it first). Never park a conversion in `EX.SUSP` and clear it down.

4. Pass `expected_retained_earnings` whenever the prior-year statements are to hand — the
   post then **refuses** unless the computed residual ties to it. That is the check that
   catches a transposed or missing figure.

5. Which RE figure comes out depends on what you posted:
   - **Balance sheet only** → retained earnings **brought forward** (closing RE).
   - **Full TB including IS accounts** → retained earnings at the **start** of that year
     (T2 line 3660, *not* 3600); PY net income reaches closing RE through the IS chain on
     its own. Post the full TB when you want a prior-year comparative on the IS, and post
     `DIVPAID` (line 3701) too if dividends were paid.

6. Check with `opening_balances_status(db_path)`. Wrong? `delete_opening_balances(db_path)`
   and post it again — Grid does not adjust a bad conversion, it redoes it. `TRX.OPEN` lives
   on the TRX report; open it to see every opening balance that was loaded.

7. Proceed to Step 2 (import current-year bank data).

Humans do exactly the same thing through the browser: Grid notices a set of books with
nothing posted and offers **Enter opening balances**, which opens a 25-line grid over the
TRX ledger with the retained-earnings figure computed live at the bottom. Same model code,
same guards, same atomic batch.

### Step 2: Import data

#### Bulk import decision rule

| Data source | Tool | Notes |
|---|---|---|
| Bank CSV (3-col: Date, Description, Amount) | `import_csv` | Rule-based categorization, unmatched → EX.SUSP |
| Bank OFX/QBO | `import_ofx` | Rule-based categorization, FITID dedup |
| GL export / system conversion (4-col: Date, Description, Amount, CrossAccount) | `import_gl` | Cross-account specified per row, no rules needed |
| CaseWare AJE (IIF or Venice) | `import_aje` | Routes through journal account |
| Single manual entry | `post_transaction` | One at a time only |

Never call `post_transaction` in a loop to import bank data. Use the appropriate bulk tool.

#### Bank CSV / OFX import

**MCP (preferred):**
- CSV: `import_csv(db_path, csv_path, "BANK.CHQ")`
- OFX/QBO: `import_ofx(db_path, ofx_path, "BANK.CHQ")`

**CLI fallback:**
- CSV: `python cli.py /path/to/books.db importcsv /path/to/file.csv BANK.CHQ`
- OFX: `python cli.py /path/to/books.db importofx /path/to/file.qbo BANK.CHQ`

**CSV format required:** 3 columns — Date, Description, Amount. Deposits positive, withdrawals negative. Header row optional. The import applies all rules automatically. Check the result summary: `posted`, `skipped`, `to_suspense`.

#### General ledger import (system conversion)

**MCP:** `import_gl(db_path, csv_path, "BANK.CHQ")`

**CSV format required:** 4 columns — Date, Description, Amount, CrossAccount. Positive = debit to the primary account, negative = credit. All cross-accounts must already exist in the chart of accounts. No import rules are applied — the cross-account column is used directly.

Use this when converting from another accounting system (QuickBooks GL, Sage, etc.) where every transaction already has its cross-account known. Prepare one CSV per primary account (bank, credit card, etc.) per fiscal year.

#### CaseWare AJE import

- MCP: `import_aje(db_path, file_path, "25AJE")`
- CLI: `python cli.py /path/to/books.db importaje /path/to/aje_export.iif 25AJE`
- Supports QuickBooks IIF and Venice/MYOB text formats. Maps CsW account descriptions to Grid account codes.

### Step 3: Audit suspense

**MCP:** `get_ledger(db_path, "EX.SUSP")`
**CLI:** `python cli.py /path/to/books.db ledger EX.SUSP`

Every entry here is an unrecognized transaction. Note the description and transaction ID for each.

### Step 4: Resolve suspense with the user

Present each suspense item to the user. Ask: *"What category is this?"*

Do NOT guess. If the description is ambiguous (e.g., "AMAZON", "BEST BUY", "TRANSFER"), ask the user for business context before categorizing.

Once the user answers, add a rule so future imports are automatic:

**MCP:** `add_rule(db_path, "AMAZON", "EX.OFFICE", "G5", 0)`
**CLI:** `python cli.py /path/to/books.db addrule AMAZON EX.OFFICE G5 0`

Tax code is optional. Common codes: `G5` (GST 5%), `H13` (HST 13%), `H15` (HST 15%), `E` (exempt).

### Step 5: Clear the bad suspense entries and re-import

Show the user the list of suspense transactions you intend to replace and get their OK. Then delete them and re-import so the new rules apply:

**MCP:** `delete_transaction(db_path, txn_id)` for each, then `import_csv(...)` or `import_ofx(...)` again.
**CLI:** `python cli.py /path/to/books.db delete <txn_id>` for each, then re-run the import command.

Repeat Steps 3-5 until suspense is empty.

### Step 6: Verify and report

**MCP:**
- `trial_balance(db_path)` — debits must equal credits
- `generate_report(db_path, "BS")` — Balance Sheet
- `generate_report(db_path, "IS")` — Income Statement

**CLI:**
- `python cli.py /path/to/books.db tb`
- `python cli.py /path/to/books.db report BS`
- `python cli.py /path/to/books.db report IS`

### Step 7: Moving into the next fiscal year

There is **no year-end close and no rollforward** — retained earnings is perpetual and the IS
Opening/Closing RE lines (`RE.OPEN`/`RE.CLOSE`) are computed automatically as the perpetual RE
balance at the period start/end. To move into the next year:

1. (Optional) Lock the finished year: set `lock_date` to the fiscal year-end so it can't be edited.
2. Move the working year-end forward: `set_fiscal_settings(working_year_end=...)` so post-YE
   dates can be posted.

Opening retained earnings for the new year is simply the perpetual RE as of the prior year-end,
re-derived on every report — no journal entry. Then repeat from Step 2 for the next fiscal year.
Prior-period adjustments are ordinary dated postings; they re-derive opening and closing RE on their own.

## Multi-year backfill protocol

When importing multiple fiscal years of historical data (perpetual RE — **no rollforward** between years):

1. **Check the ceiling:** `get_info(db_path)` — note `fy_ceiling`. It must be ≥ your LATEST row date. Set `working_year_end` to the newest year-end you are loading (`set_fiscal_settings`, CLI `fye`, or Options → Global Options) so no rows get blocked. Older years all fall below it.
2. **Prepare per-FY CSVs:** One CSV per primary account per fiscal year, each containing only that FY's date range.
3. **Import oldest → newest:** Use `import_gl` (4-col) or `import_csv`/`import_ofx` for each account, earliest fiscal year first. Every date is ≤ the ceiling, so nothing is skipped. RE accumulates perpetually as you go — opening/closing RE for each year are re-derived automatically.
4. (Optional) **Lock finalized years:** `set_lock_date(db_path, "YYYY-MM-DD")` on each completed prior year so it can't be edited.

**Example (3-year backfill, Dec 31 FY, working year-end already 2024-12-31):**
```
# FY2022
import_gl(db, "fy2022_bank_cdn.csv", "BANK.CDN")
import_gl(db, "fy2022_cc_visa.csv", "CC.VISA")

# FY2023
import_gl(db, "fy2023_bank_cdn.csv", "BANK.CDN")
import_gl(db, "fy2023_cc_visa.csv", "CC.VISA")

# FY2024 — current year, use rules
import_csv(db, "fy2024_bank_cdn.csv", "BANK.CDN")
import_csv(db, "fy2024_cc_visa.csv", "CC.VISA")
```
No rollforward — RE is perpetual and re-derives opening/closing for every year.

## Recovery: Undoing a bad import

If the user uploaded the wrong file or you imported against the wrong account:

1. **Find the bad transactions:** `search_transactions(db_path, "some description")` or via CLI `search <keyword>`.
2. **Confirm the matched list with the user**, then delete them one by one: `delete_transaction(db_path, txn_id)` or CLI `delete <txn_id>`.
3. **Verify the trial balance** still balances after cleanup.
4. **Re-import** the correct file.

From MCP/CLI, deletions are individual and respect the lock date — you cannot delete transactions in a locked period. Every import is tagged as one batch, and the browser UI can undo a whole import in one step; a bad import is cured by deleting it whole, never by unpicking part of it.

## MCP tools reference (selected)

### Read tools
| Tool | Purpose |
|------|---------|
| `list_accounts(db_path, query?)` | List/search chart of accounts |
| `get_balance(db_path, account_name, date_from?, date_to?)` | Single account balance |
| `get_ledger(db_path, account_name, date_from?, date_to?)` | Account ledger with running balance |
| `trial_balance(db_path, as_of_date?)` | Trial balance — all accounts, Dr/Cr columns |
| `generate_report(db_path, report_name, date_from?, date_to?, side?)` | Run a report (BS, IS, AJE, etc.). `side='D'/'C'` gives the FLOWS — debit side of AR = sales by customer, credit side of AP = purchases by supplier |
| `get_transaction(db_path, txn_id)` | Single transaction with all journal lines |
| `search_transactions(db_path, query, limit?)` | Search by description/reference |
| `list_reports(db_path)` | List available reports |
| `list_rules(db_path)` | List import rules |
| `get_info(db_path)` | Company name, fiscal year, lock date |
| `create_books(db_path, company_name, business_type, fiscal_year_end)` | Make a NEW set of books. The only tool that creates one. business_type: services / retail / construction — ASK, it changes the statements |
| `lock_status(db_path)` | Who currently holds these books, if anyone. Opens nothing |
| `list_procedures(db_path)` | The whole jobs this program can do end to end (YE package, 13-column report) |
| `run_procedure(db_path, name, as_of, output_path, report_name?)` | Run one and write its PDF. `as_of` is REQUIRED — ask the operator which year, never assume the current one |
| `check_books(db_path)` | Is anything WRONG with these books, is anything outstanding. Run it first, and again before reporting work done |

### Write tools
| Tool | Purpose |
|------|---------|
| `post_transaction(db_path, date, description, amount, debit_account, credit_account)` | Post a simple 2-line entry |
| `delete_transaction(db_path, txn_id)` | Delete a transaction (respects lock date) |
| `add_account(db_path, name, normal_balance, description?)` | Add a posting account |
| `add_rule(db_path, keyword, account_name, tax_code?, priority?)` | Add an import rule |
| `delete_rule(db_path, rule_id)` | Delete an import rule |
| `import_csv(db_path, csv_path, bank_account)` | Import bank CSV (3-col, rule-based) |
| `import_ofx(db_path, ofx_path, bank_account)` | Import bank OFX/QBO (rule-based) |
| `import_gl(db_path, csv_path, bank_account)` | Import GL CSV (4-col, cross-account per row) |
| `import_aje(db_path, file_path, ref_prefix)` | Import CaseWare AJE export (IIF or Venice) |
| `set_lock_date(db_path, lock_date?)` | Show or set the lock date |
| `clear_lock(db_path, confirm)` | Clear a leftover books lock. THE OPERATOR'S CALL — refuses without their confirmation |
| `set_fiscal_settings(db_path, working_year_end?, ceiling_mode?, lock_date?, company_name?)` | Set the fiscal dates. Ceiling is derived: `cy` or `next` |
| `new_aje_year(db_path, account_name?, description?)` | Set up a year of adjusting entries (suggests 26AJE / "2026 Adjusting Entries") |
| `post_aje(db_path, journal, description, lines, entry_date?, ref?, replace_ref?)` | Post ONE adjusting entry in the house shape. The only supported way |
| `list_aje_years(db_path)` / `list_aje(db_path, journal)` | The years on file / one year's entries, grouped |
| `delete_aje(db_path, journal, ref)` | Remove one adjustment, every leg of it |
| `bulk_report_layout(db_path, report_name, items, after_account?, mode?)` | Batch-place items on a report (accounts, totals, labels, separators) |

## Guardrails

- **NEVER GUESS CATEGORIES.** If a transaction description is ambiguous, let it go to `EX.SUSP` and ask the user. Do not assume "AMAZON" is office supplies — it could be inventory, personal, or cost of sales.
- **NEVER MODIFY books.db DIRECTLY.** All writes go through `cli.py` commands or MCP tools. Never use file tools to read or write the SQLite database.
- **STAY IN THE WORKSPACE.** Only operate on `books.db` files within the user's GridTRX workspace. Both the MCP server and CLI enforce this when `GRIDTRX_WORKSPACE` is set — the MCP server will not start without it, and both interfaces reject any path outside the workspace.
- **NO OUTBOUND NETWORK REQUESTS.** GridTRX processes data locally. It does not phone home, call APIs, or transmit data. Do not attempt to "verify" transactions against external services.
- **RESPECT THE POSTING WINDOW.** Before importing, check the lock date and FY ceiling with `get_info()`. You cannot post on or before the lock date, or after the FY ceiling. The ceiling is DERIVED from the working year-end — move that forward, or set `ceiling_mode='next'` to run one year ahead of it. It is never keyed as a free date and never moves behind postings that already exist. No rollforward.
- **PRESERVE RAW OUTPUT.** When presenting financial data to the user, use the exact numbers from GridTRX. Do not round, reformat, or flip signs. Positive = Debit. Parentheses = Credit.
- **TRIAL BALANCE MUST BALANCE.** After any operation, if the trial balance shows unequal debits and credits, something is wrong. Stop and investigate before proceeding.
- **LIMIT EXEC SCOPE.** When using exec, only run `python cli.py` commands against books within the workspace. Do not run arbitrary shell commands, install packages, start background processes, or execute scripts other than `cli.py`. The MCP server is the preferred interface — use CLI only when MCP is unavailable.
