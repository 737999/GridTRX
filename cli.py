#!/usr/bin/env python3
"""
GridTRX CLI — Command line interface for GridTRX accounting.
For AI agents and human accountants. Calls models.py directly.

Three modes:
    python cli.py                         # interactive mode, pick a client
    python cli.py /path/to/books.db       # open specific books
    python cli.py /path/to/books.db tb    # one-shot: run command and exit
"""
import cmd
import sys
import os
import csv
import io
import json
import shlex
import textwrap
from datetime import datetime, timedelta

# Add script directory to path so we can import models
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import models

def _check_workspace(path):
    """Enforce GRIDTRX_WORKSPACE boundary. Refuses to operate if unset."""
    ws = os.environ.get('GRIDTRX_WORKSPACE', '')
    if not ws:
        print("  Error: GRIDTRX_WORKSPACE environment variable is not set.")
        print("  Set it to the directory containing your client books, e.g.:")
        print("    export GRIDTRX_WORKSPACE=~/clients")
        return False
    ws = os.path.realpath(os.path.expanduser(ws))
    resolved = os.path.realpath(os.path.expanduser(path))
    if not resolved.startswith(ws + os.sep) and resolved != ws:
        print(f"  Access denied: '{path}' is outside the workspace ({ws}).")
        return False
    return True

# ─── Formatting helpers ──────────────────────────────────────────

def fmt(cents):
    """Format cents as dollars. Negatives in parens."""
    if cents == 0:
        return '—'
    neg = cents < 0
    c = abs(cents)
    s = f"{c // 100:,}.{c % 100:02d}"
    return f"({s})" if neg else s

def fmt_plain(cents):
    """Format cents as plain number (for CSV/piping)."""
    if cents == 0:
        return '0.00'
    neg = cents < 0
    c = abs(cents)
    s = f"{c // 100:,}.{c % 100:02d}"
    return f"-{s}" if neg else s

def parse_amount(s):
    """Parse a dollar string to cents. Handles 1500, 1500.00, 1,500.00, (500), -500, 500-"""
    return models.parse_amount(s)

def table(headers, rows, alignments=None):
    """Print a formatted text table. alignments: 'l' left, 'r' right per column."""
    if not rows:
        print("  (no data)")
        return

    # Calculate column widths
    ncols = len(headers)
    widths = [len(h) for h in headers]
    for row in rows:
        for i, cell in enumerate(row):
            if i < ncols:
                widths[i] = max(widths[i], len(str(cell)))

    if not alignments:
        alignments = 'l' * ncols

    # Header
    hdr = ''
    for i, h in enumerate(headers):
        if alignments[i] == 'r':
            hdr += str(h).rjust(widths[i])
        else:
            hdr += str(h).ljust(widths[i])
        if i < ncols - 1:
            hdr += '  '
    print(f"  {hdr}")
    print(f"  {'─' * len(hdr)}")

    # Rows
    for row in rows:
        line = ''
        for i in range(ncols):
            cell = str(row[i]) if i < len(row) else ''
            if alignments[i] == 'r':
                line += cell.rjust(widths[i])
            else:
                line += cell.ljust(widths[i])
            if i < ncols - 1:
                line += '  '
        print(f"  {line}")


def resolve_account(name_or_id):
    """Resolve an account by name (case-insensitive) or by ID number."""
    if not name_or_id or not name_or_id.strip():
        print("  No account specified.")
        print("  Use 'accounts' to list all accounts, or 'find <query>' to search.")
        return None
    name_or_id = name_or_id.strip()
    if name_or_id.isdigit():
        acct = models.get_account(int(name_or_id))
        if acct:
            return acct
    acct = models.get_account_by_name(name_or_id)
    if acct:
        return acct
    # Try partial match
    results = models.search_accounts(name_or_id)
    if len(results) == 1:
        return results[0]
    if len(results) > 1:
        print(f"  Ambiguous account '{name_or_id}'. Matches:")
        for r in results[:10]:
            print(f"    {r['name']:<20} {r['description']}")
        if len(results) > 10:
            print(f"    ... and {len(results) - 10} more")
        print(f"  Use the full account name, or try: find {name_or_id}")
        return None
    print(f"  Account not found: '{name_or_id}'")
    print("  Use 'accounts' to list all accounts, or 'find <query>' to search.")
    return None


# ─── CLI Shell ───────────────────────────────────────────────────

class GridCLI(cmd.Cmd):
    intro = None  # We print our own banner
    prompt = 'Grid> '

    def __init__(self):
        super().__init__()
        self.db_path = None

    def set_books(self, path, force=False):
        """Open a books.db file. force=True takes a stale lock (open --force)."""
        if not _check_workspace(path):
            return False
        if not os.path.exists(path):
            print(f"  File not found: {path}")
            return False
        try:
            # gated open: integrity check → daily snapshot → migrate
            models.init_db(path, force_lock=force)
        except models.BooksLocked as e:
            # v130 — a refusal must carry the way out, here as much as on the
            # lock screen. Nobody should have to go and delete a file by hand.
            print(f"\n  !! {e.headline}")
            print(f"     {e.prog} · PID {e.pid} · {e.host} · since {e.started}")
            print(f"     {e.detail}")
            if e.warn:
                print(f"     {e.warn}")
            print(f"\n     {e.door}")
            print(f"     If you are sure, open it anyway:  open {path} --force\n")
            return False
        except ValueError as e:
            print(f"  !! {e}")
            return False
        if force:
            print("  !! lock cleared by the operator")
        for line in models.re_repair_note():
            print(f"  !! {line}")
        st = models.backup_status()
        if st['error']:
            print(f"  !! {st['note']}")
        elif st['note']:
            print(f"  {st['note']}")
        self.db_path = path
        name = models.get_meta('company_name', os.path.basename(os.path.dirname(path)))
        self.prompt = f'Grid/{name}> '
        print(f"  Opened: {name} ({path})")
        return True

    def _require_books(self):
        """Check that a books.db is open."""
        if not self.db_path:
            print("  No books open. Use: open <path/to/books.db>")
            print("  Or create new books: new <folder> [\"Company Name\"]")
            return False
        return True

    def _check_lock_date(self, date_str):
        """Check if a date is on or before the lock date. Returns True if OK to post."""
        lock = models.get_meta('lock_date', '')
        if lock and date_str <= lock:
            print(f"  Cannot post on {date_str} — the books are locked through {lock}.")
            print(f"  The lock date prevents changes to transactions on or before that date.")
            print(f"  Use 'lock' to view the current lock date.")
            return False
        return True

    def _check_ceiling_date(self, date_str):
        """Check if a date is after the fiscal year end. Returns True if OK to post."""
        ceiling = models.fiscal_ceiling()
        if ceiling and date_str > ceiling:
            print(f"  Cannot post on {date_str} — fiscal year ends {ceiling}.")
            print(f"  Use 'ceiling next' to open the following year.")
            return False
        return True

    def _check_posting_account(self, acct):
        """Check that an account is a posting account. Returns True if OK."""
        if acct['account_type'] != 'posting':
            print(f"  Cannot post to '{acct['name']}' — it is a {acct['account_type']} account.")
            print(f"  You can only post to 'posting' accounts.")
            print(f"  Use 'accounts posting' to see all posting accounts.")
            return False
        return True

    # ─── help ────────────────────────────────────────────────────

    def do_help(self, arg):
        """Show available commands."""
        if arg:
            # Delegate to specific command help
            super().do_help(arg)
            return
        print("""
  Grid CLI — Commands
  ═══════════════════
  All amounts are RAW: positive = Debit, negative (parens) = Credit.
  No normal-balance sign flipping.

  NAVIGATION
    open <path>              Open a books.db file
    close                    Close current books
    info                     Show company info and stats
    library [path]           List clients in library folder

  SETUP
    new <folder> ["Company Name"] [MM-DD]
      Create new client books with default chart of accounts.
      Example: new ~/clients/acme "Acme Corp" 03-31

    addaccount <name> <D|C> <description> [posting|total]
      Add a new account. D = debit-normal, C = credit-normal.
      Example: addaccount EX.PARKING D "Parking Expense"

    editaccount <name> [--desc "text"] [--num "1000"]
      Edit an account's description or account number.
      Example: editaccount EX.RENT --desc "Office Rent" --num "5200"

  ACCOUNTS
    accounts [posting|total]  List all accounts (optional filter)
    account <name>           Show account details
    find <query>             Search accounts by name

  LEDGER
    ledger <account> [from] [to]    Show account ledger
      Examples: ledger BANK.CHQ
                ledger BANK.CHQ 2025-01-01 2025-03-31

  TRANSACTIONS
    post <date> <desc> <amount> <debit_acct> <credit_acct>
      Post a simple 2-line transaction.
      Example: post 2025-03-01 "Rent" 1500.00 EX.RENT BANK.CHQ

    postx <date> <desc>
      Post a multi-line transaction (interactive).

    importcsv <csvfile> <bank_account>
      Import a bank CSV. Applies rules automatically.
      Example: importcsv statements/jan2025.csv BANK.CHQ

    importofx <ofxfile> <bank_account>
      Import a bank OFX/QBO file. Applies rules automatically.
      Example: importofx downloads/jan2025.qbo BANK.CHQ

    importgl <csvfile> <bank_account>
      Import a pre-categorized GL CSV (Date, Description, Amount, CrossAccount).
      No rules applied — cross-account is specified per row.
      Example: importgl fy2024_bank_cdn.csv BANK.CDN

    importaje <file> [ref_prefix]
      Import CaseWare AJE export (IIF or Venice format).
      Maps CsW accounts to Grid accounts, posts all entries.
      Example: importaje cwquickb.iif 25AJE

    edit <txn_id>            Show transaction details
    delete <txn_id>          Delete a transaction
    search <query>           Search transactions

  REPORTS
    tb [as-of-date]          Trial balance
    report <name>            Run a report (BS, IS, etc.)
    reports                  List all reports
    balance <account>        Show account balance

  EXPORT
    exportcsv <report> [filename] [from_date] [to_date]
      Export a report to CSV file.
      Example: exportcsv IS income_2025.csv 2025-01-01 2025-12-31

    exporttb [filename] [as-of-date]
      Export trial balance to CSV file.
      Example: exporttb trial_balance.csv 2025-12-31

  RULES
    rules                    List import rules
    addrule <keyword> <account> [tax_code] [priority]
      Example: addrule NETFLIX EX.COMP G5 10

    editrule <id> <keyword> <account> [tax_code] [priority]
      Edit an existing import rule.
      Example: editrule 5 NETFLIX EX.COMP G5 20

    delrule <id>             Delete an import rule

  FISCAL
    aje [...]                 Adjusting entries — years, entries, post, delete
    fye [YYYY-MM-DD]          Show or set the fiscal year end being worked on
    ceiling [cy|next]         How far past that year-end posting stays open
    validate                  Validate report chain (IS→NI→RE, totals, BS balance)

  OTHER
    reconcile <account>      Show reconciliation summary
    taxcodes                 List tax codes
    lock [date]              Show or set lock date
    help                     This help
    quit                     Exit
""")

    # ─── open / close ────────────────────────────────────────────

    def do_open(self, arg):
        """Open a books.db file. Usage: open <path/to/books.db> [--force]
        --force takes the books from a stale lock — the same deliberate door
        F3 offers on the web lock screen."""
        arg = arg.strip()
        force = False
        for flag in ('--force', '-f'):
            if arg.endswith(flag):
                arg, force = arg[:-len(flag)].strip(), True
        arg = arg.strip().strip('"').strip("'")
        if not arg:
            print("  Usage: open <path/to/books.db> [--force]")
            print("  You can also point to a folder that contains books.db.")
            return
        path = os.path.expanduser(arg)
        if not _check_workspace(path):
            return
        if os.path.isdir(path):
            # Check for books.db inside
            db = os.path.join(path, 'books.db')
            if os.path.exists(db):
                path = db
            else:
                print(f"  No books.db found in {path}")
                print(f"  To create new books here: new {arg}")
                return
        if not os.path.exists(path):
            print(f"  File not found: {path}")
            folder = os.path.dirname(path) or arg
            print(f"  To create new books: new {folder}")
            return
        self.set_books(path, force=force)

    def do_close(self, arg):
        """Close the current books."""
        models.set_db_path(None)
        self.db_path = None
        self.prompt = 'Grid> '
        print("  Books closed.")

    def do_check(self, arg):
        """Check the books: is anything wrong, is anything outstanding. Usage: check"""
        if not self._require_books():
            return
        r = models.check_books()
        mark = {'ok': '  ok ', 'attention': '  ** ', 'error': ' !!! '}
        print()
        for c in r['checks']:
            print(f"{mark[c['status']]}{c['name']:<40} {c['detail']}")
        print(f"\n  {r['summary']}\n")

    def do_backup(self, arg):
        """Snapshot the open books to backups/ (checkpointed, verified). Usage: backup"""
        if not self._require_books():
            return
        try:
            path = models.backup_books(force=True)
            print(f"  Backed up: {path}")
            snaps = models.list_backups()
            print(f"  {len(snaps)} snapshot(s) on file (newest {models.BACKUP_KEEP} kept)")
        except Exception as e:
            print(f"  !! Backup failed: {e}")

    def do_openings(self, arg):
        """Opening balances (conversion). Usage:
             openings                     — where these books stand
             openings <file.csv>          — post a conversion from a CSV
             openings delete              — delete the conversion so it can be redone

        CSV columns: account, description, amount — one row per line of the prior-year
        trial balance. amount is signed as Grid displays: a plain number is a DEBIT,
        (1000.00) or -1000.00 is a CREDIT. The conversion date is the prior fiscal year
        end unless the file's first line is `date,YYYY-MM-DD`.

        Retained earnings is NEVER in the file — Grid computes it as the residual and
        posts it to RE.OB, creates TRX.OPEN, and lands the whole thing atomically.
        """
        if not self._require_books():
            return
        arg = (arg or '').strip()

        if not arg:
            st = models.openings_state()
            print(f"  Status: {st['status']}")
            if st['status'] == 'posted':
                ob = models.get_account_by_name(models.OPENING_RE_ACCT)
                re_c = -models.get_account_balance(ob['id']) if ob else 0
                conv = models.get_account_by_name(models.CONVERSION_ACCT)
                print(f"  Converted at {st['conversion_date']} — {st['entry_count']} entries")
                print(f"  Opening retained earnings: {models.fmt_amount_plain(re_c)} CR")
                print(f"  {models.CONVERSION_ACCT}: "
                      f"{models.fmt_amount_plain(models.get_account_balance(conv['id']))} "
                      f"(must be 0.00)")
            elif st['status'] == 'needed':
                print(f"  Nothing posted yet. Default conversion date: {st['default_date']}")
                print(f"  Post one with:  openings <file.csv>")
            elif st['status'] == 'declined':
                print("  Brand-new client — started at zero, no conversion.")
            else:
                print("  These books have activity but no conversion entry.")
            return

        if arg.lower() == 'delete':
            try:
                n = models.delete_opening_balances()
                print(f"  Conversion deleted — {n} entries removed.")
            except ValueError as e:
                print(f"  !! {e}")
            return

        if not os.path.exists(arg):
            print(f"  !! File not found: {arg}")
            return
        date = models.default_conversion_date()
        rows = []
        with open(arg, 'r', encoding='utf-8-sig') as f:
            for n, raw in enumerate(csv.reader(f), start=1):
                if not raw or not any(c.strip() for c in raw):
                    continue
                if n == 1 and raw[0].strip().lower() == 'date':
                    date = raw[1].strip()
                    continue
                if n == 1 and raw[0].strip().lower() in ('account', 'acct'):
                    continue                                    # header
                acct = raw[0].strip()
                desc = raw[1].strip() if len(raw) > 1 else ''
                amt = raw[2].strip() if len(raw) > 2 else ''
                try:
                    cents = models.parse_amount(amt) if amt else 0
                except ValueError:
                    print(f"  !! line {n} ({acct}): '{amt}' is not an amount")
                    return
                rows.append({'account': acct, 'description': desc, 'amount': cents})

        v = models.validate_opening_rows(date, rows)
        for r in v['rows']:
            if r['error']:
                print(f"  !! {r['account'] or '(no account)'}: {r['error']}")
            elif r['warning']:
                print(f"   ? {r['account']}: {r['warning']}")
        if not v['ok']:
            for e in v['errors']:
                print(f"  !! {e}")
            print("  Nothing posted.")
            return
        print(f"  Conversion date  : {date}")
        print(f"  Debits / Credits : {models.fmt_amount_plain(v['debit_cents'])} / "
              f"{models.fmt_amount_plain(v['credit_cents'])}")
        print(f"  Retained earnings: {models.fmt_amount_plain(v['re_credit_cents'])} CR  (computed)")
        print(f"  {v['re_note']}")
        try:
            res = models.post_opening_balances(date, rows)
        except ValueError as e:
            print(f"  !! {e}")
            return
        conv = models.get_account_by_name(models.CONVERSION_ACCT)
        print(f"  Posted {len(res['txn_ids'])} entries as {res['batch']}.")
        print(f"  {models.CONVERSION_ACCT} is now "
              f"{models.fmt_amount_plain(models.get_account_balance(conv['id']))}")

    def do_info(self, arg):
        """Show company info and database stats."""
        if not self._require_books():
            return
        name = models.get_meta('company_name', '(unnamed)')
        fye = models.get_meta('fiscal_year_end', '')
        fy = models.get_meta('fiscal_year', '')
        lock = models.get_meta('lock_date', '')
        accts = models.get_accounts()
        posting = [a for a in accts if a['account_type'] == 'posting']
        reports = models.get_reports()

        with models.get_db() as db:
            txn_count = db.execute("SELECT COUNT(*) FROM transactions").fetchone()[0]
            line_count = db.execute("SELECT COUNT(*) FROM lines").fetchone()[0]

        ceiling = models.fiscal_ceiling()
        mode = models.get_meta('fy_ceiling_mode', 'cy') or 'cy'
        anchor = models.fiscal_anchor()
        print(f"  Company:      {name}")
        print(f"  Working YE:   {anchor['cy_end'] if anchor else (fye + ' ' + fy)}")
        print(f"  Lock date:    {lock or '(none)'}")
        print(f"  FY ceiling:   {ceiling or '(none)'} ({'one year ahead' if mode == 'next' else 'same as working YE'})")
        print(f"  File:         {self.db_path}")
        print(f"  Size:         {os.path.getsize(self.db_path):,} bytes")
        print(f"  Accounts:     {len(posting)} posting, {len(accts) - len(posting)} total")
        print(f"  Transactions: {txn_count:,}")
        print(f"  Lines:        {line_count:,}")
        print(f"  Reports:      {', '.join(r['name'] for r in reports)}")

    def do_library(self, arg):
        """List clients in library folder. Usage: library [path]"""
        arg = arg.strip().strip('"').strip("'")
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'grid.json')
        if arg:
            lib_path = os.path.expanduser(arg)
        elif os.path.exists(config_path):
            with open(config_path) as f:
                cfg = json.load(f)
            lib_path = cfg.get('library_path', '')
        else:
            lib_path = ''

        if not lib_path or not os.path.isdir(lib_path):
            print("  No library path set. Usage: library <path>")
            return

        print(f"  Library: {lib_path}\n")
        rows = []
        for entry in sorted(os.listdir(lib_path)):
            client_dir = os.path.join(lib_path, entry)
            if os.path.isdir(client_dir):
                db_path = os.path.join(client_dir, 'books.db')
                if os.path.exists(db_path):
                    size = os.path.getsize(db_path)
                    mod = datetime.fromtimestamp(os.path.getmtime(db_path)).strftime('%Y-%m-%d')
                    rows.append((entry, f"{size:,}", mod, db_path))
                else:
                    rows.append((entry, '—', '—', '(no books.db)'))

        if rows:
            table(['Client', 'Size', 'Modified', 'Path'], rows, 'lrll')
        else:
            print("  (no client folders found)")

    # ─── new ──────────────────────────────────────────────────────

    def do_new(self, arg):
        """Create new client books with default chart of accounts.
        Usage: new <folder> ["Company Name"] [MM-DD]

        Creates the folder if needed, initializes books.db with the full
        default chart of accounts, reports (BS, IS, AJE, TRX),
        import rules, and tax codes.

        Examples:
          new ~/clients/acme
          new ~/clients/acme "Acme Corp"
          new ~/clients/acme "Acme Corp" 03-31
        """
        parts = _split_args(arg)
        if not parts:
            print("  Usage: new <folder> [\"Company Name\"] [MM-DD]")
            print("  Example: new ~/clients/acme \"Acme Corp\" 03-31")
            return

        folder = os.path.expanduser(parts[0])
        folder = os.path.abspath(folder)

        if not _check_workspace(folder):
            return

        # Company name: use arg or derive from folder name
        if len(parts) > 1:
            company_name = parts[1]
        else:
            company_name = os.path.basename(folder).replace('_', ' ').replace('-', ' ').title()

        # Fiscal year end: default 12-31
        fiscal_ye = '12-31'
        if len(parts) > 2:
            fye = parts[2].strip()
            # Validate MM-DD format
            try:
                datetime.strptime(f"2000-{fye}", '%Y-%m-%d')
                fiscal_ye = fye
            except ValueError:
                print(f"  Invalid fiscal year-end: '{fye}'")
                print("  Use MM-DD format, e.g. 12-31 or 03-31")
                return

        db_path = os.path.join(folder, 'books.db')

        # Check if books already exist
        if os.path.exists(db_path):
            print(f"  Books already exist: {db_path}")
            print(f"  Use 'open {folder}' to open them.")
            return

        # Create folder if needed
        try:
            os.makedirs(folder, exist_ok=True)
        except OSError as e:
            print(f"  Cannot create folder: {folder}")
            print(f"  {e}")
            return

        # Create the books
        try:
            models.create_starter_books(db_path, company_name, fiscal_ye)
        except Exception as e:
            print(f"  Error creating books: {e}")
            return

        ceiling = models.fiscal_ceiling()
        print(f"  ✓ Created new books: {company_name}")
        print(f"    Folder:     {folder}")
        print(f"    Database:   {db_path}")
        print(f"    Fiscal YE:  {fiscal_ye}")
        print(f"    FY ceiling: {ceiling}")

        # Auto-open the new books
        self.set_books(db_path)

    # ─── accounts ────────────────────────────────────────────────

    def do_accounts(self, arg):
        """List all accounts. Usage: accounts [posting|total]"""
        if not self._require_books():
            return
        accts = models.get_accounts()
        filter_type = arg.strip().lower() if arg else ''
        if filter_type and filter_type not in ('posting', 'total'):
            print(f"  Unknown filter: '{filter_type}'")
            print("  Usage: accounts [posting|total]")
            return
        rows = []
        for a in accts:
            if filter_type and a['account_type'] != filter_type:
                continue
            rows.append((
                a['name'],
                a['normal_balance'],
                a['account_type'][0].upper(),
                a['description'],
                a['account_number'] or ''
            ))
        if not rows and filter_type:
            print(f"  No {filter_type} accounts found.")
        else:
            table(['Account', 'N/B', 'T', 'Description', 'Number'], rows, 'llllr')

    def do_account(self, arg):
        """Show account details. Usage: account <name>"""
        if not self._require_books():
            return
        if not arg.strip():
            print("  Usage: account <name>")
            print("  Example: account BANK.CHQ")
            return
        acct = resolve_account(arg.strip())
        if not acct:
            return
        raw = models.get_account_balance(acct['id'])

        print(f"  Name:        {acct['name']}")
        print(f"  Description: {acct['description']}")
        print(f"  Type:        {acct['account_type']}")
        print(f"  Normal bal:  {'Debit' if acct['normal_balance'] == 'D' else 'Credit'}")
        print(f"  Number:      {acct['account_number'] or '(none)'}")
        print(f"  Raw balance: {fmt(raw)}  (+Dr / -Cr)")

        # Show which report it's on
        rpt = models.find_report_for_account(acct['id'])
        if rpt:
            print(f"  Report:      {rpt['name']}")
        else:
            print(f"  Report:      (not on any report)")

    def do_find(self, arg):
        """Search accounts. Usage: find <query>"""
        if not self._require_books():
            return
        if not arg.strip():
            print("  Usage: find <query>")
            print("  Searches account names and descriptions.")
            print("  Example: find bank")
            return
        results = models.search_accounts(arg.strip())
        if not results:
            print(f"  No accounts matching '{arg.strip()}'.")
            print("  Try a shorter or different search term, or use 'accounts' to list all.")
            return
        rows = [(a['name'], a['normal_balance'], a['description']) for a in results]
        table(['Account', 'N/B', 'Description'], rows)

    # ─── addaccount / editaccount ─────────────────────────────────

    def do_addaccount(self, arg):
        """Add a new account.
        Usage: addaccount <name> <D|C> <description> [posting|total]

        D = debit-normal (assets, expenses)
        C = credit-normal (liabilities, equity, revenue)
        Defaults to 'posting' type if not specified.

        Examples:
          addaccount EX.PARKING D "Parking Expense"
          addaccount TOTPARK D "Total Parking" total
        """
        if not self._require_books():
            return

        parts = _split_args(arg)
        if len(parts) < 3:
            print("  Usage: addaccount <name> <D|C> <description> [posting|total]")
            print("  D = debit-normal (assets, expenses)")
            print("  C = credit-normal (liabilities, equity, revenue)")
            print('  Example: addaccount EX.PARKING D "Parking Expense"')
            return

        name = parts[0].upper()
        normal_balance = parts[1].upper()
        description = parts[2]
        account_type = parts[3].lower() if len(parts) > 3 else 'posting'

        # Validate normal balance
        if normal_balance not in ('D', 'C'):
            print(f"  Invalid normal balance: '{parts[1]}'")
            print("  Use D for debit-normal (assets, expenses)")
            print("  Use C for credit-normal (liabilities, equity, revenue)")
            return

        # Validate account type
        if account_type not in ('posting', 'total'):
            print(f"  Invalid account type: '{parts[3]}'")
            print("  Use 'posting' (you can post transactions to it)")
            print("  Use 'total' (it accumulates from other accounts via reports)")
            return

        # Check if account already exists
        existing = models.get_account_by_name(name)
        if existing:
            print(f"  Account '{name}' already exists.")
            print(f"  Use 'account {name}' to view it, or 'editaccount {name}' to modify it.")
            return

        try:
            acct_id = models.add_account(name, normal_balance, description, account_type)
            print(f"  ✓ Account added: {name}")
            print(f"    Description:    {description}")
            print(f"    Normal balance: {'Debit' if normal_balance == 'D' else 'Credit'}")
            print(f"    Type:           {account_type}")

            # Warn if not on any report
            rpt = models.find_report_for_account(acct_id)
            if not rpt:
                print()
                print(f"  Note: '{name}' is not on any report yet.")
                print("  Add it to a report in the browser UI so it appears on BS or IS.")
        except Exception as e:
            print(f"  Error adding account: {e}")

    def do_setupar(self, arg):
        """Setup a Detailed Accounts Receivable subledger report.
        Usage: setupar

        Creates an AR report with 3 sample client accounts (Gretzky, Lemieux, Orr),
        a total account (ARDET), and links it to the Balance Sheet via AR.DET.

        The sample accounts use the R.XYZABC naming convention:
          R.GREWAY  Gretzky, Wayne
          R.LEMMAR  Lemieux, Mario
          R.ORRBOB  Orr, Bobby

        After setup, add real clients with: addaccount R.SMIJOH D "Smith, John"
        Then add them to the AR report with total-to set to ARDET.
        """
        if not self._require_books():
            return
        try:
            result = models.setup_detailed_ar()
            print(f"  {result}")
        except ValueError as e:
            print(f"  Error: {e}")
        except Exception as e:
            print(f"  Error setting up AR subledger: {e}")

    def do_setupap(self, arg):
        """Setup a Detailed Accounts Payable subledger report.
        Usage: setupap

        Creates an AP.SUB report with 3 sample vendor accounts (Bauer, CCM, Warrior),
        a total account (APDET), and links it to the Balance Sheet via AP.DET → AP.TOT.

        The sample accounts use the P.XYZABC naming convention:
          P.BAUEQU  Bauer, Equipment
          P.CCMSPO  CCM, Sports
          P.WARHOC  Warrior, Hockey

        After setup, add real vendors with: addaccount P.SMISUP C "Smith, Supply"
        Then add them to the AP.SUB report with total-to set to APDET.
        """
        if not self._require_books():
            return
        try:
            result = models.setup_detailed_ap()
            print(f"  {result}")
        except ValueError as e:
            print(f"  Error: {e}")
        except Exception as e:
            print(f"  Error setting up AP subledger: {e}")

    def do_editaccount(self, arg):
        """Edit an account's description or account number.
        Usage: editaccount <name> [--desc "text"] [--num "1000"]

        Examples:
          editaccount EX.RENT --desc "Office Rent"
          editaccount BANK.CHQ --num "1000"
          editaccount EX.RENT --desc "Office Rent" --num "5200"
        """
        if not self._require_books():
            return

        parts = _split_args(arg)
        if not parts:
            print('  Usage: editaccount <name> [--desc "text"] [--num "1000"]')
            print('  Example: editaccount EX.RENT --desc "Office Rent" --num "5200"')
            return

        acct_name = parts[0]
        acct = resolve_account(acct_name)
        if not acct:
            return

        # Parse --desc and --num flags
        new_desc = None
        new_num = None
        i = 1
        while i < len(parts):
            if parts[i] == '--desc' and i + 1 < len(parts):
                new_desc = parts[i + 1]
                i += 2
            elif parts[i] == '--num' and i + 1 < len(parts):
                new_num = parts[i + 1]
                i += 2
            else:
                print(f"  Unknown option: '{parts[i]}'")
                print('  Use --desc "text" to change description')
                print('  Use --num "1000" to change account number')
                return

        if new_desc is None and new_num is None:
            print("  Nothing to change. Specify --desc and/or --num.")
            print(f"  Current description: {acct['description']}")
            print(f"  Current number:      {acct['account_number'] or '(none)'}")
            return

        try:
            models.update_account(acct['id'], description=new_desc, account_number=new_num)
            print(f"  ✓ Account updated: {acct['name']}")
            if new_desc is not None:
                print(f"    Description: {new_desc}")
            if new_num is not None:
                print(f"    Number:      {new_num}")
        except Exception as e:
            print(f"  Error updating account: {e}")

    # ─── ledger ──────────────────────────────────────────────────

    def do_ledger(self, arg):
        """Show account ledger. Usage: ledger <account> [from_date] [to_date]
        Amounts are raw: positive = debit, negative (parens) = credit."""
        if not self._require_books():
            return
        parts = arg.strip().split()
        if not parts:
            print("  Usage: ledger <account> [from_date] [to_date]")
            print("  Example: ledger BANK.CHQ 2025-01-01 2025-03-31")
            return

        acct = resolve_account(parts[0])
        if not acct:
            return

        date_from = parts[1] if len(parts) > 1 else None
        date_to = parts[2] if len(parts) > 2 else None

        # Validate dates if provided
        if date_from and not _normalize_date(date_from):
            print(f"  Invalid from-date: '{date_from}'")
            print("  Use YYYY-MM-DD format, e.g. 2025-01-01")
            return
        if date_to and not _normalize_date(date_to):
            print(f"  Invalid to-date: '{date_to}'")
            print("  Use YYYY-MM-DD format, e.g. 2025-12-31")
            return
        if date_from:
            date_from = _normalize_date(date_from)
        if date_to:
            date_to = _normalize_date(date_to)

        # Get raw ledger data (bypass normal_balance sign flipping)
        with models.get_db() as db:
            sql = """
                SELECT t.id as txn_id, t.date, t.reference, t.description as txn_desc,
                       l.amount, l.description as line_desc, l.id as line_id, l.reconciled,
                       l.doc_on_file,
                       GROUP_CONCAT(DISTINCT a2.name) as cross_accounts,
                       (SELECT COUNT(*) FROM lines WHERE transaction_id = t.id) as line_count
                FROM lines l
                JOIN transactions t ON l.transaction_id = t.id
                LEFT JOIN lines l2 ON l2.transaction_id = t.id AND l2.account_id != ?
                LEFT JOIN accounts a2 ON l2.account_id = a2.id
                WHERE l.account_id = ?"""
            params = [acct['id'], acct['id']]
            if date_from: sql += " AND t.date >= ?"; params.append(date_from)
            if date_to: sql += " AND t.date <= ?"; params.append(date_to)
            sql += " GROUP BY l.id ORDER BY t.date, t.id, l.sort_order"
            entries = db.execute(sql, params).fetchall()

        print(f"\n  Ledger: {acct['name']} — {acct['description']}")
        print(f"  Normal balance: {'Debit' if acct['normal_balance'] == 'D' else 'Credit'}")
        print(f"  Display: raw (+Dr / -Cr)")
        if date_from or date_to:
            print(f"  Period: {date_from or 'start'} to {date_to or 'end'}")
        print()

        if not entries:
            print("  (no entries)")
            if date_from or date_to:
                print("  Try a wider date range, or omit dates to see all entries.")
            return

        rows = []
        running = 0
        for e in entries:
            raw_amt = e['amount']  # Raw: positive=Dr, negative=Cr
            running += raw_amt
            r_flag = '✓' if e['reconciled'] else ''
            d_flag = '◉' if e['doc_on_file'] else ''
            rows.append((
                e['date'],
                (e['reference'] or '')[:8],
                (e['line_desc'] or e['txn_desc'] or '')[:35],
                (e['cross_accounts'] or '')[:20],
                fmt(raw_amt),
                fmt(running),
                r_flag,
                d_flag
            ))

        table(
            ['Date', 'Ref', 'Description', 'Contra', 'Amount', 'Balance', 'R', 'D'],
            rows,
            'llllrrll'
        )

        print(f"\n  Ending balance: {fmt(running)}  |  {len(entries)} entries")

    # ─── transactions ────────────────────────────────────────────

    def do_post(self, arg):
        """Post a simple 2-line transaction.
        Usage: post <date> <description> <amount> <debit_acct> <credit_acct>
        Example: post 2025-03-01 "March rent" 1500.00 EX.RENT BANK.CHQ
        """
        if not self._require_books():
            return

        # Parse: handle quoted description
        parts = _split_args(arg)
        if len(parts) < 5:
            print('  Usage: post <date> <description> <amount> <debit_acct> <credit_acct>')
            print('  Example: post 2025-03-01 "March rent" 1500.00 EX.RENT BANK.CHQ')
            return

        date_str = parts[0]
        desc = parts[1]
        amount_str = parts[2]
        dr_name = parts[3]
        cr_name = parts[4]

        # Validate date
        normalized = _normalize_date(date_str)
        if not normalized:
            print(f"  Invalid date: '{date_str}'")
            print("  Use YYYY-MM-DD format, e.g. 2025-03-01")
            return
        date_str = normalized

        # Check lock date and ceiling
        if not self._check_lock_date(date_str):
            return
        if not self._check_ceiling_date(date_str):
            return

        # Validate amount
        try:
            cents = parse_amount(amount_str)
        except Exception:
            print(f"  Invalid amount: '{amount_str}'")
            print("  Valid formats: 1500, 1500.00, 1,500.00, (500), -500")
            return
        if cents <= 0:
            print(f"  Amount must be positive: {amount_str}")
            print("  The debit/credit accounts determine which side gets debited.")
            return

        # Resolve accounts
        dr_acct = resolve_account(dr_name)
        if not dr_acct:
            return
        cr_acct = resolve_account(cr_name)
        if not cr_acct:
            return

        # Check posting accounts
        if not self._check_posting_account(dr_acct):
            return
        if not self._check_posting_account(cr_acct):
            return

        if not desc or not desc.strip():
            print("  ⚠ Warning: blank description. Transactions should have descriptions for audit trail.")

        try:
            txn_id = models.add_simple_transaction(date_str, '', desc, dr_acct['id'], cr_acct['id'], cents)
            txn, lines = models.get_transaction(txn_id)
            print(f"  ✓ Posted: #{txn_id} {date_str} | {desc} | {fmt(cents)}")
            print(f"    Dr {dr_acct['name']}  {fmt(cents)}")
            print(f"    Cr {cr_acct['name']}  ({fmt(cents)})")
            print(f"    Ref: {txn['reference']}")
        except ValueError as e:
            msg = str(e)
            print(f"  Error: {msg}")
            if 'lock' in msg.lower():
                print("  Use 'lock' to check the current lock date.")
            elif 'balance' in msg.lower():
                print("  Check that your debit and credit amounts are equal.")

    def do_postx(self, arg):
        """Post a multi-line transaction (interactive).
        Usage: postx <date> <description>
        Then enter lines as: <account> <amount>
        Positive = debit, negative = credit. Type 'done' to finish.
        """
        if not self._require_books():
            return

        parts = _split_args(arg)
        if len(parts) < 2:
            print('  Usage: postx <date> <description>')
            print('  Then enter lines. Positive=debit, negative=credit.')
            print('  Example: postx 2025-03-01 "Monthly payroll"')
            return

        date_str = parts[0]
        desc = parts[1]

        # Validate date
        normalized = _normalize_date(date_str)
        if not normalized:
            print(f"  Invalid date: '{date_str}'")
            print("  Use YYYY-MM-DD format, e.g. 2025-03-01")
            return
        date_str = normalized

        # Check lock date and ceiling
        if not self._check_lock_date(date_str):
            return
        if not self._check_ceiling_date(date_str):
            return

        if not desc or not desc.strip():
            print("  ⚠ Warning: blank description. Transactions should have descriptions for audit trail.")

        lines = []
        running = 0
        print(f"  Transaction: {date_str} | {desc}")
        print(f"  Enter lines as: <account> <amount>  (positive=Dr, negative=Cr)")
        print(f"  Type 'done' to post, 'cancel' to abort.\n")

        while True:
            remaining = -running
            prompt = f"  [{fmt_plain(running)} off] line> "
            try:
                line_input = input(prompt).strip()
            except (EOFError, KeyboardInterrupt):
                print("\n  Cancelled.")
                return

            if line_input.lower() == 'cancel':
                print("  Cancelled.")
                return
            if line_input.lower() == 'done':
                break

            line_parts = _split_args(line_input)
            if len(line_parts) < 2:
                print("    Format: <account> <amount>")
                print("    Example: EX.RENT 1500  or  BANK.CHQ -1500")
                continue

            acct = resolve_account(line_parts[0])
            if not acct:
                continue

            # Check posting account
            if not self._check_posting_account(acct):
                continue

            try:
                cents = parse_amount(line_parts[1])
            except Exception:
                print(f"    Invalid amount: '{line_parts[1]}'")
                print("    Valid formats: 1500, -1500, 1500.00, (500)")
                continue

            line_desc = ' '.join(line_parts[2:]) if len(line_parts) > 2 else desc
            lines.append((acct['id'], cents, line_desc))
            running += cents
            side = 'Dr' if cents > 0 else 'Cr'
            print(f"    {side} {acct['name']}  {fmt(abs(cents))}")

        if not lines:
            print("  No lines entered. Cancelled.")
            return

        # Check balance
        if running != 0:
            print(f"\n  Transaction does not balance.")
            print(f"  Total debits minus credits = {fmt(running)}")
            if running > 0:
                print(f"  You need {fmt(running)} more in credits (negative amounts).")
            else:
                print(f"  You need {fmt(abs(running))} more in debits (positive amounts).")
            print("  Transaction not posted. Try again with 'postx'.")
            return

        try:
            txn_id = models.add_transaction(date_str, '', desc, lines)
            txn, _ = models.get_transaction(txn_id)
            print(f"\n  ✓ Posted: #{txn_id} | Ref: {txn['reference']} | {len(lines)} lines")
        except ValueError as e:
            msg = str(e)
            print(f"  Error: {msg}")
            if 'balance' in msg.lower():
                print("  Debits must equal credits. Check your amounts.")

    def do_edit(self, arg):
        """Show transaction details. Usage: edit <txn_id>"""
        if not self._require_books():
            return
        if not arg.strip():
            print("  Usage: edit <txn_id>")
            print("  Use 'search <query>' to find transaction IDs.")
            return
        if not arg.strip().isdigit():
            print(f"  Invalid transaction ID: '{arg.strip()}'")
            print("  Transaction IDs are numbers. Use 'search <query>' to find them.")
            return
        txn, lines = models.get_transaction(int(arg.strip()))
        if not txn:
            print(f"  Transaction #{arg.strip()} not found.")
            print("  Use 'search <query>' to find transactions.")
            return

        print(f"\n  Transaction #{txn['id']}")
        print(f"  Date:   {txn['date']}")
        print(f"  Ref:    {txn['reference']}")
        print(f"  Desc:   {txn['description']}")
        print()

        rows = []
        for l in lines:
            side = 'Dr' if l['amount'] > 0 else 'Cr'
            rows.append((
                side,
                l['account_name'],
                fmt(abs(l['amount'])),
                l['description'] or '',
                '✓' if l['reconciled'] else '',
                '◉' if l['doc_on_file'] else ''
            ))
        table(['D/C', 'Account', 'Amount', 'Description', 'R', 'D'], rows, 'llrlll')

    def do_delete(self, arg):
        """Delete a transaction. Usage: delete <txn_id>"""
        if not self._require_books():
            return
        if not arg.strip():
            print("  Usage: delete <txn_id>")
            return
        if not arg.strip().isdigit():
            print(f"  Invalid transaction ID: '{arg.strip()}'")
            print("  Transaction IDs are numbers. Use 'search <query>' to find them.")
            return
        txn_id = int(arg.strip())
        txn, lines = models.get_transaction(txn_id)
        if not txn:
            print(f"  Transaction #{txn_id} not found.")
            return

        # Check lock date
        if not self._check_lock_date(txn['date']):
            return

        # Check for reconciled lines
        reconciled = [l for l in lines if l['reconciled']]
        if reconciled:
            print(f"  Warning: this transaction has {len(reconciled)} reconciled line(s).")

        models.delete_transaction(txn_id)
        print(f"  ✓ Deleted #{txn_id}: {txn['date']} | {txn['description']} | {len(lines)} lines")

    def do_reclass(self, arg):
        """Reclassify a suspense transaction. Usage: reclass <txn_id> <account> [tax_code]"""
        if not self._require_books():
            return
        parts = arg.strip().split()
        if len(parts) < 2:
            print("  Usage: reclass <txn_id> <account> [tax_code]")
            print("  Example: reclass 42 EX.OFFICE G5")
            return
        if not parts[0].isdigit():
            print(f"  Invalid transaction ID: '{parts[0]}'")
            return
        txn_id = int(parts[0])
        target = parts[1]
        tax_code = parts[2] if len(parts) > 2 else ''
        try:
            r = models.reclassify_suspense(txn_id, target, tax_code)
            print(f"  ✓ Reclassified #{txn_id}: EX.SUSP → {r['new_account']}  ${r['amount_display']}")
            if r['tax_applied']:
                print(f"    Tax: {r['tax_applied']}  ${r['tax_amount_cents']/100:,.2f}  Net: ${r['net_amount_cents']/100:,.2f}")
            if r['rule_created']:
                print(f"    Auto-rule: '{r['rule_keyword']}' → {r['new_account']}")
            if r['warning']:
                print(f"    ⚠ {r['warning']}")
        except ValueError as e:
            print(f"  Error: {e}")

    def do_search(self, arg):
        """Search transactions. Usage: search <query>"""
        if not self._require_books():
            return
        if not arg.strip():
            print("  Usage: search <query>")
            print("  Searches transaction descriptions and references.")
            print("  Example: search rent")
            return
        results = models.search_transactions(arg.strip())
        if not results:
            print(f"  No transactions matching '{arg.strip()}'.")
            print("  Try a shorter or different search term.")
            return
        rows = []
        for r in results:
            rows.append((
                str(r['txn_id']),
                r['date'],
                r['reference'] or '',
                (r['description'] or '')[:35],
                r['accounts'][:30] if r['accounts'] else '',
                fmt(r['total_amount'] or 0)
            ))
        table(['ID', 'Date', 'Ref', 'Description', 'Accounts', 'Amount'], rows, 'llllrl')

    # ─── import ──────────────────────────────────────────────────

    def do_importcsv(self, arg):
        """Import a bank CSV file. Applies rules automatically.
        Usage: importcsv <csvfile> <bank_account>
        Example: importcsv jan2025.csv BANK.CHQ

        Accepts two CSV formats:
          Simple:     Date, Description, Amount (or Date, Description, Debit, Credit)
          Bank export: Multi-column (auto-detects Date, Description, Amount columns)

        Automatically repairs rows where unquoted commas in descriptions
        cause extra fields (common in bank exports). Positive = deposits,
        negative = payments.
        """
        if not self._require_books():
            return

        parts = _split_args(arg)
        if len(parts) < 2:
            print("  Usage: importcsv <csvfile> <bank_account>")
            print("  Example: importcsv statements/jan2025.csv BANK.CHQ")
            print()
            print("  CSV format: Date, Description, Amount")
            print("          or: Date, Description, Debit, Credit")
            print("          or: Multi-column bank export (auto-detected)")
            return

        csv_path = os.path.expanduser(parts[0])
        if not _check_workspace(csv_path):
            return
        if not os.path.exists(csv_path):
            print(f"  File not found: {csv_path}")
            print("  Check the file path and try again.")
            return

        bank_acct = resolve_account(parts[1])
        if not bank_acct:
            return

        if not self._check_posting_account(bank_acct):
            return

        # Read CSV
        try:
            with open(csv_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                rows_raw = list(reader)
        except Exception as e:
            print(f"  Cannot read CSV file: {e}")
            return

        if not rows_raw:
            print("  Empty CSV file.")
            return

        # Normalize CSV: detect format, repair extra fields, auto-detect columns
        has_header, data_rows, csv_repairs = _normalize_csv(rows_raw)

        if not data_rows:
            print("  No data rows in CSV (only a header).")
            return

        if csv_repairs:
            print(f"  Auto-repaired {len(csv_repairs)} row(s) with extra fields (unquoted commas).")

        # Build row dicts for import_rows: parse amounts from CSV columns
        import_data = []
        parse_errors = []
        for row_num, row in enumerate(data_rows, start=2 if has_header else 1):
            if len(row) < 3:
                parse_errors.append((row_num, "Too few columns", ' | '.join(row)))
                continue

            row_date = row[0].strip()
            row_desc = row[1].strip()

            if not row_desc:
                parse_errors.append((row_num, "Missing description", row_date))
                continue

            # Get amount
            if len(row) >= 4 and (row[2].strip() or row[3].strip()):
                try:
                    dr = parse_amount(row[2]) if row[2].strip() else 0
                    cr = parse_amount(row[3]) if row[3].strip() else 0
                    amount_cents = dr - cr
                except Exception:
                    parse_errors.append((row_num, f"Bad amount '{row[2].strip()}/{row[3].strip()}'", row_desc[:30]))
                    continue
            else:
                try:
                    amount_cents = parse_amount(row[2])
                except Exception:
                    parse_errors.append((row_num, f"Bad amount '{row[2].strip()}'", row_desc[:30]))
                    continue

            import_data.append({
                'date': row_date,
                'description': row_desc,
                'amount_cents': amount_cents,
            })

        result = models.import_rows(bank_acct['id'], import_data)
        posted = result['posted']
        skipped = result['skipped'] + len(parse_errors)
        suspense = result['to_suspense']

        # Summary
        print(f"\n  Import complete: {csv_path}")
        print(f"    Rows processed: {len(data_rows)}")
        print(f"    Posted:         {posted}")
        print(f"    Skipped:        {skipped}")
        if csv_repairs:
            print(f"    Repaired:       {len(csv_repairs)}")
        if suspense:
            print(f"    To suspense:    {suspense}")

        # Show repairs
        if csv_repairs:
            print(f"\n  Repaired rows ({len(csv_repairs)}):")
            for row_num, extra, preview in csv_repairs[:10]:
                print(f"    Row {row_num}: merged {extra} extra field(s) — {preview}")
            if len(csv_repairs) > 10:
                print(f"    ... and {len(csv_repairs) - 10} more")

        # Show errors
        all_errors = [(r, reason, detail) for r, reason, detail in parse_errors]
        if result.get('errors'):
            for e in result['errors']:
                all_errors.append((e['row'], e['reason'], ''))
        if all_errors:
            print(f"\n  Errors ({len(all_errors)}):")
            for row_num, reason, detail in all_errors[:20]:
                msg = f"    Row {row_num}: {reason}"
                if detail:
                    msg += f" — {detail}"
                print(msg)
            if len(all_errors) > 20:
                print(f"    ... and {len(all_errors) - 20} more")

        # Show possible duplicates
        dupes = result.get('possible_duplicates', [])
        if dupes:
            print(f"\n  ⚠ Possible duplicates ({len(dupes)}):")
            print(f"    These rows match existing transactions (same date + amount).")
            print(f"    They were posted — review and delete if they are actual duplicates.")
            for d in dupes[:10]:
                print(f"    Row {d['row']}: {d['date']}  {models.fmt_amount(abs(d['amount'])):>12s}  {d['description']}")
            if len(dupes) > 10:
                print(f"    ... and {len(dupes) - 10} more")

        # Next steps
        if suspense:
            print(f"\n  {suspense} items went to suspense (no matching rule).")
            print("  Review them: ledger EX.SUSP")
            print("  Add rules to prevent this: addrule <keyword> <account>")
        if posted:
            print(f"\n  Verify the import: ledger {bank_acct['name']}")

    def do_importofx(self, arg):
        """Import a bank OFX/QBO file. Applies rules automatically.
        Usage: importofx <ofxfile> <bank_account>
        Example: importofx downloads/jan2025.qbo BANK.CHQ
        """
        if not self._require_books():
            return

        parts = _split_args(arg)
        if len(parts) < 2:
            print("  Usage: importofx <ofxfile> <bank_account>")
            print("  Example: importofx downloads/jan2025.qbo BANK.CHQ")
            return

        ofx_path = os.path.expanduser(parts[0])
        if not _check_workspace(ofx_path):
            return
        if not os.path.exists(ofx_path):
            print(f"  File not found: {ofx_path}")
            return

        bank_acct = resolve_account(parts[1])
        if not bank_acct:
            return

        if not self._check_posting_account(bank_acct):
            return

        try:
            rows = models.parse_ofx(ofx_path)
        except ValueError as e:
            print(f"  Cannot read OFX file: {e}")
            return

        result = models.import_rows(bank_acct['id'], rows)

        print(f"\n  Import complete: {ofx_path}")
        print(f"    Rows processed: {result['rows_processed']}")
        print(f"    Posted:         {result['posted']}")
        print(f"    Skipped:        {result['skipped']}")
        if result.get('fitid_skipped'):
            print(f"    FITID dupes:    {result['fitid_skipped']} (skipped — already imported)")
        if result['to_suspense']:
            print(f"    To suspense:    {result['to_suspense']}")

        if result.get('errors'):
            print(f"\n  Errors ({len(result['errors'])}):")
            for e in result['errors'][:20]:
                print(f"    Row {e['row']}: {e['reason']}")
            if len(result['errors']) > 20:
                print(f"    ... and {len(result['errors']) - 20} more")

        # Show possible duplicates
        dupes = result.get('possible_duplicates', [])
        if dupes:
            print(f"\n  ⚠ Possible duplicates ({len(dupes)}):")
            print(f"    These rows match existing transactions (same date + amount).")
            print(f"    They were posted — review and delete if they are actual duplicates.")
            for d in dupes[:10]:
                print(f"    Row {d['row']}: {d['date']}  {models.fmt_amount(abs(d['amount'])):>12s}  {d['description']}")
            if len(dupes) > 10:
                print(f"    ... and {len(dupes) - 10} more")

        if result['to_suspense']:
            print(f"\n  {result['to_suspense']} items went to suspense (no matching rule).")
            print("  Review them: ledger EX.SUSP")
            print("  Add rules to prevent this: addrule <keyword> <account>")
        if result['posted']:
            print(f"\n  Verify the import: ledger {bank_acct['name']}")

    def do_importgl(self, arg):
        """Import a pre-categorized general ledger CSV.
        Usage: importgl <csvfile> <bank_account>
        Example: importgl fy2024_bank_cdn.csv BANK.CDN

        CSV format: Date, Description, Amount, CrossAccount
        Positive = debit to bank_account, negative = credit.
        Cross-accounts must already exist. No rules applied.

        Use when converting from another accounting system (a legacy GL, QuickBooks GL,
        Sage, etc.) where every transaction already has its cross-account known.
        """
        if not self._require_books():
            return

        parts = _split_args(arg)
        if len(parts) < 2:
            print("  Usage: importgl <csvfile> <bank_account>")
            print("  Example: importgl fy2024_bank_cdn.csv BANK.CDN")
            print()
            print("  CSV format: Date, Description, Amount, CrossAccount")
            return

        csv_path = os.path.expanduser(parts[0])
        if not _check_workspace(csv_path):
            return
        if not os.path.exists(csv_path):
            print(f"  File not found: {csv_path}")
            return

        bank_acct = resolve_account(parts[1])
        if not bank_acct:
            return

        if not self._check_posting_account(bank_acct):
            return

        # Read CSV
        try:
            with open(csv_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                rows_raw = list(reader)
        except Exception as e:
            print(f"  Cannot read CSV file: {e}")
            return

        if not rows_raw:
            print("  Empty CSV file.")
            return

        # Detect header
        first = rows_raw[0]
        has_header = any(h.strip().lower() in ('date', 'description', 'amount', 'crossaccount', 'cross_account')
                         for h in first)
        data_rows = rows_raw[1:] if has_header else rows_raw

        if not data_rows:
            print("  No data rows in CSV (only a header).")
            return

        # Build row dicts
        import_data = []
        parse_errors = []
        for row_num, row in enumerate(data_rows, start=2 if has_header else 1):
            if len(row) < 4:
                parse_errors.append((row_num, f"Need 4 columns, got {len(row)}", ' | '.join(row)))
                continue

            row_date = row[0].strip()
            row_desc = row[1].strip()
            cross_acct = row[3].strip()

            if not row_desc:
                parse_errors.append((row_num, "Missing description", row_date))
                continue

            if not cross_acct:
                parse_errors.append((row_num, "Missing cross-account", row_desc[:30]))
                continue

            try:
                amount_cents = parse_amount(row[2])
            except Exception:
                parse_errors.append((row_num, f"Bad amount '{row[2].strip()}'", row_desc[:30]))
                continue

            import_data.append({
                'date': row_date,
                'description': row_desc,
                'amount_cents': amount_cents,
                'cross_account': cross_acct,
            })

        result = models.import_gl_rows(bank_acct['id'], import_data)
        posted = result['posted']
        skipped = result['skipped'] + len(parse_errors)

        # Summary
        print(f"\n  GL import complete: {csv_path}")
        print(f"    Rows processed: {len(data_rows)}")
        print(f"    Posted:         {posted}")
        print(f"    Skipped:        {skipped}")

        # Show errors
        all_errors = [(r, reason, detail) for r, reason, detail in parse_errors]
        if result.get('errors'):
            for e in result['errors']:
                all_errors.append((e['row'], e['reason'], ''))
        if all_errors:
            print(f"\n  Errors ({len(all_errors)}):")
            for row_num, reason, detail in all_errors[:20]:
                msg = f"    Row {row_num}: {reason}"
                if detail:
                    msg += f" — {detail}"
                print(msg)
            if len(all_errors) > 20:
                print(f"    ... and {len(all_errors) - 20} more")

        # Show possible duplicates
        dupes = result.get('possible_duplicates', [])
        if dupes:
            print(f"\n  Possible duplicates ({len(dupes)}):")
            print(f"    These rows match existing transactions (same date + amount).")
            print(f"    They were posted — review and delete if they are actual duplicates.")
            for d in dupes[:10]:
                print(f"    Row {d['row']}: {d['date']}  {models.fmt_amount(abs(d['amount'])):>12s}  {d['description']}")
            if len(dupes) > 10:
                print(f"    ... and {len(dupes) - 10} more")

        if posted:
            print(f"\n  Verify the import: ledger {bank_acct['name']}")

    def do_importaje(self, arg):
        """Import CaseWare AJE export (IIF or Venice format).
        Usage: importaje <file> [ref_prefix]
        Example: importaje cwquickb.iif 25AJE

        Parses entries, maps CsW account names to Grid accounts,
        then posts all AJEs as multi-line transactions.
        """
        if not self._require_books():
            return

        parts = _split_args(arg)
        if len(parts) < 1:
            print("  Usage: importaje <file> [ref_prefix]")
            print("  Example: importaje cwquickb.iif 25AJE")
            print()
            print("  Supported formats:")
            print("    IIF  — QuickBooks export (TRNS/SPL/ENDTRNS)")
            print("    TXT  — Venice/MYOB export (STOP delimited)")
            return

        file_path = os.path.expanduser(parts[0])
        if not _check_workspace(file_path):
            return
        if not os.path.exists(file_path):
            print(f"  File not found: {file_path}")
            return

        # Parse file
        try:
            parsed = models.parse_csw_aje(file_path)
        except ValueError as e:
            print(f"  Cannot parse file: {e}")
            return

        n = len(parsed['entries'])
        if n == 0:
            print("  No entries found in file.")
            return

        print(f"\n  Parsed {n} entries ({parsed['format'].upper()} format)")
        print(f"  Unique CsW accounts: {len(parsed['csw_accounts'])}")

        # Auto-match accounts
        suggestions = models.auto_match_accounts(parsed['csw_accounts'])

        # Show mapping table
        unmatched = []
        print(f"\n  Account Mapping:")
        print(f"  {'CsW Account':<45} {'CsW #':<6} {'Grid Account'}")
        print(f"  {'─' * 45} {'─' * 5} {'─' * 20}")
        for csw in parsed['csw_accounts']:
            csw_name = csw['name']
            csw_num = csw.get('number', '')
            match = suggestions.get(csw_name)
            if match:
                print(f"  {csw_name:<45} {csw_num:<6} {match['name']}")
            else:
                print(f"  {csw_name:<45} {csw_num:<6} ???")
                unmatched.append(csw)

        # Prompt for unmapped accounts
        account_map = {}
        for csw_name, match in suggestions.items():
            if match:
                account_map[csw_name] = match['id']

        if unmatched:
            print(f"\n  {len(unmatched)} account(s) need manual mapping:")
            for csw in unmatched:
                csw_name = csw['name']
                csw_num = csw.get('number', '')
                hint = f" [{csw_num}]" if csw_num else ""
                while True:
                    try:
                        user_input = input(f"  Grid account for \"{csw_name}\"{hint}: ").strip()
                    except (EOFError, KeyboardInterrupt):
                        print("\n  Import cancelled.")
                        return
                    if not user_input:
                        print("    Account is required. Type 'accounts' to see list, or Ctrl+C to cancel.")
                        continue
                    acct = resolve_account(user_input)
                    if acct:
                        account_map[csw_name] = acct['id']
                        print(f"    → {acct['name']} ({acct['description']})")
                        break

        # Determine ref prefix
        if len(parts) >= 2:
            ref_prefix = parts[1]
        else:
            fy = models.get_meta('fiscal_year', '')
            default_prefix = f"{fy[-2:]}AJE" if fy else 'AJE'
            try:
                user_prefix = input(f"  AJE ref prefix [{default_prefix}]: ").strip()
            except (EOFError, KeyboardInterrupt):
                print("\n  Import cancelled.")
                return
            ref_prefix = user_prefix if user_prefix else default_prefix

        # Confirm
        print(f"\n  Ready to post {n} entries with prefix '{ref_prefix}'")
        try:
            confirm = input("  Post all entries? [y/N]: ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print("\n  Import cancelled.")
            return
        if confirm != 'y':
            print("  Import cancelled.")
            return

        # Post entries
        result = models.import_aje_entries(parsed['entries'], account_map, ref_prefix, journal_account=ref_prefix)

        print(f"\n  Import complete:")
        print(f"    Posted:   {result['posted']}")
        print(f"    Skipped:  {result['skipped']}")
        if result['errors']:
            print(f"\n  Errors ({len(result['errors'])}):")
            for e in result['errors'][:20]:
                print(f"    {e['entry']}: {e['reason']}")
            if len(result['errors']) > 20:
                print(f"    ... and {len(result['errors']) - 20} more")
        if result['posted']:
            print(f"\n  Verify: tb")

    # ─── reports ─────────────────────────────────────────────────

    def do_tb(self, arg):
        """Trial balance (raw: Dr column = positive, Cr column = negative).
        Usage: tb [as-of-date]"""
        if not self._require_books():
            return
        as_of = arg.strip() if arg.strip() else None

        if as_of:
            normalized = _normalize_date(as_of)
            if not normalized:
                print(f"  Invalid date: '{as_of}'")
                print("  Use YYYY-MM-DD format, e.g. 2025-12-31")
                return
            as_of = normalized

        # Raw trial balance: no normal_balance flipping
        with models.get_db() as db:
            accounts = db.execute("SELECT * FROM accounts WHERE account_type='posting' ORDER BY name").fetchall()

        name = models.get_meta('company_name', '')
        print(f"\n  Trial Balance{' — ' + name if name else ''}")
        if as_of:
            print(f"  As of: {as_of}")
        print(f"  Display: raw (+Dr / -Cr)")
        print()

        rows = []
        total_dr = 0
        total_cr = 0
        for acct in accounts:
            raw = models.get_account_balance(acct['id'], date_to=as_of)
            if raw == 0:
                continue
            if raw > 0:
                total_dr += raw
                rows.append((
                    acct['account_number'] or '',
                    acct['name'],
                    fmt(raw),
                    ''
                ))
            else:
                total_cr += abs(raw)
                rows.append((
                    acct['account_number'] or '',
                    acct['name'],
                    '',
                    fmt(abs(raw))
                ))

        if not rows:
            print("  (no balances)")
            return

        table(['No.', 'Account', 'Debit', 'Credit'], rows, 'llrr')

        print(f"  {'─' * 60}")
        dr_str = fmt(total_dr)
        cr_str = fmt(total_cr)
        print(f"  {'TOTALS':<35} {dr_str:>12}  {cr_str:>12}")
        diff = total_dr - total_cr
        if diff != 0:
            print(f"  *** OUT OF BALANCE BY {fmt(abs(diff))} ***")

    def do_reports(self, arg):
        """List all reports."""
        if not self._require_books():
            return
        reports = models.get_reports()
        if not reports:
            print("  No reports defined.")
            return
        rows = [(str(r['id']), r['name'], r['description']) for r in reports]
        table(['ID', 'Name', 'Description'], rows)

    def do_editreport(self, arg):
        """Edit a report's description.
        Usage: editreport <name> --desc "New Description"

        Examples:
          editreport BS --desc "Balance Sheet - TecToc Energy Ltd"
          editreport IS --desc "Income Statement 2026"
        """
        if not self._require_books():
            return

        parts = _split_args(arg)
        if not parts:
            print('  Usage: editreport <name> --desc "New Description"')
            print('  Example: editreport BS --desc "Balance Sheet - TecToc Energy Ltd"')
            return

        rpt_name = parts[0]
        rpt = models.find_report_by_name(rpt_name)
        if not rpt:
            print(f"  Report not found: '{rpt_name}'")
            reports = models.get_reports()
            if reports:
                print(f"  Available: {', '.join(r['name'] for r in reports)}")
            return

        # Parse --desc flag
        new_desc = None
        i = 1
        while i < len(parts):
            if parts[i] == '--desc' and i + 1 < len(parts):
                new_desc = parts[i + 1]
                i += 2
            else:
                print(f"  Unknown option: '{parts[i]}'")
                print('  Use --desc "text" to change description')
                return

        if new_desc is None:
            print("  Nothing to change. Specify --desc.")
            print(f"  Current description: {rpt['description']}")
            return

        try:
            old_desc = rpt['description']
            models.update_report(rpt['id'], description=new_desc)
            print(f"  Report updated: {rpt['name']}")
            print(f"    Old: {old_desc}")
            print(f"    New: {new_desc}")
        except Exception as e:
            print(f"  Error updating report: {e}")

    def do_report(self, arg):
        """Run a report. Usage: report <name> [from_date] [to_date]"""
        if not self._require_books():
            return
        parts = arg.strip().split()
        if not parts:
            print("  Usage: report <name> [from_date] [to_date]")
            reports = models.get_reports()
            if reports:
                print(f"  Available: {', '.join(r['name'] for r in reports)}")
            return

        rpt = models.find_report_by_name(parts[0])
        if not rpt:
            print(f"  Report not found: '{parts[0]}'")
            reports = models.get_reports()
            if reports:
                print(f"  Available: {', '.join(r['name'] for r in reports)}")
            return

        date_from = parts[1] if len(parts) > 1 else rpt.get('period_begin', '')
        date_to = parts[2] if len(parts) > 2 else rpt.get('period_end', '')

        data = models.compute_report_column(rpt['id'], date_from or None, date_to or None)

        name = models.get_meta('company_name', '')
        print(f"\n  {rpt['name']} — {rpt['description']}")
        if name:
            print(f"  {name}")
        if date_from or date_to:
            print(f"  Period: {date_from or 'start'} to {date_to or 'end'}")
        print()

        for item, balance in data:
            itype = item['item_type']
            desc = item['description'] or item.get('acct_name', '') or ''
            indent = item.get('indent', 0) or 0
            pad = '  ' * indent

            if itype == 'separator':
                style = item.get('sep_style', '')
                if style == 'double':
                    print(f"  {'':<40} {'═' * 14}")
                elif style == 'single':
                    print(f"  {'':<40} {'─' * 14}")
                elif style == 'blank':
                    print()
                else:
                    print()
            elif itype == 'label':
                print(f"  {pad}{desc}")
            else:
                # account or total
                bal_str = fmt(balance) if balance != 0 else '—'
                acct_name = item.get('acct_name', '') or ''
                display = desc if desc else acct_name
                print(f"  {pad}{display:<40} {bal_str:>14}")

    def do_balance(self, arg):
        """Show account balance (raw: +Dr, -Cr). Usage: balance <account> [from] [to]"""
        if not self._require_books():
            return
        parts = arg.strip().split()
        if not parts:
            print("  Usage: balance <account> [from_date] [to_date]")
            print("  Example: balance BANK.CHQ")
            return

        acct = resolve_account(parts[0])
        if not acct:
            return

        date_from = parts[1] if len(parts) > 1 else None
        date_to = parts[2] if len(parts) > 2 else None

        raw = models.get_account_balance(acct['id'], date_from, date_to)
        print(f"  {acct['name']}: {fmt(raw)}")

    # ─── export ──────────────────────────────────────────────────

    def do_exportcsv(self, arg):
        """Export a report to CSV file.
        Usage: exportcsv <report> [filename] [from_date] [to_date]

        Columns: Description, Account, Type, Balance
        If no filename given, defaults to <report_name>.csv

        Examples:
          exportcsv BS
          exportcsv IS income_2025.csv 2025-01-01 2025-12-31
        """
        if not self._require_books():
            return

        parts = _split_args(arg)
        if not parts:
            print("  Usage: exportcsv <report> [filename] [from_date] [to_date]")
            reports = models.get_reports()
            if reports:
                print(f"  Available reports: {', '.join(r['name'] for r in reports)}")
            return

        rpt = models.find_report_by_name(parts[0])
        if not rpt:
            print(f"  Report not found: '{parts[0]}'")
            reports = models.get_reports()
            if reports:
                print(f"  Available: {', '.join(r['name'] for r in reports)}")
            return

        # Parse remaining args: [filename] [from_date] [to_date]
        filename = None
        date_from = None
        date_to = None

        remaining = parts[1:]
        if remaining:
            # If first remaining arg looks like a filename (has extension or not a date)
            if remaining[0].endswith('.csv') or (not _normalize_date(remaining[0]) and len(remaining[0]) != 10):
                filename = remaining[0]
                remaining = remaining[1:]

            if remaining:
                date_from = _normalize_date(remaining[0])
                if not date_from:
                    print(f"  Invalid from-date: '{remaining[0]}'")
                    print("  Use YYYY-MM-DD format.")
                    return
            if len(remaining) > 1:
                date_to = _normalize_date(remaining[1])
                if not date_to:
                    print(f"  Invalid to-date: '{remaining[1]}'")
                    print("  Use YYYY-MM-DD format.")
                    return

        if not filename:
            filename = f"{rpt['name']}.csv"

        if not _check_workspace(os.path.abspath(filename)):
            return

        if not date_from:
            date_from = rpt.get('period_begin', '') or None
        if not date_to:
            date_to = rpt.get('period_end', '') or None

        data = models.compute_report_column(rpt['id'], date_from, date_to)

        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Description', 'Account', 'Type', 'Balance'])

                for item, balance in data:
                    itype = item['item_type']
                    if itype == 'separator':
                        continue

                    desc = item['description'] or item.get('acct_name', '') or ''
                    acct_name = item.get('acct_name', '') or ''
                    bal_str = fmt_plain(balance) if balance != 0 else '0.00'

                    writer.writerow([desc, acct_name, itype, bal_str])

            row_count = sum(1 for item, _ in data if item['item_type'] != 'separator')
            print(f"  ✓ Exported: {filename}")
            print(f"    Report: {rpt['name']} — {rpt['description']}")
            if date_from or date_to:
                print(f"    Period: {date_from or 'start'} to {date_to or 'end'}")
            print(f"    Rows:   {row_count}")
        except OSError as e:
            print(f"  Cannot write file: {e}")

    def do_exportwriteup(self, arg):
        """Export the write-up handoff (for Willy, the working-paper program).
        Usage: exportwriteup [filename]

        One JSON: CY + PY fully itemized, five fiscal years of per-account
        comparatives, the chart with leadsheet codes, the statement layouts,
        and the check_books verdict the write-up program's gate reads.
        Default filename: <Company>_<FY>_writeup.json beside the books.
        """
        if not self._require_books():
            return
        import json as _json
        try:
            payload = models.export_writeup()
        except ValueError as e:
            print(f"  {e}")
            return
        fname = arg.strip()
        if not fname:
            safe = payload['meta']['company_name'].replace(' ', '_')[:40]
            fname = os.path.join(
                os.path.dirname(models.get_db_path()),
                f"{safe}_{payload['meta']['fiscal_year']}_writeup.json")
        with open(fname, 'w') as f:
            _json.dump(payload, f, indent=1)
        m = payload['meta']
        print(f"  Wrote {fname}")
        print(f"  {m['itemized_transactions']} transactions itemized "
              f"({m['py_start']} .. {m['cy_end']}); "
              f"{m['transactions_before_window']} before / "
              f"{m['transactions_after_window']} after the window; "
              f"5 comparative years; {len(payload['accounts'])} accounts.")

    def do_exporttb(self, arg):
        """Export trial balance to CSV file.
        Usage: exporttb [filename] [as-of-date]

        Columns: Account Number, Name, Description, Normal Balance, Debit, Credit, Raw Balance

        Examples:
          exporttb
          exporttb trial_balance.csv
          exporttb trial_balance.csv 2025-12-31
        """
        if not self._require_books():
            return

        parts = _split_args(arg)
        filename = None
        as_of = None

        for p in parts:
            if p.endswith('.csv') or (not _normalize_date(p) and len(p) != 10):
                filename = p
            else:
                normalized = _normalize_date(p)
                if normalized:
                    as_of = normalized
                else:
                    print(f"  Invalid date: '{p}'")
                    print("  Use YYYY-MM-DD format, e.g. 2025-12-31")
                    return

        if not filename:
            filename = 'trial_balance.csv'

        if not _check_workspace(os.path.abspath(filename)):
            return

        with models.get_db() as db:
            accounts = db.execute("SELECT * FROM accounts WHERE account_type='posting' ORDER BY name").fetchall()

        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Account Number', 'Name', 'Description', 'Normal Balance', 'Debit', 'Credit', 'Raw Balance'])

                row_count = 0
                total_dr = 0
                total_cr = 0
                for acct in accounts:
                    raw = models.get_account_balance(acct['id'], date_to=as_of)
                    if raw == 0:
                        continue

                    if raw > 0:
                        dr_str = fmt_plain(raw)
                        cr_str = '0.00'
                        total_dr += raw
                    else:
                        dr_str = '0.00'
                        cr_str = fmt_plain(abs(raw))
                        total_cr += abs(raw)

                    writer.writerow([
                        acct['account_number'] or '',
                        acct['name'],
                        acct['description'],
                        acct['normal_balance'],
                        dr_str,
                        cr_str,
                        fmt_plain(raw)
                    ])
                    row_count += 1

            company = models.get_meta('company_name', '')
            print(f"  ✓ Exported: {filename}")
            if company:
                print(f"    Company: {company}")
            if as_of:
                print(f"    As of:   {as_of}")
            print(f"    Accounts with balances: {row_count}")
            diff = total_dr - total_cr
            if diff != 0:
                print(f"    *** OUT OF BALANCE BY {fmt(abs(diff))} ***")
        except OSError as e:
            print(f"  Cannot write file: {e}")

    # ─── rules ───────────────────────────────────────────────────

    def do_rules(self, arg):
        """List import rules."""
        if not self._require_books():
            return
        rules = models.get_import_rules()
        if not rules:
            print("  No import rules defined.")
            print("  Add one: addrule <keyword> <account> [tax_code] [priority]")
            return
        rows = [(str(r['id']), r['keyword'], r['account_name'],
                 r['tax_code'] or '', str(r['priority']), r['notes'] or '') for r in rules]
        table(['ID', 'Keyword', 'Account', 'Tax', 'Pri', 'Notes'], rows, 'llllrl')

    def do_addrule(self, arg):
        """Add an import rule. Usage: addrule <keyword> <account> [tax_code] [priority]
        Example: addrule NETFLIX EX.COMP G5 10"""
        if not self._require_books():
            return
        parts = _split_args(arg)
        if len(parts) < 2:
            print("  Usage: addrule <keyword> <account> [tax_code] [priority]")
            print("  Example: addrule NETFLIX EX.COMP G5 10")
            print()
            print("  keyword:  text to match in bank descriptions (case-insensitive)")
            print("  account:  account to post the expense/revenue to")
            print("  tax_code: tax code to apply (use 'taxcodes' to see options)")
            print("  priority: higher number = matched first (default: 0)")
            return

        keyword = parts[0]
        account = parts[1]
        tax = parts[2] if len(parts) > 2 else ''

        if len(parts) > 3:
            if parts[3].isdigit():
                pri = int(parts[3])
            else:
                print(f"  Invalid priority: '{parts[3]}'")
                print("  Priority must be a number (higher = matched first). Default is 0.")
                return
        else:
            pri = 0

        # Verify account exists
        acct = models.get_account_by_name(account)
        if not acct:
            print(f"  Warning: account '{account}' not found. Rule saved anyway.")
            print("  Use 'accounts' to check available account names.")

        # Verify tax code if provided
        if tax:
            tc = models.get_tax_code(tax) if hasattr(models, 'get_tax_code') else None
            if not tc:
                print(f"  Warning: tax code '{tax}' not found. Rule saved anyway.")
                print("  Use 'taxcodes' to see available tax codes.")

        models.save_import_rule(None, keyword, account, tax, pri)
        print(f"  ✓ Rule added: '{keyword}' → {account} (tax: {tax or 'none'}, priority: {pri})")

    def do_editrule(self, arg):
        """Edit an existing import rule.
        Usage: editrule <id> <keyword> <account> [tax_code] [priority]
        Example: editrule 5 NETFLIX EX.COMP G5 20
        """
        if not self._require_books():
            return

        parts = _split_args(arg)
        if len(parts) < 3:
            print("  Usage: editrule <id> <keyword> <account> [tax_code] [priority]")
            print("  Example: editrule 5 NETFLIX EX.COMP G5 20")
            print("  Use 'rules' to see rule IDs.")
            return

        if not parts[0].isdigit():
            print(f"  Invalid rule ID: '{parts[0]}'")
            print("  Use 'rules' to see rule IDs.")
            return

        rule_id = int(parts[0])
        keyword = parts[1]
        account = parts[2]
        tax = parts[3] if len(parts) > 3 else ''

        if len(parts) > 4:
            if parts[4].isdigit():
                pri = int(parts[4])
            else:
                print(f"  Invalid priority: '{parts[4]}'")
                print("  Priority must be a number (higher = matched first).")
                return
        else:
            pri = 0

        # Verify the rule exists
        existing_rules = models.get_import_rules()
        rule_exists = any(r['id'] == rule_id for r in existing_rules)
        if not rule_exists:
            print(f"  Rule #{rule_id} not found.")
            print("  Use 'rules' to see all rules and their IDs.")
            return

        # Verify account exists
        acct = models.get_account_by_name(account)
        if not acct:
            print(f"  Warning: account '{account}' not found. Rule saved anyway.")

        models.save_import_rule(rule_id, keyword, account, tax, pri)
        print(f"  ✓ Rule #{rule_id} updated: '{keyword}' → {account} (tax: {tax or 'none'}, priority: {pri})")

    def do_delrule(self, arg):
        """Delete an import rule. Usage: delrule <id>"""
        if not self._require_books():
            return

        arg = arg.strip()
        if not arg:
            print("  Usage: delrule <id>")
            print("  Use 'rules' to see rule IDs.")
            return

        if not arg.isdigit():
            print(f"  Invalid rule ID: '{arg}'")
            print("  Use 'rules' to see rule IDs.")
            return

        rule_id = int(arg)

        # Verify the rule exists
        existing_rules = models.get_import_rules()
        rule = None
        for r in existing_rules:
            if r['id'] == rule_id:
                rule = r
                break

        if not rule:
            print(f"  Rule #{rule_id} not found.")
            print("  Use 'rules' to see all rules and their IDs.")
            return

        models.delete_import_rule(rule_id)
        print(f"  ✓ Rule #{rule_id} deleted: '{rule['keyword']}' → {rule['account_name']}")

    # ─── reconciliation ──────────────────────────────────────────

    def do_reconcile(self, arg):
        """Show reconciliation summary (raw: +Dr, -Cr). Usage: reconcile <account>"""
        if not self._require_books():
            return
        if not arg.strip():
            print("  Usage: reconcile <account>")
            print("  Example: reconcile BANK.CHQ")
            return
        acct = resolve_account(arg.strip())
        if not acct:
            return

        # Raw reconciliation — no sign flipping
        with models.get_db() as db:
            book = db.execute(
                "SELECT COALESCE(SUM(l.amount),0) FROM lines l JOIN transactions t ON l.transaction_id=t.id WHERE l.account_id=?",
                (acct['id'],)).fetchone()[0]
            cleared = db.execute(
                "SELECT COALESCE(SUM(l.amount),0) FROM lines l JOIN transactions t ON l.transaction_id=t.id WHERE l.account_id=? AND l.reconciled=1",
                (acct['id'],)).fetchone()[0]

        print(f"\n  Reconciliation: {acct['name']}  (raw +Dr / -Cr)")
        print(f"  Book balance:     {fmt(book)}")
        print(f"  Cleared balance:  {fmt(cleared)}")
        print(f"  Uncleared:        {fmt(book - cleared)}")

    # ─── tax codes ───────────────────────────────────────────────

    def do_taxcodes(self, arg):
        """List tax codes."""
        if not self._require_books():
            return
        codes = models.get_tax_codes()
        if not codes:
            print("  No tax codes defined.")
            return
        rows = [(c['id'], c['description'], f"{c['rate_percent']}%",
                 c['collected_account'] or '', c['paid_account'] or '') for c in codes]
        table(['Code', 'Description', 'Rate', 'Collected', 'Paid'], rows)

    # ─── lock date ───────────────────────────────────────────────

    def do_lock(self, arg):
        """Show or set lock date. Usage: lock [YYYY-MM-DD]"""
        if not self._require_books():
            return
        if arg.strip():
            date_str = arg.strip()
            normalized = _normalize_date(date_str)
            if not normalized:
                print(f"  Invalid date: '{date_str}'")
                print("  Use YYYY-MM-DD format, e.g. 2025-12-31")
                return
            try:
                models.set_fiscal_settings(lock_date=normalized)
            except ValueError as e:
                print(f"  {e}")
                return
            print(f"  ✓ Lock date set: {normalized}")
            print("  Transactions on or before this date cannot be posted, edited, or deleted.")
        else:
            lock = models.get_meta('lock_date', '')
            if lock:
                print(f"  Lock date: {lock}")
                print("  Transactions on or before this date are protected.")
            else:
                print("  Lock date: (none)")
                print("  Set one with: lock YYYY-MM-DD")

    # ─── Validate ─────────────────────────────────────────────

    def do_validate(self, arg):
        """Validate the report total-to chain.
        Usage: validate

        Checks that the IS chain reaches RE on the BS, that RE.OB totals into RE
        and RE.OPEN/RE.CLOSE are computed display lines, that total report items
        have accounts linked, and that the BS balances (and BS RE == IS Closing RE).
        """
        if not self._require_books():
            return

        issues = models.validate_report_chain()
        errors = [i for i in issues if i['level'] == 'error']
        warnings = [i for i in issues if i['level'] == 'warning']
        ok = [i for i in issues if i['level'] == 'ok']

        if ok:
            print(f"\n  {ok[0]['message']}")

        if warnings:
            print(f"\n  Warnings ({len(warnings)}):")
            for w in warnings:
                print(f"    ⚠ {w['message']}")

        if errors:
            print(f"\n  Errors ({len(errors)}):")
            for e in errors:
                print(f"    ✗ {e['message']}")

        if not errors and not warnings:
            print("  All checks passed.")
        elif errors:
            print(f"\n  {len(errors)} error(s) found. The report chain has problems.")
        print()

    def do_trace(self, arg):
        """Trace the accumulation tree for a report account.
        Usage: trace <account> [as-of-date]

        Shows what feeds into the account total, with amounts and sources.
        Useful for understanding why a total shows a particular number.

        Examples:
          trace RE                      # All-time RE accumulation
          trace RE 2025-05-31           # RE as of a specific date
          trace NETEARN 2024-06-01 2025-05-31  # NETEARN for a date range
        """
        if not self._require_books():
            return

        parts = _split_args(arg)
        if not parts:
            print("  Usage: trace <account> [date_to | date_from date_to]")
            return

        acct_name = parts[0].upper()
        date_from = None
        date_to = None
        if len(parts) == 2:
            date_to = models.normalize_date(parts[1])
        elif len(parts) >= 3:
            date_from = models.normalize_date(parts[1])
            date_to = models.normalize_date(parts[2])

        result = models.trace_account(acct_name, date_from, date_to)

        if result['display'] == 0 and not result['contributors']:
            print(f"\n  {acct_name}: not found on any report or has no activity.")
            return

        nb = result['normal_balance']
        period = ""
        if date_from and date_to:
            period = f" ({date_from} to {date_to})"
        elif date_to:
            period = f" (as of {date_to})"

        print(f"\n  {acct_name} ({nb}-normal){period}")
        print(f"  Display: {models.fmt_amount(result['display'])}")
        if result['own_raw'] != 0:
            sign = 1 if nb == 'D' else -1
            print(f"  Own postings: {models.fmt_amount(result['own_raw'] * sign)}")

        if result['feeds_into']:
            print(f"  Feeds into: {', '.join(result['feeds_into'])}")

        if result['contributors']:
            print(f"\n  {'Account':<20s} {'Report':<6s} {'Own':>14s} {'+ Received':>14s} {'= Dumped':>14s}")
            print(f"  {'─'*20} {'─'*6} {'─'*14} {'─'*14} {'─'*14}")
            for c in result['contributors']:
                sign = 1 if nb == 'D' else -1
                own_d = models.fmt_amount(c['own_raw'] * sign) if c['own_raw'] else '—'
                acc_d = models.fmt_amount(c['accumulated'] * sign) if c['accumulated'] else '—'
                val_d = models.fmt_amount(c['display'])
                print(f"  {c['name']:<20s} {c['report']:<6s} {own_d:>14s} {acc_d:>14s} {val_d:>14s}")

            print(f"  {'─'*20} {'─'*6} {'─'*14} {'─'*14} {'─'*14}")
            total_dumped = sum(c['value_dumped'] for c in result['contributors'])
            sign = 1 if nb == 'D' else -1
            print(f"  {'Total':<27s} {'':>14s} {'':>14s} {models.fmt_amount(total_dumped * sign):>14s}")
        print()

    # ─── YE rollover ───────────────────────────────────────────




    def do_aje(self, arg):
        """Adjusting entries. Usage:
             aje                          — the years on file
             aje <journal>                — the adjustments in one year
             aje new [NAME] [description] — set a year up (suggests if omitted)
             aje post <journal> <file.csv>  — post ONE entry from a CSV
             aje del <journal> <ref>      — remove one entry, every leg of it

        The post CSV carries the entry's header on its first lines and then the
        legs — this is one adjustment, not a batch:

            ref,26AJE03
            date,2026-12-31
            description,Rcd 2026 tax provision
            EX.TAX,100.00
            GST.PAY,(80.00)
            AP,(20.00)

        ref and date may be omitted — the next reference in the journal's own
        sequence and the working year-end are used. Amount is signed as Grid
        displays: plain is a DEBIT, brackets or a leading minus is a CREDIT.
        NEVER include the journal account itself; Grid posts that side.
        """
        if not self._require_books():
            return
        parts = (arg or '').strip().split()

        if not parts:
            js = models.aje_journals()
            if not js:
                print("  No years of adjusting entries yet.")
                print("  Set one up with:  aje new")
                return
            for j in js:
                n = len(models.aje_groups(j['id']))
                bal = models.get_account_balance(j['id'])
                flag = '' if bal == 0 else f"  !! journal is {models.fmt_amount_plain(bal)}, should be 0.00"
                print(f"  {j['name']:<10} {(j['description'] or ''):<28} {n:>3} entr"
                      f"{'y' if n == 1 else 'ies'}   next {models.next_aje_ref(j['id'])}{flag}")
            return

        cmd = parts[0].lower()

        if cmd == 'new':
            sug = models.suggest_aje_batch()
            name = parts[1] if len(parts) > 1 else sug['account']
            desc = ' '.join(parts[2:]) if len(parts) > 2 else sug['description']
            try:
                acct = models.create_aje_batch(name, desc)
            except ValueError as e:
                print(f"  !! {e}")
                return
            print(f"  ✓ {acct['name']} — {acct['description']}")
            print(f"    First entry will be {models.next_aje_ref(acct['id'])}")
            return

        if cmd == 'del':
            if len(parts) < 3:
                print("  Usage: aje del <journal> <ref>")
                return
            acct = models.get_account_by_name(parts[1])
            if not acct or not models.is_aje_journal(acct['id']):
                print(f"  '{parts[1]}' is not a year of adjusting entries.")
                return
            try:
                n = models.delete_aje(acct['id'], parts[2])
                print(f"  ✓ {parts[2]} removed ({n} lines).")
            except ValueError as e:
                print(f"  !! {e}")
            return

        if cmd == 'post':
            if len(parts) < 3:
                print("  Usage: aje post <journal> <file.csv>")
                return
            acct = models.get_account_by_name(parts[1])
            if not acct or not models.is_aje_journal(acct['id']):
                print(f"  '{parts[1]}' is not a year of adjusting entries.")
                return
            path = os.path.expanduser(' '.join(parts[2:]))
            if not os.path.exists(path):
                print(f"  File not found: {path}")
                return
            ref = date_str = desc = ''
            rows = []
            import csv as _csv
            with open(path, newline='', encoding='utf-8-sig') as fh:
                for raw in _csv.reader(fh):
                    if not raw or not (raw[0] or '').strip():
                        continue
                    key = raw[0].strip().lower()
                    val = (raw[1].strip() if len(raw) > 1 else '')
                    if key == 'ref':          ref = val
                    elif key == 'date':       date_str = val
                    elif key == 'description': desc = val
                    else:
                        try:
                            rows.append({'account': raw[0].strip(),
                                         'amount': models.parse_amount(val) if val else 0})
                        except ValueError:
                            print(f"  !! '{val}' is not an amount (line for {raw[0].strip()})")
                            return
            anchor = models.fiscal_anchor()
            try:
                v = models.post_aje(acct['id'], ref or models.next_aje_ref(acct['id']),
                                    date_str or (anchor['cy_end'] if anchor else ''),
                                    desc, rows)
            except ValueError as e:
                print(f"  !! {e}")
                return
            print(f"  ✓ {v['ref']} posted — {len(v['txn_ids'])} legs, "
                  f"Dr {models.fmt_amount_plain(v['debit_cents'])} "
                  f"Cr {models.fmt_amount_plain(v['credit_cents'])}")
            print(f"    {acct['name']} balance: "
                  f"{models.fmt_amount_plain(models.get_account_balance(acct['id']))} (must be 0.00)")
            print(f"    Next reference: {models.next_aje_ref(acct['id'])}")
            return

        acct = models.get_account_by_name(parts[0])
        if not acct or not models.is_aje_journal(acct['id']):
            print(f"  '{parts[0]}' is not a year of adjusting entries. Try:  aje")
            return
        groups = models.aje_groups(acct['id'])
        if not groups:
            print(f"  {acct['name']} has no adjustments yet. First will be "
                  f"{models.next_aje_ref(acct['id'])}.")
            return
        for g in groups:
            flag = '' if g['balanced'] else '   !! out of balance'
            print(f"  {g['date']}  {g['ref']:<10} {g['description']}{flag}")
            for l in g['lines']:
                print(f"      {(l['account'] or '—'):<16} "
                      f"{models.fmt_amount_plain(l['amount']):>14}")
        print(f"  {len(groups)} entr{'y' if len(groups) == 1 else 'ies'}; "
              f"next {models.next_aje_ref(acct['id'])}")

    def do_fye(self, arg):
        """Show or set the fiscal year end being WORKED ON — the year the
        statements report. Usage: fye [YYYY-MM-DD]"""
        if not self._require_books():
            return
        arg = arg.strip()
        if arg:
            normalized = _normalize_date(arg)
            if not normalized:
                print(f"  Invalid date: '{arg}'")
                print("  Use YYYY-MM-DD format, e.g. 2026-05-31")
                return
            try:
                r = models.set_fiscal_settings(working_ye=normalized)
            except ValueError as e:
                print(f"  {e}")
                return
            print(f"  ✓ Working year-end: {r['working_ye']}")
            print(f"    Posting closes after {r['ceiling']}")
        else:
            a = models.fiscal_anchor()
            if not a:
                print("  Fiscal year end: (not set)")
                print("  Set one with: fye YYYY-MM-DD")
                return
            print(f"  Working year-end: {a['cy_end']}")
            print(f"  Current period:   {a['cy_start']} to {a['cy_end']}")
            print(f"  Comparative:      {a['py_start']} to {a['py_end']}")
            print(f"  Posting closes:   {models.fiscal_ceiling()}")

    def do_ceiling(self, arg):
        """Show or set how far past the working year-end posting stays open.
        Usage: ceiling [cy|next]   (cy = stop at the working YE, next = one year on)"""
        if not self._require_books():
            return

        arg = arg.strip().lower()
        if arg:
            if arg not in ('cy', 'next'):
                print(f"  Unknown setting: '{arg}'")
                print("  Use 'ceiling cy' (stop at the working year-end) or")
                print("      'ceiling next' (open the following year).")
                return
            try:
                r = models.set_fiscal_settings(ceiling_mode=arg)
            except ValueError as e:
                print(f"  {e}")
                return
            print(f"  ✓ Posting now closes after {r['ceiling']}")
        else:
            mode = models.get_meta('fy_ceiling_mode', 'cy') or 'cy'
            anchor = models.fiscal_anchor()
            print(f"  Working year-end:   {anchor['cy_end'] if anchor else '(not set)'}")
            print(f"  Posting closes:     {models.fiscal_ceiling() or '(none)'}"
                  f" ({'one year ahead' if mode == 'next' else 'same as working YE'})")

    # ─── quit ────────────────────────────────────────────────────

    def do_quit(self, arg):
        """Exit Grid CLI."""
        print("  Bye.")
        return True

    do_exit = do_quit
    do_EOF = do_quit  # Ctrl+D

    # ─── default / error handling ────────────────────────────────

    def default(self, line):
        """Handle unknown commands."""
        cmd_word = line.split()[0] if line.split() else line
        print(f"  Unknown command: '{cmd_word}'")
        print("  Type 'help' for available commands.")

    def emptyline(self):
        """Do nothing on empty input."""
        pass


# ─── Argument parsing helper ─────────────────────────────────────

def _normalize_csv(rows_raw):
    """Delegates to models.normalize_csv — ONE csv normalizer for all interfaces
    (content-based header detection, extra-field repair, signed Debit/Credit
    netting on multi-column exports)."""
    return models.normalize_csv(rows_raw)

def _split_args(s):
    """Split command arguments, respecting quoted strings."""
    try:
        return shlex.split(s)
    except ValueError:
        return s.split()

def _normalize_date(s):
    """Try to normalize a date string to YYYY-MM-DD."""
    return models.normalize_date(s)


# ─── Main ────────────────────────────────────────────────────────

def main():
    cli = GridCLI()

    args = sys.argv[1:]

    if not args:
        # Interactive mode, no file specified
        print("\n  Grid CLI — Simple. Analog. Robust.")
        print("  Type 'help' for commands, 'open <path>' to load books.\n")
        cli.cmdloop()
        return

    # First arg = db path
    force = False
    for flag in ('--force', '-f'):
        if flag in args:
            args = [a for a in args if a != flag]
            force = True

    db_path = os.path.expanduser(args[0])
    if os.path.isdir(db_path):
        db_path = os.path.join(db_path, 'books.db')

    if not cli.set_books(db_path, force=force):
        sys.exit(1)

    if len(args) > 1:
        # One-shot mode: run command and exit
        command = ' '.join(args[1:])
        cli.onecmd(command)
    else:
        # Interactive mode with file pre-opened
        print("  Type 'help' for commands.\n")
        cli.cmdloop()


if __name__ == '__main__':
    main()
