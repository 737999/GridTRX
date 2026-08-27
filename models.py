"""
GridTRX — Data Model & Database Layer (v2)
NV-style architecture: reports contain ordered items.
Items can be posting accounts, total accounts, labels, or separators.
6 total-to columns enable flexible report arithmetic.
All amounts stored as integers (cents). Double-entry enforced.
One data layer — CLI, MCP server, and browser UI all call these functions.
"""
import sqlite3, os, sys, re as _re
from datetime import datetime, date, timedelta
from contextlib import contextmanager

# ─── GIFI Code Reference (CRA 2026 full chart) ──────────────────────────
# Strict validation: only codes in this dict are allowed.
GIFI_CODES = {
    # ══════════════════════════════════════════════════════════════════════
    # S100 — BALANCE SHEET: ASSETS
    # ══════════════════════════════════════════════════════════════════════
    # Cash & deposits (1000-1059)
    '1000': 'Cash and deposits',
    '1001': 'Cash',
    '1002': 'Deposits in Canadian banks',
    '1003': 'Deposits in foreign banks',
    '1006': 'Term deposits and guaranteed investment certificates',
    # Short-term investments (1060-1119 — note: CRA groups these differently)
    # Accounts receivable (1060-1119)
    '1060': 'Accounts receivable',
    '1061': 'Trade accounts receivable',
    '1062': 'Allowance for doubtful accounts',
    '1063': 'Unbilled revenue',
    '1066': 'Accrued revenue',
    # Inventories (1120-1179)
    '1120': 'Inventories',
    '1121': 'Raw materials',
    '1122': 'Work in progress',
    '1123': 'Finished goods',
    '1124': 'Goods in transit',
    # Short-term investments (1180-1239)
    '1180': 'Canadian short-term investments',
    '1181': 'Canadian term deposits',
    '1187': 'Short-term foreign investments',
    # Loans & notes receivable (1240-1299)
    '1240': 'Loans and notes receivable',
    '1241': 'Due from related parties',
    '1244': 'Due from shareholders',
    # Other current assets (1300-1479)
    '1300': 'Tax instalments',
    '1301': 'Income tax receivable',
    '1302': 'GST/HST receivable',
    # Prepaid expenses (1480-1499)
    '1480': 'Prepaid expenses',
    '1482': 'Prepaid insurance',
    '1484': 'Prepaid rent',
    # Total current assets
    '1599': 'Total current assets',
    # ── Tangible capital assets ──
    # Land (1600-1619)
    '1600': 'Land',
    # Buildings (1680-1699)
    '1680': 'Buildings',
    '1681': 'Accumulated amortization — buildings',
    # Machinery, equipment, furniture (1740-1779)
    '1740': 'Machinery, equipment, furniture, fixtures',
    '1741': 'Accumulated amortization — machinery/equipment',
    '1742': 'Furniture and fixtures',
    '1743': 'Accumulated amortization — furniture/fixtures',
    # Computer equipment (1772-1779)
    '1772': 'Computer equipment',
    '1773': 'Accumulated amortization — computer equipment',
    '1774': 'Computer software',
    '1775': 'Accumulated amortization — computer software',
    # Leasehold improvements (1780-1789)
    '1780': 'Leasehold improvements',
    '1781': 'Accumulated amortization — leasehold improvements',
    # Other tangible assets
    '1783': 'Other tangible capital assets',
    '1784': 'Accumulated amortization — other tangible',
    # Vehicles
    '1786': 'Automobiles',
    '1787': 'Accumulated amortization — automobiles',
    # Total tangible capital assets
    '1889': 'Total tangible capital assets (net)',
    # ── Intangible capital assets ──
    '2008': 'Goodwill',
    '2009': 'Accumulated amortization — goodwill',
    '2010': 'Patents',
    '2011': 'Accumulated amortization — patents',
    '2060': 'Trademarks',
    '2061': 'Accumulated amortization — trademarks',
    '2178': 'Other intangible assets',
    '2179': 'Accumulated amortization — other intangible',
    '2189': 'Total intangible capital assets (net)',
    # ── Long-term assets ──
    '2300': 'Long-term investments',
    '2301': 'Investments in Canadian corporations',
    '2302': 'Investments in foreign corporations',
    '2310': 'Investments in associated corporations',
    '2320': 'Investments in joint ventures',
    '2360': 'Long-term receivables',
    '2420': 'Due from related parties — long-term',
    '2500': 'Other long-term assets',
    '2510': 'Security deposits',
    '2589': 'Total long-term assets',
    '2599': 'Total assets',
    # ══════════════════════════════════════════════════════════════════════
    # S100 — BALANCE SHEET: LIABILITIES
    # ══════════════════════════════════════════════════════════════════════
    # Bank overdraft
    '2600': 'Bank overdraft',
    '2601': 'Bank indebtedness',
    # Accounts payable (2620-2679)
    '2620': 'Accounts payable and accrued liabilities',
    '2621': 'Trade accounts payable',
    '2622': 'Accrued liabilities',
    '2624': 'Employee deductions payable',
    '2626': 'Accrued wages and salaries',
    '2628': 'Customer deposits and prepayments',
    '2630': 'Royalties payable',
    '2640': 'Current portion of deferred revenue',
    # Taxes payable (2680-2699)
    '2680': 'Taxes payable — federal',
    '2681': 'Income tax payable — federal',
    '2682': 'Income tax payable — provincial',
    '2683': 'GST/HST payable',
    '2684': 'Provincial sales tax payable',
    '2685': 'Payroll taxes payable',
    # Current debt (2700-2779)
    '2700': 'Current portion of long-term debt',
    '2701': 'Bank loans — current',
    '2710': 'Notes payable — current',
    '2740': 'Mortgage payable — current portion',
    # Due to shareholders / related (2780-2819)
    '2780': 'Due to shareholder(s)',
    '2781': 'Due to shareholder(s)',
    '2782': 'Due to directors',
    '2789': 'Due to related parties',
    '2790': 'Due to related corporations',
    # Other current liabilities
    '2800': 'Deferred revenue — current',
    '2810': 'Current portion of obligations under capital lease',
    # Total current liabilities
    '3139': 'Total current liabilities',
    # ── Long-term liabilities ──
    '3140': 'Long-term bank loans',
    '3141': 'Mortgages payable',
    '3142': 'Bonds and debentures payable',
    '3143': 'Due to shareholder(s) — long-term',
    '3144': 'Notes payable — long-term',
    '3145': 'Obligations under capital lease — long-term',
    '3146': 'Due to related corporations — long-term',
    '3148': 'Deferred income taxes',
    '3300': 'Deferred revenue — long-term',
    '3400': 'Other long-term liabilities',
    '3450': 'Future income taxes',
    '3499': 'Total liabilities',
    # ══════════════════════════════════════════════════════════════════════
    # S100 — BALANCE SHEET: EQUITY
    # ══════════════════════════════════════════════════════════════════════
    '3500': 'Common shares',
    '3520': 'Preferred shares',
    '3540': 'Contributed surplus',
    '3560': 'Other paid-in capital',
    '3600': 'Retained earnings/deficit',
    '3620': 'Total equity',
    '3640': 'Total liabilities and equity',
    # Retained earnings detail
    '3660': 'Retained earnings — start of year',
    '3680': 'Net income/loss',
    '3700': 'Dividends declared',
    '3701': 'Cash dividends',
    '3720': 'Other adjustments to retained earnings',
    '3849': 'Retained earnings — end of year',
    # ══════════════════════════════════════════════════════════════════════
    # S125 — INCOME STATEMENT: REVENUE
    # ══════════════════════════════════════════════════════════════════════
    '8000': 'Trade sales of goods and services',
    '8020': 'Sales of natural resources',
    '8040': 'Real estate sales',
    '8070': 'Other trade revenue',
    '8089': 'Rental revenue',
    '8090': 'Interest and investment revenue',
    '8091': 'Interest from foreign sources',
    '8092': 'Interest from Canadian government bonds',
    '8094': 'Interest from other Canadian sources',
    '8095': 'Dividend income',
    '8096': 'Dividends from taxable Canadian corporations',
    '8097': 'Dividends from foreign corporations',
    '8100': 'Royalty income',
    '8102': 'Securities interest',
    '8110': 'Commission revenue',
    '8140': 'Net rental income/loss',
    '8141': 'Gross rental revenue',
    '8142': 'Rental expenses',
    '8150': 'Management fees earned',
    '8160': 'Service revenue',
    '8210': 'Realized gains/losses on disposition of assets',
    '8211': 'Realized gains/losses on sale of investments',
    '8212': 'Gain on disposition of capital property',
    '8230': 'Foreign exchange gains/losses',
    '8231': 'Foreign exchange gain',
    '8232': 'Foreign exchange loss',
    '8235': 'Income/loss from partnerships',
    '8240': 'Wage subsidies received',
    '8242': 'Government assistance',
    '8244': 'Research and development grants',
    '8249': 'Expense recoveries',
    '8250': 'Insurance proceeds',
    '8260': 'Bad debts recovered',
    '8270': 'Other revenue',
    '8299': 'Total revenue',
    # ══════════════════════════════════════════════════════════════════════
    # S125 — INCOME STATEMENT: COST OF GOODS SOLD
    # ══════════════════════════════════════════════════════════════════════
    '8300': 'Cost of goods sold',
    '8301': 'Opening inventory',
    '8302': 'Purchases, raw materials',
    '8303': 'Direct wages',
    '8304': 'Manufacturing overhead',
    '8305': 'Closing inventory',
    '8320': 'Subcontracts — COGS',
    '8340': 'Direct material costs',
    '8360': 'Direct labour costs',
    '8380': 'Manufacturing overhead costs',
    '8500': 'Gross profit',
    '8518': 'Total cost of goods sold',
    '8519': 'Gross profit / net of COGS',
    # ══════════════════════════════════════════════════════════════════════
    # S125 — INCOME STATEMENT: OPERATING EXPENSES
    # ══════════════════════════════════════════════════════════════════════
    '8520': 'Advertising and promotion',
    '8521': 'Advertising',
    '8522': 'Donations and gifts',
    '8523': 'Meals and entertainment',
    '8570': 'Bad debts expense',
    '8571': 'Provision for bad debts',
    '8590': 'Employee benefits',
    '8620': 'Delivery, shipping, express',
    '8621': 'Freight and cartage',
    '8640': 'Fuel costs',
    '8670': 'Amortization of tangible capital assets',
    '8671': 'Amortization of buildings',
    '8672': 'Amortization of equipment',
    '8673': 'Amortization of vehicles',
    '8680': 'Amortization of intangible capital assets',
    '8690': 'Insurance',
    '8691': 'Life insurance premiums',
    '8692': 'Business insurance',
    '8693': 'Vehicle insurance',
    '8710': 'Interest and bank charges',
    '8711': 'Interest on short-term debt',
    '8712': 'Interest on long-term debt',
    '8713': 'Interest on mortgage',
    '8714': 'Interest on capital leases',
    '8715': 'Bank charges',
    '8716': 'Credit card charges',
    '8760': 'Business taxes, licences, memberships',
    '8761': 'Business licence fees',
    '8762': 'Professional membership dues',
    '8764': 'Penalties and fines',
    '8810': 'Office expenses',
    '8811': 'Office supplies',
    '8812': 'Postage and courier',
    '8813': 'Printing and stationery',
    '8814': 'Computer and IT expenses',
    '8860': 'Professional fees',
    '8861': 'Accounting and auditing',
    '8862': 'Legal fees',
    '8863': 'Consulting fees',
    '8869': 'Brokerage and custodial fees',
    '8870': 'Management and administration fees',
    '8871': 'Management salaries',
    '8910': 'Rent',
    '8911': 'Rent — real property',
    '8912': 'Equipment rental',
    '8960': 'Repairs and maintenance',
    '8961': 'Building maintenance',
    '8962': 'Equipment repairs',
    '9010': 'Research and development',
    '9060': 'Salaries, wages and benefits',
    '9061': 'Salaries and wages',
    '9062': 'Commissions paid',
    '9063': 'Employee benefits — employer portion',
    '9064': 'CPP contributions — employer',
    '9065': 'EI premiums — employer',
    '9066': 'Workers compensation',
    '9068': 'Director fees',
    '9100': 'Security',
    '9130': 'Subcontracts',
    '9131': 'Subcontractor fees',
    '9150': 'Supplies',
    '9180': 'Property taxes',
    '9200': 'Travel expenses',
    '9201': 'Travel — transportation',
    '9202': 'Travel — accommodation',
    '9203': 'Travel — meals',
    '9220': 'Utilities',
    '9221': 'Electricity',
    '9222': 'Water',
    '9223': 'Natural gas',
    '9224': 'Telephone and telecommunications',
    '9225': 'Internet',
    '9270': 'Motor vehicle expenses (not CCA)',
    '9271': 'Fuel and oil — vehicles',
    '9272': 'Vehicle repairs and maintenance',
    '9273': 'Vehicle licence and registration',
    '9274': 'Vehicle insurance',
    '9275': 'Capital cost allowance (CCA)',
    '9276': 'Terminal loss',
    '9277': 'Recapture of CCA',
    '9281': 'Motor vehicle expenses — total',
    '9284': 'Other expenses',
    '9285': 'Miscellaneous expenses',
    '9286': 'Warranty expense',
    '9290': 'Loss on disposal of assets',
    # Totals
    '9367': 'Total operating expenses',
    '9368': 'Total expenses',
    '9369': 'Net non-farming income',
    # ══════════════════════════════════════════════════════════════════════
    # S125 — OTHER / TAX
    # ══════════════════════════════════════════════════════════════════════
    '9370': 'Net farming income',
    '9600': 'Net partnership income',
    '9898': 'Extraordinary items',
    '9970': 'Net income/loss before taxes and extraordinary items',
    '9990': 'Current income taxes',
    '9991': 'Federal income tax',
    '9992': 'Provincial income tax',
    '9995': 'Future/deferred income tax expense',
    '9999': 'Net income/loss after taxes and extraordinary items',
}

DB_PATH = None
_wal_set = set()        # db paths already switched to WAL this process (WAL is persistent on the file)
_balances_cache = {}    # {'gen':, 'path':, 'by_args': {(from,to,side): {acct_id: cents}}}
BALANCE_CACHE_MAX = 24  # statement columns per book; plenty, and bounded
def get_db_path(): return DB_PATH
def set_db_path(path):
    global DB_PATH
    if not path:
        release_books_lock()   # closing the books frees the single-instance lock
    DB_PATH = path
    # Any open/close/switch drops the generation mirror: while the books were
    # closed another process may have written them (the lock only guards the
    # OPEN books), so the counters must be re-read from the file.
    _gens_mirror['path'] = None

@contextmanager
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    # WAL mode is persistent on the db FILE — setting it once per process is enough.
    # Re-running it per connection costs ~0.2ms each (a WAL re-negotiation) for no benefit.
    if DB_PATH not in _wal_set:
        # CHECK the answer. `PRAGMA journal_mode=WAL` reports the mode it ended
        # up in, and it does NOT switch while another connection holds the file
        # — it just returns 'delete' and raises nothing. Marking the path done
        # regardless meant one failed attempt left the book in rollback-journal
        # mode for the whole session, and there a long write transaction BLOCKS
        # READERS: a conversion would lock out the very progress poll meant to
        # show it running. Only remember it once it is really WAL; otherwise the
        # next connection tries again.
        mode = conn.execute("PRAGMA journal_mode=WAL").fetchone()
        if mode and str(mode[0]).lower() == 'wal':
            _wal_set.add(DB_PATH)
    conn.execute("PRAGMA foreign_keys=ON")   # per-connection — required for FK enforcement
    conn.execute("PRAGMA busy_timeout=5000")
    try:
        yield conn
        conn.commit()
    except:
        conn.rollback()
        raise
    finally:
        # Refresh the in-process generation mirror after any block that wrote,
        # so warm cache asks stay connection-free (the v148 guarantee). Safe as
        # in-process state ONLY because db.lock guarantees a single opener per
        # set of books: nothing else can move them behind us. On any doubt the
        # mirror is dropped and the next ask reads the file.
        if conn.total_changes:
            try:
                rows = conn.execute("SELECT key, value FROM meta WHERE key IN "
                                    "('balance_gen','chain_gen')").fetchall()
                d = {r[0]: int(r[1] or 0) for r in rows}
                _gens_mirror.update({'path': DB_PATH,
                                     'gens': (d.get('balance_gen', 0),
                                              d.get('chain_gen', 0))})
            except Exception:
                _gens_mirror['path'] = None
        conn.close()

def init_db(path, force_lock=False):
    """THE single way to open a set of books — all three interfaces call this.
    LAP shell discipline, in order: Check Books (integrity gate), Backup Books
    (daily snapshot of the file exactly as it was opened, BEFORE any migration
    touches it), then open/migrate.

    force_lock=True takes the books from a stale lock. It is never automatic:
    something must have shown the operator who holds them and been told to go
    in anyway (the F3 door on the web lock screen, --force on the CLI)."""
    global _backup_note, _re_repair
    _backup_note = ''
    _re_repair = []
    acquire_books_lock(path, force=force_lock)   # ONE opener per set of books
    set_db_path(path)
    existing = os.path.exists(path) and os.path.getsize(path) > 0
    if existing:
        # Integrity gate (LAP file check): a damaged file must be refused BEFORE
        # anything writes to it — working on it makes recovery worse. Badly
        # mangled files make quick_check itself raise — same refusal.
        try:
            with get_db() as db:
                ok = db.execute("PRAGMA quick_check").fetchone()[0]
        except sqlite3.OperationalError as e:
            # COULD NOT OPEN — SQLite never got as far as reading the contents,
            # so this says NOTHING about whether the books are healthy. Telling
            # the operator to restore a snapshot here would be a lie, and the
            # dangerous kind: it invites replacing good books with older ones.
            # OperationalError must be caught BEFORE DatabaseError, its parent.
            set_db_path(None)
            raise ValueError(
                f"'{os.path.basename(path)}' could not be OPENED: {e}. "
                f"The books themselves have NOT been checked and may be "
                f"perfectly healthy — this is about reaching the file, not its "
                f"contents. Close anything else that has it open, check the "
                f"folder still exists and is writable, then try again. "
                f"(SQLite also reports this when the volume cannot host the "
                f"-wal/-shm files it must create beside the books.)")
        except sqlite3.DatabaseError as e:
            ok = str(e)          # genuinely damaged: 'file is not a database' &c
        if ok != 'ok':
            set_db_path(None)
            snaps = list_backups(path)
            where = (f"restore the newest snapshot beside it "
                     f"({os.path.basename(snaps[-1])})" if snaps else
                     "there is NO snapshot beside it yet — recover this file "
                     "from your own backups")
            raise ValueError(
                f"'{os.path.basename(path)}' FAILED its integrity check: {ok}. "
                f"Do NOT keep working on this file — {where}.")
        try:
            snap = backup_books()          # daily; no-op if today's exists
            if snap:
                _backup_note = f"Snapshot taken: backups/{os.path.basename(snap)}"
        except Exception as e:
            # A failed backup must be LOUD but must not brick the books —
            # they are still healthy; the operator just isn't protected today.
            _backup_note = f"BACKUP FAILED: {e}"
    try:
        with get_db() as db:
            db.executescript(SCHEMA)
        _ensure_columns()
        _ensure_data_gens()
        migrate_re_computed()   # LAP-style RE rewire; idempotent, no-op once clean / on a fresh db.
                                # Lives here so MCP/CLI/Flask all see the SAME migrated structure.
        ensure_engagement_folders()   # Engagement File root (+ standard sections on fresh files)
        ensure_trx_layout()           # TRX head: heading, TRX.OPEN, RE.OB — forced, self-healing
        ensure_fiscal_settings()      # ceiling derived from the year-end; legacy fy_end_date folded in
        ensure_aje_layout()           # AJE head: the heading; year journals below, operator's order
    except sqlite3.OperationalError as e:
        if 'readonly' in str(e).lower() or 'locked' in str(e).lower():
            set_db_path(None)
            raise ValueError(
                f"'{os.path.basename(path)}' cannot be written ({e}). If this file is on a "
                f"network mount (SMB/gvfs), open the books on the machine where the folder is "
                f"LOCAL — SQLite locking is not reliable over network shares.")
        raise

# ─── Single-instance lock (LAP "books already in use") ─────────────
# DOS-style: ONE opener per set of books, ever. A plain-text lock file sits
# beside the db while it is open; a second opener — another Grid window,
# the CLI, an MCP agent, another machine — is refused with an error that
# says exactly who holds it. Same-host stale locks (crashed process) self-
# heal; locks from another machine must be removed by hand, because their
# liveness cannot be verified — and that is precisely the dangerous case.
_lock_held = None    # path of the lock file THIS process holds

def _lock_path_for(db_path):
    return os.path.abspath(db_path) + '.lock'

def _read_lock(lp):
    try:
        info = {}
        for line in open(lp):
            if '=' in line:
                k, v = line.split('=', 1)
                info[k.strip()] = v.strip()
        return info
    except OSError:
        return {}

def _pid_alive(pid):
    """Is this PID still running? ASKS ONLY — never signals or touches it.
    Unknown answer = alive: refusing to open books that might be live is safe,
    stealing them is not. Windows has no signal 0 (os.kill THERE terminates the
    target with that exit code), so it gets its own OpenProcess probe."""
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return False
    if pid <= 0:
        return False
    if os.name == 'nt':
        import ctypes
        PROCESS_QUERY_LIMITED_INFORMATION, STILL_ACTIVE, ACCESS_DENIED = 0x1000, 259, 5
        try:
            k = ctypes.WinDLL('kernel32', use_last_error=True)
            k.OpenProcess.restype = ctypes.c_void_p
            k.OpenProcess.argtypes = [ctypes.c_ulong, ctypes.c_int, ctypes.c_ulong]
            k.GetExitCodeProcess.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong)]
            k.CloseHandle.argtypes = [ctypes.c_void_p]
            h = k.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
            if not h:
                # 5 = someone else's process (exists); anything else = no such process.
                return ctypes.get_last_error() == ACCESS_DENIED
            try:
                code = ctypes.c_ulong()
                if k.GetExitCodeProcess(h, ctypes.byref(code)):
                    return code.value == STILL_ACTIVE
                return True
            finally:
                k.CloseHandle(h)
        except OSError:
            return True
    try:
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return True

def books_lock_info(db_path):
    """Holder info for a books file, or None if not locked (or locked by us)."""
    lp = _lock_path_for(db_path)
    if not os.path.exists(lp) or lp == _lock_held:
        return None
    return _read_lock(lp) or {'prog': 'unknown'}

# What a locked file may be, and what actually fixes each one. A refusal with
# no door is what sends staff to delete files by hand over a phone, so every
# state here names the ONE thing that gets the operator in.
LOCK_RUNNING = 'running'   # live PID, this machine → close that session properly
LOCK_FOREIGN = 'foreign'   # another machine's lock (incl. one synced in by Box)
# A dead PID on this machine is not a state: it self-heals and never reaches here.

class BooksLocked(ValueError):
    """These books are held by someone else — WITH the way in.

    Still a ValueError, so every existing caller keeps catching it; the extra
    fields let an interface draw the classic lock screen instead of a dead end."""
    def __init__(self, headline, detail, door, state, info, lock_path, warn=''):
        pid, host = info.get('pid', '?'), info.get('host', '?')
        prog, started = info.get('prog', 'Grid'), info.get('started', '?')
        # The flat message carries the holder too. The screen shows that on its
        # own line, so the sentences stay clean — but anything that only ever
        # sees str(e) still learns exactly who has the books.
        super().__init__(' '.join(x for x in (
            headline, f"[{prog} · PID {pid} · {host} · since {started}]",
            detail, door, warn) if x))
        self.headline, self.detail, self.door, self.warn = headline, detail, door, warn
        self.state, self.info, self.lock_path = state, info, lock_path
        self.pid, self.host, self.prog, self.started = pid, host, prog, started
        try:
            self.port = int(info.get('port', 0)) or None
        except ValueError:
            self.port = None

def _locked_error(info, lp, host):
    """Build the refusal for a lock we cannot take. ONE place, so the browser,
    the CLI and an agent all tell the operator exactly the same story."""
    prog = info.get('prog') or 'Grid'
    pid = info.get('pid', '?')
    started = info.get('started', '?')
    who = info.get('host', '?')
    if who != host:
        return BooksLocked(
            headline=f"These books show as OPEN on another computer ({who}).",
            detail=f"Grid cannot see across machines, so it cannot tell whether {who} "
                   f"is really in them or crashed and left this behind. Check that "
                   f"nobody is working in them.",
            door="Nobody is in them — clear the lock and open the books.",
            warn="If someone IS in them, you will each be working on your own copy, and "
                 "your file sync will keep one and rename the other.",
            state=LOCK_FOREIGN, info=info, lock_path=lp)
    return BooksLocked(
        headline="These books are already open on this computer.",
        detail="Another Grid session still has them. Closing it puts the books away "
               "properly first — nothing is lost.",
        door="Close that session and open the books here.",
        state=LOCK_RUNNING, info=info, lock_path=lp)

def books_lock_state(db_path):
    """Who is in the way of opening these books, or None if nothing is.

    READ-ONLY: takes no lock, heals nothing, touches no file. A crash leftover
    (this machine, dead PID) answers None because acquiring WILL self-heal it —
    the operator should never see a screen for the ordinary crash."""
    import socket
    lp = _lock_path_for(db_path)
    if not os.path.exists(lp) or lp == _lock_held:
        return None
    host = socket.gethostname()
    info = _read_lock(lp)
    if info.get('host', host) == host and not _pid_alive(info.get('pid', -1)):
        return None
    return _locked_error(info, lp, host)

def clear_books_lock(db_path):
    """Remove someone else's lock file, deliberately. Returns what it removed.

    The ONE door — nothing calls this on its own. Its caller has already shown
    the operator who holds the books and been told to go in anyway."""
    lp = _lock_path_for(db_path)
    info = _read_lock(lp) if os.path.exists(lp) else {}
    try:
        os.remove(lp)
    except FileNotFoundError:
        pass
    except OSError as e:
        raise ValueError(f"Could not clear the lock file ({e}): {lp}")
    return info

LOCK_EXTRA = {}   # extra lines an interface wants in its lock file (app.py: port=)

def acquire_books_lock(path, force=False):
    """Take the exclusive lock on a set of books, or raise BooksLocked with who
    has it AND the way in. Acquiring books B while holding books A releases A
    first — one set of books per process, always.

    force=True takes a lock that is not ours. Only ever reached because someone
    was shown the holder and said go in anyway."""
    import socket
    global _lock_held
    lp = _lock_path_for(path)
    if _lock_held == lp and os.path.exists(lp):
        return   # already ours (re-open of the same books in this process)
    host = socket.gethostname()
    if os.path.exists(lp):
        # ONE question, asked in ONE place — so the lock screen and the open
        # that follows it can never disagree about who holds the books.
        # None means nothing is in the way: free, or a crash leftover from this
        # machine, which self-heals here, silently, as it always has.
        blocked = books_lock_state(path)
        if blocked is not None and not force:
            raise blocked
        try: os.remove(lp)
        except OSError: pass
    release_books_lock()   # one set of books per process
    prog = os.path.basename(sys.argv[0] or 'grid') if hasattr(sys, 'argv') else 'grid'
    extra = ''.join(f"{k}={v}\n" for k, v in LOCK_EXTRA.items())
    try:
        fd = os.open(lp, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        with os.fdopen(fd, 'w') as f:
            f.write(f"pid={os.getpid()}\nhost={host}\nprog={prog}\n"
                    f"started={datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n{extra}")
    except FileExistsError:
        raise ValueError(f"These books were just opened by another instance. Try again, "
                         f"or check the lock file: {lp}")
    except OSError as e:
        raise ValueError(f"Cannot create the lock file beside the books ({e}). "
                         f"Read-only folder or network mount? Open the books on the "
                         f"machine where the folder is local.")
    _lock_held = lp

def release_books_lock():
    """Drop the lock this process holds (close, switch, or exit)."""
    global _lock_held
    if _lock_held:
        try:
            info = _read_lock(_lock_held)
            if str(info.get('pid')) == str(os.getpid()):   # only ever delete our own
                os.remove(_lock_held)
        except OSError:
            pass
        _lock_held = None

import atexit as _atexit
_atexit.register(release_books_lock)

# ─── Backups (LAP "Backup Books" / "Check Books") ──────────────────
# Every snapshot is a complete, checkpointed, integrity-verified, SINGLE-FILE
# copy of the books in backups/ beside the db. Restore = open (or copy back)
# a snapshot. Box carries the folder offsite; snapshots are the restore points
# that don't depend on Box catching the db+WAL pair at a consistent moment.
BACKUP_KEEP = 30        # ≈ a month of daily restore points per set of books
_backup_note = ''       # set at open; UIs surface it as a status line

# A repair to RETAINED EARNINGS must never be silent. migrate_re_computed runs
# on every open and heals a file from any earlier shape — which is right — but
# rewiring RE can MOVE THE NUMBER ON A CLIENT'S BALANCE SHEET. That is not a
# housekeeping detail to slip in quietly; it is the single figure a reader
# checks. So the migration records what it did and every interface says so.
_re_repair = []         # [] = nothing to report; else lines describing the repair

def _backup_dir(db_path):
    return os.path.join(os.path.dirname(os.path.abspath(db_path)), 'backups')

def list_backups(db_path=None):
    """Snapshot paths for a set of books, oldest→newest."""
    path = db_path or DB_PATH
    if not path:
        return []
    base = os.path.splitext(os.path.basename(path))[0]
    bdir = _backup_dir(path)
    if not os.path.isdir(bdir):
        return []
    paths = [os.path.join(bdir, fn) for fn in os.listdir(bdir)
             if fn.startswith(base + '.') and fn.endswith('.db')]
    # mtime order, name as tiebreak — collision-suffixed same-second snapshots
    # would otherwise sort as older than they are
    return sorted(paths, key=lambda p: (os.path.getmtime(p), p))

def backup_books(force=False):
    """Snapshot the open books to backups/<name>.<yyyymmdd-hhmmss>.db.

    Checkpoints the WAL first (so the copy is the WHOLE state, not the state
    minus recent commits), snapshots via VACUUM INTO (compact + consistent),
    verifies the snapshot with quick_check before keeping it, then rotates to
    the newest BACKUP_KEEP. Auto-runs once per day at open; force=True for a
    manual backup. Returns the snapshot path, or None if today's exists."""
    path = DB_PATH
    if not path:
        raise ValueError("No books open")
    base = os.path.splitext(os.path.basename(path))[0]
    stamp = datetime.now().strftime('%Y%m%d-%H%M%S')
    if not force:
        today_prefix = f"{base}.{stamp[:8]}"
        if any(os.path.basename(p).startswith(today_prefix) for p in list_backups(path)):
            return None
    bdir = _backup_dir(path)
    os.makedirs(bdir, exist_ok=True)
    dest = os.path.join(bdir, f"{base}.{stamp}.db")
    n = 1
    while os.path.exists(dest):   # two snapshots in one second must not collide
        dest = os.path.join(bdir, f"{base}.{stamp}-{n}.db")
        n += 1
    tmp = dest + '.tmp'
    if os.path.exists(tmp):
        os.remove(tmp)
    try:
        with get_db() as db:
            db.execute("PRAGMA wal_checkpoint(TRUNCATE)")
            db.execute("VACUUM INTO ?", (tmp,))
        chk = sqlite3.connect(tmp)
        ok = chk.execute("PRAGMA quick_check").fetchone()[0]
        chk.close()
        if ok != 'ok':
            raise ValueError(f"snapshot failed its integrity check ({ok})")
        os.replace(tmp, dest)
    except Exception:
        if os.path.exists(tmp):
            try: os.remove(tmp)
            except OSError: pass
        raise
    # Rotate — newest BACKUP_KEEP stay
    snaps = list_backups(path)
    for old in snaps[:-BACKUP_KEEP]:
        try: os.remove(old)
        except OSError: pass
    return dest

def re_repair_note():
    """What the retained-earnings repair did when this file was opened, if it
    ran. Empty list = nothing to report. Cleared on every open."""
    return list(_re_repair)


def backup_status():
    """Status line data for the UIs: last snapshot, count, any open-time error."""
    snaps = list_backups()
    return {
        'last': os.path.basename(snaps[-1]) if snaps else '',
        'count': len(snaps),
        'note': _backup_note,
        'error': _backup_note.startswith('BACKUP FAILED'),
    }

def checkpoint_books():
    """Fold the WAL into the main db file (best-effort). Called on close so the
    cloud-synced single .db file is complete at rest."""
    if not DB_PATH:
        return
    try:
        with get_db() as db:
            db.execute("PRAGMA wal_checkpoint(TRUNCATE)")
    except Exception:
        pass

def _ensure_columns():
    """Add columns that may be missing from older database files."""
    with get_db() as db:
        cols = {r[1] for r in db.execute("PRAGMA table_info(lines)").fetchall()}
        if 'doc_on_file' not in cols:
            db.execute("ALTER TABLE lines ADD COLUMN doc_on_file INTEGER DEFAULT 0")
        txn_cols = {r[1] for r in db.execute("PRAGMA table_info(transactions)").fetchall()}
        if 'import_batch' not in txn_cols:
            # '' = manual entry; 'imp-YYYYMMDD-HHMMSS' = tagged file import.
            # The tag is what makes an import a deletable unit (delete-and-redo).
            db.execute("ALTER TABLE transactions ADD COLUMN import_batch TEXT DEFAULT ''")
        # Indexed HERE, not in SCHEMA: the column arrives by ALTER above, and a
        # CREATE INDEX on a column that does not exist yet aborts the whole
        # schema script. opening_batch() probes this on every home page load.
        db.execute("CREATE INDEX IF NOT EXISTS idx_txn_batch "
                   "ON transactions(import_batch)")
        acct_cols = {r[1] for r in db.execute("PRAGMA table_info(accounts)").fetchall()}
        if 'gifi_code' not in acct_cols:
            db.execute("ALTER TABLE accounts ADD COLUMN gifi_code TEXT DEFAULT ''")
        if 'next_ref' not in acct_cols:
            # LAP [Next Ref#]: >0 = auto-numbered references for manual entry on
            # this account (cheque/invoice numbers); 0 = off (random refs).
            db.execute("ALTER TABLE accounts ADD COLUMN next_ref INTEGER DEFAULT 0")
        if 'leadsheet' not in acct_cols:
            # Working-paper lead-sheet code (CaseWare-style: A, B-1, E-1…) —
            # groups accounts under a lead sheet; mirrors the GIFI mapping idea.
            db.execute("ALTER TABLE accounts ADD COLUMN leadsheet TEXT DEFAULT ''")
        wp_cols = {r[1] for r in db.execute("PRAGMA table_info(workpapers)").fetchall()}
        if wp_cols and 'folder_id' not in wp_cols:
            db.execute("ALTER TABLE workpapers ADD COLUMN folder_id INTEGER DEFAULT 0")
        ri_cols = {r[1] for r in db.execute("PRAGMA table_info(report_items)").fetchall()}
        if 'ref_mark' not in ri_cols:
            # The red pencil: working-paper index mark (E-1, B-2.1) shown beside
            # the numbers on statements — the manual reference legacy programs never had.
            db.execute("ALTER TABLE report_items ADD COLUMN ref_mark TEXT DEFAULT ''")
        if 'system' not in acct_cols:
            # Grid MAINTAINS this account: it is created, positioned and posted
            # to by a module (conversion, opening RE), not by the operator. Not a
            # lock — you can still open it and read it — but the interface says
            # so, because casually posting to one corrupts what it represents.
            db.execute("ALTER TABLE accounts ADD COLUMN system INTEGER DEFAULT 0")
        if 'computed' not in acct_cols:
            # '' = normal; 'open:<ACCT>'/'close:<ACCT>' = a derived, off-book line whose
            # value is the perpetual balance of <ACCT> as of the report period start/end
            # (used for the computed Opening/Closing Retained Earnings lines).
            db.execute("ALTER TABLE accounts ADD COLUMN computed TEXT DEFAULT ''")

def _ensure_data_gens():
    """Two in-file generation counters, bumped by TRIGGERS so no write path can
    forget them: 'balance_gen' moves when MONEY moves (lines, transaction
    dates/deletes), 'chain_gen' when the accumulation WIRING moves (total-tos,
    what sits on a report, an account's name/side/computed mode). The balance
    cache invalidates on balance_gen alone — so moving a line on a statement,
    saving a column layout or pencilling a ref mark no longer throws away every
    computed column and re-adds the book (on a 318k-transaction file that was
    seconds of dead time per click). Position-only moves deliberately bump
    NEITHER: an UPDATE that only sets position/indent/description is furniture."""
    with get_db() as db:
        db.execute("INSERT OR IGNORE INTO meta(key, value) VALUES('balance_gen','0')")
        db.execute("INSERT OR IGNORE INTO meta(key, value) VALUES('chain_gen','0')")
        bump_b = ("UPDATE meta SET value = CAST(value AS INTEGER) + 1 "
                  "WHERE key='balance_gen';")
        bump_c = ("UPDATE meta SET value = CAST(value AS INTEGER) + 1 "
                  "WHERE key='chain_gen';")
        for name, when in (
            ('trg_balgen_lines_i', 'AFTER INSERT ON lines'),
            ('trg_balgen_lines_u', 'AFTER UPDATE ON lines'),
            ('trg_balgen_lines_d', 'AFTER DELETE ON lines'),
            ('trg_balgen_txn_u',   'AFTER UPDATE OF date ON transactions'),
            ('trg_balgen_txn_d',   'AFTER DELETE ON transactions'),
        ):
            db.execute(f"CREATE TRIGGER IF NOT EXISTS {name} {when} "
                       f"BEGIN {bump_b} END")
        for name, when in (
            ('trg_chaingen_ri_i', 'AFTER INSERT ON report_items'),
            ('trg_chaingen_ri_d', 'AFTER DELETE ON report_items'),
            ('trg_chaingen_ri_u', 'AFTER UPDATE OF account_id, item_type, '
             'total_to_1, total_to_2, total_to_3, total_to_4, total_to_5, '
             'total_to_6 ON report_items'),
            ('trg_chaingen_acct_u', 'AFTER UPDATE OF name, normal_balance, '
             'computed, account_type ON accounts'),
        ):
            db.execute(f"CREATE TRIGGER IF NOT EXISTS {name} {when} "
                       f"BEGIN {bump_c} END")


_gens_mirror = {'path': None, 'gens': (0, 0)}

def _data_gens():
    """(balance_gen, chain_gen) as ints. Served from the in-process mirror when
    it is current (warm cache asks open no connection); read from the file and
    mirrored otherwise."""
    if _gens_mirror['path'] == DB_PATH:
        return _gens_mirror['gens']
    with get_db() as db:
        rows = db.execute("SELECT key, value FROM meta WHERE key IN "
                          "('balance_gen','chain_gen')").fetchall()
    d = {r[0]: int(r[1] or 0) for r in rows}
    gens = (d.get('balance_gen', 0), d.get('chain_gen', 0))
    _gens_mirror.update({'path': DB_PATH, 'gens': gens})
    return gens


SCHEMA = """
CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT);

CREATE TABLE IF NOT EXISTS reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    description TEXT DEFAULT '',
    sort_order INTEGER DEFAULT 0,
    period_begin TEXT DEFAULT '',
    period_end TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS report_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    report_id INTEGER NOT NULL REFERENCES reports(id),
    position INTEGER NOT NULL DEFAULT 0,
    item_type TEXT NOT NULL DEFAULT 'account'
        CHECK(item_type IN ('account','total','label','separator')),
    description TEXT DEFAULT '',
    account_id INTEGER REFERENCES accounts(id),
    indent INTEGER DEFAULT 0,
    total_to_1 TEXT DEFAULT '',
    total_to_2 TEXT DEFAULT '',
    total_to_3 TEXT DEFAULT '',
    total_to_4 TEXT DEFAULT '',
    total_to_5 TEXT DEFAULT '',
    total_to_6 TEXT DEFAULT '',
    sep_style TEXT DEFAULT '' CHECK(sep_style IN ('','single','double','blank'))
);

CREATE TABLE IF NOT EXISTS accounts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    description TEXT DEFAULT '',
    normal_balance TEXT NOT NULL CHECK(normal_balance IN ('D','C')),
    account_type TEXT DEFAULT 'posting' CHECK(account_type IN ('posting','total')),
    account_number TEXT DEFAULT '',
    notes TEXT DEFAULT '',
    computed TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS transactions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT NOT NULL,
    reference TEXT DEFAULT '',
    description TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS lines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    transaction_id INTEGER NOT NULL REFERENCES transactions(id) ON DELETE CASCADE,
    account_id INTEGER NOT NULL REFERENCES accounts(id),
    amount INTEGER NOT NULL,
    description TEXT DEFAULT '',
    reconciled INTEGER DEFAULT 0,
    doc_on_file INTEGER DEFAULT 0,
    sort_order INTEGER DEFAULT 0
);

CREATE INDEX IF NOT EXISTS idx_lines_txn ON lines(transaction_id);
CREATE INDEX IF NOT EXISTS idx_lines_acct_amount ON lines(account_id, amount);
-- superseded by the covering index above, which serves everything it did
DROP INDEX IF EXISTS idx_lines_acct;
CREATE INDEX IF NOT EXISTS idx_txn_date ON transactions(date);

CREATE TABLE IF NOT EXISTS tax_codes (
    id TEXT PRIMARY KEY,
    description TEXT DEFAULT '',
    rate_percent REAL NOT NULL DEFAULT 0,
    collected_account TEXT DEFAULT '',
    paid_account TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS import_rules (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    keyword TEXT NOT NULL,
    account_name TEXT NOT NULL,
    tax_code TEXT DEFAULT '',
    priority INTEGER DEFAULT 0,
    notes TEXT DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_rules_kw ON import_rules(keyword);

CREATE TABLE IF NOT EXISTS workpapers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    fy TEXT DEFAULT '',
    ref TEXT NOT NULL,
    description TEXT DEFAULT '',
    path TEXT DEFAULT '',
    prep_by TEXT DEFAULT '',
    rev_by TEXT DEFAULT '',
    to_print INTEGER DEFAULT 0,
    folder_id INTEGER DEFAULT 0,
    UNIQUE(fy, ref)
);

CREATE TABLE IF NOT EXISTS wp_folders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    parent_id INTEGER,
    sort INTEGER DEFAULT 0
);
"""

# ─── Meta ──────────────────────────────────────────────────────────
def get_meta(key, default=''):
    with get_db() as db:
        row = db.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
        return row['value'] if row else default

def set_meta(key, value):
    with get_db() as db:
        db.execute("INSERT OR REPLACE INTO meta(key,value) VALUES(?,?)", (key, value))

def fiscal_anchor():
    """Single source of truth for the current/prior fiscal-year boundaries,
    derived from the BOOK's settings (fiscal_year + fiscal_year_end) — NOT from
    wall-clock time. Returns a dict of ISO date strings, or None if fiscal_year
    is unusable (callers should fall back).

      CY = the fiscal year that ENDS on <fiscal_year>-<fiscal_year_end>.
      PY = the prior fiscal year (the comparative).
    Post-YE activity is anything dated after cy_end — still valid, just after cutoff.
    """
    from datetime import date, timedelta
    import calendar
    fye_md = get_meta('fiscal_year_end', '') or '12-31'
    try:
        m, d = int(fye_md.split('-')[0]), int(fye_md.split('-')[1])
    except (ValueError, IndexError):
        m, d = 12, 31
    try:
        fy = int(get_meta('fiscal_year', ''))
    except (ValueError, TypeError):
        return None

    def year_end(y):
        last = calendar.monthrange(y, m)[1]  # clamp (e.g. Feb-29 on a non-leap year)
        return date(y, m, min(d, last))

    def year_start(y):
        return year_end(y - 1) + timedelta(days=1)  # day after prior year-end

    return {
        'fy': fy,
        'cy_start': year_start(fy).isoformat(),
        'cy_end':   year_end(fy).isoformat(),
        'py_start': year_start(fy - 1).isoformat(),
        'py_end':   year_end(fy - 1).isoformat(),
        'next_end': year_end(fy + 1).isoformat(),
    }

# ─── Fiscal dates: two knobs, not one ─────────────────────────────────────
# The year-end being WORKED ON (fiscal_year + fiscal_year_end) is what the
# statements report. How far posting stays open past it is a separate choice
# with exactly two settings: the same year-end, or one year ahead. The
# ceiling is DERIVED from those, never keyed — a free-form ceiling drifts
# away from the year-end the moment either one is edited on its own.

def year_end_on(y, m, d):
    """The year-end date in year <y>, clamped for short months (Feb-29)."""
    import calendar
    return date(y, m, min(d, calendar.monthrange(y, m)[1]))

def _fye_md():
    """(month, day) of the fiscal year end, defaulting to 31 December."""
    try:
        md = (get_meta('fiscal_year_end', '') or '12-31').split('-')
        return int(md[0]), int(md[1])
    except (ValueError, IndexError):
        return 12, 31

def ceiling_for(fy, m, d, mode):
    """The ceiling a given fiscal setting would produce. Pure — no reads."""
    return year_end_on(fy + (1 if mode == 'next' else 0), m, d).isoformat()

def fiscal_ceiling():
    """The last date that may be posted."""
    a = fiscal_anchor()
    if not a:
        return get_meta('fy_end_date', '')   # no usable fiscal year: legacy value stands
    m, d = _fye_md()
    return ceiling_for(a['fy'], m, d, get_meta('fy_ceiling_mode', 'cy'))

def latest_transaction_date():
    """MAX posted date, or '' on empty books."""
    with get_db() as db:
        row = db.execute("SELECT MAX(date) FROM transactions").fetchone()
    return (row[0] or '') if row else ''

def transactions_after(date_str):
    """How many postings sit past a proposed ceiling."""
    with get_db() as db:
        return db.execute("SELECT COUNT(*) FROM transactions WHERE date > ?", (date_str,)).fetchone()[0]

def ensure_fiscal_settings():
    """Migration + healer for the fiscal dates, run from the init_db funnel so
    web/CLI/MCP converge on one answer.

    The legacy `fy_end_date` key held a free-form ceiling stamped ONCE from the
    wall clock when the file was created and never recomputed, so it drifted
    away from the year-end being worked on (a 31-May-2026 file carrying a
    31-Dec-2026 ceiling). It is folded into `fy_ceiling_mode` here and the key
    deleted — one stored fact, not two that can disagree."""
    legacy = get_meta('fy_end_date', '')

    # A file with no fiscal year but a legacy ceiling: adopt the ceiling's date.
    if not get_meta('fiscal_year', '') and legacy:
        try:
            y, m, d = (int(x) for x in legacy.split('-'))
            date(y, m, d)
            set_meta('fiscal_year', str(y))
            if not get_meta('fiscal_year_end', ''):
                set_meta('fiscal_year_end', '%02d-%02d' % (m, d))
        except (ValueError, TypeError):
            pass

    a = fiscal_anchor()
    if not a:
        return   # nothing to derive from; leave the legacy value alone

    if not get_meta('fy_ceiling_mode', ''):
        # NEVER narrow an existing book on migration. A legacy ceiling past the
        # year-end, or postings already dated past it, both mean the operator is
        # working a year ahead.
        mode = 'cy'
        if legacy and legacy > a['cy_end']:
            mode = 'next'
        latest = latest_transaction_date()
        if latest and latest > a['cy_end']:
            mode = 'next'
        set_meta('fy_ceiling_mode', mode)

    if legacy:
        with get_db() as db:
            db.execute("DELETE FROM meta WHERE key='fy_end_date'")

def set_fiscal_settings(company_name=None, working_ye=None, ceiling_mode=None, lock_date=None):
    """THE guarded writer for the fiscal dates — every interface goes through it.
    The whole set is validated together and NOTHING is written unless it all
    holds, so a refused save can never leave half a fiscal year behind.

    working_ye   the year-end being worked on, YYYY-MM-DD
    ceiling_mode 'cy' (posting stops at that year-end) or 'next' (one year on)
    Omitted arguments keep their current value."""
    a = fiscal_anchor()
    cur_m, cur_d = _fye_md()

    if not (working_ye or '').strip():
        working_ye = None
    if working_ye is None:
        if not a:
            raise ValueError("This file has no fiscal year end set. Enter the year-end being worked on.")
        fy, m, d = a['fy'], cur_m, cur_d
    else:
        working_ye = (working_ye or '').strip()
        try:
            y, m, d = (int(x) for x in working_ye.split('-'))
            ye = date(y, m, d)
        except (ValueError, TypeError):
            raise ValueError(f"'{working_ye}' is not a date. Enter the year-end as YYYY-MM-DD, e.g. 2026-05-31.")
        fy, m, d = ye.year, ye.month, ye.day

    if ceiling_mode is None:
        ceiling_mode = get_meta('fy_ceiling_mode', 'cy') or 'cy'
    if ceiling_mode not in ('cy', 'next'):
        raise ValueError("System fiscal year must be 'cy' (same as the year-end) or 'next' (one year ahead).")

    ceiling = ceiling_for(fy, m, d, ceiling_mode)

    if lock_date is None:
        lock_date = get_meta('lock_date', '')
    lock_date = (lock_date or '').strip()
    if lock_date:
        try:
            date(*(int(x) for x in lock_date.split('-')))
        except (ValueError, TypeError):
            raise ValueError(f"'{lock_date}' is not a date. Enter the lock date as YYYY-MM-DD, or leave it blank.")

    # THE self-check: a ceiling can move forward freely, but never back behind
    # postings that already exist — those entries would fall outside the system.
    latest = latest_transaction_date()
    if latest and latest > ceiling:
        n = transactions_after(ceiling)
        raise ValueError(
            f"Refused: that setting ends the fiscal system at {ceiling}, but "
            f"{n} transaction{'s' if n != 1 else ''} {'are' if n != 1 else 'is'} dated after it "
            f"(latest {latest}). Move the year-end forward, set the system fiscal year one year "
            f"ahead, or remove those entries first.")

    if lock_date and lock_date >= ceiling:
        raise ValueError(
            f"Refused: a lock date of {lock_date} closes everything up to the ceiling "
            f"({ceiling}), leaving nothing postable.")

    if company_name is not None:
        set_meta('company_name', company_name)
    set_meta('fiscal_year', str(fy))
    set_meta('fiscal_year_end', '%02d-%02d' % (m, d))
    set_meta('fy_ceiling_mode', ceiling_mode)
    set_meta('lock_date', lock_date)
    return {'fy': fy, 'working_ye': year_end_on(fy, m, d).isoformat(),
            'ceiling': ceiling, 'ceiling_mode': ceiling_mode, 'lock_date': lock_date}

# ─── Reports ──────────────────────────────────────────────────────
def get_reports():
    with get_db() as db:
        return db.execute("SELECT * FROM reports ORDER BY sort_order, id").fetchall()

def get_report(report_id):
    with get_db() as db:
        return db.execute("SELECT * FROM reports WHERE id=?", (report_id,)).fetchone()

def add_report(name, description='', sort_order=0):
    with get_db() as db:
        cur = db.execute("INSERT INTO reports(name, description, sort_order) VALUES(?,?,?)",
            (name, description, sort_order))
        return cur.lastrowid

def update_report(report_id, description=None, sort_order=None):
    """Update a report's description and/or sort order."""
    with get_db() as db:
        rpt = db.execute("SELECT * FROM reports WHERE id=?", (report_id,)).fetchone()
        if not rpt:
            raise ValueError(f"Report ID {report_id} not found.")
        new_desc = description if description is not None else rpt['description']
        new_sort = sort_order if sort_order is not None else rpt['sort_order']
        db.execute("UPDATE reports SET description=?, sort_order=? WHERE id=?",
                   (new_desc, new_sort, report_id))
        return dict(db.execute("SELECT * FROM reports WHERE id=?", (report_id,)).fetchone())

# ─── Accounts ─────────────────────────────────────────────────────
def get_account(account_id):
    with get_db() as db:
        return db.execute("SELECT * FROM accounts WHERE id=?", (account_id,)).fetchone()

def get_account_by_name(name):
    with get_db() as db:
        return db.execute("SELECT * FROM accounts WHERE name=? COLLATE NOCASE", (name,)).fetchone()

def get_accounts():
    with get_db() as db:
        return db.execute("SELECT * FROM accounts ORDER BY name").fetchall()

def add_account(name, normal_balance='D', description='', account_type='posting', account_number=''):
    with get_db() as db:
        cur = db.execute(
            "INSERT INTO accounts(name, normal_balance, description, account_type, account_number) VALUES(?,?,?,?,?)",
            (name, normal_balance, description, account_type, account_number))
        return cur.lastrowid

def search_accounts(query):
    with get_db() as db:
        q = f"%{query}%"
        return db.execute("SELECT * FROM accounts WHERE name LIKE ? OR description LIKE ? ORDER BY name", (q, q)).fetchall()

def accounts_on_any_report():
    """Set of account ids that appear on at least one report. The post autocomplete uses
    this to avoid offering orphan accounts (on no report = not a real chart account, so
    posting to it would never show on a statement)."""
    with get_db() as db:
        return {r['account_id'] for r in db.execute(
            "SELECT DISTINCT account_id FROM report_items WHERE account_id IS NOT NULL")}


# ─── GIFI Mapping ─────────────────────────────────────────────────
def set_gifi(account_name, gifi_code):
    """Set the GIFI code on a posting account. Code must exist in GIFI_CODES."""
    if gifi_code and gifi_code not in GIFI_CODES:
        raise ValueError(f"Invalid GIFI code '{gifi_code}'. Must be a valid CRA GIFI code.")
    with get_db() as db:
        acct = db.execute("SELECT id, account_type FROM accounts WHERE name = ?",
                          (account_name,)).fetchone()
        if not acct:
            raise ValueError(f"Account not found: {account_name}")
        db.execute("UPDATE accounts SET gifi_code = ? WHERE id = ?",
                   (gifi_code or '', acct['id']))


def get_gifi_map():
    """Return all accounts with GIFI codes set."""
    with get_db() as db:
        return db.execute(
            "SELECT name, description, gifi_code FROM accounts WHERE gifi_code != '' ORDER BY name"
        ).fetchall()


def gifi_export(date_from=None, date_to=None):
    """Roll up GL balances by GIFI code for T2 S100/S125 output.

    Returns both a structured breakdown (schedule_100/schedule_125) and
    a t2engine_input dict ready to write as gifi.json for T2Engine.

    T2Engine sign convention:
      Assets, expenses = positive (debit-normal)
      Liabilities, equity, revenue = negative (credit-normal)
    """
    with get_db() as db:
        accounts = db.execute(
            "SELECT id, name, description, normal_balance, gifi_code FROM accounts "
            "WHERE gifi_code != '' AND account_type = 'posting'"
        ).fetchall()

    results = {}
    for acct in accounts:
        bal_raw = get_account_balance(acct['id'], date_from=date_from, date_to=date_to)
        # bal_raw is in cents, positive = debit, negative = credit
        bal_dollars = bal_raw / 100

        gifi = acct['gifi_code']
        if gifi not in results:
            results[gifi] = {
                'gifi': gifi,
                'description': GIFI_CODES.get(gifi, ''),
                'amount': 0,
                'amount_natural': 0,
                'accounts': [],
            }
        # T2Engine sign: debit-normal stays positive, credit-normal stays negative
        results[gifi]['amount'] += bal_dollars
        # Natural sign for human-readable output
        sign = 1 if acct['normal_balance'] == 'D' else -1
        results[gifi]['amount_natural'] += bal_dollars * sign
        results[gifi]['accounts'].append(acct['name'])

    # Round amounts
    for v in results.values():
        v['amount'] = round(v['amount'], 2)
        v['amount_natural'] = round(v['amount_natural'], 2)

    # Split into S100 (BS) and S125 (IS)
    s100 = [v for k, v in sorted(results.items()) if int(k) < 4000]
    s125 = [v for k, v in sorted(results.items()) if int(k) >= 4000]

    # T2Engine input format: {gifi_code: whole_dollars, ...}
    t2_input = {}
    for k, v in results.items():
        amt = round(v['amount'])
        if amt != 0:
            t2_input[k] = amt

    return {
        'schedule_100': s100,
        'schedule_125': s125,
        'period': f'{date_from} to {date_to}',
        't2engine_input': t2_input,
    }


# ─── Report Items ─────────────────────────────────────────────────
def get_report_items(report_id):
    with get_db() as db:
        return db.execute(
            "SELECT ri.*, a.name as acct_name, a.description as acct_desc, "
            "a.normal_balance, a.account_type, a.account_number, a.computed, "
            "COALESCE(a.system,0) AS system "
            "FROM report_items ri LEFT JOIN accounts a ON ri.account_id = a.id "
            "WHERE ri.report_id=? ORDER BY ri.position", (report_id,)).fetchall()

def add_report_item(report_id, item_type='account', description='', account_id=None,
                    indent=0, position=None, total_to_1='', total_to_2='',
                    total_to_3='', total_to_4='', total_to_5='', total_to_6='',
                    sep_style=''):
    if item_type == 'total' and account_id is None:
        raise ValueError(
            f"Total report items require an account_id (the total account that accumulates). "
            f"Description: '{description}'. Create a total-type account first, then pass its id.")
    with get_db() as db:
        if position is None:
            row = db.execute("SELECT MAX(position) as mx FROM report_items WHERE report_id=?", (report_id,)).fetchone()
            position = (row['mx'] or 0) + 10
        cur = db.execute(
            "INSERT INTO report_items(report_id, position, item_type, description, account_id, "
            "indent, total_to_1, total_to_2, total_to_3, total_to_4, total_to_5, total_to_6, sep_style) "
            "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (report_id, position, item_type, description, account_id,
             indent, total_to_1, total_to_2, total_to_3, total_to_4, total_to_5, total_to_6, sep_style))
        new_id = cur.lastrowid
        # Resequence to clean up any collisions
        _resequence(db, report_id)
        return new_id

def _resequence(db, report_id):
    """Resequence all items in a report to positions 10, 20, 30, ..."""
    items = db.execute(
        "SELECT id FROM report_items WHERE report_id=? ORDER BY position, id",
        (report_id,)).fetchall()
    for i, item in enumerate(items):
        db.execute("UPDATE report_items SET position=? WHERE id=?", ((i + 1) * 10, item['id']))

def resequence_report(report_id):
    """Public wrapper for resequencing."""
    with get_db() as db:
        _resequence(db, report_id)

def move_report_item(item_id, direction):
    """Move a report item up (-1) or down (+1). Returns True if moved."""
    with get_db() as db:
        item = db.execute("SELECT id, report_id, position FROM report_items WHERE id=?", (item_id,)).fetchone()
        if not item:
            return False
        report_id = item['report_id']
        # Get all items in order
        items = db.execute(
            "SELECT id, position FROM report_items WHERE report_id=? ORDER BY position, id",
            (report_id,)).fetchall()
        # Find current index
        idx = None
        for i, it in enumerate(items):
            if it['id'] == item_id:
                idx = i
                break
        if idx is None:
            return False
        swap_idx = idx + direction
        if swap_idx < 0 or swap_idx >= len(items):
            return False
        # Swap the two items' positions
        my_pos = items[idx]['position']
        other_pos = items[swap_idx]['position']
        other_id = items[swap_idx]['id']
        # If positions are the same (collision), assign distinct values first
        if my_pos == other_pos:
            # Resequence everything, then re-find and swap
            _resequence(db, report_id)
            items = db.execute(
                "SELECT id, position FROM report_items WHERE report_id=? ORDER BY position, id",
                (report_id,)).fetchall()
            for i, it in enumerate(items):
                if it['id'] == item_id:
                    idx = i
                    break
            swap_idx = idx + direction
            if swap_idx < 0 or swap_idx >= len(items):
                return False
            my_pos = items[idx]['position']
            other_pos = items[swap_idx]['position']
            other_id = items[swap_idx]['id']
        db.execute("UPDATE report_items SET position=? WHERE id=?", (other_pos, item_id))
        db.execute("UPDATE report_items SET position=? WHERE id=?", (my_pos, other_id))
        return True

def update_report_item(item_id, **kwargs):
    """Update any fields on a report item. Pass field=value pairs."""
    allowed = {'description','indent','total_to_1','total_to_2','total_to_3',
               'total_to_4','total_to_5','total_to_6','sep_style','position','item_type'}
    fields = {k: v for k, v in kwargs.items() if k in allowed}
    if not fields:
        return
    with get_db() as db:
        sets = ', '.join(f'{k}=?' for k in fields)
        vals = list(fields.values()) + [item_id]
        db.execute(f"UPDATE report_items SET {sets} WHERE id=?", vals)

def delete_report_item(item_id):
    """Delete a report item. Refuses if it's a total account that other items reference."""
    with get_db() as db:
        item = db.execute("SELECT ri.*, a.name as acct_name FROM report_items ri "
            "LEFT JOIN accounts a ON ri.account_id=a.id WHERE ri.id=?", (item_id,)).fetchone()
        if not item:
            raise ValueError("Item not found")
        check_trx_pinned(item_id)
        # If it's an account/total with a name, check if anything totals TO it
        if item['acct_name']:
            refs = db.execute("SELECT COUNT(*) as cnt FROM report_items WHERE "
                "total_to_1=? OR total_to_2=? OR total_to_3=? OR total_to_4=? OR total_to_5=? OR total_to_6=?",
                tuple([item['acct_name']] * 6)).fetchone()
            if refs['cnt'] > 0:
                raise ValueError(f"Cannot delete: {refs['cnt']} item(s) total to {item['acct_name']}")
        # If it's a posting account with transactions, refuse
        if item['account_id'] and item['item_type'] == 'account':
            acct = db.execute("SELECT account_type FROM accounts WHERE id=?", (item['account_id'],)).fetchone()
            if acct and acct['account_type'] == 'posting':
                txns = db.execute("SELECT COUNT(*) as cnt FROM lines WHERE account_id=?", (item['account_id'],)).fetchone()
                if txns['cnt'] > 0:
                    raise ValueError(f"Cannot delete: account has {txns['cnt']} transaction line(s)")
        db.execute("DELETE FROM report_items WHERE id=?", (item_id,))
        # If that was the account's last appearance on ANY report and it carries no
        # transactions, delete the account record too — so a "deleted" account is truly
        # gone: no orphan lingering as a post option, and its name can be re-added.
        if item['account_id']:
            still = db.execute("SELECT COUNT(*) AS c FROM report_items WHERE account_id=?",
                               (item['account_id'],)).fetchone()['c']
            ntxn = db.execute("SELECT COUNT(*) AS c FROM lines WHERE account_id=?",
                              (item['account_id'],)).fetchone()['c']
            if still == 0 and ntxn == 0:
                db.execute("DELETE FROM accounts WHERE id=?", (item['account_id'],))

def clear_report_items(report_id):
    """Delete ALL items from a report. Used by bulk_report_layout replace mode."""
    with get_db() as db:
        db.execute("DELETE FROM report_items WHERE report_id=?", (report_id,))

def update_account(account_id, description=None, account_number=None):
    """Update account description and/or account number."""
    with get_db() as db:
        if description is not None:
            db.execute("UPDATE accounts SET description=? WHERE id=?", (description, account_id))
        if account_number is not None:
            db.execute("UPDATE accounts SET account_number=? WHERE id=?", (account_number, account_id))

def find_report_for_account(account_id):
    """Find which report contains this account (returns first match)."""
    with get_db() as db:
        row = db.execute("""
            SELECT r.* FROM reports r
            JOIN report_items ri ON ri.report_id = r.id
            WHERE ri.account_id = ? AND ri.item_type = 'account'
            ORDER BY r.id LIMIT 1
        """, (account_id,)).fetchone()
        return dict(row) if row else None

def get_report_accounts(report_id):
    """Get all posting accounts belonging to a report, in position order."""
    with get_db() as db:
        return db.execute("""
            SELECT a.id, a.name, a.description, a.normal_balance
            FROM report_items ri
            JOIN accounts a ON ri.account_id = a.id
            WHERE ri.report_id = ? AND ri.item_type = 'account'
              AND a.account_type = 'posting'
            ORDER BY ri.position
        """, (report_id,)).fetchall()

def find_report_by_name(name):
    """Find a report by name (case-insensitive partial match)."""
    with get_db() as db:
        row = db.execute("SELECT * FROM reports WHERE name LIKE ? COLLATE NOCASE ORDER BY id LIMIT 1",
            (f'%{name}%',)).fetchone()
        return dict(row) if row else None

# ─── Transactions & Lines ─────────────────────────────────────────

def generate_ref():
    """Generate a random 5-char alphanumeric reference (lowercase + digits)."""
    import random, string
    chars = string.ascii_lowercase + string.digits
    return ''.join(random.choice(chars) for _ in range(5))

def _meta_on(db, key, default=''):
    """get_meta() against an ALREADY-OPEN connection.

    get_meta() opens its own connection, which is right for a one-off read and
    ruinous inside a posting loop: a fresh sqlite3.connect + two PRAGMAs +
    commit + close for EVERY transaction. Profiled on a 12,000-row batch it was
    88% of the entire cost of the import, and connection-open is far dearer on
    Windows than it is here — a real conversion measured 20x slower per row on
    the operator's machine than on this one."""
    row = db.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    return row['value'] if row else default


def _entry_error(date_str, lines, totals, lock, ceiling):
    """The posting rules, as pure checks over data already in hand.

    ONE definition, called by BOTH doors — the single-entry door
    (_post_transaction_db) and the bulk door (post_bulk_batch) — so the fast
    path and the careful path can never drift apart on what is allowed.
    Returns the refusal text, or None if the entry may be posted.

    totals: {account_id: name} for total accounts. It only ever needs to cover
    the ids in `lines`; the single door looks up just those, the bulk door
    fetches the whole chart once.
    """
    total = sum(l[1] for l in lines)
    if total != 0:
        return f"Transaction does not balance: off by {total/100:.2f}"
    if len(lines) < 2:
        return "Transaction must have at least 2 lines"
    for acct_id, amount, desc in lines:
        if acct_id in totals:
            return (f"Cannot post to '{totals[acct_id]}' — it is a total account. "
                    f"Post to a detail account instead.")
    if lock and date_str <= lock:
        return (f"Date {date_str} is on or before the lock date ({lock}). "
                f"Posting is not allowed.")
    if ceiling and date_str > ceiling:
        return (f"Date {date_str} is after the fiscal year end ({ceiling}). "
                f"Check the date, or open the next year in Options → System Fiscal Year.")
    return None


def _post_transaction_db(db, date_str, reference, description, lines, _bypass_ceiling=False, batch=''):
    """Validate and insert ONE transaction on an already-open connection.
    No commit here — the caller owns the transaction boundary, which is what
    makes batch imports atomic. Every posting guard lives here so all paths
    (single post, batch import) share identical rules."""
    total = sum(l[1] for l in lines)
    if total != 0:
        raise ValueError(f"Transaction does not balance: off by {total/100:.2f}")
    if len(lines) < 2:
        raise ValueError("Transaction must have at least 2 lines")
    # Auto-assign reference if blank
    if not reference or not reference.strip():
        reference = generate_ref()
    # Block posting to total accounts
    # ONE question for the whole entry, not one per line. Restricted to the ids
    # actually being posted — _entry_error only ever looks up those — so this is
    # the same dict the bulk door builds from the whole chart.
    slots = ','.join('?' * len(lines))
    totals = {r['id']: r['name'] for r in db.execute(
        f"SELECT id, name FROM accounts WHERE id IN ({slots}) AND account_type='total'",
        [l[0] for l in lines])}
    err = _entry_error(date_str, lines, totals,
                       _meta_on(db, 'lock_date'),
                       '' if _bypass_ceiling else fiscal_ceiling())
    if err:
        raise ValueError(err)
    cur = db.execute("INSERT INTO transactions(date, reference, description, import_batch) VALUES(?,?,?,?)",
        (date_str, reference, description, batch))
    txn_id = cur.lastrowid
    for i, (acct_id, amount, desc) in enumerate(lines):
        db.execute("INSERT INTO lines(transaction_id, account_id, amount, description, sort_order) VALUES(?,?,?,?,?)",
            (txn_id, acct_id, amount, desc, i))
    return txn_id

def add_transaction(date_str, reference, description, lines, _bypass_ceiling=False):
    with get_db() as db:
        return _post_transaction_db(db, date_str, reference, description, lines, _bypass_ceiling)

def add_simple_transaction(date_str, reference, description, debit_acct_id, credit_acct_id, amount_cents):
    return add_transaction(date_str, reference, description, [
        (debit_acct_id, amount_cents, description),
        (credit_acct_id, -amount_cents, description)])

# ─── Import batches ────────────────────────────────────────────────
# An import is a UNIT: it lands whole or not at all, and it can be deleted
# whole. That is the entire robustness story for bulk data — any mistake
# (wrong file, wrong account, wrong mapping) is cured by delete-and-redo,
# never by hand-picking rows out of a half-landed import.

def _post_batch_db(db, txns, prefix='imp', avoid=''):
    """Post a batch on an ALREADY-OPEN connection. No commit — the caller owns
    the transaction boundary, which is what lets a conversion be REPLACED
    (delete + repost) without a window where the client has no opening position."""
    if not txns:
        raise ValueError("Nothing to import")
    batch_id = prefix + '-' + datetime.now().strftime('%Y%m%d-%H%M%S')
    # Two batches inside the same second must not share a tag
    n = db.execute("SELECT COUNT(*) FROM transactions WHERE import_batch=?", (batch_id,)).fetchone()[0]
    if n or (avoid and batch_id == avoid):
        # `avoid` is the batch being REPLACED in this same transaction: its rows
        # are already gone, so a row count cannot see the clash.
        batch_id += '-' + generate_ref()[:3]
    ids = []
    for row_num, (d, ref, desc, lines) in enumerate(txns, start=1):
        try:
            ids.append(_post_transaction_db(db, d, ref, desc, lines, batch=batch_id))
        except ValueError as e:
            raise ValueError(f"row {row_num} ('{str(desc)[:40]}'): {e}")
    return batch_id, ids


BULK_CHUNK = 5000        # rows per executemany; keeps memory flat on a big book


def post_bulk_batch(txns, batch, bypass_ceiling=False, progress=None, total=0):
    """Post a whole prepared batch in ONE pass. The bulk door.

    Same rules as the single-entry door — they share `_entry_error` — but asked
    with the chart and the lock date read ONCE, and the rows written with
    executemany instead of a round trip per line. Still one db transaction, so
    it is still all-or-nothing: a row that breaks a rule refuses the WHOLE
    batch, naming the row, with nothing written.

    Why it exists: the per-entry door re-read `meta` for every transaction, and
    `get_meta` opens its own connection. On a converted legacy book that was a
    million fresh OS handles on books.db — profiled at 88% of import cost even
    on a quiet machine, and far worse on Windows, where a scanner inspects each
    open. Bulk data gets a bulk door. See [feedback] GridTRX is a bulk data
    processor, not an audit-trail ledger.

    txns: an iterable of (date_str, reference, description,
          [(account_id, amount_cents, line_desc), ...]). Consumed lazily, so a
          caller may hand over a generator and never build the second list.
    total: how many rows are coming, when the caller knows — a generator cannot
          be measured, and a counter with no denominator reads as a hang.
    Returns (batch, n_posted).
    """
    tell = progress or (lambda *a: None)
    posted = 0
    with get_db() as db:
      try:
        totals = {r['id']: r['name'] for r in
                  db.execute("SELECT id, name FROM accounts WHERE account_type='total'")}
        known = {r['id'] for r in db.execute("SELECT id FROM accounts")}
        lock = _meta_on(db, 'lock_date')
        ceiling = '' if bypass_ceiling else fiscal_ceiling()
        # AUTOINCREMENT keeps its own counter, which sits ABOVE max(id) once rows
        # have been deleted. Take the higher of the two or the explicit ids
        # collide with a rowid SQLite has already handed out.
        base = db.execute("SELECT COALESCE(MAX(id), 0) FROM transactions").fetchone()[0]
        seq = db.execute("SELECT seq FROM sqlite_sequence WHERE name='transactions'").fetchone()
        base = max(base, seq[0] if seq else 0)

        tbuf, lbuf = [], []
        def flush():
            if tbuf:
                db.executemany("INSERT INTO transactions(id, date, reference, "
                               "description, import_batch) VALUES(?,?,?,?,?)", tbuf)
                db.executemany("INSERT INTO lines(transaction_id, account_id, "
                               "amount, description, sort_order) VALUES(?,?,?,?,?)", lbuf)
                tbuf.clear(); lbuf.clear()
                # COMMIT each chunk. Holding one transaction over a million rows
                # meant nothing could fold out of the write-ahead log until the
                # very end: on a real conversion the -wal reached 211 MB, and the
                # commit-plus-checkpoint that followed took an exclusive lock for
                # minutes, freezing every page in the app with no counter moving.
                # Committing as we go keeps the log small and the writes steady.
                db.commit()

        for ix, (date_str, reference, description, lines) in enumerate(txns):
            for acct_id, amount, desc in lines:
                if acct_id not in known:
                    raise ValueError(
                        f"Row {ix + 1} ({date_str}): account id {acct_id} is not "
                        f"in this chart. Nothing was posted.")
            err = _entry_error(date_str, lines, totals, lock, ceiling)
            if err:
                raise ValueError(f"Row {ix + 1} ({date_str}): {err} Nothing was posted.")
            if not reference or not str(reference).strip():
                reference = generate_ref()
            base += 1
            tbuf.append((base, date_str, reference, description, batch))
            for i, (acct_id, amount, desc) in enumerate(lines):
                lbuf.append((base, acct_id, amount, desc, i))
            posted += 1
            if len(tbuf) >= BULK_CHUNK:
                tell('Posting transactions', posted, total)
                flush()
        flush()
      except BaseException:
        # Chunked commits mean earlier chunks are already on disk, so
        # "nothing was posted" has to be MADE true rather than left to a
        # rollback. The batch tag is what makes that a single sweep — the same
        # tag Undo uses. Anything short of the process being killed outright
        # still leaves the books exactly as they were.
        db.rollback()
        db.execute("DELETE FROM lines WHERE transaction_id IN "
                   "(SELECT id FROM transactions WHERE import_batch=?)", (batch,))
        db.execute("DELETE FROM transactions WHERE import_batch=?", (batch,))
        db.commit()
        raise
    return batch, posted


def post_import_batch(txns, prefix='imp'):
    """Atomically post a prepared list of transactions as one tagged batch.

    txns: [(date_str, reference, description, lines), ...]
    Returns (batch_id, txn_ids). ANY row failing rolls back the ENTIRE batch,
    so a re-run after fixing the file can never create partial duplicates.
    """
    with get_db() as db:
        return _post_batch_db(db, txns, prefix)

# SQL predicate: is a line reconciled? The reconciled column holds 0 (open) or a
# TAG (LAP [Rec] value — statement date, payment ref, or plain 1). Text tags mean
# `reconciled=1` no longer matches; use this expression instead.
REC_SQL = "(l.reconciled IS NOT NULL AND l.reconciled NOT IN (0,''))"

def list_import_batches(limit=10):
    """Recent import batches with enough context to recognize and undo them."""
    with get_db() as db:
        return db.execute(f"""
            SELECT t.import_batch AS batch_id,
                   COUNT(*) AS txn_count,
                   MIN(t.date) AS date_from, MAX(t.date) AS date_to,
                   MAX(t.created_at) AS imported_at,
                   (SELECT a.name FROM lines l JOIN accounts a ON a.id = l.account_id
                      JOIN transactions t2 ON t2.id = l.transaction_id
                     WHERE t2.import_batch = t.import_batch
                     GROUP BY a.name ORDER BY COUNT(*) DESC, a.name LIMIT 1) AS main_account,
                   (SELECT COUNT(*) FROM lines l JOIN transactions t3 ON t3.id = l.transaction_id
                     WHERE t3.import_batch = t.import_batch AND {REC_SQL}) AS reconciled_count
            FROM transactions t
            WHERE t.import_batch != ''
            GROUP BY t.import_batch
            ORDER BY imported_at DESC, batch_id DESC LIMIT ?""", (limit,)).fetchall()

def delete_import_batch(batch_id):
    """Delete every transaction in an import batch, atomically. The undo half of
    delete-and-redo: fix the file or the rules, wipe the batch, import again.
    Refuses if any line is reconciled or any date is on/before the lock —
    at that point it's no longer 'just an import'."""
    if not batch_id or not str(batch_id).strip():
        raise ValueError("No batch id given")
    with get_db() as db:
        n = db.execute("SELECT COUNT(*) FROM transactions WHERE import_batch=?", (batch_id,)).fetchone()[0]
        if not n:
            raise ValueError(f"No transactions found for batch '{batch_id}'")
        rec = db.execute(
            "SELECT COUNT(*) FROM lines l JOIN transactions t ON t.id = l.transaction_id "
            f"WHERE t.import_batch=? AND {REC_SQL}", (batch_id,)).fetchone()[0]
        if rec:
            raise ValueError(f"Batch has {rec} reconciled line(s) — unreconcile them first if you really mean to delete it.")
        lock = get_meta('lock_date', '')
        if lock:
            locked = db.execute("SELECT COUNT(*) FROM transactions WHERE import_batch=? AND date<=?",
                                (batch_id, lock)).fetchone()[0]
            if locked:
                raise ValueError(f"Batch has {locked} transaction(s) on or before the lock date ({lock}).")
        db.execute("DELETE FROM lines WHERE transaction_id IN "
                   "(SELECT id FROM transactions WHERE import_batch=?)", (batch_id,))
        db.execute("DELETE FROM transactions WHERE import_batch=?", (batch_id,))
        return n

# ─── Opening balances (conversion) ─────────────────────────────────
# ONE door for putting a client's opening position into Grid. Before this, the
# conversion was a written procedure (create TRX.OPEN, hand-post a 2-line entry
# per account, work out retained earnings yourself, post it to RE.OB, check
# TRX.OPEN nets to zero) — and every one of those steps is a place to drift.
# Staff were skipping TRX.OPEN, keying retained earnings by hand, and parking
# the whole conversion in EX.SUSP and clearing it down. None of that is possible
# through here: retained earnings is COMPUTED as the residual and never keyed,
# TRX.OPEN is created and used for you, and the whole conversion lands as ONE
# atomic batch that deletes and redoes in a single action.

OPENING_PREFIX  = 'open'      # transactions.import_batch tag → 'open-<stamp>'
OPENING_REF     = 'OPEN'      # every conversion entry carries this reference
OPENING_DESC    = 'Opening balance'
CONVERSION_ACCT = 'TRX.OPEN'
OPENING_RE_ACCT = 'RE.OB'

# Accounts that may never appear in a conversion grid, and why. The message is
# the whole point: it teaches the rule at the moment the rule is broken.
OPENING_BLOCKED = {
    'RE':       "Retained earnings is a total account — it accumulates, you cannot post to it.",
    'RE.OB':    "Opening retained earnings is COMPUTED for you from the balances you enter — "
                "check the figure at the bottom against the prior-year statements.",
    'RE.OPEN':  "That is a computed display line on the income statement, not a posting account.",
    'RE.CLOSE': "That is a computed display line on the income statement, not a posting account.",
    'EX.SUSP':  "Suspense is for transactions you cannot identify yet — never for opening "
                "balances. Enter the real account this balance belongs to.",
    CONVERSION_ACCT: "Grid posts the conversion contra for you — just enter the real accounts.",
}


def opening_batch():
    """The conversion batch tag, or '' if opening balances were never posted.
    Derived from the data, not a flag: delete the batch and the books are back
    to needing openings, with nothing left over to contradict that."""
    # A prefix RANGE, not LIKE: SQLite will not use an index for LIKE, so the
    # old form scanned every transaction — on a converted legacy book that is a
    # million rows, on EVERY home page load. Same rows, 700x less work.
    lo = OPENING_PREFIX + '-'
    hi = OPENING_PREFIX + '.'      # '.' is the next byte up from '-', so
                                   # [lo, hi) is exactly the lo-prefixed tags
    with get_db() as db:
        r = db.execute("SELECT import_batch FROM transactions "
                       "WHERE import_batch >= ? AND import_batch < ? "
                       "ORDER BY id LIMIT 1", (lo, hi)).fetchone()
    return r['import_batch'] if r else ''


def openings_state():
    """Where this set of books stands on opening balances.

      needed   — nothing posted at all; these are new books (offer the module)
      posted   — a conversion batch is in place
      declined — brand-new client, told us to start at zero
      later    — has activity but no conversion (don't nag; the TRX report has
                 a permanent way in)
    """
    batch = opening_batch()
    with get_db() as db:
        txns = db.execute("SELECT COUNT(*) FROM transactions").fetchone()[0]
        conv = date = ''
        lines = 0
        if batch:
            row = db.execute("SELECT MIN(date) d, COUNT(*) n FROM transactions "
                             "WHERE import_batch=?", (batch,)).fetchone()
            date, lines = row['d'] or '', row['n'] or 0
    if batch:
        status = 'posted'
    elif get_meta('openings_declined', '') == '1':
        status = 'declined'
    elif txns == 0:
        status = 'needed'
    else:
        status = 'later'
    return {'status': status, 'batch': batch, 'conversion_date': date,
            'entry_count': lines, 'txn_count': txns,
            'default_date': default_conversion_date()}


def default_conversion_date():
    """Opening balances are the PRIOR year's closing balances, so the default is
    the prior fiscal year end — from the book's fiscal settings, never the clock."""
    a = fiscal_anchor()
    return a['py_end'] if a else ''


def ensure_conversion_account():
    """TRX.OPEN — the contra every conversion entry posts against. It nets to
    ZERO once the trial balance is complete, which is exactly why it totals to
    nothing and lives off-statement. Its existence and its place on the TRX
    report are the healer's job, so this just makes sure the head is in shape
    and hands the account back."""
    ensure_trx_layout()
    return get_account_by_name(CONVERSION_ACCT)


def opening_account_choices():
    """The accounts offerable in a conversion grid: real posting accounts that
    live on a report, minus the ones the module owns or forbids. Same rule as
    the validator, in one place, so the picker can never suggest a dead end."""
    on_report = accounts_on_any_report()
    out = []
    with get_db() as db:
        for a in db.execute("SELECT id, name, description FROM accounts "
                            "WHERE account_type='posting' AND COALESCE(computed,'')='' "
                            "ORDER BY name"):
            if a['id'] in on_report and a['name'].upper() not in OPENING_BLOCKED:
                out.append({'name': a['name'], 'description': a['description'] or ''})
    return out


# ─── The TRX report: fixed head, open body ─────────────────────────
# TRX is NOT the conversion report, whatever its old title said. It is the
# off-statement workings ledger inherited from LAP practice: the place for anything that
# belongs on no statement and in no subledger — recurring postings like
# "25.SHTRX Expenses paid by shareholder 2025", holding accounts, workings. The
# conversion happens to live at the top of it because it had nowhere better to
# be, not because TRX belongs to it.
#
# So the head is fixed and the body is the operator's:
#   1  Transactions Ledger          (heading)
#   2  Opening conversion balances  (TRX.OPEN)   — pinned
#   3  System Opening Retained Earnings (RE.OB)  — pinned
#   4  … everything else, in whatever order the operator wants
# These print. A report line is not the place for an explanation — the account
# says WHAT it is in as few words as carry the meaning, and the ledger banner and
# tooltip carry the WHY. Keep every one of these short enough to read at a glance
# in a column beside a number.
TRX_HEADING     = 'Transactions Ledger'
TRX_REPORT_DESC = 'Transactions Ledger'
CONVERSION_DESC = 'Opening conversion balances'
OPENING_RE_DESC = 'System Opening RE'
TRX_BODY_LABEL  = 'Workings — recurring entries'
TRX_PINNED_POS  = {'heading': 10, CONVERSION_ACCT: 20, OPENING_RE_ACCT: 30}
TRX_BODY_START  = 40

# ─── Adjusting entries ────────────────────────────────────────────────────
# An AJE is ONE accountant's adjustment — a reference, a date, a description and
# two or more account legs. It lands as N separate 2-line transactions, each leg
# paired against the year's JOURNAL account, which therefore nets to zero and
# shows every leg of every adjustment in one ledger. That shape is not special to
# AJE: it is how the reports and the ledger read everywhere, which is exactly why
# the AJE module must produce it rather than invent its own.
#
# The description is the load-bearing field — it prints on EVERY leg — so it is
# capped at report-line length, the [[feedback-security-identity]] "a name is a
# report line" rule applied to an entry.
AJE_HEADING     = 'Adjusting Entries'
AJE_REPORT_DESC = 'Adjusting Entries'
AJE_DESC_MAX    = 50          # prints on every leg — keep it to a report line
AJE_PREFIX      = 'aje'       # transactions.import_batch tag → 'aje-<stamp>'
AJE_PINNED_POS  = {'heading': 10}
AJE_BODY_START  = 20
AJE_REF_RE      = _re.compile(r'^(\d{2})AJE(\d+)$', _re.I)

def aje_report(db=None):
    """The AJE report row, created if a very old file somehow lacks it."""
    def _get(d):
        r = d.execute("SELECT * FROM reports WHERE name='AJE'").fetchone()
        if not r:
            d.execute("INSERT INTO reports(name, description, sort_order) VALUES('AJE',?,30)",
                      (AJE_REPORT_DESC,))
            r = d.execute("SELECT * FROM reports WHERE name='AJE'").fetchone()
        return r
    if db is not None:
        return _get(db)
    with get_db() as d:
        return _get(d)

def ensure_aje_layout():
    """Force the AJE head into shape on EVERY open, for every file however old.

    Line 1 is the heading. Everything below is the operator's own year journals,
    in their existing order — Grid never reorders them, exactly as with TRX.
    Idempotent; runs from init_db so web, CLI and MCP converge."""
    with get_db() as db:
        rpt = aje_report(db)
        if (rpt['description'] or '') != AJE_REPORT_DESC:
            db.execute("UPDATE reports SET description=? WHERE id=?", (AJE_REPORT_DESC, rpt['id']))

        items = db.execute("SELECT * FROM report_items WHERE report_id=? ORDER BY position, id",
                           (rpt['id'],)).fetchall()
        # Identify the heading by IDENTITY (a label carrying the heading text),
        # never by position — an operator's own line may already sit at 10.
        hits = [i for i in items if i['item_type'] == 'label'
                and (i['description'] or '').strip().lower() == AJE_HEADING.lower()]
        for extra in hits[1:]:
            db.execute("DELETE FROM report_items WHERE id=?", (extra['id'],))
        pinned = set()
        if hits:
            db.execute("UPDATE report_items SET position=?, description=?, indent=0 WHERE id=?",
                       (AJE_PINNED_POS['heading'], AJE_HEADING, hits[0]['id']))
            pinned.add(hits[0]['id'])
        else:
            cur = db.execute("INSERT INTO report_items(report_id, position, item_type, description, indent) "
                             "VALUES(?,?,'label',?,0)", (rpt['id'], AJE_PINNED_POS['heading'], AJE_HEADING))
            pinned.add(cur.lastrowid)

        pos = AJE_BODY_START
        for it in db.execute("SELECT id FROM report_items WHERE report_id=? ORDER BY position, id",
                             (rpt['id'],)).fetchall():
            if it['id'] in pinned:
                continue
            db.execute("UPDATE report_items SET position=? WHERE id=?", (pos, it['id']))
            pos += 10

def aje_journals():
    """The year journal accounts on the AJE report, in report order."""
    rpt = aje_report()
    out = []
    with get_db() as db:
        rows = db.execute(
            "SELECT ri.id AS item_id, a.* FROM report_items ri JOIN accounts a ON a.id=ri.account_id "
            "WHERE ri.report_id=? AND ri.account_id IS NOT NULL AND a.account_type='posting' "
            "ORDER BY ri.position, ri.id", (rpt['id'],)).fetchall()
    return [dict(r) for r in rows]

def is_aje_journal(account_id):
    """Is this account a year journal on the AJE report?"""
    return any(j['id'] == account_id for j in aje_journals())

def suggest_aje_batch():
    """What to pre-fill the 'new year of adjustments' box with. Suggestions only —
    the operator retypes either of them if the house style differs."""
    a = fiscal_anchor()
    y = a['fy'] if a else date.today().year
    name = f"{str(y)[-2:]}AJE"
    n, suffix = name, 1
    while get_account_by_name(n) and not is_aje_journal(get_account_by_name(n)['id']):
        suffix += 1
        n = f"{name}-{suffix}"          # the plain name is taken by something else
    return {'account': n, 'description': f"{y} Adjusting Entries", 'year': y}

def create_aje_batch(account_name, description=''):
    """Create a year's journal account and put it on the AJE report.

    This is the ONE door — it guarantees the account is a posting account, is
    credit-normal (it nets to zero, so the side is arbitrary; C matches the
    existing import path), and actually appears on the report where the operator
    will look for it."""
    name = (account_name or '').strip().upper()
    desc = (description or '').strip()
    if not name:
        raise ValueError("Give the year's entries an account name (e.g. 26AJE).")
    if len(name) > 20:
        raise ValueError("Account name is too long — keep it short, like 26AJE.")
    if not desc:
        raise ValueError("Give the year's entries a description (e.g. 2026 Adjusting Entries).")
    if len(desc) > 40:
        raise ValueError(f"Description is {len(desc)} characters — keep it to 40 or fewer; "
                         f"it prints beside a number.")
    existing = get_account_by_name(name)
    if existing:
        if existing['account_type'] != 'posting':
            raise ValueError(f"'{name}' already exists and is a {existing['account_type']} account.")
        if not is_aje_journal(existing['id']):
            raise ValueError(f"'{name}' is already an account in this file, and it is not on the "
                             f"AJE report. Pick another name.")
        raise ValueError(f"'{name}' already exists on the AJE report — open it, or pick another name.")
    rpt = aje_report()
    acct_id = add_account(name, 'C', desc, 'posting')
    add_report_item(rpt['id'], 'account', desc, acct_id, indent=1)
    ensure_aje_layout()
    return get_account(acct_id)

def next_aje_ref(journal_id):
    """The next reference in this journal's own sequence — 26AJE01, 26AJE02, …

    Sequence is per JOURNAL, so each year numbers from 01 and a reference names
    the year it belongs to without anyone having to remember which."""
    acct = get_account(journal_id)
    if not acct:
        raise ValueError("Journal account not found")
    name = (acct['name'] or '').strip().upper()
    # A journal named 26AJE numbers its own entries 26AJE01, 26AJE02, … A journal
    # the operator named something else still gets a house-format reference, taken
    # from the fiscal year.
    if _re.match(r'^\d{2}AJE$', name):
        stem = name
    else:
        a = fiscal_anchor()
        stem = f"{str(a['fy'])[-2:]}AJE" if a else '00AJE'
    top = 0
    with get_db() as db:
        rows = db.execute(
            "SELECT DISTINCT t.reference FROM transactions t JOIN lines l ON l.transaction_id=t.id "
            "WHERE l.account_id=?", (journal_id,)).fetchall()
    for r in rows:
        mm = AJE_REF_RE.match((r['reference'] or '').strip())
        if mm and mm.group(1) == stem[:2]:
            top = max(top, int(mm.group(2)))
    return f"{stem}{top + 1:02d}"

def validate_aje(journal_id, ref, entry_date, description, rows):
    """Check one adjusting entry WITHOUT posting it.

    rows: [{'account': name, 'amount': cents}, ...] — blank rows ignored.
    Returns per-row errors plus the totals the screen renders. Errors block;
    the entry must balance, because an adjustment that does not balance is not
    an adjustment.
    """
    out = {'ok': False, 'rows': [], 'errors': [], 'debit_cents': 0, 'credit_cents': 0,
           'net_cents': 0, 'entry_count': 0, 'ref': (ref or '').strip().upper(),
           'description': (description or '').strip()}

    journal = get_account(journal_id)
    if not journal:
        out['errors'].append("Journal account not found — open the year's entries again.")
        return out
    if not is_aje_journal(journal_id):
        out['errors'].append(f"'{journal['name']}' is not a year journal on the AJE report.")
        return out

    if not out['ref']:
        out['errors'].append('Enter a reference (e.g. %s).' % next_aje_ref(journal_id))
    elif not AJE_REF_RE.match(out['ref']):
        out['errors'].append(f"'{out['ref']}' is not an AJE reference. Use the house format "
                             f"xxAJEyy — two-digit year, then the sequence, e.g. "
                             f"{next_aje_ref(journal_id)}.")

    if not entry_date:
        out['errors'].append('Enter the date the adjustment is posted at.')
    else:
        try:
            date.fromisoformat(entry_date)
        except ValueError:
            out['errors'].append(f"'{entry_date}' is not a date (use yyyy-mm-dd).")

    if not out['description']:
        out['errors'].append('Enter a description — it prints on every line of the entry.')
    elif len(out['description']) > AJE_DESC_MAX:
        out['errors'].append(f"The description is {len(out['description'])} characters. "
                             f"Keep it to {AJE_DESC_MAX} — it prints on every line of the entry.")

    for i, raw in enumerate(rows):
        name = (raw.get('account') or '').strip()
        amt = int(raw.get('amount') or 0)
        if not name and not amt:
            continue
        r = {'index': i, 'account': name, 'amount': amt, 'error': '', 'account_id': None}
        if not name:
            r['error'] = 'Amount with no account — which account is this leg?'
        elif not amt:
            r['error'] = 'Account with no amount — enter the debit or credit, or clear the line.'
        else:
            acct = get_account_by_name(name)
            if not acct:
                r['error'] = (f"'{name}' is not in the chart of accounts. Add it first, "
                              f"then come back — the entry keeps what you typed.")
            elif acct['account_type'] == 'total':
                r['error'] = (f"'{name}' is a total account — it accumulates other accounts and "
                              f"cannot be posted to. Post to the account underneath it.")
            elif acct['id'] == journal_id:
                r['error'] = (f"'{name}' is the journal itself — Grid posts that side for you. "
                              f"Enter only the accounts being adjusted.")
            elif is_aje_journal(acct['id']):
                r['error'] = f"'{name}' is another year's AJE journal, not an account to adjust."
            else:
                r['account_id'] = acct['id']
        out['rows'].append(r)

    live = [r for r in out['rows'] if not r['error']]
    out['entry_count'] = len(live)
    out['debit_cents'] = sum(r['amount'] for r in live if r['amount'] > 0)
    out['credit_cents'] = -sum(r['amount'] for r in live if r['amount'] < 0)
    out['net_cents'] = sum(r['amount'] for r in live)

    if any(r['error'] for r in out['rows']):
        out['errors'].append(f"{sum(1 for r in out['rows'] if r['error'])} line(s) need fixing "
                             f"before this entry can be posted.")
    elif not out['rows']:
        out['errors'].append('Nothing to post yet — enter the accounts being adjusted.')
    elif len(live) < 2:
        out['errors'].append('An adjusting entry needs at least two lines.')
    elif out['net_cents'] != 0:
        out['errors'].append(
            f"The entry is out of balance by {fmt_amount_plain(out['net_cents'])} — "
            f"debits {fmt_amount_plain(out['debit_cents'])}, "
            f"credits {fmt_amount_plain(out['credit_cents'])}.")
    out['ok'] = not out['errors']
    return out

def aje_error_text(v):
    """A failed AJE flattened into ONE actionable string, per-row reasons and all
    — the interfaces with no grid to paint into see only this."""
    parts = list(v['errors'])
    for r in v['rows']:
        if r['error']:
            parts.append(f"Line {r['index'] + 1} ({r['account'] or 'no account'}): {r['error']}")
    return ' '.join(parts)

def _aje_txn_ids(db, journal_id, ref):
    return [r['id'] for r in db.execute(
        "SELECT DISTINCT t.id FROM transactions t JOIN lines l ON l.transaction_id=t.id "
        "WHERE t.reference=? COLLATE NOCASE AND l.account_id=? ORDER BY t.id",
        ((ref or '').strip(), journal_id)).fetchall()]

def _delete_aje_db(db, journal_id, ref):
    """Remove every leg of one adjustment on an ALREADY-OPEN connection."""
    ids = _aje_txn_ids(db, journal_id, ref)
    for tid in ids:
        rec = db.execute("SELECT COUNT(*) FROM lines WHERE transaction_id=? AND reconciled=1",
                         (tid,)).fetchone()[0]
        if rec:
            raise ValueError(f"{ref} has a reconciled line — unreconcile it before changing the entry.")
        d = db.execute("SELECT date FROM transactions WHERE id=?", (tid,)).fetchone()['date']
        lock = get_meta('lock_date', '')
        if lock and d <= lock:
            raise ValueError(f"{ref} is dated {d}, on or before the lock date ({lock}).")
        db.execute("DELETE FROM lines WHERE transaction_id=?", (tid,))
        db.execute("DELETE FROM transactions WHERE id=?", (tid,))
    return len(ids)

def post_aje(journal_id, ref, entry_date, description, rows, replace_ref=None):
    """Post ONE adjusting entry as the house shape, atomically.

    Each leg becomes its own 2-line transaction — the account being adjusted, and
    the year journal on the other side — all sharing the reference, the date and
    the description. The journal nets to zero; the ledger and every report read
    it as ordinary postings, because that is what they are.

    replace_ref re-posts over an existing entry in ONE transaction (old legs
    deleted and new ones written together), so an edit never leaves the books
    holding half an adjustment.
    """
    v = validate_aje(journal_id, ref, entry_date, description, rows)
    if not v['ok']:
        raise ValueError(aje_error_text(v))

    ref = v['ref']
    desc = v['description']
    if not replace_ref:
        with get_db() as db:
            if _aje_txn_ids(db, journal_id, ref):
                raise ValueError(f"{ref} already exists in this journal. Edit that entry, or "
                                 f"use the next reference ({next_aje_ref(journal_id)}).")

    txns = [(entry_date, ref, desc,
             [(r['account_id'], r['amount'], desc), (journal_id, -r['amount'], desc)])
            for r in v['rows'] if not r['error']]

    with get_db() as db:
        replaced = _delete_aje_db(db, journal_id, replace_ref) if replace_ref else 0
        batch, ids = _post_batch_db(db, txns, prefix=AJE_PREFIX)
    v['batch'], v['txn_ids'], v['replaced'] = batch, ids, replaced
    return v

def aje_groups(journal_id):
    """Every adjustment in one journal, grouped the way an accountant reads them:
    one block per reference, each block carrying its legs."""
    groups, order = {}, []
    # A leg's cross-account by NAME → its row, so each leg can carry the full
    # account name and the client's own number, not just Grid's code.
    by_name = {(a['name'] or '').upper(): a for a in get_accounts()}
    for e in get_ledger(journal_id):
        ref = (e['reference'] or '').strip()
        key = ref or f"__{e['txn_id']}"
        if key not in groups:
            groups[key] = {'ref': ref, 'date': e['date'], 'description': e['description'],
                           'lines': [], 'debit_cents': 0, 'credit_cents': 0}
            order.append(key)
        g = groups[key]
        # The journal's own leg is C-normal, so get_ledger has already flipped it:
        # e['amount'] IS the debit/credit of the account being adjusted.
        xa = by_name.get((e['cross_accounts'] or '').strip().upper())
        g['lines'].append({'txn_id': e['txn_id'], 'line_id': e['line_id'],
                           'account': e['cross_accounts'], 'amount': e['amount'],
                           'reconciled': e['reconciled'],
                           # What a CLIENT can read comes first; Grid's own code
                           # is reference, and sits out on the right.
                           'account_id': xa['id'] if xa else None,
                           'account_desc': (xa['description'] or '') if xa else '',
                           'account_number': (xa['account_number'] or '') if xa else ''})
        if e['amount'] > 0: g['debit_cents'] += e['amount']
        else:               g['credit_cents'] -= e['amount']
        if e['date'] < g['date']:
            g['date'] = e['date']
    out = []
    for k in order:
        g = groups[k]
        g['balanced'] = g['debit_cents'] == g['credit_cents']
        out.append(g)
    return out

def aje_rows_from_ref(journal_id, ref):
    """Read one posted adjustment back into entry rows, so it can be edited
    instead of re-keyed."""
    for g in aje_groups(journal_id):
        if g['ref'].upper() == (ref or '').strip().upper():
            return {'ref': g['ref'], 'date': g['date'], 'description': g['description'],
                    'rows': [{'account': l['account'], 'amount': l['amount']} for l in g['lines']]}
    return None

def delete_aje(journal_id, ref):
    """Remove one adjustment entirely — all of its legs, or none of them."""
    with get_db() as db:
        n = _delete_aje_db(db, journal_id, ref)
    if not n:
        raise ValueError(f"No entry {ref} in this journal.")
    return n



def ensure_trx_layout():
    """Force the TRX head into shape, on EVERY open, for every file however old.

    Idempotent, and it never touches the body: whatever the operator has added
    below line 3 keeps its order and simply sits after it. Runs from init_db, so
    web, CLI and MCP all converge on the identical layout."""
    with get_db() as db:
        trx = db.execute("SELECT * FROM reports WHERE name='TRX'").fetchone()
        if not trx:
            db.execute("INSERT INTO reports(name, description, sort_order) VALUES('TRX',?,40)",
                       (TRX_REPORT_DESC,))
            trx = db.execute("SELECT * FROM reports WHERE name='TRX'").fetchone()
        elif (trx['description'] or '') != TRX_REPORT_DESC:
            # One name for one thing. Older files carry titles like "Conversion &
            # workings" (which framed the whole report as conversion-only) or
            # "Transactions Journal" (which then contradicted the heading on line
            # 1). TRX is a system report with a fixed head, so Grid owns what it
            # is called — the operator owns everything below line 3.
            db.execute("UPDATE reports SET description=? WHERE id=?", (TRX_REPORT_DESC, trx['id']))

        # The two system accounts. TRX.OPEN used to appear only once a conversion
        # had been posted; it is part of the furniture now, so it always exists.
        for name, desc, bal in ((CONVERSION_ACCT, CONVERSION_DESC, 'C'),
                                (OPENING_RE_ACCT, OPENING_RE_DESC, 'C')):
            row = db.execute("SELECT * FROM accounts WHERE name=? COLLATE NOCASE", (name,)).fetchone()
            if not row:
                db.execute("INSERT INTO accounts(name, description, normal_balance, account_type, system) "
                           "VALUES(?,?,?,'posting',1)", (name, desc, bal))
            else:
                db.execute("UPDATE accounts SET description=?, system=1 WHERE id=?", (desc, row['id']))
        conv = db.execute("SELECT * FROM accounts WHERE name=? COLLATE NOCASE", (CONVERSION_ACCT,)).fetchone()
        reob = db.execute("SELECT * FROM accounts WHERE name=? COLLATE NOCASE", (OPENING_RE_ACCT,)).fetchone()

        items = db.execute("SELECT * FROM report_items WHERE report_id=? ORDER BY position, id",
                           (trx['id'],)).fetchall()

        def pinned_row(kind):
            """The existing line for a pinned slot, if any — extras are pruned."""
            if kind == 'heading':
                hits = [i for i in items if i['item_type'] == 'label'
                        and (i['description'] or '').strip().lower() == TRX_HEADING.lower()]
            else:
                hits = [i for i in items if i['account_id'] == kind['id']]
            for extra in hits[1:]:
                db.execute("DELETE FROM report_items WHERE id=?", (extra['id'],))
            return hits[0] if hits else None

        # Track the head's ids as they are written. Deriving "pinned" from the
        # POSITION afterwards is wrong: an operator's own line may already be
        # sitting at 10 or 30, and it would be mistaken for part of the head and
        # left there — stranded on top of the line that belongs in that slot.
        pinned_ids = set()

        head = pinned_row('heading')
        if head:
            db.execute("UPDATE report_items SET position=?, description=?, indent=0 WHERE id=?",
                       (TRX_PINNED_POS['heading'], TRX_HEADING, head['id']))
            pinned_ids.add(head['id'])
        else:
            cur = db.execute("INSERT INTO report_items(report_id, position, item_type, description, indent) "
                             "VALUES(?,?,'label',?,0)", (trx['id'], TRX_PINNED_POS['heading'], TRX_HEADING))
            pinned_ids.add(cur.lastrowid)

        for acct, desc in ((conv, CONVERSION_DESC), (reob, OPENING_RE_DESC)):
            row = pinned_row(acct)
            pos = TRX_PINNED_POS[acct['name'].upper()]
            # RE.OB feeds retained earnings; TRX.OPEN totals to nothing — it nets
            # to zero when a conversion is complete, which is the whole point.
            tt = 'RE' if acct['name'].upper() == OPENING_RE_ACCT else ''
            if row:
                db.execute("UPDATE report_items SET position=?, description=?, indent=1, total_to_1=? "
                           "WHERE id=?", (pos, desc, tt, row['id']))
                pinned_ids.add(row['id'])
            else:
                cur = db.execute("INSERT INTO report_items(report_id, position, item_type, account_id, "
                                 "indent, description, total_to_1) VALUES(?,?,'account',?,1,?,?)",
                                 (trx['id'], pos, acct['id'], desc, tt))
                pinned_ids.add(cur.lastrowid)

        # Everything else keeps its relative order and sits below the head.
        pos = TRX_BODY_START
        for it in db.execute("SELECT id FROM report_items WHERE report_id=? ORDER BY position, id",
                             (trx['id'],)).fetchall():
            if it['id'] in pinned_ids:
                continue
            db.execute("UPDATE report_items SET position=? WHERE id=?", (pos, it['id']))
            pos += 10

    # A one-time nudge about what the body is FOR. Seeded once and remembered, so
    # deleting it sticks (same rule as the engagement folders).
    if get_meta('trx_body_seeded', '') != '1':
        with get_db() as db:
            trx = db.execute("SELECT * FROM reports WHERE name='TRX'").fetchone()
            n = db.execute("SELECT COUNT(*) FROM report_items WHERE report_id=? AND position>=?",
                           (trx['id'], TRX_BODY_START)).fetchone()[0]
            if not n:
                db.execute("INSERT INTO report_items(report_id, position, item_type, description, indent) "
                           "VALUES(?,?,'separator','',0)", (trx['id'], TRX_BODY_START))
                db.execute("INSERT INTO report_items(report_id, position, item_type, description, indent) "
                           "VALUES(?,?,'label',?,0)",
                           (trx['id'], TRX_BODY_START + 10, TRX_BODY_LABEL))
        set_meta('trx_body_seeded', '1')


def trx_pinned_item_ids():
    """The TRX head — the lines that may not be moved, renamed or deleted.
    Identified by WHAT they are, never by where they sit."""
    with get_db() as db:
        trx = db.execute("SELECT id FROM reports WHERE name='TRX'").fetchone()
        if not trx:
            return set()
        out = {r['id'] for r in db.execute(
            "SELECT ri.id FROM report_items ri JOIN accounts a ON a.id = ri.account_id "
            "WHERE ri.report_id=? AND a.name IN (?,?) COLLATE NOCASE",
            (trx['id'], CONVERSION_ACCT, OPENING_RE_ACCT)).fetchall()}
        head = db.execute("SELECT id FROM report_items WHERE report_id=? AND item_type='label' "
                          "AND description=? ORDER BY position LIMIT 1",
                          (trx['id'], TRX_HEADING)).fetchone()
        if head:
            out.add(head['id'])
        return out


PINNED_MSG = ('That line is part of the fixed head of the Transactions Ledger — the heading, '
              'the opening conversion account and system opening retained earnings always sit '
              'at the top, in that order. Everything below it is yours to arrange.')


def check_trx_pinned(item_id):
    """Raise if this report item is part of the TRX head."""
    if item_id in trx_pinned_item_ids():
        raise ValueError(PINNED_MSG)


def validate_opening_rows(conversion_date, rows, expected_re_cents=None):
    """Check a conversion grid WITHOUT posting anything.

    rows: [{'account': name, 'description': text, 'amount': cents}, ...] — blank
    rows are ignored, so a 25-line grid with 6 entries is normal.

    Returns a dict the entry screen renders directly: per-row errors and
    warnings, the running totals, and the computed opening retained earnings.
    Errors block the post; warnings never do — they are the things that are
    usually wrong but legitimately might not be.
    """
    from datetime import date as _date
    out = {'ok': False, 'rows': [], 'errors': [], 'warnings': [],
           'debit_cents': 0, 'credit_cents': 0, 'net_cents': 0,
           're_line_cents': 0, 're_credit_cents': 0, 'entry_count': 0}

    if not conversion_date:
        out['errors'].append('Enter the conversion date — the balances are as at that date.')
    else:
        try:
            _date.fromisoformat(conversion_date)
        except ValueError:
            out['errors'].append(f"'{conversion_date}' is not a date (use yyyy-mm-dd).")
    lock = get_meta('lock_date', '')
    if conversion_date and lock and conversion_date <= lock:
        out['errors'].append(
            f"The conversion date ({conversion_date}) is on or before the lock date ({lock}). "
            f"Move the lock date back before posting opening balances.")

    # Which accounts report on the income statement — one scan, memoised for
    # this call, used only for the "that belongs inside opening RE" warning.
    _stmt_cache = {}
    def _stmt_of():
        if not _stmt_cache:
            _stmt_cache.update(statement_type_map() or {'': ''})
        return _stmt_cache

    seen = {}
    for i, raw in enumerate(rows):
        name = (raw.get('account') or '').strip()
        desc = (raw.get('description') or '').strip()
        amt  = int(raw.get('amount') or 0)
        if not name and not amt:
            continue                      # an untouched line on the grid
        r = {'index': i, 'account': name, 'description': desc, 'amount': amt,
             'error': '', 'warning': '', 'account_id': None}
        if not name:
            r['error'] = 'Amount with no account — which account is this balance?'
        elif not amt:
            r['error'] = 'Account with no amount — enter the balance, or clear the line.'
        else:
            acct = get_account_by_name(name)
            blocked = OPENING_BLOCKED.get(name.upper())
            if blocked:
                r['error'] = blocked
            elif not acct:
                r['error'] = (f"'{name}' is not in the chart of accounts. Add it first "
                              f"(+ Account), then come back — the grid keeps what you typed.")
            elif acct['account_type'] == 'total':
                r['error'] = (f"'{name}' is a total account — it accumulates other accounts "
                              f"and cannot be posted to. Enter the accounts underneath it.")
            elif (acct['computed'] if 'computed' in acct.keys() else ''):
                r['error'] = f"'{name}' is a computed display line, not a posting account."
            else:
                r['account_id'] = acct['id']
                r['account'] = acct['name']          # canonical case
                if not desc:
                    r['description'] = acct['description'] or OPENING_DESC
                if name.upper() in seen:
                    r['warning'] = (f"{acct['name']} is on line {seen[name.upper()] + 1} as well — "
                                    f"two balances for one account will be added together.")
                seen[name.upper()] = i
                # The sign nudge. Not an error: a bank account CAN be overdrawn
                # and a shareholder loan CAN swing the other way.
                if not r['warning']:
                    if acct['normal_balance'] == 'C' and amt > 0:
                        r['warning'] = (f"{acct['name']} normally carries a credit balance — "
                                        f"you entered a debit. Correct?")
                    elif acct['normal_balance'] == 'D' and amt < 0:
                        r['warning'] = (f"{acct['name']} normally carries a debit balance — "
                                        f"you entered a credit. Correct?")
        out['rows'].append(r)

    live = [r for r in out['rows'] if not r['error']]
    out['entry_count'] = len(live)
    out['debit_cents']  = sum(r['amount'] for r in live if r['amount'] > 0)
    out['credit_cents'] = -sum(r['amount'] for r in live if r['amount'] < 0)
    out['net_cents']    = sum(r['amount'] for r in live)
    # Retained earnings is the residual — the figure that makes the entry balance.
    # It is NEVER typed. Positive re_credit_cents = a credit balance = the normal
    # case; negative = an accumulated deficit.
    out['re_line_cents']   = -out['net_cents']
    out['re_credit_cents'] =  out['net_cents']

    # What the residual means depends on WHAT was entered, and the difference is
    # worth stating plainly:
    #   balance sheet only        → retained earnings brought forward
    #   full trial balance (+ IS) → retained earnings at the START of that year,
    #                               because the year's income is in the grid and
    #                               flows to closing RE through the IS chain
    stmt = _stmt_of()
    out['is_count'] = sum(1 for r in live if stmt.get(r['account']) == 'IS')
    if out['is_count']:
        out['re_note'] = (
            f"{out['is_count']} income-statement account(s) are included, so this is the whole "
            f"trial balance for that year — the figure above is retained earnings at the START "
            f"of it. The year's income flows through to closing retained earnings on its own. "
            f"(On a T2 that is line 3660, not 3600.)")
    else:
        out['re_note'] = ("Balance-sheet accounts only, so this is retained earnings brought "
                          "forward — the closing figure from the prior-year statements.")

    if any(r['error'] for r in out['rows']):
        out['errors'].append(f"{sum(1 for r in out['rows'] if r['error'])} line(s) need fixing "
                             f"before this can be posted.")
    if not live and not out['rows']:
        out['errors'].append('Nothing to post yet — enter the balances as at the conversion date.')
    # A check figure measured against the SURVIVING rows reads as "your balances
    # are wrong" when the real problem is a rejected line. Say the one true thing.
    if expected_re_cents is not None and live and not any(r['error'] for r in out['rows']):
        diff = out['re_credit_cents'] - int(expected_re_cents)
        if diff:
            out['errors'].append(
                f"Opening retained earnings comes to {fmt_amount_plain(out['re_credit_cents'])} but the "
                f"prior-year statements say {fmt_amount_plain(int(expected_re_cents))} — "
                f"a difference of {fmt_amount_plain(diff)}. Something in the balances above is wrong, "
                f"missing, or on the wrong side.")
    out['ok'] = not out['errors']
    return out


def opening_error_text(v):
    """Flatten a failed validation into ONE actionable string.

    The entry screen paints each row's reason beside the row, so a human sees
    exactly which line is wrong and why. CLI and MCP callers have no grid — they
    only ever see the exception text, and a bare "3 line(s) need fixing" leaves a
    headless agent with nothing to act on. So the per-row reasons travel with it.
    """
    parts = list(v['errors'])
    for r in v['rows']:
        if r['error']:
            parts.append(f"Line {r['index'] + 1} ({r['account'] or 'no account'}): {r['error']}")
    return ' '.join(parts)


def post_opening_balances(conversion_date, rows, expected_re_cents=None, replace=False):
    """Post a whole conversion as ONE atomic batch — the only supported way to
    put opening balances into a set of books.

    What lands: one 2-line entry per account against TRX.OPEN (same date, same
    reference OPEN, so the TRX ledger reads as a clean conversion list), then the
    computed residual as Cr RE.OB / Dr TRX.OPEN — which leaves TRX.OPEN at zero
    and feeds opening retained earnings up to the balance sheet the way the
    engine expects. Any row failing rolls the whole thing back.

    replace=True re-posts over an existing conversion in ONE transaction (the old
    one is deleted and the new one written together) — that is what the Edit path
    on the opening-balances screen uses.

    Returns the validation dict with 'batch', 'txn_ids' and 'replaced' added.
    """
    existing = opening_batch()
    if existing and not replace:
        raise ValueError(
            "Opening balances have already been posted for these books. Edit them, or "
            "delete the existing conversion and re-enter it — Grid does not layer a "
            "second set of openings on top of the first.")

    v = validate_opening_rows(conversion_date, rows, expected_re_cents)
    if not v['ok']:
        raise ValueError(opening_error_text(v))

    conv = ensure_conversion_account()
    re_ob = get_account_by_name(OPENING_RE_ACCT)
    if not re_ob:
        raise ValueError(
            f"{OPENING_RE_ACCT} (opening retained earnings) is missing from these books. "
            f"Close and reopen them — the chart repairs itself on open — then try again.")

    txns = []
    for r in v['rows']:
        if r['error']:
            continue
        txns.append((conversion_date, OPENING_REF, r['description'],
                     [(r['account_id'], r['amount'], r['description']),
                      (conv['id'], -r['amount'], r['description'])]))
    if v['re_line_cents']:
        desc = 'Opening retained earnings (b/f)'
        txns.append((conversion_date, OPENING_REF, desc,
                     [(re_ob['id'], v['re_line_cents'], desc),
                      (conv['id'], -v['re_line_cents'], desc)]))

    # Replacing is ONE transaction: the old conversion goes and the new one lands
    # together, or neither does. There is never a moment where the client has no
    # opening position because a repost failed halfway.
    with get_db() as db:
        replaced = _delete_batch_db(db, existing) if existing else 0
        batch, ids = _post_batch_db(db, txns, prefix=OPENING_PREFIX, avoid=existing)
    set_meta('openings_declined', '')       # they clearly weren't starting at zero
    v['batch'], v['txn_ids'], v['replaced'] = batch, ids, replaced
    return v


def opening_rows_from_batch():
    """Read a posted conversion back into grid rows, so the operator can EDIT it
    instead of re-keying 40 lines. The TRX.OPEN contra legs are dropped (Grid
    owns those) and so is the RE.OB line (retained earnings is recomputed from
    whatever the edited grid says — it is never carried forward as a number)."""
    batch = opening_batch()
    if not batch:
        return []
    skip = {CONVERSION_ACCT.upper(), OPENING_RE_ACCT.upper()}
    out = []
    with get_db() as db:
        for r in db.execute(
                "SELECT a.name, l.description, l.amount FROM lines l "
                "JOIN transactions t ON t.id = l.transaction_id "
                "JOIN accounts a ON a.id = l.account_id "
                "WHERE t.import_batch=? ORDER BY t.id, l.sort_order", (batch,)):
            if r['name'].upper() in skip:
                continue
            out.append({'account': r['name'], 'description': r['description'] or '',
                        'amount': r['amount']})
    return out


def _delete_batch_db(db, batch_id):
    """Delete a batch on an open connection, with the same refusals as
    delete_import_batch. Returns the number of transactions removed."""
    n = db.execute("SELECT COUNT(*) FROM transactions WHERE import_batch=?", (batch_id,)).fetchone()[0]
    if not n:
        raise ValueError(f"No transactions found for batch '{batch_id}'")
    rec = db.execute(
        "SELECT COUNT(*) FROM lines l JOIN transactions t ON t.id = l.transaction_id "
        f"WHERE t.import_batch=? AND {REC_SQL}", (batch_id,)).fetchone()[0]
    if rec:
        raise ValueError(f"{rec} line(s) in this conversion are reconciled — unreconcile them first.")
    lock = get_meta('lock_date', '')
    if lock:
        locked = db.execute("SELECT COUNT(*) FROM transactions WHERE import_batch=? AND date<=?",
                            (batch_id, lock)).fetchone()[0]
        if locked:
            raise ValueError(f"{locked} entry(s) in this conversion are on or before the "
                             f"lock date ({lock}). Move the lock date back first.")
    db.execute("DELETE FROM lines WHERE transaction_id IN "
               "(SELECT id FROM transactions WHERE import_batch=?)", (batch_id,))
    db.execute("DELETE FROM transactions WHERE import_batch=?", (batch_id,))
    return n


def delete_opening_balances():
    """Delete the whole conversion — the redo half of delete-and-redo. Refuses
    for the same reasons any batch does (reconciled lines, locked dates)."""
    batch = opening_batch()
    if not batch:
        raise ValueError('No opening balances have been posted for these books.')
    return delete_import_batch(batch)


def decline_openings():
    """Brand-new client: no history, everything starts at zero. Remembered so
    the prompt stops asking."""
    set_meta('openings_declined', '1')


def is_opening_txn(txn_id, db=None):
    """Is this transaction part of the conversion? Drives the soft lock."""
    def _q(d):
        r = d.execute("SELECT import_batch FROM transactions WHERE id=?", (txn_id,)).fetchone()
        return bool(r and (r['import_batch'] or '').startswith(OPENING_PREFIX + '-'))
    if db is not None:
        return _q(db)
    with get_db() as d:
        return _q(d)


# The soft lock's message. Recognisable by its prefix so the browser can turn it
# into "are you sure?" rather than a dead end — the operator is allowed to edit
# opening balances, they just may not do it by accident.
OPENING_EDIT_WARNING = (
    'OPENING: This is an opening balance from the conversion. Changing it moves the '
    "client's opening position, and retained earnings moves with it.")


def get_transaction(txn_id):
    with get_db() as db:
        txn = db.execute("SELECT * FROM transactions WHERE id=?", (txn_id,)).fetchone()
        if not txn: return None, []
        lines = db.execute(
            "SELECT l.*, a.name as account_name, a.normal_balance, a.description as acct_desc "
            "FROM lines l JOIN accounts a ON l.account_id=a.id WHERE l.transaction_id=? ORDER BY l.sort_order",
            (txn_id,)).fetchall()
        return txn, lines

def update_transaction(txn_id, date_str, reference, description, lines, allow_opening=False):
    """Update a transaction. lines = [(acct_id, amount, desc), ...] or
    [(acct_id, amount, desc, reconciled, doc_on_file), ...] to preserve flags.

    allow_opening: opening balances are editable, but never by accident — the
    caller must have asked the operator first (see OPENING_EDIT_WARNING)."""
    total = sum(l[1] for l in lines)
    if total != 0:
        raise ValueError(f"Transaction does not balance: off by {total/100:.2f}")
    lock = get_meta('lock_date', '')
    if lock and date_str <= lock:
        raise ValueError(f"Date {date_str} is on or before the lock date ({lock}).")
    with get_db() as db:
        if not allow_opening and is_opening_txn(txn_id, db):
            raise ValueError(OPENING_EDIT_WARNING)
        # Reconciled lines are settled against an external statement — changing
        # them silently breaks the reconciliation (LAP rule: unreconcile first).
        rec = reconciled_count(txn_id, db)
        if rec:
            raise ValueError(f"Transaction has {rec} reconciled line(s) — "
                             "unreconcile them first (Reconcile screen or the ✓ flag on the ledger).")
        # Block posting to total accounts
        for line in lines:
            acct_id = line[0]
            acct = db.execute("SELECT name, account_type FROM accounts WHERE id=?",
                              (acct_id,)).fetchone()
            if acct and acct['account_type'] == 'total':
                raise ValueError(
                    f"Cannot post to '{acct['name']}' — it is a total account. "
                    "Post to a detail account instead.")
        db.execute("UPDATE transactions SET date=?, reference=?, description=? WHERE id=?",
            (date_str, reference, description, txn_id))
        db.execute("DELETE FROM lines WHERE transaction_id=?", (txn_id,))
        for i, line in enumerate(lines):
            acct_id, amount, desc = line[0], line[1], line[2]
            reconciled = line[3] if len(line) > 3 else 0
            doc_flag = line[4] if len(line) > 4 else 0
            db.execute("INSERT INTO lines(transaction_id, account_id, amount, description, reconciled, doc_on_file, sort_order) VALUES(?,?,?,?,?,?,?)",
                (txn_id, acct_id, amount, desc, reconciled, doc_flag, i))

def delete_transaction(txn_id, allow_opening=False):
    with get_db() as db:
        if not allow_opening and is_opening_txn(txn_id, db):
            raise ValueError(OPENING_EDIT_WARNING + ' To replace the whole conversion, '
                             'delete it from the opening-balances screen and re-enter it.')
        lock = get_meta('lock_date', '')
        if lock:
            txn = db.execute("SELECT date FROM transactions WHERE id=?", (txn_id,)).fetchone()
            if txn and txn['date'] <= lock:
                raise ValueError(f"Cannot delete: transaction date {txn['date']} is on or before lock date ({lock}).")
        rec = reconciled_count(txn_id, db)
        if rec:
            raise ValueError(f"Cannot delete: {rec} reconciled line(s) — unreconcile them first.")
        db.execute("DELETE FROM transactions WHERE id=?", (txn_id,))

def bulk_delete_transactions(txn_ids, allow_opening=False):
    """Delete multiple transactions at once. Skips locked, reconciled, and
    conversion entries — a sweep selection must never take the client's opening
    position with it."""
    with get_db() as db:
        lock = get_meta('lock_date', '')
        skipped = 0
        deleted = 0
        for tid in txn_ids:
            if not allow_opening and is_opening_txn(tid, db):
                skipped += 1
                continue
            if lock:
                txn = db.execute("SELECT date FROM transactions WHERE id=?", (tid,)).fetchone()
                if txn and txn['date'] <= lock:
                    skipped += 1
                    continue
            if reconciled_count(tid, db):
                skipped += 1
                continue
            db.execute("DELETE FROM transactions WHERE id=?", (tid,))
            deleted += 1
        return deleted, skipped

def _rec_tag(tag):
    """Normalize a reconcile tag: blank/0 → plain 1, else the tag itself
    (LAP [Rec] semantics — the value groups items by statement/application)."""
    tag = str(tag).strip() if tag is not None else ''
    return tag if tag not in ('', '0', '1') else 1

def toggle_reconcile(line_id, tag=None):
    """Toggle a line open/closed. Closing stamps `tag` (statement date, payment
    ref) into the reconciled field; opening sets it back to 0."""
    with get_db() as db:
        row = db.execute("SELECT reconciled FROM lines WHERE id=?", (line_id,)).fetchone()
        new_val = 0 if row['reconciled'] else _rec_tag(tag)
        db.execute("UPDATE lines SET reconciled=? WHERE id=?", (new_val, line_id))
        return new_val

def reconciled_count(txn_id, db=None):
    """How many lines of this transaction are reconciled (tag or 1)."""
    sql = f"SELECT COUNT(*) FROM lines l WHERE l.transaction_id=? AND {REC_SQL}"
    if db is not None:
        return db.execute(sql, (txn_id,)).fetchone()[0]
    with get_db() as db2:
        return db2.execute(sql, (txn_id,)).fetchone()[0]

def toggle_doc_on_file(line_id):
    """Toggle doc_on_file for ALL lines in the same transaction."""
    with get_db() as db:
        row = db.execute("SELECT doc_on_file, transaction_id FROM lines WHERE id=?", (line_id,)).fetchone()
        new_val = 0 if row['doc_on_file'] else 1
        db.execute("UPDATE lines SET doc_on_file=? WHERE transaction_id=?", (new_val, row['transaction_id']))
        return new_val

def batch_reconcile(line_ids, value=1):
    """Set reconciled flag/tag on multiple lines at once (0 = reopen)."""
    value = 0 if str(value).strip() in ('', '0') else _rec_tag(value)
    with get_db() as db:
        for lid in line_ids:
            db.execute("UPDATE lines SET reconciled=? WHERE id=?", (value, lid))

def get_reconcile_summary(account_id):
    """Get reconciliation totals for an account."""
    with get_db() as db:
        acct = get_account(account_id)
        sign = 1 if acct['normal_balance'] == 'D' else -1
        # Total of all lines (= book balance)
        book = db.execute(
            "SELECT COALESCE(SUM(l.amount),0) FROM lines l JOIN transactions t ON l.transaction_id=t.id WHERE l.account_id=?",
            (account_id,)).fetchone()[0] * sign
        # Total of reconciled lines
        cleared = db.execute(
            f"SELECT COALESCE(SUM(l.amount),0) FROM lines l JOIN transactions t ON l.transaction_id=t.id WHERE l.account_id=? AND {REC_SQL}",
            (account_id,)).fetchone()[0] * sign
        return {'book_balance': book, 'cleared_balance': cleared, 'uncleared': book - cleared}

def account_history(account_id):
    """LAP-style account history: per-month debits, credits, net and cumulative
    balance (display sign), gap months filled, oldest first. Also flags fiscal
    year-end months so the view can rule them off."""
    acct = get_account(account_id)
    sign = 1 if acct['normal_balance'] == 'D' else -1
    with get_db() as db:
        rows = db.execute("""
            SELECT substr(t.date,1,7) AS ym,
                   COALESCE(SUM(CASE WHEN l.amount > 0 THEN l.amount END),0) AS dr,
                   COALESCE(SUM(CASE WHEN l.amount < 0 THEN -l.amount END),0) AS cr
            FROM lines l JOIN transactions t ON l.transaction_id = t.id
            WHERE l.account_id = ?
            GROUP BY ym ORDER BY ym""", (account_id,)).fetchall()
    if not rows:
        return []
    by_ym = {r['ym']: (r['dr'], r['cr']) for r in rows}
    fye_mm = (get_meta('fiscal_year_end', '') or '-').split('-')[0]

    def next_ym(ym):
        y, m = int(ym[:4]), int(ym[5:7])
        return f"{y + (m == 12):04d}-{(m % 12) + 1:02d}"

    out, cum = [], 0
    ym, last = rows[0]['ym'], rows[-1]['ym']
    while ym <= last:
        dr, cr = by_ym.get(ym, (0, 0))
        net = (dr - cr) * sign
        cum += net
        out.append({'ym': ym, 'dr': dr, 'cr': cr, 'net': net, 'cum': cum,
                    'fye': ym[5:7] == fye_mm})
        ym = next_ym(ym)
    return out

# ─── Working papers (the weave: index ↔ client-folder documents) ────
# Grid NEVER stores documents. The client folder is the stack of paper; the
# workpapers table is the map: index ref → RELATIVE path inside the folder
# (relative is the fix for CaseWare's absolute-path fragility — archive or
# move the whole folder, links survive; works the same on Linux/Windows).

WP_DOC_EXTS = {'.pdf', '.xlsx', '.xls', '.docx', '.doc', '.csv', '.txt',
               '.jpg', '.jpeg', '.png', '.eml', '.msg', '.rtf'}

def client_dir():
    return os.path.dirname(os.path.abspath(DB_PATH)) if DB_PATH else ''

def wp_base():
    """The containment root for every workpaper link — the client folder with
    symlinks resolved, so it compares equal to the realpath of anything inside
    it (a books folder reached through a symlinked home must not read as an
    escape)."""
    b = client_dir()
    return os.path.realpath(b) if b else ''

def wp_resolve(rel_path):
    """Absolute path of a workpaper link, REFUSED if it escapes the client
    folder (containment is what makes serving/opening these safe)."""
    base = wp_base()
    if not base:
        raise ValueError("No books open")
    p = os.path.realpath(os.path.join(base, str(rel_path)))
    if not (p == base or p.startswith(base + os.sep)):
        raise ValueError(f"Path escapes the client folder: {rel_path}")
    return p

def wp_relativize(abs_path):
    """Turn an absolute path — what the OS file navigator hands back — into the
    client-folder-relative link we store. Refuses anything outside the client
    folder: links are relative BY DESIGN so the engagement can be moved,
    archived or Box-synced whole (the CaseWare absolute-path fragility we are
    deliberately not repeating), and a file living elsewhere has no such link.
    Stored with forward slashes on every platform."""
    base = wp_base()
    if not base:
        raise ValueError("No books open")
    p = os.path.realpath(str(abs_path))
    if not p.startswith(base + os.sep):
        raise ValueError(
            "That file is outside the client folder, so there is no stable "
            "link to store.\n\n"
            f"Client folder:  {base}\n"
            f"You picked:     {p}\n\n"
            "Working-paper links are kept relative to the client folder so the "
            "whole engagement can be moved or archived without breaking. Copy "
            "the file into the client folder (a WPFdocs sub-folder is the usual "
            "spot) and browse to it there.")
    return os.path.relpath(p, base).replace(os.sep, '/')

# ─── The OS file navigator ──────────────────────────────────────────
# Browse pops the operating system's OWN file chooser (Explorer/Finder/GTK)
# on the machine running Grid, and gets back a real path. Always as a
# SUBPROCESS, never an in-process GUI toolkit: a dialog the operator leaves
# open must never wedge the Flask worker or fight it for a UI main loop.
# Same localhost/single-operator premise that lets Open call xdg-open.

FILE_DIALOG_TIMEOUT = 600      # 10 min, then the child is killed

# A chooser that cannot reach a display exits exactly like Cancel does
# (rc=1, nothing on stdout) — only stderr tells them apart, so we read it.
_NO_DISPLAY = ('cannot open display', 'unable to init server', 'no display',
               'failed to open display', 'could not open display',
               'cannot connect to wayland', 'no protocol specified')

def _file_dialog_cmds(start, title, directory=False):
    """Native file-chooser command lines for this platform, best first, with a
    tkinter subprocess as the portable last resort.

    directory=True picks a FOLDER instead of a file — how an operator points at
    an old client's books rather than hunting for the right file inside them."""
    import platform
    sysname = platform.system()
    if sysname == 'Windows':
        d = start.replace("'", "''")
        if directory:
            ps = ("Add-Type -AssemblyName System.Windows.Forms;"
                  "$d = New-Object System.Windows.Forms.FolderBrowserDialog;"
                  f"$d.Description = '{title}';"
                  f"$d.SelectedPath = '{d}';"
                  "if ($d.ShowDialog() -eq 'OK') { [Console]::Out.Write($d.SelectedPath) }")
        else:
            ps = ("Add-Type -AssemblyName System.Windows.Forms;"
                  "$d = New-Object System.Windows.Forms.OpenFileDialog;"
                  f"$d.Title = '{title}';"
                  f"$d.InitialDirectory = '{d}';"
                  "$d.Filter = 'Documents|*.pdf;*.xlsx;*.xls;*.docx;*.doc;*.csv;"
                  "*.txt;*.jpg;*.jpeg;*.png;*.eml;*.msg;*.rtf|All files (*.*)|*.*';"
                  "if ($d.ShowDialog() -eq 'OK') { [Console]::Out.Write($d.FileName) }")
        cmds = [['powershell', '-NoProfile', '-STA', '-Command', ps]]
    elif sysname == 'Darwin':
        loc = start.replace('\\', '\\\\').replace('"', '\\"')
        what = 'choose folder' if directory else 'choose file'
        cmds = [['osascript', '-e',
                 f'POSIX path of ({what} with prompt "{title}" '
                 f'default location POSIX file "{loc}")']]
    else:
        zen = ['--file-selection', f'--title={title}',
               '--filename=' + start.rstrip(os.sep) + os.sep]
        if directory:
            zen.append('--directory')
        kd = '--getexistingdirectory' if directory else '--getopenfilename'
        cmds = [['zenity'] + zen, ['qarma'] + zen,
                ['kdialog', '--title', title, kd, start]]
    cmds.append([sys.executable, '-c',
                 _TK_PICK_DIR if directory else _TK_PICK, title, start])
    return cmds


_TK_PICK_DIR = """
import sys
try:
    import tkinter
    from tkinter import filedialog
except Exception:
    sys.exit(2)
r = tkinter.Tk(); r.withdraw()
try: r.attributes('-topmost', True)
except Exception: pass
sys.stdout.write(filedialog.askdirectory(title=sys.argv[1],
                                         initialdir=sys.argv[2]) or '')
"""

_TK_PICK = """
import sys
try:
    import tkinter
    from tkinter import filedialog
except Exception:
    sys.exit(2)
r = tkinter.Tk(); r.withdraw()
try: r.attributes('-topmost', True)
except Exception: pass
sys.stdout.write(filedialog.askopenfilename(title=sys.argv[1],
                                            initialdir=sys.argv[2]) or '')
"""

def pick_file_dialog(start_dir=None, title='Select a document', directory=False):
    """Pop the OS file navigator and return the absolute path chosen, or ''
    if the operator cancelled. Raises if no chooser exists on this machine."""
    import subprocess
    start = start_dir or wp_base() or os.path.expanduser('~')
    if not os.path.isdir(start):
        start = os.path.expanduser('~')
    tried = []
    for cmd in _file_dialog_cmds(start, title, directory=directory):
        try:
            r = subprocess.run(cmd, capture_output=True, text=True,
                               timeout=FILE_DIALOG_TIMEOUT)
        except FileNotFoundError:
            continue                       # chooser not installed — try the next
        except subprocess.TimeoutExpired:
            raise RuntimeError("The file dialog sat open too long and was "
                               "closed. Click Browse again.")
        out = (r.stdout or '').strip()
        if out:
            return out.splitlines()[0].strip()
        err = (r.stderr or '').strip()
        if r.returncode in (0, 1) and not any(m in err.lower() for m in _NO_DISPLAY):
            return ''                      # Cancel
        tried.append(f"{os.path.basename(cmd[0])}: {err.splitlines()[-1][:90]}"
                     if err else f"{os.path.basename(cmd[0])}: rc={r.returncode}")
    detail = ('\n\nTried — ' + '; '.join(tried)) if tried else ''
    raise RuntimeError(
        "No file navigator could be opened on the machine running Grid. Type "
        "the path into the Link box instead (relative to the client folder), "
        "or install a file chooser — on Linux: sudo apt install zenity." + detail)

def wp_fy(fy=None):
    return str(fy or get_meta('fiscal_year', '') or '')

def ensure_engagement_folders():
    """Every file always has the 'Engagement File' root folder; a FRESH file
    also gets the three standard sections (Assets / Liabilities and Equity /
    Income Statement — user's standing order). User deletions of the three
    stick: they are only seeded when the root itself is created. Orphan papers
    are adopted into the root (error anticipation: no selection = root)."""
    with get_db() as db:
        root = db.execute("SELECT id FROM wp_folders WHERE parent_id IS NULL "
                          "AND name='Engagement File'").fetchone()
        if not root:
            cur = db.execute("INSERT INTO wp_folders(name, parent_id, sort) "
                             "VALUES('Engagement File', NULL, 0)")
            rid = cur.lastrowid
            for i, n in enumerate(('Assets', 'Liabilities and Equity', 'Income Statement'), 1):
                db.execute("INSERT INTO wp_folders(name, parent_id, sort) VALUES(?,?,?)",
                           (n, rid, i))
        else:
            rid = root['id']
        db.execute("UPDATE workpapers SET folder_id=? WHERE folder_id IS NULL OR folder_id=0 "
                   "OR folder_id NOT IN (SELECT id FROM wp_folders)", (rid,))
        return rid

def wp_root_id():
    with get_db() as db:
        r = db.execute("SELECT id FROM wp_folders WHERE parent_id IS NULL "
                       "AND name='Engagement File'").fetchone()
    return r['id'] if r else ensure_engagement_folders()

def list_wp_folders():
    with get_db() as db:
        return [dict(r) for r in db.execute(
            "SELECT * FROM wp_folders ORDER BY parent_id IS NOT NULL, parent_id, sort, name")]

def add_wp_folder(name, parent_id=None):
    name = str(name or '').strip()
    if not name:
        raise ValueError("Folder name required")
    parent = int(parent_id) if parent_id else wp_root_id()
    with get_db() as db:
        if not db.execute("SELECT 1 FROM wp_folders WHERE id=?", (parent,)).fetchone():
            parent = wp_root_id()
        cur = db.execute("INSERT INTO wp_folders(name, parent_id, sort) VALUES(?,?,99)",
                         (name, parent))
        return cur.lastrowid

def rename_wp_folder(folder_id, name):
    name = str(name or '').strip()
    if not name:
        raise ValueError("Folder name required")
    with get_db() as db:
        row = db.execute("SELECT parent_id FROM wp_folders WHERE id=?", (folder_id,)).fetchone()
        if not row:
            raise ValueError("Folder not found")
        if row['parent_id'] is None:
            raise ValueError("The Engagement File root cannot be renamed")
        db.execute("UPDATE wp_folders SET name=? WHERE id=?", (name, folder_id))

def delete_wp_folder(folder_id):
    """Analog rule: empty it first. Refuses the root, non-empty folders,
    and folders with subfolders."""
    with get_db() as db:
        row = db.execute("SELECT parent_id FROM wp_folders WHERE id=?", (folder_id,)).fetchone()
        if not row:
            raise ValueError("Folder not found")
        if row['parent_id'] is None:
            raise ValueError("The Engagement File root cannot be deleted")
        if db.execute("SELECT 1 FROM wp_folders WHERE parent_id=?", (folder_id,)).fetchone():
            raise ValueError("Folder has subfolders — remove them first")
        if db.execute("SELECT 1 FROM workpapers WHERE folder_id=?", (folder_id,)).fetchone():
            raise ValueError("Folder has working papers — move or remove them first")
        db.execute("DELETE FROM wp_folders WHERE id=?", (folder_id,))

def find_workpaper_by_ref(ref, fy=None):
    with get_db() as db:
        return db.execute("SELECT * FROM workpapers WHERE fy=? AND ref=? COLLATE NOCASE",
                          (wp_fy(fy), str(ref or '').strip())).fetchone()

def list_workpapers(fy=None):
    fy = wp_fy(fy)
    with get_db() as db:
        rows = db.execute("SELECT * FROM workpapers WHERE fy=? "
                          "ORDER BY ref COLLATE NOCASE", (fy,)).fetchall()
    out = []
    for r in rows:
        d = dict(r)
        d['file_exists'] = False
        if d['path']:
            try:
                d['file_exists'] = os.path.exists(wp_resolve(d['path']))
            except ValueError:
                pass
        out.append(d)
    return out

def workpaper_fys():
    with get_db() as db:
        return [r[0] for r in db.execute(
            "SELECT DISTINCT fy FROM workpapers ORDER BY fy DESC")]

def add_workpaper(ref, description='', path='', fy=None, folder_id=None):
    ref = str(ref or '').strip()
    if not ref:
        raise ValueError("Reference required (e.g. E-1, B-2.1)")
    path = str(path or '').strip()
    if path:
        wp_resolve(path)          # containment check up front
    fy = wp_fy(fy)
    fid = int(folder_id) if folder_id else wp_root_id()
    with get_db() as db:
        if not db.execute("SELECT 1 FROM wp_folders WHERE id=?", (fid,)).fetchone():
            fid = wp_root_id()    # no selection / dead folder → Engagement File root
        try:
            cur = db.execute(
                "INSERT INTO workpapers(fy, ref, description, path, folder_id) VALUES(?,?,?,?,?)",
                (fy, ref, str(description or '').strip(), path, fid))
            return cur.lastrowid
        except sqlite3.IntegrityError:
            raise ValueError(f"Reference {ref} already exists for FY{fy}")

def update_workpaper(wp_id, field, value):
    if field not in ('ref', 'description', 'path', 'to_print', 'prep_by', 'rev_by', 'folder_id'):
        raise ValueError(f"Field not editable: {field}")
    if field == 'path':
        value = str(value or '').strip()
        if value:
            wp_resolve(value)
    if field == 'to_print':
        value = 1 if str(value).strip().lower() in ('1', 'true', 'on') else 0
    if field == 'ref':
        value = str(value).strip()
        if not value:
            raise ValueError("Reference cannot be blank")
    with get_db() as db:
        if not db.execute("SELECT 1 FROM workpapers WHERE id=?", (wp_id,)).fetchone():
            raise ValueError("Working paper not found")
        try:
            db.execute(f"UPDATE workpapers SET {field}=? WHERE id=?", (value, wp_id))
        except sqlite3.IntegrityError:
            raise ValueError("That reference already exists for this year")

def delete_workpaper(wp_id):
    with get_db() as db:
        db.execute("DELETE FROM workpapers WHERE id=?", (wp_id,))

def get_workpaper(wp_id):
    with get_db() as db:
        return db.execute("SELECT * FROM workpapers WHERE id=?", (wp_id,)).fetchone()

def scan_client_docs(include_linked=False):
    """Document files in the client folder. Default: only those not linked by
    ANY workpaper (the unindexed pile). include_linked=True returns everything
    (the Browse picker). Skips backups/, book/lock files, hidden files."""
    base = wp_base()
    if not base or not os.path.isdir(base):
        return []
    linked = set()
    with get_db() as db:
        for r in db.execute("SELECT path FROM workpapers WHERE path != ''"):
            linked.add(os.path.normpath(r['path']).lower())
    out = []
    for root, dirs, files in os.walk(base):
        dirs[:] = [d for d in dirs if not d.startswith('.')
                   and d.lower() not in ('backups', '__pycache__')]
        for fn in files:
            if fn.startswith('.') or fn.startswith('$'):
                continue
            if os.path.splitext(fn)[1].lower() not in WP_DOC_EXTS:
                continue
            rel = os.path.relpath(os.path.join(root, fn), base).replace(os.sep, '/')
            if not include_linked and os.path.normpath(rel).lower() in linked:
                continue
            out.append(rel)
    return sorted(out)

def verify_workpapers(fy=None):
    """The old-school completeness discipline as one call: every statement mark
    has a paper, every paper has a present file, every doc file is indexed,
    every paper is prepared. (Sub-papers like E-1.1 need no statement mark —
    the check is marks ⊆ papers, not the reverse.)"""
    fy = wp_fy(fy)
    papers = list_workpapers(fy)
    refs = {p['ref'].strip().upper() for p in papers}
    with get_db() as db:
        marks = {r[0].strip().upper() for r in db.execute(
            "SELECT DISTINCT ref_mark FROM report_items WHERE ref_mark != ''")}
    return {
        'fy': fy,
        'marks_without_paper': sorted(marks - refs),
        'papers_without_file': sorted(p['ref'] for p in papers
                                      if not p['path'] or not p['file_exists']),
        'unindexed_files': scan_client_docs(),
        'unprepared': sorted(p['ref'] for p in papers if not p['prep_by']),
    }

# ─── Working-paper references & lead sheets ────────────────────────

def set_ref_mark(item_id, mark):
    """Set/clear the working-paper index mark on a report line (the red pencil)."""
    with get_db() as db:
        if not db.execute("SELECT 1 FROM report_items WHERE id=?", (item_id,)).fetchone():
            raise ValueError(f"Report line {item_id} not found")
        db.execute("UPDATE report_items SET ref_mark=? WHERE id=?",
                   (str(mark or '').strip(), item_id))

LEADSHEET_MAX = 3          # 'A', 'B-1', '20', '30' — an index mark, not a name

def set_leadsheet(account_name, code):
    """Assign an account to a lead sheet (blank clears). Deliberately dumb: any
    three characters the operator wants — a letter (A, B, V) or the old section
    numbers (20 revenue, 30 cost of sales). Upper-cased so 'a' and 'A' are the
    same sheet."""
    code = str(code or '').strip().upper()
    if len(code) > LEADSHEET_MAX:
        raise ValueError(f"Lead sheet code is at most {LEADSHEET_MAX} characters: {code!r}")
    with get_db() as db:
        cur = db.execute("UPDATE accounts SET leadsheet=? WHERE name=? COLLATE NOCASE",
                         (code, account_name))
        if cur.rowcount == 0:
            raise ValueError(f"Account not found: {account_name}")

def statement_type_map():
    """account name → 'BS' | 'IS' | '' by walking total-to chains to the
    statement reports. Decides whether a lead-sheet number is a closing balance
    (BS: perpetual) or period activity (IS: fiscal-year window)."""
    rpt_of, feeds = {}, {}
    with get_db() as db:
        for r in db.execute("""SELECT a.name an, rp.name rn,
                    ri.total_to_1 t1, ri.total_to_2 t2, ri.total_to_3 t3,
                    ri.total_to_4 t4, ri.total_to_5 t5, ri.total_to_6 t6
                FROM report_items ri
                JOIN accounts a ON a.id = ri.account_id
                JOIN reports rp ON rp.id = ri.report_id
                WHERE ri.account_id IS NOT NULL"""):
            if r['rn'] in ('BS', 'IS') or r['an'] not in rpt_of:
                rpt_of[r['an']] = r['rn']
            s = feeds.setdefault(r['an'], set())
            for k in ('t1', 't2', 't3', 't4', 't5', 't6'):
                if r[k]:
                    s.add(r[k])

    resolved = {}
    def resolve(name, seen):
        if name in resolved:
            return resolved[name]
        if rpt_of.get(name) in ('BS', 'IS'):
            return rpt_of[name]
        for t in feeds.get(name, ()):
            if t not in seen:
                r = resolve(t, seen | {t})
                if r:
                    return r
        return ''
    for n in rpt_of:
        resolved[n] = resolve(n, {n})
    return resolved

def leadsheet_index():
    """All lead-sheet codes with account counts (for the index page)."""
    with get_db() as db:
        return db.execute(
            "SELECT leadsheet AS code, COUNT(*) AS n FROM accounts "
            "WHERE leadsheet != '' GROUP BY leadsheet ORDER BY leadsheet").fetchall()

def leadsheet_data(code):
    """One lead sheet: its accounts with CY / PY / $chg / %chg (display sign),
    BS accounts as closing balances, IS accounts as fiscal-year activity."""
    anchor = fiscal_anchor()
    stypes = statement_type_map()
    with get_db() as db:
        accts = db.execute(
            "SELECT * FROM accounts WHERE leadsheet=? COLLATE NOCASE "
            "AND account_type != 'total' ORDER BY name", (code,)).fetchall()
    rows = []
    for a in accts:
        sign = 1 if a['normal_balance'] == 'D' else -1
        st = stypes.get(a['name'], '') or 'BS'
        if anchor and st == 'IS':
            cy = get_account_balance(a['id'], anchor['cy_start'], anchor['cy_end'])
            py = get_account_balance(a['id'], anchor['py_start'], anchor['py_end'])
        elif anchor:
            cy = get_account_balance(a['id'], None, anchor['cy_end'])
            py = get_account_balance(a['id'], None, anchor['py_end'])
        else:
            cy, py = get_account_balance(a['id']), 0
        cy, py = cy * sign, py * sign
        chg = cy - py
        rows.append({'id': a['id'], 'name': a['name'], 'description': a['description'],
                     'type': st, 'cy': cy, 'py': py, 'chg': chg,
                     'pct': (chg / py * 100) if py else None})
    tcy = sum(r['cy'] for r in rows)
    tpy = sum(r['py'] for r in rows)
    return {'code': code, 'rows': rows, 'cy': tcy, 'py': tpy, 'chg': tcy - tpy,
            'pct': ((tcy - tpy) / tpy * 100) if tpy else None, 'anchor': anchor}

def normalize_csv(rows_raw):
    """THE CSV pre-processor for the CLI and MCP import paths (one copy — the
    interfaces used to carry diverging duplicates). Detects the header by
    CONTENT (row 0's date cell isn't a date ⇒ header), repairs rows with extra
    fields from unquoted commas, normalizes >4-column bank exports to
    [date, description, amount] — and nets separate Debit/Credit columns into
    ONE SIGNED amount (a credit is not just 'the other column').

    Returns (has_header, data_rows, repairs)."""
    if not rows_raw:
        return False, [], []
    first_row = rows_raw[0]
    header = [str(h).strip().lower() for h in first_row]
    has_header, sniff_date, sniff_desc, sniff_amt = sniff_csv_columns(rows_raw)
    start = 1 if has_header else 0
    data_rows = [list(r) for r in rows_raw[start:]]
    expected = len(first_row)
    repairs = []

    if expected > 4:
        # ── Multi-column bank export → 3 columns ──
        date_col, amt_cols, desc_cols = None, [], []
        dr_col = cr_col = None
        if has_header:
            for i, h in enumerate(header):
                if 'date' in h and date_col is None:
                    date_col = i
                elif '$' in h or h in ('amount', 'debit', 'credit'):
                    amt_cols.append(i)
                    if h == 'debit': dr_col = i
                    elif h == 'credit': cr_col = i
                elif any(kw in h for kw in ['description', 'desc', 'memo',
                                            'payee', 'detail', 'narrative']):
                    desc_cols.append(i)
        if date_col is None or not amt_cols:
            # No usable header names — fall back to content-sniffed columns
            if sniff_date < 0 or sniff_amt < 0:
                return has_header, data_rows, repairs
            date_col, amt_cols = sniff_date, [sniff_amt]
            desc_cols = [sniff_desc] if sniff_desc >= 0 else []
            dr_col = cr_col = None

        def _amount(vals_by_col):
            if dr_col is not None and cr_col is not None:
                try:
                    dr = parse_amount(vals_by_col.get(dr_col, '') or '0')
                    cr = parse_amount(vals_by_col.get(cr_col, '') or '0')
                    return f"{(dr - cr) / 100:.2f}"
                except ValueError:
                    pass   # unparseable → fall through; downstream flags the row
            for c in amt_cols:
                v = str(vals_by_col.get(c, '') or '').strip()
                if v:
                    return v
            return ''

        normalized = []
        for idx, row in enumerate(data_rows):
            n = len(row)
            row_num = idx + start + 1
            date_val = str(row[date_col]).strip() if date_col < n else ''
            if n > expected:
                # Repair: unquoted commas split the description; amounts sit at
                # the row's END in amt_cols order.
                extra = n - expected
                amt_start = n - len(amt_cols)
                vals = {amt_cols[k]: row[amt_start + k] for k in range(len(amt_cols))
                        if amt_start + k < n}
                desc_parts = [str(row[i]).strip() for i in range(date_col + 1, amt_start)
                              if str(row[i]).strip()]
                desc_joined = ': '.join(desc_parts)
                repairs.append((row_num, extra, desc_joined[:50]))
            else:
                vals = {c: row[c] for c in amt_cols if c < n}
                desc_parts = [str(row[c]).strip() for c in desc_cols
                              if c < n and str(row[c]).strip()]
                desc_joined = ': '.join(desc_parts)
            normalized.append([date_val, desc_joined, _amount(vals)])
        return has_header, normalized, repairs

    # ── 3-4 column Grid format: repair rows with extra fields ──
    for i, row in enumerate(data_rows):
        if len(row) > expected:
            row_num = i + start + 1
            extra = len(row) - expected
            amt_count = expected - 2
            date_val = row[0]
            desc_fields = row[1: len(row) - amt_count]
            amt_fields = row[len(row) - amt_count:]
            merged = ', '.join(str(f).strip() for f in desc_fields)
            data_rows[i] = [date_val, merged] + amt_fields
            repairs.append((row_num, extra, merged[:50]))
    return has_header, data_rows, repairs

def bump_next_ref(account_id):
    """Increment [Next Ref#] after a successful auto-numbered post. Manual
    references never bump the counter (LAP rule)."""
    with get_db() as db:
        db.execute("UPDATE accounts SET next_ref = next_ref + 1 WHERE id=? AND next_ref > 0",
                   (account_id,))

def set_account_next_ref(account_id, value):
    with get_db() as db:
        db.execute("UPDATE accounts SET next_ref=? WHERE id=?", (max(0, int(value)), account_id))

# ─── Ledger ───────────────────────────────────────────────────────
def get_ledger(account_id, date_from=None, date_to=None, opening_balance=0):
    with get_db() as db:
        sql = """
            WITH txn_counts AS (
                SELECT transaction_id, COUNT(*) as line_count
                FROM lines
                WHERE transaction_id IN (SELECT transaction_id FROM lines WHERE account_id = ?)
                GROUP BY transaction_id
            )
            SELECT t.id as txn_id, t.date, t.reference, t.description as txn_desc,
                   l.amount, l.description as line_desc, l.id as line_id, l.reconciled,
                   l.doc_on_file,
                   GROUP_CONCAT(DISTINCT a2.name) as cross_accounts,
                   tc.line_count
            FROM lines l
            JOIN transactions t ON l.transaction_id = t.id
            JOIN txn_counts tc ON tc.transaction_id = t.id
            LEFT JOIN lines l2 ON l2.transaction_id = t.id AND l2.account_id != ?
            LEFT JOIN accounts a2 ON l2.account_id = a2.id
            WHERE l.account_id = ?"""
        params = [account_id, account_id, account_id]   # CTE subquery, l2 != filter, WHERE l.account_id
        if date_from: sql += " AND t.date >= ?"; params.append(date_from)
        if date_to: sql += " AND t.date <= ?"; params.append(date_to)
        sql += " GROUP BY l.id ORDER BY t.date, t.id, l.sort_order"
        rows = db.execute(sql, params).fetchall()
        acct = get_account(account_id)
        sign = 1 if acct['normal_balance'] == 'D' else -1
        result, balance = [], opening_balance
        for row in rows:
            display_amount = row['amount'] * sign  # Flip for credit-normal accounts
            balance += display_amount
            result.append({
                'txn_id': row['txn_id'], 'line_id': row['line_id'],
                'date': row['date'], 'reference': row['reference'],
                'description': row['line_desc'] or row['txn_desc'],
                'amount': display_amount, 'raw_amount': row['amount'],
                'cross_accounts': row['cross_accounts'] or '',
                'running_balance': balance, 'reconciled': row['reconciled'],
                'doc_on_file': row['doc_on_file'],
                'line_count': row['line_count']})
        return result

# ─── Balance Computation ──────────────────────────────────────────
def get_account_balance(account_id, date_from=None, date_to=None):
    """Raw sum of lines (D positive, C negative) in date range."""
    with get_db() as db:
        if not date_from and not date_to:
            # No date filter — skip the JOIN to transactions (much faster on large files)
            return db.execute(
                "SELECT COALESCE(SUM(amount),0) as total FROM lines WHERE account_id=?",
                (account_id,)).fetchone()['total']
        sql = "SELECT COALESCE(SUM(l.amount),0) as total FROM lines l JOIN transactions t ON l.transaction_id=t.id WHERE l.account_id=?"
        params = [account_id]
        if date_from: sql += " AND t.date >= ?"; params.append(date_from)
        if date_to: sql += " AND t.date <= ?"; params.append(date_to)
        return db.execute(sql, params).fetchone()['total']

# ─── Flows: one side of the account, not the net ───────────────────
# LAP doctrine: "debits and credits ARE the flows." The DEBIT side of AR is
# sales by customer; the CREDIT side of AP is purchases by supplier; the debit
# side of inventory is purchases and the credit side is COGS. Netting throws
# all of that away, which is why a net-only report cannot answer the questions
# people actually ask of a sub-ledger.
FLOW_DEBIT, FLOW_CREDIT = 'D', 'C'

# A boomerang (an item and its cross-item on the SAME account — how a partial
# payment leaves the invoice remainder open, ai.txt "BOOMERANG") nets to zero on
# the account and is NOT a flow. LAP practice excludes it from the debit and credit views
# for exactly that reason (p.182); counting it would inflate both sides by the
# same invented amount. Detected structurally: within ONE transaction, the same
# account carrying both signs.
_NOT_A_BOOMERANG = ("""
        AND NOT EXISTS (SELECT 1 FROM lines x
                        WHERE x.transaction_id = l.transaction_id
                          AND x.account_id = l.account_id
                          AND (x.amount > 0) <> (l.amount > 0))""")


def get_all_account_balances(date_from=None, date_to=None, side=None):
    """Bulk fetch: raw balance for ALL accounts in one query. Returns {account_id: balance}.

    side=None  the net balance (what a statement column shows)
    side='D'   debits only  — the flow IN, boomerangs excluded
    side='C'   credits only — the flow OUT, boomerangs excluded (stays negative;
               the caller's normal-balance sign flip presents it)"""
    if side not in (None, FLOW_DEBIT, FLOW_CREDIT):
        raise ValueError(f"side must be None, 'D' or 'C' — got {side!r}")
    whole_book = not date_from and not date_to and not side
    # A statement column is "as at" a date — no floor. Expressed as `t.date <= ?`
    # alone, SQLite plans it as a scan of every LINE with a primary-key probe
    # into transactions for each one: 1.9M random lookups. Given a lower bound
    # it uses idx_txn_date as a range instead. Measured on a 973k-row book:
    # 3.57s -> 1.07s. '' is the safe floor — every string sorts at or above it,
    # so the result set cannot change (verified identical).
    if date_to and not date_from:
        date_from = ''
    # Cached against the BALANCE generation (trigger-maintained, v155):
    # recomputed the moment money moves, free until then — and a layout edit
    # (move a line, save columns, ref-mark) no longer discards it. Covers the
    # dated statement columns too, which is what made opening a balance sheet
    # cost seconds every single time.
    key = (date_from or '', date_to or '', side or '')
    bal_gen, _chain = _data_gens()
    if _balances_cache.get('gen') != bal_gen or _balances_cache.get('path') != DB_PATH:
        _balances_cache.clear()
        _balances_cache.update({'gen': bal_gen, 'path': DB_PATH, 'by_args': {}})
    hit = _balances_cache['by_args'].get(key)
    if hit is not None:
        return dict(hit)
    with get_db() as db:
        if whole_book:
            # No date filter — skip the JOIN to transactions (12x faster on large files)
            sql = "SELECT account_id, COALESCE(SUM(amount),0) as total FROM lines GROUP BY account_id"
            rows = db.execute(sql).fetchall()
        else:
            sql = ("SELECT l.account_id, COALESCE(SUM(l.amount),0) as total "
                   "FROM lines l JOIN transactions t ON l.transaction_id=t.id WHERE 1=1")
            params = []
            # `is not None`, not truthiness: '' is a deliberate floor (see above)
            # and it is falsy, so testing truth silently dropped the predicate
            # that makes SQLite choose the range plan.
            if date_from is not None: sql += " AND t.date >= ?"; params.append(date_from)
            if date_to: sql += " AND t.date <= ?"; params.append(date_to)
            if side:
                sql += " AND l.amount > 0" if side == FLOW_DEBIT else " AND l.amount < 0"
                sql += _NOT_A_BOOMERANG
            sql += " GROUP BY l.account_id"
            rows = db.execute(sql, params).fetchall()
        out = {r['account_id']: r['total'] for r in rows}
    # Bounded: a report with many columns must not grow this without limit.
    if len(_balances_cache['by_args']) < BALANCE_CACHE_MAX:
        _balances_cache['by_args'][key] = dict(out)   # a copy — callers must not edit it
    return out


def proof_cents():
    """The LAP PROOF line: every line amount in the book, added up. Zero, always.

    Derived from the cached whole-book balances rather than its own
    `SELECT SUM(amount) FROM lines`, so the header costs ONE scan of the book
    per write instead of two per page render."""
    return sum(get_all_account_balances().values())


def get_all_report_items():
    """Get all report items across ALL reports. Used for building the global total-to chain."""
    with get_db() as db:
        return db.execute(
            "SELECT ri.*, a.name as acct_name, a.description as acct_desc, "
            "a.normal_balance, a.account_type, a.account_number, a.computed "
            "FROM report_items ri LEFT JOIN accounts a ON ri.account_id = a.id "
            "ORDER BY ri.report_id, ri.position").fetchall()

def compute_report_column(report_id, date_from=None, date_to=None,
                          _display_items=None, _all_items=None, side=None):
    """
    Compute one analysis column for a report.
    The total-to chain is GLOBAL across all reports (BS, IS, TRX all cross-talk).
    Uses raw DB balances for accumulation, applies display sign at the end.
    Each account is processed ONCE (first occurrence with total-to wins) to avoid
    double-counting when the same account appears on multiple reports.
    """
    display_items = _display_items or get_report_items(report_id)
    all_items = _all_items or get_all_report_items()
    tt_fields = ['total_to_1','total_to_2','total_to_3','total_to_4','total_to_5','total_to_6']

    # Step 1: Get RAW balances for all posting accounts in ONE query.
    # side='D'/'C' makes this a FLOWS column — one side of each account instead
    # of the net. The total-to arithmetic below is untouched: a flow ripples up
    # the chain exactly as a net balance does, so "sales by customer" totals to
    # the same controlling account its net would.
    bulk_bal = get_all_account_balances(date_from, date_to, side=side)
    raw_bal = {}
    seen = set()
    for it in all_items:
        if it['account_id'] and it['account_type'] == 'posting' and it['acct_name'] not in seen:
            seen.add(it['acct_name'])
            raw_bal[it['acct_name']] = bulk_bal.get(it['account_id'], 0)

    # Step 2: Deduplicate items by account name.
    # For each account name, we only process it ONCE for total-to purposes.
    # If the same account appears on multiple reports with different total-to's,
    # merge all total-to targets from all occurrences.
    # For items without an account name (labels, separators), skip.
    acct_tt = {}  # name -> set of (field, target) pairs
    acct_meta = {}  # name -> first item's metadata
    keys0 = all_items[0].keys() if all_items else ()
    for it in all_items:
        name = it['acct_name']
        if not name:
            continue
        # A COMPUTED line (Opening/Closing RE) is PRESENTATION. Its value is read
        # off another account as at a date; it owns no postings and it must never
        # feed the chain. Enforced here, structurally, rather than relying on
        # every migration to have blanked its total-to: a legacy file whose
        # RE.OPEN still pointed at RE went on double-counting the opening balance
        # right through the repair, because the repair only knew how to unhook
        # the one target it expected.
        if ('computed' in keys0) and (it['computed'] or ''):
            continue
        if name not in acct_meta:
            acct_meta[name] = it
        # Collect all total-to targets from all occurrences
        if name not in acct_tt:
            acct_tt[name] = set()
        for ttf in tt_fields:
            target = it[ttf]
            if target:
                acct_tt[name].add(target)

    # Step 3: Multi-pass accumulation until stable.
    # An account's value = its own raw balance (if posting) + anything accumulated into it.
    # This handles posting accounts that are also accumulation targets
    # (e.g., AR receives from AR.ALL sub-ledger but also has its own postings).
    accumulated = {}
    
    for _pass in range(100):  # iterate to convergence; cap guards a cyclic total_to misconfig
        prev = dict(accumulated)
        accumulated = {}
        
        for name, targets in acct_tt.items():
            if not targets:
                continue
            # Value = own raw balance + anything accumulated into this account
            own_raw = raw_bal.get(name, 0)
            acc_into = prev.get(name, 0)
            val = own_raw + acc_into
            
            # Dump into each unique target
            for target in targets:
                accumulated[target] = accumulated.get(target, 0) + val
        
        if accumulated == prev:
            break

    # Step 4: Merge — each account's display value is raw + accumulated
    merged = {}
    all_names = set(raw_bal.keys()) | set(accumulated.keys())
    for name in all_names:
        merged[name] = raw_bal.get(name, 0) + accumulated.get(name, 0)

    # Build normal_balance map
    nb_map = {}
    for it in all_items:
        if it['acct_name'] and it['normal_balance']:
            nb_map[it['acct_name']] = it['normal_balance']

    # Step 5: Return display items with sign-adjusted balances.
    # Computed lines (Opening/Closing RE) are OFF-BOOK: their value is the perpetual
    # balance of a source account as of the period start/end — not their own postings,
    # and they total_to nothing, so they never feed the balance sheet.
    def _day_before(iso):
        if not iso: return None
        y, m, d = (int(x) for x in iso.split('-'))
        return (date(y, m, d) - timedelta(days=1)).isoformat()
    result = []
    for it in display_items:
        keys = it.keys()
        comp = (it['computed'] if 'computed' in keys else '') or ''
        if comp:
            mode, _, src = comp.partition(':')
            src = src or 'RE'
            asof = _day_before(date_from) if mode == 'open' else date_to
            val = trace_account(src, date_from=None, date_to=asof)['display'] if asof else 0
            result.append((dict(it), val))
            continue
        name = it['acct_name']
        raw = merged.get(name, 0) if name else 0
        nb = nb_map.get(name, it['normal_balance'] or 'D')
        sign = 1 if nb == 'D' else -1
        result.append((dict(it), raw * sign))
    return result


_trace_cache = {}   # (gens, path, name, from, to) -> result; bounded below

def trace_account(account_name, date_from=None, date_to=None):
    """Trace the full accumulation tree for a report account.

    Returns a dict with:
      name, normal_balance, own_raw, accumulated, merged, display,
      contributors: [{name, own_raw, accumulated, value_dumped, display, targets, report}]

    The contributors list shows every account that feeds into this one
    (directly via total_to), with the amounts from the current accumulation pass.

    Memoised on (balance_gen, chain_gen): the computed Opening/Closing RE lines
    call this once per statement column, and each call re-ran the full-chart
    accumulation. A layout move bumps neither generation, so re-rendering after
    one is cache hits all the way down.
    """
    key = (_data_gens(), DB_PATH, account_name, date_from or '', date_to or '')
    hit = _trace_cache.get(key)
    if hit is not None:
        return {**hit, 'contributors': [dict(c) for c in hit['contributors']]}
    all_items = get_all_report_items()
    tt_fields = ['total_to_1','total_to_2','total_to_3','total_to_4','total_to_5','total_to_6']

    bulk_bal = get_all_account_balances(date_from, date_to)
    raw_bal = {}
    seen = set()
    for it in all_items:
        if it['account_id'] and it['account_type'] == 'posting' and it['acct_name'] not in seen:
            seen.add(it['acct_name'])
            raw_bal[it['acct_name']] = bulk_bal.get(it['account_id'], 0)

    acct_tt = {}
    for it in all_items:
        name = it['acct_name']
        if not name:
            continue
        if name not in acct_tt:
            acct_tt[name] = set()
        for ttf in tt_fields:
            target = it[ttf]
            if target:
                acct_tt[name].add(target)

    # Multi-pass accumulation (same as compute_report_column)
    accumulated = {}
    for _pass in range(100):  # iterate to convergence; cap guards a cyclic total_to misconfig
        prev = dict(accumulated)
        accumulated = {}
        for name, targets in acct_tt.items():
            if not targets:
                continue
            val = raw_bal.get(name, 0) + prev.get(name, 0)
            for target in targets:
                accumulated[target] = accumulated.get(target, 0) + val
        if accumulated == prev:
            break

    # Build normal_balance map and report map
    nb_map = {}
    report_map = {}
    for it in all_items:
        if it['acct_name']:
            if it['normal_balance']:
                nb_map[it['acct_name']] = it['normal_balance']
            if it['acct_name'] not in report_map:
                report_map[it['acct_name']] = it['report_id']

    target_name = account_name.upper()
    target_nb = nb_map.get(target_name, 'D')
    sign = 1 if target_nb == 'D' else -1
    own_raw = raw_bal.get(target_name, 0)
    acc_into = accumulated.get(target_name, 0)
    merged = own_raw + acc_into
    display = merged * sign

    # Find direct contributors: accounts whose total_to includes target_name
    contributors = []
    for name, targets in sorted(acct_tt.items()):
        if target_name in targets:
            c_own = raw_bal.get(name, 0)
            c_acc = accumulated.get(name, 0)
            c_val = c_own + c_acc  # value dumped into target
            c_nb = nb_map.get(name, 'D')
            c_sign = 1 if c_nb == 'D' else -1
            # Get report name for this account
            rpt_id = report_map.get(name)
            rpt_name = ''
            if rpt_id:
                rpt = get_report(rpt_id)
                if rpt:
                    rpt_name = rpt['name']
            contributors.append({
                'name': name,
                'normal_balance': c_nb,
                'own_raw': c_own,
                'accumulated': c_acc,
                'value_dumped': c_val,
                'display': c_val * sign,  # show in target's sign convention
                'targets': sorted(targets),
                'report': rpt_name,
            })

    out = {
        'name': target_name,
        'normal_balance': target_nb,
        'own_raw': own_raw,
        'accumulated': acc_into,
        'merged': merged,
        'display': display,
        'date_from': date_from,
        'date_to': date_to,
        'contributors': sorted(contributors, key=lambda c: abs(c['value_dumped']), reverse=True),
        'feeds_into': sorted(acct_tt.get(target_name, set())),
    }
    if len(_trace_cache) >= BALANCE_CACHE_MAX * 2:
        _trace_cache.clear()   # bounded; stale keys die with their generations
    _trace_cache[key] = {**out, 'contributors': [dict(c) for c in out['contributors']]}
    return out


def subledger_report_for_total(account_name):
    """For a TOTAL (accumulator) account, return the report_id that holds its detail —
    the report whose accounts total_to this account. Returns that report_id only if
    exactly ONE such report exists; otherwise None (caller shows a read-only notice
    instead of jumping). Matches the engine's total_to_1..6 accumulation fields.
    Lets a click on e.g. 'Detailed AR' jump straight to the AR subledger."""
    with get_db() as db:
        rows = db.execute(
            "SELECT DISTINCT ri.report_id FROM report_items ri "
            "WHERE ? IN (ri.total_to_1, ri.total_to_2, ri.total_to_3, "
            "             ri.total_to_4, ri.total_to_5, ri.total_to_6)",
            (account_name,)).fetchall()
        rids = [r['report_id'] for r in rows]
        return rids[0] if len(rids) == 1 else None


def migrate_re_computed():
    """One-time, idempotent rewire to LAP-style retained earnings.

    Target (the LAP model — one perpetual RE figure, no closing):
      - ONE Retained Earnings line on the BS (the perpetual `RE` total). No separate
        balance-sheet "CONVERSION" section.
      - A SINGLE opening-balance posting account `RE.OB` (LAP "Opening Retained
        Earnings") holds the conversion opening RE and totals into `RE` from the
        off-statement TRX workings report.
      - NI / DIVPAID feed `RE` directly (LAP "Change to Retained Earnings" -> BS-RE).
      - RE.OPEN / RE.CLOSE become computed display-only IS lines (open:RE / close:RE)
        that total_to nothing, so the BS RE stays independent & perpetual.
      - The obsolete RE.OFS "totals-to-nothing" plug and its self-cancelling pickup
        transactions are deleted whole (balanced txns), and the legacy conversion
        cruft (PY.CONV section, PY.CLOSE, RE.CONV, RE.WORK, TRX.OPNE typo) is pruned.

    Self-healing & idempotent: runs from any prior state (un-migrated, or an older
    v92/v93 RE.CONV/RE.WORK migration) and is a no-op once the clean shape is in place.
    Returns True if it changed anything. (prune-don't-layer)
    """
    # Accounts that are pure RE/conversion presentation scaffolding — a transaction
    # built only from these is a self-cancelling plug we can delete whole.
    RE_SCAFFOLD = {'RE.OFS', 'RE.OPEN', 'RE.CLOSE', 'RE', 'RE.CONV', 'RE.OB',
                   'PY.CLOSE', 'PY.CONV'}

    global _re_repair

    def _re_now():
        """Retained earnings as the balance sheet currently reports it."""
        try:
            return trace_account('RE')['display']
        except Exception:
            return None

    before = _re_now()
    changed = _migrate_re_body(RE_SCAFFOLD)
    if not changed:
        return False

    after = _re_now()
    lines = ['Retained earnings wiring was REPAIRED when this file was opened '
             '(it was in an older shape).']
    if before is None or after is None:
        lines.append('Retained earnings could not be read before the repair — check the '
                     'balance sheet against your last statements.')
    elif before == after:
        lines.append(f'Retained earnings is UNCHANGED at {fmt_amount_plain(after)}.')
    else:
        lines.append(f'Retained earnings MOVED: {fmt_amount_plain(before)} -> '
                     f'{fmt_amount_plain(after)} '
                     f'(a difference of {fmt_amount_plain(after - before)}). '
                     f'This is the figure the old wiring was reporting wrongly — but '
                     f'check it against your last statements before you rely on it.')
    _re_repair = lines
    return True


def _migrate_re_body(RE_SCAFFOLD):
    """The repair itself. Called only by migrate_re_computed, which measures
    retained earnings either side of it so the change can be reported."""
    with get_db() as db:
        def acct(name):
            return db.execute("SELECT * FROM accounts WHERE name=?", (name,)).fetchone()
        def report(name):
            return db.execute("SELECT * FROM reports WHERE name=?", (name,)).fetchone()

        re, re_open, re_close = acct('RE'), acct('RE.OPEN'), acct('RE.CLOSE')
        if not (re and re_open and re_close):
            return False                       # not the standard RE chain — leave it alone

        # Already in the clean LAP shape? -> nothing to do.
        #
        # This must test for the TARGET SHAPE, not merely for the absence of the
        # legacy cruft. It used to test only the latter, and since v95 built
        # starter books lean — RE.OB created directly, no RE.OFS/PY.CONV zoo —
        # a brand-new file matched "clean" on its very first migration and left
        # here with RE.OPEN still an ordinary posting account. The wiring never
        # happened, on that open or any later one, so the Statement of Retained
        # Earnings opening line stayed dead and the balance sheet's RE differed
        # from the income statement's closing RE by the whole opening balance.
        # Found by check_books the first time it was ever run, on a file one
        # function call old. Every set of books created since v95 has it.
        if (acct('RE.OB') and not acct('RE.OFS') and not acct('RE.CONV')
                and not acct('PY.CONV') and not acct('PY.CLOSE')
                and not acct('TRX.OPNE') and not report('RE.WORK') and not report('RE.OFS')
                and (re_open['computed'] or '') == 'open:RE'
                and (re_close['computed'] or '') == 'close:RE'):
            return False

        def repoint(account_id, old, new):
            """Replace total-to target `old` with `new` ('' drops it) in every slot."""
            for r in db.execute("SELECT id,total_to_1,total_to_2,total_to_3,total_to_4,total_to_5,total_to_6 "
                                "FROM report_items WHERE account_id=?", (account_id,)).fetchall():
                vals = [r['total_to_1'], r['total_to_2'], r['total_to_3'],
                        r['total_to_4'], r['total_to_5'], r['total_to_6']]
                if old in vals:
                    vals = [new if v == old else v for v in vals]
                    db.execute("UPDATE report_items SET total_to_1=?,total_to_2=?,total_to_3=?,"
                               "total_to_4=?,total_to_5=?,total_to_6=? WHERE id=?", (*vals, r['id']))

        def ensure_susp():
            s = acct('EX.SUSP')
            if not s:
                db.execute("INSERT INTO accounts(name,description,normal_balance,account_type) "
                           "VALUES('EX.SUSP','Suspense — needs clearing','D','posting')")
                s = acct('EX.SUSP')
            return s

        # 1. Retire the RE.OFS plug. Every txn touching RE.OFS is obsolete presentation
        #    scaffolding. If the whole txn is RE-family (the standard self-cancelling
        #    pickup, e.g. RE.OFS/RE.OPEN), delete the WHOLE balanced transaction (this
        #    removes the double-counted opening RE). If RE.OFS is mixed with real
        #    accounts, don't guess — strand just the RE.OFS leg to Suspense for review.
        ofs = acct('RE.OFS')
        if ofs:
            txn_ids = [r['transaction_id'] for r in db.execute(
                "SELECT DISTINCT transaction_id FROM lines WHERE account_id=?", (ofs['id'],)).fetchall()]
            for tid in txn_ids:
                names = [r['name'] for r in db.execute(
                    "SELECT a.name FROM lines l JOIN accounts a ON a.id=l.account_id "
                    "WHERE l.transaction_id=?", (tid,)).fetchall()]
                if all(n in RE_SCAFFOLD for n in names):
                    db.execute("DELETE FROM transactions WHERE id=?", (tid,))   # cascades its lines
                else:
                    susp = ensure_susp()
                    db.execute("UPDATE lines SET account_id=? WHERE transaction_id=? AND account_id=?",
                               (susp['id'], tid, ofs['id']))

        # 2. Single opening-RE account RE.OB (LAP "Opening Retained Earnings").
        ob = acct('RE.OB')
        if not ob:
            db.execute("INSERT INTO accounts(name,description,normal_balance,account_type) "
                       "VALUES('RE.OB','Opening retained earnings (b/f)','C','posting')")
            ob = acct('RE.OB')
        # Genuine opening pickups (remaining RE.OPEN postings + any legacy RE.CONV) -> RE.OB
        db.execute("UPDATE lines SET account_id=? WHERE account_id=?", (ob['id'], re_open['id']))
        conv = acct('RE.CONV')
        if conv:
            db.execute("UPDATE lines SET account_id=? WHERE account_id=?", (ob['id'], conv['id']))

        # 3. RE.OB feeds RE from the off-statement TRX workings report (not the BS face).
        feeds = db.execute("SELECT 1 FROM report_items WHERE account_id=? AND 'RE' IN "
                           "(total_to_1,total_to_2,total_to_3,total_to_4,total_to_5) LIMIT 1",
                           (ob['id'],)).fetchone()
        if not feeds:
            trx = report('TRX')
            if not trx:
                db.execute("INSERT INTO reports(name,description,sort_order) "
                           "VALUES('TRX','Conversion workings (off-statement)',99)")
                trx = report('TRX')
            pos = db.execute("SELECT COALESCE(MAX(position),0)+10 p FROM report_items WHERE report_id=?",
                             (trx['id'],)).fetchone()['p']
            db.execute("INSERT INTO report_items(report_id,position,item_type,account_id,indent,total_to_1,description) "
                       "VALUES(?,?,'account',?,2,'RE','Opening retained earnings (b/f)')", (trx['id'], pos, ob['id']))

        # 4. NI / DIVPAID feed RE directly (were feeding the presentation RE.CLOSE).
        for nm in ('NI', 'DIVPAID'):
            a = acct(nm)
            if a:
                repoint(a['id'], 'RE.CLOSE', 'RE')

        # 5. Decouple the presentation lines from the BS chain + mark them computed.
        # Clear EVERY total-to slot on both presentation lines — not just the one
        # target this migration happens to expect. They feed nothing, ever.
        for _pid in (re_open['id'], re_close['id']):
            db.execute("UPDATE report_items SET total_to_1='',total_to_2='',total_to_3='',"
                       "total_to_4='',total_to_5='',total_to_6='' WHERE account_id=?", (_pid,))
        db.execute("UPDATE accounts SET computed='open:RE', account_type='total' WHERE id=?", (re_open['id'],))
        db.execute("UPDATE accounts SET computed='close:RE' WHERE id=?", (re_close['id'],))

        # 6. Prune the conversion zoo. report_items FIRST (FK has no cascade), then the
        #    account/report. A doomed account with unexpected postings is stranded to
        #    Suspense rather than silently destroyed.
        def drop_account(name):
            a = acct(name)
            if not a:
                return
            n = db.execute("SELECT COUNT(*) c FROM lines WHERE account_id=?", (a['id'],)).fetchone()['c']
            if n:
                susp = ensure_susp()
                db.execute("UPDATE lines SET account_id=? WHERE account_id=?", (susp['id'], a['id']))
            db.execute("DELETE FROM report_items WHERE account_id=?", (a['id'],))
            db.execute("DELETE FROM accounts WHERE id=?", (a['id'],))

        def drop_report(name):
            r = report(name)
            if not r:
                return
            db.execute("DELETE FROM report_items WHERE report_id=?", (r['id'],))
            db.execute("DELETE FROM reports WHERE id=?", (r['id'],))

        bs = report('BS')
        if bs:                                   # drop the BS "CONVERSION" section label
            db.execute("DELETE FROM report_items WHERE report_id=? AND item_type='label' "
                       "AND UPPER(TRIM(description))='CONVERSION'", (bs['id'],))
        drop_account('PY.CONV')                  # was the BS CONVERSION section line
        drop_account('PY.CLOSE')                 # was on TRX -> RE
        drop_account('TRX.OPNE')                 # typo orphan
        drop_account('RE.CONV')                  # legacy v92 opening holder (lines already moved)
        drop_report('RE.WORK')                   # legacy v92 off-statement report
        drop_report('RE.OFS')                    # obsolete plug report
        drop_account('RE.OFS')                   # now empty (or stranded) -> safe to delete

        # 7. Tidy the gap the CONVERSION section left between Total Equity and the
        #    TOTAL L&E grand total: keep a single rule, drop redundant blanks.
        if bs:
            def _pos(name):
                r = db.execute("SELECT position FROM report_items WHERE report_id=? AND item_type='total' "
                               "AND account_id=(SELECT id FROM accounts WHERE name=?)", (bs['id'], name)).fetchone()
                return r['position'] if r else None
            eqp, tlp = _pos('EQ'), _pos('TL')
            if eqp is not None and tlp is not None:
                kept_sep = False
                for g in db.execute("SELECT id,item_type,description FROM report_items WHERE report_id=? "
                                    "AND position>? AND position<? ORDER BY position",
                                    (bs['id'], eqp, tlp)).fetchall():
                    if g['item_type'] == 'separator' and not kept_sep:
                        kept_sep = True
                        continue
                    if g['item_type'] in ('separator', 'label') and not (g['description'] or '').strip():
                        db.execute("DELETE FROM report_items WHERE id=?", (g['id'],))
        return True


# ─── Trial Balance ────────────────────────────────────────────────
def get_trial_balance(as_of_date=None):
    with get_db() as db:
        accounts = db.execute("SELECT * FROM accounts WHERE account_type='posting' ORDER BY name").fetchall()
        # Build map of account_id -> report names
        report_map = {}
        rows = db.execute("""
            SELECT ri.account_id, GROUP_CONCAT(DISTINCT r.name) as report_names
            FROM report_items ri JOIN reports r ON ri.report_id = r.id
            WHERE ri.account_id IS NOT NULL
            GROUP BY ri.account_id""").fetchall()
        for row in rows:
            report_map[row['account_id']] = row['report_names']
        result, total_dr, total_cr = [], 0, 0
        for acct in accounts:
            raw = get_account_balance(acct['id'], date_to=as_of_date)
            if raw == 0: continue
            sign = 1 if acct['normal_balance'] == 'D' else -1
            bal = raw * sign
            dr = bal if bal > 0 and acct['normal_balance'] == 'D' else (abs(bal) if bal < 0 and acct['normal_balance'] == 'C' else 0)
            cr = bal if bal > 0 and acct['normal_balance'] == 'C' else (abs(bal) if bal < 0 and acct['normal_balance'] == 'D' else 0)
            total_dr += dr; total_cr += cr
            result.append({'id': acct['id'], 'name': acct['name'], 'description': acct['description'],
                'normal_balance': acct['normal_balance'], 'account_number': acct['account_number'] or '',
                'balance': bal, 'debit': dr, 'credit': cr,
                'report_name': report_map.get(acct['id'], '')})
        return result, total_dr, total_cr

# ─── Search ───────────────────────────────────────────────────────
def search_transactions(query, limit=100):
    with get_db() as db:
        q = f"%{query}%"
        return db.execute("""
            SELECT DISTINCT t.id as txn_id, t.date, t.reference, t.description,
                   GROUP_CONCAT(DISTINCT a.name) as accounts,
                   (SELECT SUM(ABS(l2.amount)) FROM lines l2 WHERE l2.transaction_id=t.id AND l2.amount > 0) as total_amount
            FROM transactions t JOIN lines l ON l.transaction_id = t.id
            JOIN accounts a ON l.account_id = a.id
            WHERE t.description LIKE ? OR t.reference LIKE ? OR a.name LIKE ? OR l.description LIKE ?
            GROUP BY t.id ORDER BY t.date DESC LIMIT ?""", (q, q, q, q, limit)).fetchall()

# ─── Tax Codes ───────────────────────────────────────────────────
def get_tax_codes():
    with get_db() as db:
        return db.execute("SELECT * FROM tax_codes ORDER BY id").fetchall()

def get_tax_code(code_id):
    with get_db() as db:
        return db.execute("SELECT * FROM tax_codes WHERE id=?", (code_id,)).fetchone()

def save_tax_code(code_id, description, rate_percent, collected_account='', paid_account=''):
    with get_db() as db:
        db.execute("INSERT OR REPLACE INTO tax_codes(id, description, rate_percent, collected_account, paid_account) VALUES(?,?,?,?,?)",
            (code_id, description, rate_percent, collected_account, paid_account))

def delete_tax_code(code_id):
    with get_db() as db:
        db.execute("DELETE FROM tax_codes WHERE id=?", (code_id,))

# ─── Import Rules ────────────────────────────────────────────────
def get_import_rules():
    with get_db() as db:
        # Ties break by LONGEST keyword first — "cpc scp" must beat "cpc";
        # alphabetical was letting less-specific rules win.
        return db.execute("SELECT * FROM import_rules "
                          "ORDER BY priority DESC, LENGTH(keyword) DESC, keyword").fetchall()

def save_import_rule(rule_id, keyword, account_name, tax_code='', priority=0, notes=''):
    with get_db() as db:
        if rule_id:
            db.execute("UPDATE import_rules SET keyword=?, account_name=?, tax_code=?, priority=?, notes=? WHERE id=?",
                (keyword, account_name, tax_code, priority, notes, rule_id))
        else:
            db.execute("INSERT INTO import_rules(keyword, account_name, tax_code, priority, notes) VALUES(?,?,?,?,?)",
                (keyword, account_name, tax_code, priority, notes))

def delete_import_rule(rule_id):
    with get_db() as db:
        db.execute("DELETE FROM import_rules WHERE id=?", (rule_id,))

def apply_rules(description, amount_cents):
    """Apply import rules to a description. Returns (account_name, tax_code, lines).
    lines is the list of (account_id, amount, desc) tuples ready for posting.
    If no rule matches, returns ('EX.SUSP', '', simple_lines)."""
    rules = get_import_rules()
    desc_lower = description.lower()
    desc_norm = desc_lower.replace('-', ' ')

    matched_rule = None
    for rule in rules:
        kw = rule['keyword'].lower().replace('-', ' ')
        if len(kw) <= 4:
            # Short keywords require word boundary match to avoid
            # false positives like NSF matching traNSFer
            if _re.search(r'(?<![a-z])' + _re.escape(kw) + r'(?![a-z])', desc_norm):
                matched_rule = rule
                break
        elif kw in desc_norm:
            matched_rule = rule
            break  # rules are priority-sorted, first match wins
    
    if not matched_rule:
        return 'EX.SUSP', '', None
    
    acct_name = matched_rule['account_name']
    tax_id = matched_rule['tax_code']
    
    if tax_id:
        tc = get_tax_code(tax_id)
        if tc and tc['rate_percent'] > 0:
            rate = tc['rate_percent']
            # Amount is tax-inclusive. Split: tax = amount * rate / (100 + rate)
            tax_cents = round(abs(amount_cents) * rate / (100 + rate))
            net_cents = abs(amount_cents) - tax_cents
            
            # Determine which tax account to use
            if amount_cents > 0:
                # Money coming IN (revenue) → GST collected
                tax_acct = tc['collected_account'] or 'GST.OUT'
            else:
                # Money going OUT (expense) → GST paid (ITC)
                tax_acct = tc['paid_account'] or 'GST.IN'
            
            return acct_name, tax_id, {'net': net_cents, 'tax': tax_cents, 'tax_acct': tax_acct}
    
    return acct_name, tax_id, None

# ─── Suspense Reclassification ────────────────────────────────────

GENERIC_TERMS = {
    # Banking operations
    'cheque', 'check', 'chq', 'transfer', 'trf', 'tfr', 'xfer',
    'deposit', 'payment', 'withdrawal', 'debit', 'credit',
    'e-transfer', 'etransfer', 'interac', 'eft',
    'wire', 'draft', 'preauthorized', 'pre-authorized', 'pap',
    'pos', 'point of sale',
    # Generic descriptors
    'invoice', 'inv', 'receipt', 'refund', 'reversal',
    'fee', 'charge', 'interest', 'service charge',
    'monthly', 'annual', 'quarterly', 'weekly',
    'personal', 'business', 'commercial',
    'online', 'mobile', 'telephone', 'phone',
    # Catch-all banking
    'misc', 'miscellaneous', 'sundry', 'other', 'general',
    'adjustment', 'correction', 'void', 'acct', 'account',
    # Payroll
    'payroll', 'salary', 'wages', 'pay',
    # Too common
    'purchase', 'buy', 'sale', 'order',
    # Shareholder / owner (case-by-case)
    'sh draw', 'shareholder', 'owner', 'director',
    'draw', 'advance', 'loan',
    # Ambiguous
    'rent', 'lease',
    # Locations (not vendors)
    'downtown', 'uptown', 'midtown', 'core', 'central',
    'north', 'south', 'east', 'west', 'northwest', 'northeast',
    'southwest', 'southeast', 'sent', 'received',
}

def _extract_rule_keyword(description):
    """Extract a specific vendor/trade keyword from a transaction description.
    Returns the keyword string or '' if too generic."""
    if not description:
        return ''
    s = description.upper().strip()
    # Remove noise suffixes: terminal IDs, store numbers, dates, cities
    s = _re.sub(r'\s+STN\s*\d+', '', s)
    s = _re.sub(r'\s+STORE\s*\d+', '', s)
    s = _re.sub(r'\s+UNIT\s*\d+', '', s)
    s = _re.sub(r'\s*#\d+', '', s)
    s = _re.sub(r'\s+-\s+[A-Z]{2,3}$', '', s)  # trailing province codes
    s = _re.sub(r'\s+-\s+[A-Z]+\s+[A-Z]{2}$', '', s)  # "- CALGARY AB"
    s = _re.sub(r'\s+(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\s*\d*', '', s)
    s = _re.sub(r'\s+(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s*\d*', '', s)
    s = _re.sub(r'\s+Q[1-4]$', '', s)
    s = _re.sub(r'\s+(MONTHLY|ANNUAL|QUARTERLY|WEEKLY|ONLINE|MOBILE)$', '', s)
    s = _re.sub(r'\s+\d{4,}$', '', s)  # trailing long numbers
    # Strip generic prefix/suffix words and separators
    s = _re.sub(r'\s*-\s*', ' ', s)  # normalize dashes to spaces
    s = s.strip()
    if not s or len(s) < 3:
        return ''
    # Strip leading generic/short words (POS PURCHASE, E-TRANSFER DEPOSIT, etc.)
    words = s.split()
    while words and (words[0].lower() in GENERIC_TERMS or len(words[0]) <= 2):
        words.pop(0)
    # Strip trailing generic/short words
    while words and (words[-1].lower() in GENERIC_TERMS or len(words[-1]) <= 2):
        words.pop()
    # If nothing left after stripping generic words, reject
    if not words:
        return ''
    s = ' '.join(words)
    if len(s) < 3:
        return ''
    # Check if every remaining word is generic or too short
    remaining = s.lower().split()
    if all(w in GENERIC_TERMS or len(w) <= 2 for w in remaining):
        return ''
    # Check if the whole phrase is generic
    if s.lower() in GENERIC_TERMS:
        return ''
    return s

def _rule_already_exists(keyword):
    """Check if a rule exists that would cover this keyword."""
    kw_lower = keyword.lower()
    rules = get_import_rules()
    for rule in rules:
        rk = rule['keyword'].lower()
        if rk == kw_lower:
            return True
        if rk in kw_lower or kw_lower in rk:
            return True
    return False

def reclassify_suspense(txn_id, target_account_name, tax_code=''):
    """Reclassify a suspense transaction to the correct account in-place.

    Args:
        txn_id: Transaction ID to reclassify
        target_account_name: Account name to reclassify to (e.g. 'EX.OFFICE')
        tax_code: Optional tax code (e.g. 'G5'). Splits amount into net + tax.

    Returns:
        dict with: txn_id, old_account, new_account, amount_cents, amount_display,
        tax_applied, tax_amount_cents, net_amount_cents, rule_created, rule_keyword, warning
    """
    if target_account_name.upper() == 'EX.SUSP':
        raise ValueError("Cannot reclassify to EX.SUSP — that's a no-op.")

    txn, lines = get_transaction(txn_id)
    if not txn:
        raise ValueError(f"Transaction {txn_id} not found.")
    if len(lines) != 2:
        raise ValueError(
            f"Transaction has {len(lines)} lines — expected 2. "
            "Use update_transaction() directly for complex reclassifications.")

    # Find the EX.SUSP line and the bank line
    susp_line = bank_line = None
    for line in lines:
        if line['account_name'].upper() == 'EX.SUSP':
            susp_line = line
        else:
            bank_line = line
    if not susp_line:
        raise ValueError(f"Transaction {txn_id} has no EX.SUSP line — not a suspense transaction.")
    if not bank_line:
        raise ValueError(f"Transaction {txn_id} has no bank line.")

    # Resolve target account
    target_acct = get_account_by_name(target_account_name)
    if not target_acct:
        raise ValueError(f"Account '{target_account_name}' not found.")

    susp_amount = susp_line['amount']  # signed cents
    line_desc = susp_line['description'] or ''

    # Preserve bank line flags
    bank_tuple = (bank_line['account_id'], bank_line['amount'], bank_line['description'] or '',
                  bank_line['reconciled'], bank_line['doc_on_file'])

    # Build new lines
    tax_applied = ''
    tax_amount_cents = 0
    net_amount_cents = susp_amount

    if tax_code:
        tc = get_tax_code(tax_code)
        if tc and tc['rate_percent'] > 0:
            rate = tc['rate_percent']
            tax_cents = round(abs(susp_amount) * rate / (100 + rate))
            net_cents = abs(susp_amount) - tax_cents
            # Reapply sign
            sign = 1 if susp_amount > 0 else -1
            net_signed = net_cents * sign
            tax_signed = tax_cents * sign
            # Determine tax account by direction
            if bank_line['amount'] < 0:
                # Bank credit = money going out = expense = ITCs
                tax_acct_name = tc['paid_account'] or 'GST.IN'
            else:
                # Bank debit = money coming in = revenue = GST collected
                tax_acct_name = tc['collected_account'] or 'GST.OUT'
            tax_acct = get_account_by_name(tax_acct_name)
            if not tax_acct:
                raise ValueError(f"Tax account '{tax_acct_name}' not found.")

            new_lines = [
                bank_tuple,
                (target_acct['id'], net_signed, line_desc, 0, 0),
                (tax_acct['id'], tax_signed, line_desc, 0, 0),
            ]
            tax_applied = tax_code
            tax_amount_cents = tax_signed
            net_amount_cents = net_signed
        else:
            # Tax code exists but rate is 0 (exempt) — no split
            new_lines = [
                bank_tuple,
                (target_acct['id'], susp_amount, line_desc, 0, 0),
            ]
    else:
        new_lines = [
            bank_tuple,
            (target_acct['id'], susp_amount, line_desc, 0, 0),
        ]

    # Update transaction in place
    try:
        update_transaction(txn_id, txn['date'], txn['reference'] or '', txn['description'] or '', new_lines)
    except ValueError as e:
        if 'lock date' in str(e).lower():
            raise ValueError(f"Cannot reclassify: transaction is in a locked period.")
        raise

    # Auto-rule learning
    rule_created = False
    rule_keyword = ''
    txn_desc = txn['description'] or ''
    candidate = _extract_rule_keyword(txn_desc)
    if candidate and not _rule_already_exists(candidate):
        save_import_rule(None, candidate, target_account_name, tax_code, 0, 'auto')
        rule_created = True
        rule_keyword = candidate

    # Check if target account is on a report
    warning = ''
    rpt = find_report_for_account(target_acct['id'])
    if not rpt:
        warning = f"{target_account_name} is not on any report. Add it to IS so it appears on reports."

    return {
        'txn_id': txn_id,
        'old_account': 'EX.SUSP',
        'new_account': target_account_name,
        'amount_cents': susp_amount,
        'amount_display': f"{abs(susp_amount)/100:,.2f}",
        'tax_applied': tax_applied,
        'tax_amount_cents': tax_amount_cents,
        'net_amount_cents': net_amount_cents,
        'rule_created': rule_created,
        'rule_keyword': rule_keyword,
        'warning': warning,
    }


def batch_reclassify_suspense(items):
    """Reclassify multiple suspense transactions.

    Args:
        items: list of dicts, each with:
            - txn_id: int
            - target_account: str
            - tax_code: str (optional, default '')

    Returns:
        dict with: processed, failed, rules_created, results
    """
    processed = failed = rules_created = 0
    results = []
    for item in items:
        try:
            r = reclassify_suspense(
                item['txn_id'],
                item['target_account'],
                item.get('tax_code', '')
            )
            processed += 1
            if r['rule_created']:
                rules_created += 1
            results.append(r)
        except (ValueError, Exception) as e:
            failed += 1
            results.append({
                'txn_id': item.get('txn_id'),
                'error': str(e),
            })
    return {
        'processed': processed,
        'failed': failed,
        'rules_created': rules_created,
        'results': results,
    }

# ─── Formatting ───────────────────────────────────────────────────
def fmt_amount(cents):
    if cents == 0: return '—'
    neg = cents < 0; c = abs(cents)
    s = f"{c // 100:,}.{c % 100:02d}"
    return f"({s})" if neg else s

def fmt_amount_plain(cents):
    if cents == 0: return '0.00'
    neg = cents < 0; c = abs(cents)
    s = f"{c // 100:,}.{c % 100:02d}"
    return f"-{s}" if neg else s

def parse_amount(s):
    """Parse a money string to integer cents. Accepts commas, $, parentheses or a
    trailing '-' for negatives. Raises ValueError on malformed strings ('1.2.3',
    '--5') instead of silently guessing — an import surfaces that as a row error,
    which beats posting wrong money. 3+ decimals round half-up (1.005 → 1.01)."""
    s = str(s).strip().replace(',', '').replace('$', '').replace('−', '-')
    neg = False
    if s.startswith('(') and s.endswith(')'): neg = True; s = s[1:-1]
    if s.startswith('-'): neg = not neg; s = s[1:]
    if s.endswith('-'): neg = not neg; s = s[:-1]
    s = s.strip()
    if not s: return 0
    if not _re.match(r'^(?:\d+\.?\d*|\.\d+)$', s):
        raise ValueError(f"Not a valid amount: '{s}'")
    from decimal import Decimal, ROUND_HALF_UP
    cents = int((Decimal(s) * 100).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
    return -cents if neg else cents

def normalize_date(s):
    """Normalize a date string to YYYY-MM-DD. Handles OFX (YYYYMMDD), spelled-out
    months (16-Jul-26), spreadsheet timestamps (2026-07-16 00:00:00) and the
    common numeric formats. Returns None when the string is not a date —
    callers report that as a bad row, so guessing here is worse than failing."""
    s = str(s).strip().strip('"').strip("'")
    if not s: return None

    # Drop a trailing time component — xlsx/xls date cells and some bank
    # exports carry one ("2026-07-16 00:00:00", "7/16/2026 2:30 PM").
    m = _re.match(r'^(.*?)[ T]\d{1,2}:\d{2}(?::\d{2})?(?:\.\d+)?\s*(?:[AaPp]\.?[Mm]\.?)?$', s)
    if m and m.group(1).strip():
        s = m.group(1).strip()

    if len(s) == 10 and s[4] == '-' and s[7] == '-': return s

    # OFX format: YYYYMMDD or YYYYMMDDHHMMSS. Only when the leading 4 digits are
    # a plausible year — otherwise "06012025" (an 8-digit MMDDYYYY) used to be
    # sliced into the year 0601.
    if len(s) >= 8 and s[:8].isdigit() and 1950 <= int(s[:4]) <= 2099:
        try: return datetime(int(s[0:4]), int(s[4:6]), int(s[6:8])).strftime('%Y-%m-%d')
        except ValueError: return None

    for fmt_str in ('%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d',
                    '%m-%d-%Y', '%d-%m-%Y', '%b %d, %Y', '%B %d, %Y',
                    '%m/%d/%y', '%d/%m/%y', '%m-%d-%y', '%d-%m-%y',
                    # spelled-out months — Amex and most card exports use these
                    '%d-%b-%y', '%d-%b-%Y', '%d %b %Y', '%d %b %y',
                    '%d-%B-%y', '%d-%B-%Y', '%d %B %Y', '%d %B %y',
                    '%b %d %Y', '%B %d %Y', '%b-%d-%Y', '%b-%d-%y',
                    '%Y-%b-%d', '%Y %b %d'):
        try: return datetime.strptime(s, fmt_str).strftime('%Y-%m-%d')
        except ValueError: continue
    return None


def sniff_csv_columns(rows):
    """Find the date / description / amount columns in a bank export by looking at
    the DATA, not the header names. Returns (has_header, date_col, desc_col, amount_col)
    with -1 for anything not found.

    Header names are unreliable — a headerless export makes row 1 look like a
    header, and a description containing the word "Amount" makes a data row look
    like one. Content decides: a header row is simply a first row whose date
    column does not hold a date."""
    if not rows: return False, -1, -1, -1
    ncols = max(len(r) for r in rows[:51])

    def cells(ci, source):
        return [str(r[ci]).strip() for r in source if ci < len(r) and str(r[ci]).strip()]

    # Date column = most values that parse as a date, scored over rows 2+ so a
    # header row cannot influence the choice.
    body = rows[1:51] if len(rows) > 1 else rows[:1]
    date_col, best = -1, 0
    for ci in range(ncols):
        score = sum(1 for v in cells(ci, body) if normalize_date(v))
        if score > best: best, date_col = score, ci
    if date_col < 0: return True, -1, -1, -1

    row0 = rows[0]
    has_header = not (date_col < len(row0) and normalize_date(str(row0[date_col]).strip()))
    body = rows[1:51] if has_header else rows[:51]

    # Amount column: mostly numeric and not constant (rules out account/card
    # numbers). Prefer a column carrying negatives over a running balance.
    amount_col, amount_key = -1, None
    for ci in range(ncols):
        if ci == date_col: continue
        vals = cells(ci, body)
        if not vals: continue
        nums = []
        for v in vals:
            try: nums.append(float(v.replace(',', '').replace('$', '')))
            except ValueError: pass
        if len(nums) < len(vals) * 0.8 or len(set(nums)) < 2: continue
        key = (any(n < 0 for n in nums), any(abs(n) % 1 for n in nums), -ci)
        if amount_key is None or key > amount_key:
            amount_key, amount_col = key, ci

    # Description column: the wordiest remaining column.
    desc_col, best_len = -1, 0.0
    for ci in range(ncols):
        if ci in (date_col, amount_col): continue
        vals = cells(ci, body)
        if not vals: continue
        wordy = [v for v in vals if any(ch.isalpha() for ch in v)]
        if len(wordy) < len(vals) * 0.5: continue
        avg = sum(len(v) for v in wordy) / len(vals)
        if avg > best_len: best_len, desc_col = avg, ci

    return has_header, date_col, desc_col, amount_col

def _ofx_sgml_to_xml(content):
    """Convert OFX SGML to valid XML by closing unclosed tags."""
    import re
    # Container/aggregate tags that wrap children — do NOT self-close these
    aggregates = {
        'OFX', 'SIGNONMSGSRSV1', 'SONRS', 'STATUS', 'FI',
        'BANKMSGSRSV1', 'STMTTRNRS', 'STMTRS', 'BANKACCTFROM',
        'BANKTRANLIST', 'STMTTRN', 'LEDGERBAL', 'AVAILBAL',
        'CREDITCARDMSGSRSV1', 'CCSTMTTRNRS', 'CCSTMTRS', 'CCACCTFROM',
    }
    agg_lower = {t.lower() for t in aggregates}

    def close_tags(match):
        tag = match.group(1)
        value = match.group(2).strip()
        if tag.lower() in agg_lower:
            return match.group(0)  # leave aggregates alone
        return f"<{tag}>{value}</{tag}>"

    # Match <TAG>value where value is non-empty text (not starting with <)
    return re.sub(r'<([A-Za-z0-9_.]+)>([^<\r\n]+)', close_tags, content)

def parse_ofx(file_path):
    """Parse an OFX/QBO file and return a list of row dicts for import_rows().

    Each dict has: date, description, amount_cents, reference (FITID).
    Uses stdlib only (xml.etree.ElementTree).
    """
    import xml.etree.ElementTree as ET

    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()

    # Strip OFX headers (everything before <OFX>)
    idx = content.upper().find('<OFX>')
    if idx < 0:
        raise ValueError("Not a valid OFX file: no <OFX> tag found")
    content = content[idx:]

    # Try parsing as valid XML first, then fall back to SGML conversion
    root = None
    for attempt_content in [content, _ofx_sgml_to_xml(content)]:
        for suffix in ['', '</OFX>']:
            try:
                root = ET.fromstring(attempt_content + suffix)
                break
            except ET.ParseError:
                continue
        if root is not None:
            break
    if root is None:
        raise ValueError("Cannot parse OFX file: invalid XML/SGML structure")

    # Find all STMTTRN elements (works for both bank and credit card statements)
    transactions = root.iter('STMTTRN')

    rows = []
    for txn in transactions:
        dt_el = txn.find('DTPOSTED')
        amt_el = txn.find('TRNAMT')
        name_el = txn.find('NAME')
        memo_el = txn.find('MEMO')
        fitid_el = txn.find('FITID')

        if dt_el is None or amt_el is None:
            continue

        # Build description from NAME + MEMO
        name = (name_el.text or '').strip() if name_el is not None else ''
        memo = (memo_el.text or '').strip() if memo_el is not None else ''
        if memo and memo.lower() != name.lower():
            description = f"{name} — {memo}" if name else memo
        else:
            description = name or memo

        if not description:
            continue

        amount_cents = parse_amount(amt_el.text)
        fitid = (fitid_el.text or '').strip() if fitid_el is not None else ''

        rows.append({
            'date': (dt_el.text or '').strip(),
            'description': description,
            'amount_cents': amount_cents,
            'reference': fitid,
        })

    if not rows:
        raise ValueError("No transactions found in OFX file")

    return rows

def import_rows(bank_account_id, rows):
    """Shared posting loop for CSV and OFX imports.

    Args:
        bank_account_id: int — the bank account to post against
        rows: list of dicts with keys: date, description, amount_cents, reference (optional)

    Returns:
        dict with: rows_processed, posted, skipped, to_suspense, errors, possible_duplicates
    """
    posted = 0
    skipped = 0
    suspense = 0
    errors = []
    possible_duplicates = []
    lock = get_meta('lock_date', '')

    # Pre-scan for possible duplicates: existing transactions on the bank account.
    # Two detection methods:
    #   1. FITID (reference) — deterministic for OFX imports. Same FITID = same txn.
    #   2. Date + amount — soft match for CSV imports (warns but still posts).
    existing = set()
    existing_refs = set()
    fitid_skipped = 0
    with get_db() as db:
        for r in db.execute(
                "SELECT t.date, t.reference, l.amount FROM lines l "
                "JOIN transactions t ON l.transaction_id = t.id "
                "WHERE l.account_id = ?", (bank_account_id,)).fetchall():
            existing.add((r['date'], r['amount']))
            ref = r['reference']
            if ref and ref.strip():
                existing_refs.add(ref.strip())

    for row_num, row in enumerate(rows, start=1):
        row_date = normalize_date(row['date'])
        row_desc = row['description']
        amount_cents = row['amount_cents']
        reference = row.get('reference', '')

        if not row_desc:
            errors.append({'row': row_num, 'reason': 'Missing description'})
            skipped += 1
            continue

        if not row_date:
            errors.append({'row': row_num, 'reason': f"Bad date '{row['date']}'"})
            skipped += 1
            continue

        if lock and row_date <= lock:
            errors.append({'row': row_num, 'reason': f'Before lock date {lock}'})
            skipped += 1
            continue

        ceiling = fiscal_ceiling()
        if ceiling and row_date > ceiling:
            errors.append({'row': row_num, 'reason': f'After fiscal year end {ceiling}'})
            skipped += 1
            continue

        if amount_cents == 0:
            errors.append({'row': row_num, 'reason': 'Zero amount'})
            skipped += 1
            continue

        # FITID duplicate detection: if reference matches an existing transaction, skip entirely.
        # FITID is a unique bank-assigned ID — same FITID = guaranteed duplicate.
        if reference and reference.strip() and reference.strip() in existing_refs:
            fitid_skipped += 1
            skipped += 1
            continue

        # Soft duplicate detection: check if this date+amount already exists on the bank account
        if (row_date, amount_cents) in existing:
            possible_duplicates.append({
                'row': row_num, 'date': row_date,
                'amount': amount_cents, 'description': row_desc[:60]})

        matched_acct, tax_code, tax_info = apply_rules(row_desc, amount_cents)

        target_acct = get_account_by_name(matched_acct)
        if not target_acct:
            # A rule pointing at a dead account must not DROP the row — land it
            # in suspense (money always arrives; a human clears it), and say why.
            susp = get_account_by_name('EX.SUSP')
            if not susp:
                errors.append({'row': row_num, 'reason': f"Account '{matched_acct}' not found and no EX.SUSP to fall back to"})
                skipped += 1
                continue
            errors.append({'row': row_num, 'reason': f"Rule account '{matched_acct}' not found — posted to EX.SUSP instead"})
            target_acct = susp
            matched_acct = 'EX.SUSP'
            tax_info = None

        if matched_acct == 'EX.SUSP':
            suspense += 1

        try:
            if tax_info and tax_info.get('tax'):
                tax_acct = get_account_by_name(tax_info['tax_acct'])
                if not tax_acct:
                    add_simple_transaction(
                        row_date, reference, row_desc,
                        target_acct['id'] if amount_cents < 0 else bank_account_id,
                        bank_account_id if amount_cents < 0 else target_acct['id'],
                        abs(amount_cents))
                else:
                    net = tax_info['net']
                    tax = tax_info['tax']
                    if amount_cents < 0:
                        txn_lines = [
                            (target_acct['id'], net, row_desc),
                            (tax_acct['id'], tax, f"{tax_code} on {row_desc[:30]}"),
                            (bank_account_id, -(net + tax), row_desc),
                        ]
                    else:
                        txn_lines = [
                            (bank_account_id, net + tax, row_desc),
                            (target_acct['id'], -net, row_desc),
                            (tax_acct['id'], -tax, f"{tax_code} on {row_desc[:30]}"),
                        ]
                    add_transaction(row_date, reference, row_desc, txn_lines)
            else:
                if amount_cents < 0:
                    add_simple_transaction(
                        row_date, reference, row_desc,
                        target_acct['id'], bank_account_id, abs(amount_cents))
                else:
                    add_simple_transaction(
                        row_date, reference, row_desc,
                        bank_account_id, target_acct['id'], abs(amount_cents))
            posted += 1
            # Track for within-batch duplicate detection
            existing.add((row_date, amount_cents))
            if reference and reference.strip():
                existing_refs.add(reference.strip())
        except ValueError as e:
            errors.append({'row': row_num, 'reason': str(e)})
            skipped += 1

    result = {
        'rows_processed': len(rows),
        'posted': posted,
        'skipped': skipped,
        'fitid_skipped': fitid_skipped,
        'to_suspense': suspense,
        'errors': errors[:20] if errors else [],
        'possible_duplicates': possible_duplicates[:20] if possible_duplicates else [],
    }
    return result


def import_gl_rows(bank_account_id, rows):
    """Posting loop for pre-categorized general ledger imports.

    Like import_rows but each row specifies its cross-account directly —
    no rule matching, no suspense routing. Used when importing from another
    accounting system (a legacy GL, QuickBooks GL, Sage, etc.) where cross-accounts
    are already known.

    Args:
        bank_account_id: int — the primary account for these rows
        rows: list of dicts with keys: date, description, amount_cents, cross_account

    Returns:
        dict with: rows_processed, posted, skipped, errors, possible_duplicates
    """
    posted = 0
    skipped = 0
    errors = []
    possible_duplicates = []
    lock = get_meta('lock_date', '')

    # Pre-scan for duplicates on the primary account
    existing = set()
    with get_db() as db:
        for r in db.execute(
                "SELECT t.date, l.amount FROM lines l "
                "JOIN transactions t ON l.transaction_id = t.id "
                "WHERE l.account_id = ?", (bank_account_id,)).fetchall():
            existing.add((r['date'], r['amount']))

    for row_num, row in enumerate(rows, start=1):
        row_date = normalize_date(row['date'])
        row_desc = row['description']
        amount_cents = row['amount_cents']
        cross_account = row['cross_account']

        if not row_desc:
            errors.append({'row': row_num, 'reason': 'Missing description'})
            skipped += 1
            continue

        if not row_date:
            errors.append({'row': row_num, 'reason': f"Bad date '{row['date']}'"})
            skipped += 1
            continue

        if lock and row_date <= lock:
            errors.append({'row': row_num, 'reason': f'Before lock date {lock}'})
            skipped += 1
            continue

        ceiling = fiscal_ceiling()
        if ceiling and row_date > ceiling:
            errors.append({'row': row_num, 'reason': f'After fiscal year end {ceiling}'})
            skipped += 1
            continue

        if amount_cents == 0:
            errors.append({'row': row_num, 'reason': 'Zero amount'})
            skipped += 1
            continue

        if not cross_account:
            errors.append({'row': row_num, 'reason': 'Missing cross-account'})
            skipped += 1
            continue

        target_acct = get_account_by_name(cross_account)
        if not target_acct:
            errors.append({'row': row_num, 'reason': f"Cross-account '{cross_account}' not found"})
            skipped += 1
            continue

        # Soft duplicate detection
        if (row_date, amount_cents) in existing:
            possible_duplicates.append({
                'row': row_num, 'date': row_date,
                'amount': amount_cents, 'description': row_desc[:60]})

        try:
            if amount_cents < 0:
                add_simple_transaction(
                    row_date, '', row_desc,
                    target_acct['id'], bank_account_id, abs(amount_cents))
            else:
                add_simple_transaction(
                    row_date, '', row_desc,
                    bank_account_id, target_acct['id'], abs(amount_cents))
            posted += 1
            existing.add((row_date, amount_cents))
        except ValueError as e:
            errors.append({'row': row_num, 'reason': str(e)})
            skipped += 1

    result = {
        'rows_processed': len(rows),
        'posted': posted,
        'skipped': skipped,
        'errors': errors[:20] if errors else [],
        'possible_duplicates': possible_duplicates[:20] if possible_duplicates else [],
    }
    return result


# ─── Report Chain Validation ──────────────────────────────────────
# ─── Check Books: one question, one answer ─────────────────────────
# LAP made correctness VISIBLE — the GL-PROOF line sat permanently on the report
# showing 0.00, and the LAP file check verified it before you worked in it. GridTRX
# had both halves (the PROOF chip, quick_check at open) and no way to ASK. This
# is the ask: everything that can be wrong with a set of books, or outstanding
# in them, answered in one call for a person or an agent.
#
# Three states, and the difference matters:
#   error      the books are WRONG. Fix this before doing anything else.
#   attention  real work outstanding. Not a fault — a to-do.
#   ok         nothing to say.
CHECK_ERROR, CHECK_ATTENTION, CHECK_OK = 'error', 'attention', 'ok'


def check_books():
    """Everything that can be wrong with these books, or waiting in them.

    Returns {ok, errors, attention, summary, checks: [{name, status, detail}]}.
    Read-only: it opens nothing, posts nothing and heals nothing."""
    checks = []

    def add(name, status, detail=''):
        checks.append({'name': name, 'status': status, 'detail': detail})

    # 1. The file itself. Everything below is meaningless if this fails.
    try:
        with get_db() as db:
            integrity = db.execute('PRAGMA quick_check').fetchone()[0]
        add('File integrity', CHECK_OK if integrity == 'ok' else CHECK_ERROR,
            'sqlite quick_check: ok' if integrity == 'ok' else str(integrity))
    except Exception as e:
        add('File integrity', CHECK_ERROR, str(e))

    # 2. The trial balance. Every transaction balances, so the whole file must
    #    sum to zero — this is the PROOF line, asked as a question.
    with get_db() as db:
        proof = db.execute('SELECT COALESCE(SUM(amount),0) FROM lines').fetchone()[0]
        txns = db.execute('SELECT COUNT(*) FROM transactions').fetchone()[0]
        unbal = db.execute("""SELECT t.id, t.date, t.description, SUM(l.amount) AS off
                              FROM transactions t JOIN lines l ON l.transaction_id=t.id
                              GROUP BY t.id HAVING off != 0 LIMIT 10""").fetchall()
    add('Trial balance ties', CHECK_OK if proof == 0 else CHECK_ERROR,
        'PROOF 0.00' if proof == 0 else f'out by {fmt_amount_plain(proof)}')
    if unbal:
        add('Every transaction balances', CHECK_ERROR,
            '; '.join(f"#{r['id']} {r['date']} {r['description'][:30]} off by "
                      f"{fmt_amount_plain(r['off'])}" for r in unbal))
    else:
        add('Every transaction balances', CHECK_OK, f'{txns} transactions')

    # 2b. Did opening this file REPAIR it? A silent repair to retained earnings
    #     is exactly the thing that must not be silent.
    for line in re_repair_note():
        add('Retained earnings repaired on open', CHECK_ATTENTION, line)

    # 3. Was this file ever SET UP? Empty books have nothing wrong with them,
    #    which is exactly the trap: an import that landed nothing leaves a file
    #    that passes every other check and prints a confident blank statement.
    #    "Sound" and "finished" are not the same word.
    have = {r['name'] for r in get_reports()}
    missing = [n for n in ('BS', 'IS') if n not in have]
    if missing:
        add('Set up as a client', CHECK_ERROR,
            f"no {' or '.join(missing)} report — this file has no chart of accounts. "
            f"It was opened, not created. Build it with create_starter_books "
            f"(MCP: create_books) rather than working in it.")
    elif txns == 0:
        add('Set up as a client', CHECK_ATTENTION,
            'the chart is in place but NOTHING has been posted — no opening '
            'balances and no transactions. Statements will print blank.')
    else:
        add('Set up as a client', CHECK_OK, f'{len(have)} reports, {txns} transactions')

    # 4. The report chain — do the statements actually add up to each other.
    try:
        # validate_report_chain reports its own clean bill of health as a
        # level 'ok' issue — that is a result, not something outstanding.
        LEVELS = {'error': CHECK_ERROR, 'warning': CHECK_ATTENTION, 'ok': CHECK_OK}
        for issue in validate_report_chain():
            add('Report chain', LEVELS.get(issue.get('level'), CHECK_ATTENTION),
                issue.get('message', ''))
    except Exception as e:
        add('Report chain', CHECK_ERROR, f'could not be validated: {e}')

    # 4. Suspense — the whole point of the import pipeline is that nothing is
    #    guessed; what it could not place is sitting here waiting for a human.
    susp = get_account_by_name('EX.SUSP')
    if susp:
        bal = get_account_balance(susp['id'])
        with get_db() as db:
            n = db.execute('SELECT COUNT(*) FROM lines WHERE account_id=?',
                           (susp['id'],)).fetchone()[0]
        add('Suspense cleared', CHECK_OK if bal == 0 else CHECK_ATTENTION,
            'nothing parked' if bal == 0
            else f'{fmt_amount_plain(bal)} across {n} lines still to classify')

    # 5. The posting window. Rows outside it are how a year quietly goes wrong.
    ceiling = fiscal_ceiling()
    if ceiling:
        after = transactions_after(ceiling)
        add('Nothing posted past the year-end', CHECK_OK if not after else CHECK_ATTENTION,
            f'ceiling {ceiling}' if not after
            else f'{after} transaction(s) dated after {ceiling}')
    lock = get_meta('lock_date', '')
    if lock:
        with get_db() as db:
            inside = db.execute('SELECT COUNT(*) FROM transactions WHERE date <= ?',
                                (lock,)).fetchone()[0]
        add('Locked period', CHECK_OK, f'{inside} transaction(s) locked on or before {lock}')

    # 6. Opening balances — new books with nothing brought forward.
    st = openings_state()
    add('Opening balances',
        CHECK_ATTENTION if st.get('status') == 'needed' else CHECK_OK,
        {'needed': 'no conversion posted — these books start from nothing',
         'posted': f"conversion posted {st.get('conversion_date') or ''}".strip(),
         'declined': 'client starts at zero (declined)',
         'later': 'has activity, no conversion on file'}.get(st.get('status'), str(st.get('status'))))

    # 7. Reconciliation continuity (LAP rule). An item dated inside a period
    #    you already reconciled, still open, was missed — the statement and the
    #    ledger have quietly stopped agreeing.
    with get_db() as db:
        rows = db.execute(f"""
            SELECT a.name,
                   MAX(CASE WHEN {REC_SQL} THEN t.date END) AS last_rec,
                   SUM(CASE WHEN NOT {REC_SQL} THEN 1 ELSE 0 END) AS open_items
            FROM lines l
            JOIN accounts a ON a.id = l.account_id
            JOIN transactions t ON t.id = l.transaction_id
            GROUP BY a.id HAVING last_rec IS NOT NULL""").fetchall()
        missed = []
        for r in rows:
            n = db.execute(f"""SELECT COUNT(*) FROM lines l
                               JOIN transactions t ON t.id=l.transaction_id
                               JOIN accounts a ON a.id=l.account_id
                               WHERE a.name=? AND t.date <= ? AND NOT {REC_SQL}""",
                           (r['name'], r['last_rec'])).fetchone()[0]
            if n:
                missed.append(f"{r['name']}: {n} item(s) on or before {r['last_rec']}")
    if rows:
        add('Reconciliation continuity', CHECK_OK if not missed else CHECK_ATTENTION,
            f'{len(rows)} account(s) reconciled' if not missed
            else 'unreconciled items inside a reconciled period — ' + '; '.join(missed))

    # 8. Accounts on no report cannot appear on a statement. Usually harmless
    #    leftovers, occasionally a real account nobody can see.
    on_report = accounts_on_any_report()          # ids, not names
    with get_db() as db:
        orphans = [r['name'] for r in db.execute(
            """SELECT a.id, a.name FROM accounts a
               WHERE a.account_type!='total'
                 AND EXISTS (SELECT 1 FROM lines l WHERE l.account_id = a.id)""").fetchall()
            if r['id'] not in on_report]
    add('Every account with postings is on a report',
        CHECK_OK if not orphans else CHECK_ATTENTION,
        'yes' if not orphans else 'off every statement: ' + ', '.join(sorted(orphans)[:10]))

    # 9. Today's snapshot.
    try:
        bs = backup_status()
        add('Snapshot on file', CHECK_ERROR if bs.get('error') else CHECK_OK,
            bs.get('note', ''))
    except Exception as e:
        add('Snapshot on file', CHECK_ERROR, str(e))

    errors = sum(1 for c in checks if c['status'] == CHECK_ERROR)
    attention = sum(1 for c in checks if c['status'] == CHECK_ATTENTION)
    if errors:
        summary = f"{errors} thing(s) WRONG with these books" + (
            f", {attention} outstanding" if attention else "")
    elif attention:
        summary = f"The books are sound. {attention} thing(s) outstanding."
    else:
        summary = "The books are sound and there is nothing outstanding."
    return {'ok': errors == 0, 'errors': errors, 'attention': attention,
            'summary': summary, 'checks': checks}


def validate_report_chain():
    """Validate the total-to chain across all reports.

    Checks:
      1. Every total report item has an account_id linked
      2. IS net-income chain reaches the perpetual RE on the BS (IS totals → NI → RE)
      3. RE is a total accumulator; RE.OB (opening) totals into it; RE.OPEN/RE.CLOSE
         are computed display-only lines (open:RE / close:RE)
      4. BS balances (Total Assets = Total L&E) and BS RE == IS Closing RE
      5. No orphan total accounts (accounts referenced by total_to but not on any report)

    Returns: list of {level: 'error'|'warning', message: str}
    """
    issues = []

    all_items = get_all_report_items()
    if not all_items:
        issues.append({'level': 'error', 'message': 'No report items found. Run create_starter_books() first.'})
        return issues

    # 1. Total items without account_id
    for it in all_items:
        if it['item_type'] == 'total' and not it['account_id']:
            rpt = it['report_id']
            issues.append({'level': 'error',
                'message': f"Total item '{it['description']}' (report {rpt}) has no account_id linked. "
                           f"It will always show zero."})

    # 2. Collect the total-to graph
    tt_fields = ['total_to_1','total_to_2','total_to_3','total_to_4','total_to_5','total_to_6']
    # Map: account_name -> set of target names
    feeds_into = {}
    all_acct_names = set()
    for it in all_items:
        name = it['acct_name'] if 'acct_name' in it.keys() else None
        if not name:
            continue
        all_acct_names.add(name)
        for ttf in tt_fields:
            target = it[ttf] if ttf in it.keys() else ''
            if target:
                feeds_into.setdefault(name, set()).add(target)
                all_acct_names.add(target)

    # 3. Check IS net-income chain reaches the perpetual RE on the BS.
    #    LAP-style model: NI/DIVPAID and the single opening account RE.OB total
    #    straight into RE; RE.OPEN/RE.CLOSE are computed display-only lines.
    bs_report = find_report_by_name('BS')
    is_report = find_report_by_name('IS')

    # RE total accumulator on the BS
    re_acct = get_account_by_name('RE')
    if not re_acct:
        issues.append({'level': 'error',
            'message': "RE total account missing. Retained earnings has no accumulator on BS."})
    elif re_acct['account_type'] != 'total':
        issues.append({'level': 'warning',
            'message': "RE is a posting account, not a total. It won't accumulate the IS net-income chain."})

    # Opening retained earnings: single posting account RE.OB totalling into RE
    re_ob = get_account_by_name('RE.OB')
    if re_ob and 'RE' not in feeds_into.get('RE.OB', set()):
        issues.append({'level': 'warning',
            'message': "RE.OB (opening retained earnings) does not total-to RE. Opening RE won't reach the BS."})

    # RE.OPEN / RE.CLOSE should be computed display lines, not posting accounts
    for nm in ('RE.OPEN', 'RE.CLOSE'):
        a = get_account_by_name(nm)
        if a and a['account_type'] == 'posting':
            issues.append({'level': 'warning',
                'message': f"{nm} is a posting account; expected a computed open:/close:RE display "
                           f"line. Open the file once to run the RE migration."})

    # NI net-income accumulator must exist
    ni_acct = get_account_by_name('NI') or get_account_by_name('NETINC')
    if not ni_acct:
        issues.append({'level': 'error',
            'message': "NI (or NETINC) total account missing. No net income accumulator."})

    # Check that at least one IS total feeds through to RE on the BS
    if is_report:
        is_items = get_report_items(is_report['id'])
        is_totals = [it['acct_name'] for it in is_items
                     if it['item_type'] == 'total' and it['acct_name']]
        # Trace from each IS total to see if it reaches RE
        def reaches(name, target, visited=None):
            if visited is None:
                visited = set()
            if name in visited:
                return False
            visited.add(name)
            targets = feeds_into.get(name, set())
            if target in targets:
                return True
            return any(reaches(t, target, visited) for t in targets)

        chain_ok = any(reaches(t, 'RE') for t in is_totals)
        if not chain_ok:
            issues.append({'level': 'error',
                'message': "IS total-to chain does not reach RE on BS. "
                           "Net income won't flow to equity. Need: IS totals → NI → RE."})

    # 4. Check BS balances
    if bs_report:
        data = compute_report_column(bs_report['id'])
        ta_bal = tle_bal = None
        for item, bal in data:
            name = item.get('acct_name')
            if name == 'TA':
                ta_bal = bal
            elif name == 'TLE' or name == 'TL':
                tle_bal = bal
        if ta_bal is not None and tle_bal is not None:
            diff = ta_bal - tle_bal
            if diff != 0:
                issues.append({'level': 'error',
                    'message': f"BS out of balance: Total Assets ({fmt_amount(ta_bal)}) != "
                               f"Liabilities & Equity ({fmt_amount(tle_bal)}). "
                               f"Difference: {fmt_amount(diff)}."})
        elif ta_bal is None:
            issues.append({'level': 'warning',
                'message': "Cannot find TA (Total Assets) on BS to verify balance."})

    # 5. BS RE vs IS RE.CLOSE cross-check
    #    Computes both values for the current FY and flags any discrepancy.
    #    This catches duplicate transactions, chain wiring mismatches, and
    #    conversion artifacts — the single most valuable data-integrity check.
    if bs_report and is_report and re_acct:
        try:
            # FY dates come from fiscal_anchor() — the single source of truth,
            # which already clamps oddities like a Feb-29 year end (the old
            # hand-rolled math here raised on those and silently disabled
            # this whole cross-check).
            anchor = fiscal_anchor()
            if anchor:
                fy_start, fy_end = anchor['cy_start'], anchor['cy_end']

                # BS RE: all-time up to FY end
                bs_data = compute_report_column(bs_report['id'], date_to=fy_end)
                bs_re_val = None
                for item, bal in bs_data:
                    if item.get('acct_name') == 'RE':
                        bs_re_val = bal
                        break

                # IS RE.CLOSE: for the FY period
                is_data = compute_report_column(is_report['id'],
                                                date_from=fy_start, date_to=fy_end)
                is_reclose_val = None
                # Try RE.CLOSE first, fall back to any account containing 'RE' at IS tail
                for item, bal in is_data:
                    if item.get('acct_name') == 'RE.CLOSE':
                        is_reclose_val = bal
                        break
                if is_reclose_val is None:
                    for item, bal in is_data:
                        if item.get('acct_name') == 'RE':
                            is_reclose_val = bal
                            break

                if bs_re_val is not None and is_reclose_val is not None:
                    diff = bs_re_val - is_reclose_val
                    if diff != 0:
                        issues.append({'level': 'error',
                            'message': f"BS Retained Earnings ({fmt_amount(bs_re_val)}) != "
                                       f"IS Closing RE ({fmt_amount(is_reclose_val)}) "
                                       f"for FY ending {fy_end}. "
                                       f"Difference: {fmt_amount(diff)}. "
                                       f"Check for duplicate transactions or chain wiring issues."})
                elif bs_re_val is None:
                    issues.append({'level': 'warning',
                        'message': "Cannot find RE on BS to cross-check against IS."})
                elif is_reclose_val is None:
                    issues.append({'level': 'warning',
                        'message': "Cannot find RE.CLOSE (or RE) on IS to cross-check against BS."})
        except (ValueError, IndexError, TypeError):
            pass  # FY dates not configured — skip this check

    # 6. Orphan total-to targets (referenced but not on any report)
    on_report = set()
    for it in all_items:
        if it['acct_name']:
            on_report.add(it['acct_name'])
    for name, targets in feeds_into.items():
        for t in targets:
            if t not in on_report:
                acct = get_account_by_name(t)
                if not acct:
                    issues.append({'level': 'error',
                        'message': f"'{name}' total-to '{t}' but account '{t}' does not exist."})
                else:
                    issues.append({'level': 'warning',
                        'message': f"'{name}' total-to '{t}' but '{t}' is not on any report. "
                                   f"Balance will accumulate but never display."})

    # 7. Orphan posting accounts — non-zero balance but not on any report
    all_balances = get_all_account_balances()
    accounts_on_reports = set()
    with get_db() as db:
        rows = db.execute(
            "SELECT DISTINCT account_id FROM report_items WHERE account_id IS NOT NULL"
        ).fetchall()
        for r in rows:
            accounts_on_reports.add(r['account_id'])
    all_accts = get_accounts()
    for acct in all_accts:
        if acct['account_type'] != 'posting':
            continue
        bal = all_balances.get(acct['id'], 0)
        if bal != 0 and acct['id'] not in accounts_on_reports:
            issues.append({'level': 'warning',
                'message': f"Account '{acct['name']}' has a non-zero balance but is not on any report. "
                           f"It will not appear on BS or IS."})

    if not issues:
        issues.append({'level': 'ok', 'message': 'Report chain is valid. BS balances.'})

    return issues


# ─── Export to write-up (Willy) ────────────────────────────────────
# The handoff to the working-paper program: ONE JSON-ready dict, frozen-
# snapshot discipline — Willy imports it and never reaches back into these
# books. Contents follow the operator's spec (2026-08-26): TWO full years
# itemized (CY + PY, so the agent compares this year's postings against last
# year's) + FIVE fiscal years of per-account comparatives (materiality /
# comparative testing) + the chart + the statement layouts + the check_books
# verdict Willy's front-door gate reads. Nothing is silently excluded: the
# meta block counts what fell outside the itemized window.

def export_writeup():
    """Build the write-up handoff payload. Raises if no fiscal year is set —
    the export is anchored to the book's own year, never wall-clock."""
    anchor = fiscal_anchor()
    if not anchor:
        raise ValueError(
            "Set the fiscal year first (Settings) — the write-up export is "
            "anchored to the book's working year-end.")
    m, d = _fye_md()
    fy = anchor['fy']
    company = get_meta('company_name', 'My Books')

    with get_db() as db:
        accounts = [{
            'name': a['name'],
            'description': a['description'] or '',
            'normal_balance': a['normal_balance'],
            'account_type': a['account_type'],
            'account_number': a['account_number'] or '',
            'leadsheet': (a['leadsheet'] if 'leadsheet' in a.keys() else '') or '',
            'computed': a['computed'] if 'computed' in a.keys() else '',
            'system': a['system'] if 'system' in a.keys() else 0,
        } for a in db.execute(
            "SELECT * FROM accounts ORDER BY name").fetchall()]
        acct_name = {a['id']: a['name'] for a in
                     db.execute("SELECT id, name FROM accounts").fetchall()}

        # CY + PY, fully itemized. Window = [py_start, cy_end]; anything
        # outside is COUNTED so the reader knows it exists.
        txns = []
        for t in db.execute(
                "SELECT * FROM transactions WHERE date >= ? AND date <= ? "
                "ORDER BY date, id",
                (anchor['py_start'], anchor['cy_end'])).fetchall():
            lines = [{
                'account_name': acct_name.get(ln['account_id'],
                                              f"UNKNOWN_{ln['account_id']}"),
                'amount': ln['amount'],
                'description': ln['description'] or '',
                'reconciled': ln['reconciled'],
            } for ln in db.execute(
                "SELECT * FROM lines WHERE transaction_id=? ORDER BY sort_order",
                (t['id'],)).fetchall()]
            txns.append({
                'date': t['date'],
                'reference': t['reference'] or '',
                'description': t['description'] or '',
                'fiscal_year': fy if t['date'] > anchor['py_end'] else fy - 1,
                'import_batch': (t['import_batch']
                                 if 'import_batch' in t.keys() else '') or '',
                'lines': lines,
            })
        n_before = db.execute(
            "SELECT COUNT(*) FROM transactions WHERE date < ?",
            (anchor['py_start'],)).fetchone()[0]
        n_after = db.execute(
            "SELECT COUNT(*) FROM transactions WHERE date > ?",
            (anchor['cy_end'],)).fetchone()[0]

        # Five fiscal years of comparatives: per-account balance AT each
        # year-end (the BS view) and activity WITHIN each year (the IS view).
        # Willy derives whatever it wants from the pair; nothing else needed.
        comparatives = []
        for k in range(5):
            y = fy - k
            ye = year_end_on(y, m, d).isoformat()
            ys = (year_end_on(y - 1, m, d) + timedelta(days=1)).isoformat()
            bal = {acct_name[r['account_id']]: r['s'] for r in db.execute(
                "SELECT l.account_id, SUM(l.amount) AS s FROM lines l "
                "JOIN transactions t ON t.id = l.transaction_id "
                "WHERE t.date <= ? GROUP BY l.account_id HAVING s != 0",
                (ye,)).fetchall() if r['account_id'] in acct_name}
            act = {acct_name[r['account_id']]: r['s'] for r in db.execute(
                "SELECT l.account_id, SUM(l.amount) AS s FROM lines l "
                "JOIN transactions t ON t.id = l.transaction_id "
                "WHERE t.date >= ? AND t.date <= ? GROUP BY l.account_id "
                "HAVING s != 0", (ys, ye)).fetchall()
                if r['account_id'] in acct_name}
            comparatives.append({'fiscal_year': y, 'year_end': ye,
                                 'balances': bal, 'activity': act})

        # Statement layouts: every report, rows in position order.
        reports = []
        for rep in db.execute(
                "SELECT * FROM reports ORDER BY sort_order, id").fetchall():
            items = [{
                'type': ri['item_type'],
                'description': ri['description'] or '',
                'account_name': acct_name.get(ri['account_id'], ''),
                'indent': ri['indent'],
                'ref_mark': (ri['ref_mark']
                             if 'ref_mark' in ri.keys() else '') or '',
                'total_tos': [ri[f'total_to_{i}'] or ''
                              for i in range(1, 7)
                              if (f'total_to_{i}' in ri.keys())],
            } for ri in db.execute(
                "SELECT * FROM report_items WHERE report_id=? ORDER BY position",
                (rep['id'],)).fetchall()]
            reports.append({'name': rep['name'],
                            'description': rep['description'] or '',
                            'items': items})

    return {
        '_grid_export': 'writeup',
        '_version': 1,
        '_exported': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'meta': {
            'company_name': company,
            'fiscal_year': fy,
            'fiscal_year_end': get_meta('fiscal_year_end', '') or '12-31',
            'lock_date': get_meta('lock_date', ''),
            'cy_start': anchor['cy_start'], 'cy_end': anchor['cy_end'],
            'py_start': anchor['py_start'], 'py_end': anchor['py_end'],
            'itemized_transactions': len(txns),
            'transactions_before_window': n_before,
            'transactions_after_window': n_after,
        },
        'accounts': accounts,
        'transactions': txns,
        'comparatives': comparatives,
        'reports': reports,
        'check': check_books(),
    }


# ─── Starter Books ───────────────────────────────────────────────






# ─── Starter Template ─────────────────────────────────────────────
# ─── The three shapes a set of books comes in ──────────────────────
# The first question anyone asks opening a new client: what KIND of business is
# this? It decides one thing and one thing only — how the top of the income
# statement is built, and what current assets go with it.
#
#   services      no cost of sales at all. Revenue goes straight to expenses;
#                 there is no gross profit line because there is no gross profit.
#   retail        goods bought and resold: purchases, freight-in, and the
#                 inventory adjustment, against Inventory on the balance sheet.
#   construction  jobs: materials, subcontractors, direct labour, equipment,
#                 against Work in Progress and holdbacks both ways.
#
# Everything below the gross profit line is the same in all three. This is NOT a
# template SYSTEM — it is three shapes, because those are the three the firm
# actually takes on.
BUSINESS_TYPES = ('services', 'retail', 'construction')

BUSINESS_TYPE_NOTE = {
    'services': 'No cost of sales — revenue less operating expenses.',
    'retail': 'Goods bought and resold — purchases, freight-in and the inventory '
              'adjustment give gross profit, against Inventory on the balance sheet.',
    'construction': 'Jobs — materials, subcontract, direct labour and equipment give '
                    'gross profit, against Work in Progress and holdbacks.',
}


def client_folder_name(legal_name, fiscal_year):
    """THE house name for a new client folder: LegalNameCompact-YYYYgx.

    The operator keys the legal name and the fiscal year; the folder name is
    DERIVED — never typed. Letters and digits only (spaces and punctuation are
    trouble on Linux and in shells); 'gx' marks the folder as a GridTRX
    client. 'Meridian Cabinets Ltd' + 2026 ->
    'MeridianCabinetsLtd-2026gx'."""
    compact = _re.sub(r'[^A-Za-z0-9]', '', legal_name or '')
    if not compact:
        raise ValueError('The legal name needs at least one letter or digit.')
    year = str(fiscal_year).strip()
    if not _re.fullmatch(r'\d{4}', year):
        raise ValueError('Give the fiscal year as YYYY.')
    return f'{compact}-{year}gx'


def create_starter_books(path, company_name='My Company', fiscal_ye='12-31',
                         business_type='services', working_year=None):
    init_db(path)
    set_meta('company_name', company_name)
    set_meta('fiscal_year_end', fiscal_ye)
    business_type = (business_type or 'services').strip().lower()
    if business_type not in BUSINESS_TYPES:
        raise ValueError(
            f"business_type must be one of {', '.join(BUSINESS_TYPES)} — got "
            f"'{business_type}'. {' '.join(f'{k}: {v}' for k, v in BUSINESS_TYPE_NOTE.items())}")
    set_meta('business_type', business_type)

    # The year-end being worked on = the most recent one that has PASSED. A file
    # opened in August for a 31-May client is that May's engagement, not next
    # May's. Posting stops at that year-end until the operator opens the next.
    if working_year:
        # The operator SAID which year-end is being worked on — the same year
        # that names the client folder. The books agree with the label.
        fy_end = year_end_on(int(working_year), int(fiscal_ye.split('-')[0]),
                             int(fiscal_ye.split('-')[1]))
    else:
        today = date.today()
        fy_mm, fy_dd = int(fiscal_ye.split('-')[0]), int(fiscal_ye.split('-')[1])
        fy_end = year_end_on(today.year, fy_mm, fy_dd)
        if fy_end > today:
            fy_end = year_end_on(today.year - 1, fy_mm, fy_dd)
    set_meta('fiscal_year', str(fy_end.year))
    set_meta('fy_ceiling_mode', 'cy')

    bs = add_report('BS', 'Balance Sheet', 10)
    is_ = add_report('IS', 'Income Statement', 20)
    aje = add_report('AJE', 'Adjusting Journal Entries', 30)
    trx = add_report('TRX', TRX_REPORT_DESC, 40)

    a = {}
    def ac(name, bal, desc, atype='posting'):
        a[name] = add_account(name, bal, desc, atype)

    # ── BS accounts ──
    ac('CASH','D','Petty Cash'); ac('BANK.CHQ','D','Bank - Chequing'); ac('BANK.SAV','D','Bank - Savings')
    ac('CLR.TSF','D','Clearing - Account Transfers')
    ac('TOTBANK','D','Total Bank Accounts','total')
    ac('AR','D','Accounts Receivable'); ac('AR.TOT','D','Total AR','total')
    ac('PREPAIDS','D','Prepaid Expenses'); ac('DEP','D','Deposits')
    if business_type == 'retail':
        ac('INVENTORY','D','Inventory')
    elif business_type == 'construction':
        ac('WIP','D','Work in Progress')
        ac('HOLDBACK.AR','D','Holdbacks Receivable')
        ac('HOLDBACK.AP','C','Holdbacks Payable')
    ac('CA','D','Total Current Assets','total')
    ac('EQUIP','D','Equipment'); ac('FURN','D','Furniture'); ac('COMP','D','Computer Equipment')
    ac('TOTFA','D','Total Capital Assets','total')
    ac('EQUIP.DEP','C','Accum Amort - Equipment'); ac('FURN.DEP','C','Accum Amort - Furniture'); ac('COMP.DEP','C','Accum Amort - Computer')
    ac('TOTDEP','C','Total Accum Amortization','total')
    ac('NETFA','D','Net Capital Assets','total')
    ac('TA','D','TOTAL ASSETS','total')
    ac('AP','C','Accounts Payable'); ac('AP.CC','C','Credit Card Payable')
    ac('GST.OUT','C','GST Collected'); ac('GST.IN','D','GST Paid (ITCs)')
    ac('GST.REMIT','C','GST Remittance'); ac('GST.PAY','C','GST Payable')
    ac('TOTGST','C','Total GST','total')
    ac('FEDTAX','C','Federal Tax Payable'); ac('PROTAX','C','Provincial Tax Payable')
    ac('TOT.TAX','C','Total Tax Payable','total')
    ac('CL','C','Total Current Liabilities','total')
    ac('LOAN','C','Bank Loan'); ac('SH.LOAN','C','Shareholder Loan'); ac('TOTTERM','C','Total LT Debt','total')
    ac('LTL','C','Total Long-Term Liabilities','total')
    ac('CAPITAL','C','Share Capital'); ac('RE','C','Retained Earnings','total')
    ac('EQ','C','Total Equity','total')
    ac('TL','C','TOTAL LIABILITIES & EQUITY','total')

    # ── IS accounts ──
    ac('REV','C','Revenue - Sales'); ac('REV.SVC','C','Revenue - Services')
    ac('TOTREV','C','Total Revenue','total')
    # Cost of sales — the part that depends on what the business DOES.
    COST_ACCOUNTS = {
        'services': [],
        'retail': [('CS.PURCH','Purchases'), ('CS.FREIGHT','Freight-in'),
                   ('CS.INVADJ','Inventory Adjustment')],
        'construction': [('CS.MAT','Materials'), ('CS.SUB','Subcontractors'),
                         ('CS.LAB','Direct Labour'), ('CS.EQUIP','Equipment & Rentals'),
                         ('CS.WIPADJ','Work in Progress Adjustment')],
    }[business_type]
    for cname, cdesc in COST_ACCOUNTS:
        ac(cname, 'D', f'Cost of Sales - {cdesc}')
    if COST_ACCOUNTS:
        ac('GROSS','C','Gross Profit','total')
    ac('EX.SAL','D','Salaries & Wages'); ac('EX.RENT','D','Rent'); ac('EX.OFFICE','D','Office & General')
    ac('EX.COMP','D','Computer & IT'); ac('EX.ADV','D','Advertising'); ac('EX.INS','D','Insurance')
    ac('EX.PHONE','D','Telephone'); ac('EX.TRAVEL','D','Travel'); ac('EX.MEALS','D','Meals & Entertainment')
    ac('EX.AUTO','D','Vehicle'); ac('EX.POST','D','Postage & Courier'); ac('EX.FEES','D','Professional Fees')
    ac('EX.SC','D','Service Charges'); ac('EX.AMORT','D','Amortization'); ac('EX.SUSP','D','Suspense')
    ac('TOTEX','D','Total Operating Expenses','total')
    ac('OPINC','C','Operating Income','total')
    ac('EX.LIFE','D','Life Insurance'); ac('EX.LTINT','D','Interest on LT Debt'); ac('EX.INTAX','D','Income Tax Expense')
    ac('TAXINC','C','Income Before Taxes','total'); ac('NETINC','C','Net Income','total')
    ac('NI','C','Net Income for Year','total')
    ac('RE.OPEN','C','Retained Earnings - Open'); ac('DIVPAID','C','Dividends Paid')
    ac('RE.CLOSE','C','Retained Earnings - Close','total')

    # ── BS report items ──
    p = [0]
    def bi(itype, desc='', an=None, ind=0, tt1='', sep=''):
        p[0] += 10
        add_report_item(bs, itype, desc, a.get(an), ind, p[0], tt1, sep_style=sep)

    bi('label','CURRENT ASSETS')
    bi('label','Bank Accounts:')
    bi('account','','CASH',2,'TOTBANK'); bi('account','','BANK.CHQ',2,'TOTBANK'); bi('account','','BANK.SAV',2,'TOTBANK')
    bi('account','','CLR.TSF',2,'TOTBANK')
    bi('separator',sep='single'); bi('total','','TOTBANK',3,'CA')
    bi('label','')
    bi('label','Accounts Receivable:')
    bi('account','','AR',2,'AR.TOT')
    bi('separator',sep='single'); bi('total','','AR.TOT',3,'CA')
    bi('label','')
    if business_type == 'retail':
        bi('label',''); bi('label','Inventory:')
        bi('account','','INVENTORY',2,'CA')
    elif business_type == 'construction':
        bi('label',''); bi('label','Work in Progress:')
        bi('account','','WIP',2,'CA'); bi('account','','HOLDBACK.AR',2,'CA')
    bi('label','')
    bi('label','Other Current Assets:')
    bi('account','','PREPAIDS',2,'CA'); bi('account','','DEP',2,'CA')
    bi('separator',sep='single'); bi('total','Total Current Assets','CA',3,'TA')
    bi('separator',sep='single'); bi('label','')
    bi('label','Capital Assets')
    bi('account','','EQUIP',2,'TOTFA'); bi('account','','FURN',2,'TOTFA'); bi('account','','COMP',2,'TOTFA')
    bi('separator',sep='single'); bi('total','','TOTFA',3,'NETFA')
    bi('label','')
    bi('label','Accumulated Amortization')
    bi('account','','EQUIP.DEP',2,'TOTDEP'); bi('account','','FURN.DEP',2,'TOTDEP'); bi('account','','COMP.DEP',2,'TOTDEP')
    bi('separator',sep='single'); bi('total','','TOTDEP',3,'NETFA')
    bi('separator',sep='single'); bi('total','Net Capital Assets','NETFA',3,'TA')
    bi('separator',sep='single'); bi('label','')
    bi('total','TOTAL ASSETS','TA',0); bi('separator',sep='double'); bi('label','')
    bi('label','CURRENT LIABILITIES')
    bi('account','','AP',2,'CL'); bi('account','','AP.CC',2,'CL')
    if business_type == 'construction':
        bi('account','','HOLDBACK.AP',2,'CL')
    bi('label',''); bi('label','GST:')
    bi('account','','GST.OUT',2,'TOTGST'); bi('account','','GST.IN',2,'TOTGST')
    bi('account','','GST.REMIT',2,'TOTGST'); bi('account','','GST.PAY',2,'TOTGST')
    bi('separator',sep='single'); bi('total','','TOTGST',3,'CL')
    bi('label','')
    bi('account','','FEDTAX',2,'TOT.TAX'); bi('account','','PROTAX',2,'TOT.TAX')
    bi('separator',sep='single'); bi('total','','TOT.TAX',3,'CL')
    bi('separator',sep='single'); bi('total','Total Current Liabilities','CL',3,'TL')
    bi('separator',sep='single'); bi('label','')
    bi('label','Long-Term Liabilities')
    bi('account','','LOAN',2,'TOTTERM')
    bi('account','','SH.LOAN',2,'TOTTERM')
    bi('separator',sep='single'); bi('total','','TOTTERM',3,'LTL')
    bi('total','Total Long-Term Liabilities','LTL',3,'TL')
    bi('separator',sep='single'); bi('label','')
    bi('label','Equity')
    bi('account','','CAPITAL',2,'EQ'); bi('account','','RE',2,'EQ')
    bi('separator',sep='single'); bi('total','Total Equity','EQ',3,'TL')
    bi('separator',sep='single'); bi('label','')
    bi('total','TOTAL LIABILITIES & EQUITY','TL',0); bi('separator',sep='double')

    # ── IS report items ──
    p[0] = 0
    def ii(itype, desc='', an=None, ind=0, tt1='', sep=''):
        p[0] += 10
        add_report_item(is_, itype, desc, a.get(an), ind, p[0], tt1, sep_style=sep)

    ii('label','REVENUE')
    ii('account','','REV',2,'TOTREV'); ii('account','','REV.SVC',2,'TOTREV')
    # With no cost of sales there is no gross profit line to total INTO, so
    # revenue totals straight to operating income. The arithmetic must follow
    # the shape of the business, not be papered over with an empty section.
    if COST_ACCOUNTS:
        ii('separator',sep='single'); ii('total','Total Revenue','TOTREV',3,'GROSS'); ii('label','')
        ii('label','COST OF SALES')
        for cname, _ in COST_ACCOUNTS:
            ii('account','',cname,2,'GROSS')
        ii('separator',sep='single'); ii('label','')
        ii('total','Gross Profit','GROSS',3,'OPINC'); ii('separator',sep='single'); ii('label','')
    else:
        ii('separator',sep='single'); ii('total','Total Revenue','TOTREV',3,'OPINC')
        ii('separator',sep='single'); ii('label','')
    ii('label','EXPENSES')
    ii('account','','EX.SAL',2,'TOTEX'); ii('account','','EX.RENT',2,'TOTEX'); ii('account','','EX.OFFICE',2,'TOTEX')
    ii('account','','EX.COMP',2,'TOTEX'); ii('account','','EX.ADV',2,'TOTEX'); ii('account','','EX.INS',2,'TOTEX')
    ii('account','','EX.PHONE',2,'TOTEX'); ii('account','','EX.TRAVEL',2,'TOTEX'); ii('account','','EX.MEALS',2,'TOTEX')
    ii('account','','EX.AUTO',2,'TOTEX'); ii('account','','EX.POST',2,'TOTEX'); ii('account','','EX.FEES',2,'TOTEX')
    ii('account','','EX.SC',2,'TOTEX'); ii('account','','EX.AMORT',2,'TOTEX'); ii('account','','EX.SUSP',2,'TOTEX')
    ii('separator',sep='single'); ii('total','Total Operating Expenses','TOTEX',3,'OPINC')
    ii('separator',sep='single'); ii('label','')
    ii('total','Operating Income','OPINC',3,'TAXINC'); ii('separator',sep='single'); ii('label','')
    ii('label','Other Items:')
    ii('account','','EX.LIFE',2,'TAXINC'); ii('account','','EX.LTINT',2,'TAXINC')
    ii('separator',sep='single')
    ii('total','Income Before Taxes','TAXINC',3,'NETINC'); ii('label','')
    ii('account','','EX.INTAX',2,'NETINC')
    ii('separator',sep='single'); ii('total','Net Income (Loss)','NETINC',3,'NI'); ii('separator',sep='double')
    ii('label','')
    ii('account','Retained Earnings - Open','RE.OPEN',2,'RE.CLOSE')
    ii('total','Net Income for Year','NI',2,'RE.CLOSE')
    ii('account','Dividends Paid','DIVPAID',2,'RE.CLOSE')
    ii('separator',sep='single'); ii('total','Retained Earnings - Close','RE.CLOSE',3,'RE'); ii('separator',sep='double')

    # Retained-earnings finalization (LAP model): the opening account RE.OB — which
    # totals into RE from the off-statement TRX report — and the computed open:/close:RE
    # display lines are created by migrate_re_computed() at the end of this function,
    # the SINGLE source of truth for the RE structure. No RE.OFS / PY.CONV / PY.CLOSE.

    # AJE and TRX start empty — the user adds TRX.OPEN during onboarding; the backslash
    # menu builds AJE entries.

    # ── Default import rules (minimal set — shows naming conventions) ──
    # Add client-specific rules via add_rule or reclassify_suspense (auto-learns).
    rules = [
        # Banking — always EX.SC
        ('SERVICE CHARGE','EX.SC',  'E',  10, ''),
        ('BANK FEE',   'EX.SC',     'E',  10, ''),
        ('NSF',        'EX.SC',     'E',  10, ''),
        ('MONTHLY FEE','EX.SC',     'E',  10, ''),
        # Transfers — through clearing
        ('TRANSFER IN','CLR.TSF',   'E',  12, 'Account transfer'),
        ('TRANSFER OUT','CLR.TSF',  'E',  12, 'Account transfer'),
        # Payroll
        ('PAYROLL',    'EX.SAL',    'E',  5,  ''),
        # Rent
        ('RENT',       'EX.RENT',   'E',  5,  ''),
    ]
    for kw, acct, tax, pri, notes in rules:
        save_import_rule(None, kw, acct, tax, pri, notes)

    # Default tax codes for Canada
    save_tax_code('G5', 'GST 5%', 5.0, 'GST.OUT', 'GST.IN')
    save_tax_code('H13', 'HST 13% (Ontario)', 13.0, 'GST.OUT', 'GST.IN')
    save_tax_code('H15', 'HST 15% (Atlantic)', 15.0, 'GST.OUT', 'GST.IN')
    save_tax_code('E', 'Exempt (no tax)', 0, '', '')

    # Default GIFI mappings for standard accounts
    gifi_defaults = {
        'CASH':     '1000', 'BANK.CHQ':  '1000', 'BANK.SAV': '1000',
        'AR':       '1060', 'PREPAIDS':  '1480', 'DEP':      '1180',
        'EQUIP':    '1680', 'FURN':      '1680', 'COMP':     '1680',
        'EQUIP.DEP':'1740', 'FURN.DEP':  '1740', 'COMP.DEP': '1740',
        'AP':       '2620', 'AP.CC':     '2620',
        'FEDTAX':   '2680', 'PROTAX':    '2680',
        'LOAN':     '3140', 'SH.LOAN':   '2781',
        'CAPITAL':  '3500', 'RE.OPEN':   '3600', 'DIVPAID':  '3700',
        'REV':      '8000', 'REV.SVC':   '8000',
        'EX.ADV':   '8520', 'EX.INS':    '8690', 'EX.SC':    '8715',
        'EX.OFFICE':'8760', 'EX.FEES':   '8860', 'EX.RENT':  '9270',
        'EX.SAL':   '9060', 'EX.MEALS':  '9284', 'EX.PHONE': '9220',
        'EX.TRAVEL':'9200', 'EX.AUTO':   '9270', 'EX.POST':  '8760',
        'EX.COMP':  '8760', 'EX.AMORT':  '9275', 'EX.INTAX': '9990',
    }
    with get_db() as db:
        for acct_name, gifi in gifi_defaults.items():
            db.execute("UPDATE accounts SET gifi_code = ? WHERE name = ?", (gifi, acct_name))

    # The chart above is built in the legacy RE shape; converge it to the canonical
    # LAP structure (RE.OB + computed open:/close:RE; no RE.OFS / PY.CONV / PY.CLOSE)
    # so new files are born clean. Both are idempotent.
    migrate_re_computed()
    # Building a file is not REPAIRING one. The migration is what wires a fresh
    # chart into the modern RE shape, so it always reports a change here — and
    # telling an operator their brand-new books were repaired is alarming and
    # untrue. Only a file that arrived already built can have been repaired.
    global _re_repair
    _re_repair = []
    ensure_trx_layout()     # heading / TRX.OPEN / RE.OB, in that order, from birth
    return path


# ─── Setup Detailed AR Subledger ─────────────────────────────────
def setup_detailed_ar():
    """Scaffold a Detailed Accounts Receivable subledger report.

    Creates:
      - AR report on the home screen with 3 sample client accounts
      - ARDET total account that accumulates client balances
      - AR.DET total account on the BS that receives the cross-report total
      - Cross-report link: client R. accounts → ARDET → AR.DET → AR.TOT → CA → TA

    Raises ValueError if the AR report already exists or BS report is missing.
    Returns a summary string on success.
    """
    # ── Guard: AR report must not already exist ──
    existing = find_report_by_name('AR')
    if existing:
        raise ValueError("AR report already exists. Cannot create duplicate.")

    # ── Guard: BS report must exist ──
    bs = find_report_by_name('BS')
    if not bs:
        raise ValueError("Balance Sheet (BS) report not found. Create books first.")

    # ── Create accounts ──
    samples = [
        ('R.GREWAY', 'D', 'Gretzky, Wayne'),
        ('R.LEMMAR', 'D', 'Lemieux, Mario'),
        ('R.ORRBOB', 'D', 'Orr, Bobby'),
    ]
    sample_ids = {}
    for name, nb, desc in samples:
        ex = get_account_by_name(name)
        if ex:
            sample_ids[name] = ex['id']
        else:
            sample_ids[name] = add_account(name, nb, desc, 'posting')

    # ARDET — total account on AR report (receives client balances)
    ex = get_account_by_name('ARDET')
    if ex:
        ardet_id = ex['id']
    else:
        ardet_id = add_account('ARDET', 'D', 'Detailed AR - Total', 'total')

    # AR.DET — total account on BS (receives cross-report total from ARDET)
    ex = get_account_by_name('AR.DET')
    if ex:
        ardet_bs_id = ex['id']
    else:
        ardet_bs_id = add_account('AR.DET', 'D', 'Detailed AR', 'total')

    # ── Create the AR report ──
    # Sort after BS(10), IS(20), before AJE(30) — use 25
    ar_report = add_report('AR', 'Accounts Receivable - Detailed', 25)

    # ── Populate AR report items ──
    p = [0]
    def ai(itype, desc='', acct_id=None, ind=0, tt1='', sep=''):
        p[0] += 10
        add_report_item(ar_report, itype, desc, acct_id, ind, p[0], tt1, sep_style=sep)

    ai('label', 'Accounts Receivable - Detailed')
    ai('label', '')
    ai('label', 'Client Receivables:', ind=1)
    ai('account', '', sample_ids['R.GREWAY'], 2, 'ARDET')
    ai('account', '', sample_ids['R.LEMMAR'], 2, 'ARDET')
    ai('account', '', sample_ids['R.ORRBOB'], 2, 'ARDET')
    ai('separator', sep='single')
    ai('total', '', ardet_id, 3, 'AR.DET')  # cross-report link
    ai('separator', sep='double')

    # ── Insert AR.DET on the BS near existing AR account ──
    bs_items = get_report_items(bs['id'])

    # Find existing AR account on BS
    ar_item = None
    ar_position = None
    ar_total_to = 'CA'  # fallback
    for item in bs_items:
        if item['acct_name'] and item['acct_name'].upper() == 'AR':
            ar_item = item
            ar_position = item['position']
            ar_total_to = item['total_to_1'] or 'AR.TOT'
            break

    if ar_item:
        # Insert AR.DET just after the existing AR account
        insert_pos = ar_position + 1
    else:
        # No AR account found — look for Current Assets total (CA) and insert before it
        for item in bs_items:
            if item['acct_name'] and item['acct_name'].upper() == 'CA':
                insert_pos = item['position'] - 1
                ar_total_to = 'CA'
                break
        else:
            # Last resort: append to end of BS
            insert_pos = None
            ar_total_to = 'CA'

    # AR.DET totals to the same target as AR (typically AR.TOT)
    add_report_item(bs['id'], 'account', '', ardet_bs_id, 2, insert_pos, ar_total_to)

    created = ', '.join(s[0] for s in samples)
    return (f"Detailed AR subledger created. "
            f"Report: AR ({len(samples)} sample clients: {created}). "
            f"Total account ARDET → AR.DET on BS → {ar_total_to}.")


# ─── Setup Detailed AP Subledger ─────────────────────────────────
def setup_detailed_ap():
    """Scaffold a Detailed Accounts Payable subledger report.

    Creates:
      - AP.SUB report on the home screen with 3 sample vendor accounts
      - APDET total account that accumulates vendor balances
      - AP.DET total account on the BS that receives the cross-report total
      - AP.TOT total account on the BS that subtotals AP + AP.CC + AP.DET → CL
      - Cross-report link: vendor P. accounts → APDET → AP.DET → AP.TOT → CL → TL

    Raises ValueError if the AP.SUB report already exists or BS report is missing.
    Returns a summary string on success.
    """
    # ── Guard: AP.SUB report must not already exist ──
    existing = find_report_by_name('AP.SUB')
    if existing:
        raise ValueError("AP.SUB report already exists. Cannot create duplicate.")

    # ── Guard: BS report must exist ──
    bs = find_report_by_name('BS')
    if not bs:
        raise ValueError("Balance Sheet (BS) report not found. Create books first.")

    # ── Create vendor accounts ──
    samples = [
        ('P.BAUEQU', 'C', 'Bauer, Equipment'),
        ('P.CCMSPO', 'C', 'CCM, Sports'),
        ('P.WARHOC', 'C', 'Warrior, Hockey'),
    ]
    sample_ids = {}
    for name, nb, desc in samples:
        ex = get_account_by_name(name)
        if ex:
            sample_ids[name] = ex['id']
        else:
            sample_ids[name] = add_account(name, nb, desc, 'posting')

    # APDET — total account on AP.SUB report (receives vendor balances)
    ex = get_account_by_name('APDET')
    if ex:
        apdet_id = ex['id']
    else:
        apdet_id = add_account('APDET', 'C', 'Detailed AP - Total', 'total')

    # AP.DET — total account on BS (receives cross-report total from APDET)
    ex = get_account_by_name('AP.DET')
    if ex:
        apdet_bs_id = ex['id']
    else:
        apdet_bs_id = add_account('AP.DET', 'C', 'Detailed AP', 'total')

    # AP.TOT — total account on BS (subtotals AP + AP.CC + AP.DET → CL)
    ex = get_account_by_name('AP.TOT')
    if ex:
        aptot_id = ex['id']
    else:
        aptot_id = add_account('AP.TOT', 'C', 'Total Accounts Payable', 'total')

    # ── Create the AP.SUB report ──
    # Sort after AR(25), before AJE(30) — use 26
    ap_report = add_report('AP.SUB', 'Accounts Payable - Detailed', 26)

    # ── Populate AP.SUB report items ──
    p = [0]
    def ai(itype, desc='', acct_id=None, ind=0, tt1='', sep=''):
        p[0] += 10
        add_report_item(ap_report, itype, desc, acct_id, ind, p[0], tt1, sep_style=sep)

    ai('label', 'Accounts Payable - Detailed')
    ai('label', '')
    ai('label', 'Vendor Payables:', ind=1)
    ai('account', '', sample_ids['P.BAUEQU'], 2, 'APDET')
    ai('account', '', sample_ids['P.CCMSPO'], 2, 'APDET')
    ai('account', '', sample_ids['P.WARHOC'], 2, 'APDET')
    ai('separator', sep='single')
    ai('total', '', apdet_id, 3, 'AP.DET')  # cross-report link
    ai('separator', sep='double')

    # ── Restructure BS: add AP.DET and AP.TOT near existing AP accounts ──
    bs_items = get_report_items(bs['id'])

    # Find existing AP accounts on BS and the target they total to
    ap_item = None
    ap_cc_item = None
    ap_total_to = 'CL'  # fallback
    for item in bs_items:
        name = (item['acct_name'] or '').upper()
        if name == 'AP':
            ap_item = item
            ap_total_to = item['total_to_1'] or 'CL'
        elif name == 'AP.CC':
            ap_cc_item = item

    # Re-point existing AP and AP.CC to total to AP.TOT instead of CL
    if ap_item:
        update_report_item(ap_item['id'], total_to_1='AP.TOT')
    if ap_cc_item:
        update_report_item(ap_cc_item['id'], total_to_1='AP.TOT')

    # Determine insertion position — after AP.CC (or after AP if no AP.CC)
    # _resequence() runs after each add_report_item, so re-read positions between inserts.
    def _pos_after(acct_name):
        """Find current position of an account on the BS and return pos+1."""
        for it in get_report_items(bs['id']):
            if (it['acct_name'] or '').upper() == acct_name.upper():
                return it['position'] + 1
        return None

    last_ap = ap_cc_item or ap_item
    last_ap_name = last_ap['acct_name'] if last_ap else ''

    # Insert AP.DET after last AP account
    pos = _pos_after(last_ap_name) if last_ap_name else None
    add_report_item(bs['id'], 'account', '', apdet_bs_id, 2, pos, 'AP.TOT')

    # Insert separator after AP.DET (re-read position since resequence ran)
    pos = _pos_after('AP.DET')
    add_report_item(bs['id'], 'separator', '', None, 0, pos, sep_style='single')

    # Insert AP.TOT after the separator — find AP.DET pos, skip 2 (AP.DET + sep)
    items_now = get_report_items(bs['id'])
    for i, it in enumerate(items_now):
        if (it['acct_name'] or '').upper() == 'AP.DET':
            # AP.DET is at index i, separator is i+1, insert AP.TOT at i+2
            pos = items_now[i + 1]['position'] + 1 if i + 1 < len(items_now) else None
            break
    else:
        pos = None
    add_report_item(bs['id'], 'total', '', aptot_id, 3, pos, ap_total_to)

    created = ', '.join(s[0] for s in samples)
    return (f"Detailed AP subledger created. "
            f"Report: AP.SUB ({len(samples)} sample vendors: {created}). "
            f"Total account APDET → AP.DET on BS → AP.TOT → {ap_total_to}.")


# ─── CaseWare AJE Import ────────────────────────────────────────

def parse_csw_aje(file_path):
    """Parse CaseWare AJE export (IIF or Venice format).

    Auto-detects format from file content.
    Returns:
      {
        'format': 'iif' | 'venice',
        'entries': [
          {
            'num': '01',
            'date': '2024-12-01',
            'description': '...',
            'lines': [
              {'csw_account': 'Dividends paid-taxable', 'csw_number': '', 'amount_cents': 10000},
              ...
            ]
          }, ...
        ],
        'csw_accounts': [
          {'name': 'Dividends paid-taxable', 'number': ''}, ...
        ]
      }
    """
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        raw = f.read()

    if '!TRNS' in raw or 'ENDTRNS' in raw:
        return _parse_iif_aje(raw)
    elif 'STOP' in raw:
        return _parse_venice_aje(raw)
    else:
        raise ValueError("Unrecognized file format. Expected IIF (TRNS/ENDTRNS) or Venice (STOP delimited).")


def _parse_iif_aje(raw):
    """Parse QuickBooks IIF format AJE export."""
    import re
    lines = raw.split('\n')
    entries = []
    current_lines = []
    seen_accounts = {}

    for line in lines:
        stripped = line.strip()
        if not stripped or stripped.startswith('!'):
            continue
        cols = stripped.split('\t')
        marker = cols[0].strip().upper() if cols else ''

        if marker == 'TRNS' or marker == 'SPL':
            # Col indices: 0=marker, 1=ID, 2=TYPE, 3=DATE, 4=ACCNT, 5=NAME, 6=CLASS, 7=AMOUNT, 8=DOCNUM, 9=MEMO
            acct_name = cols[4].strip() if len(cols) > 4 else ''
            amount_str = cols[7].strip() if len(cols) > 7 else '0'
            docnum = cols[8].strip() if len(cols) > 8 else ''
            memo = cols[9].strip() if len(cols) > 9 else ''
            date_str = cols[3].strip() if len(cols) > 3 else ''

            amount_cents = parse_amount(amount_str)

            current_lines.append({
                'csw_account': acct_name,
                'csw_number': '',
                'amount_cents': amount_cents,
                'date': date_str,
                'docnum': docnum,
                'memo': memo,
                'is_trns': marker == 'TRNS',
            })

            if acct_name and acct_name not in seen_accounts:
                seen_accounts[acct_name] = {'name': acct_name, 'number': ''}

        elif marker == 'ENDTRNS':
            if current_lines:
                # Extract entry info from the TRNS line
                trns = current_lines[0]
                # Parse AJE number from DOCNUM (e.g. "AJE01" -> "01")
                num_match = re.search(r'AJE\s*(\d+)', trns['docnum'], re.IGNORECASE)
                aje_num = num_match.group(1) if num_match else trns['docnum']

                # Normalize date (IIF uses M/D/YY or M/D/YYYY)
                entry_date = normalize_date(trns['date'])
                if not entry_date:
                    entry_date = trns['date']

                entries.append({
                    'num': aje_num,
                    'date': entry_date,
                    'description': trns['memo'],
                    'lines': [{'csw_account': l['csw_account'],
                               'csw_number': l['csw_number'],
                               'amount_cents': l['amount_cents']}
                              for l in current_lines],
                })
            current_lines = []

    return {
        'format': 'iif',
        'entries': entries,
        'csw_accounts': list(seen_accounts.values()),
    }


def _parse_venice_aje(raw):
    """Parse Venice/MYOB format AJE export."""
    import re
    blocks = raw.split('STOP')
    entries = []
    seen_accounts = {}

    for block in blocks:
        block = block.strip()
        if not block:
            continue
        lines = block.split('\n')
        if not lines:
            continue

        # Header line: chars 0-10 = date (DD/MM/YYYY), chars 10-15 = ref (AJEnn), rest = description
        header = lines[0]
        if len(header) < 15:
            continue

        date_str = header[0:10].strip()
        ref_str = header[10:15].strip()
        description = header[15:].strip()

        # Parse date explicitly as DD/MM/YYYY (Venice format)
        try:
            entry_date = datetime.strptime(date_str, '%d/%m/%Y').strftime('%Y-%m-%d')
        except ValueError:
            entry_date = normalize_date(date_str) or date_str

        # Parse AJE number
        num_match = re.search(r'AJE\s*(\d+)', ref_str, re.IGNORECASE)
        aje_num = num_match.group(1) if num_match else ref_str

        # Detail lines: account name, account number, amount
        entry_lines = []
        for dline in lines[1:]:
            dline = dline.rstrip()
            if not dline.strip():
                continue
            # Pattern: account name (with spaces), then 2+ spaces, then 4-digit number, then spaces, then signed amount
            m = re.match(r'^(.+?)\s{2,}(\d{4})\s+([-\d,.]+)\s*$', dline)
            if m:
                acct_name = m.group(1).strip()
                acct_num = m.group(2).strip()
                amount_cents = parse_amount(m.group(3))

                entry_lines.append({
                    'csw_account': acct_name,
                    'csw_number': acct_num,
                    'amount_cents': amount_cents,
                })

                if acct_name not in seen_accounts:
                    seen_accounts[acct_name] = {'name': acct_name, 'number': acct_num}

        if entry_lines:
            entries.append({
                'num': aje_num,
                'date': entry_date,
                'description': description,
                'lines': entry_lines,
            })

    return {
        'format': 'venice',
        'entries': entries,
        'csw_accounts': list(seen_accounts.values()),
    }


def auto_match_accounts(csw_accounts):
    """Try to match CsW account names to Grid accounts.

    Match strategies (in order):
      1. Exact name match (case-insensitive)
      2. Grid account description contains CsW name (case-insensitive)
      3. CsW number matches Grid account_number field
      4. None — user must assign

    Args:
        csw_accounts: list of {'name': str, 'number': str}

    Returns:
        dict: {csw_name: {'id': int, 'name': str} or None}
    """
    all_accounts = get_accounts()
    suggestions = {}

    for csw in csw_accounts:
        csw_name = csw['name']
        csw_num = csw.get('number', '')
        match = None

        # Strategy 1: exact name match
        for a in all_accounts:
            if a['name'].lower() == csw_name.lower():
                match = {'id': a['id'], 'name': a['name']}
                break

        # Strategy 2: Grid description contains CsW name
        if not match:
            csw_lower = csw_name.lower()
            for a in all_accounts:
                if csw_lower in (a['description'] or '').lower():
                    match = {'id': a['id'], 'name': a['name']}
                    break

        # Strategy 3: CsW number matches Grid account_number
        if not match and csw_num:
            for a in all_accounts:
                if a['account_number'] and a['account_number'] == csw_num:
                    match = {'id': a['id'], 'name': a['name']}
                    break

        suggestions[csw_name] = match

    return suggestions


def ensure_journal_account(name, report_name):
    """Ensure a journal account exists and is on the given report.
    Creates the account (C-normal, posting) and places it if needed.
    Returns account_id.
    """
    acct = get_account_by_name(name)
    if acct:
        acct_id = acct['id']
    else:
        acct_id = add_account(name, 'C', f'Journal - {name}', 'posting')

    # Check if already on the report
    reports = get_reports()
    report = None
    for r in reports:
        if r['name'] == report_name:
            report = r
            break
    if report:
        items = get_report_items(report['id'])
        already_on = any(item['account_id'] == acct_id for item in items)
        if not already_on:
            add_report_item(report['id'], 'account', '', acct_id, indent=2)

    return acct_id


def import_aje_entries(entries, account_map, ref_prefix, journal_account=None):
    """Post parsed AJE entries as transactions, routing through a journal account.

    Each CaseWare AJE with N lines becomes N separate 2-line transactions:
    one line to the real account, one to the journal account (opposite sign).
    The journal account nets to zero but shows every entry in one ledger.

    Args:
        entries: list from parse_csw_aje()['entries']
        account_map: {csw_account_name: grid_account_id}
        ref_prefix: e.g. '25AJE'
        journal_account: name of journal account (e.g. '25AJE'). If provided,
            routes all legs through this account. If None, posts as compound txns.

    Returns:
        dict: {posted: int, skipped: int, errors: list}
    """
    posted = 0
    skipped = 0
    errors = []

    # Set up journal account if requested
    journal_acct_id = None
    if journal_account:
        journal_acct_id = ensure_journal_account(journal_account, 'AJE')

    for i, entry in enumerate(entries):
        ref = f"{ref_prefix}{entry['num']}"
        entry_date = entry['date']
        desc = entry['description']

        # Build transaction lines
        txn_lines = []
        unmapped = False
        for line in entry['lines']:
            grid_acct_id = account_map.get(line['csw_account'])
            if not grid_acct_id:
                errors.append({'entry': ref, 'reason': f"Unmapped account: {line['csw_account']}"})
                unmapped = True
                break
            txn_lines.append((grid_acct_id, line['amount_cents'], desc))

        if unmapped:
            skipped += 1
            continue

        # Validate balance
        total = sum(l[1] for l in txn_lines)
        if total != 0:
            errors.append({'entry': ref, 'reason': f"Entry does not balance: off by {total/100:.2f}"})
            skipped += 1
            continue

        try:
            if journal_acct_id:
                # Route each leg through the journal account as separate 2-line txns
                for acct_id, amount_cents, line_desc in txn_lines:
                    add_transaction(entry_date, ref, desc, [
                        (acct_id, amount_cents, line_desc),
                        (journal_acct_id, -amount_cents, line_desc),
                    ])
            else:
                # Legacy: post as one compound transaction
                add_transaction(entry_date, ref, desc, txn_lines)
            posted += 1
        except ValueError as e:
            errors.append({'entry': ref, 'reason': str(e)})
            skipped += 1

    return {'posted': posted, 'skipped': skipped, 'errors': errors}


def process_aje_file(file_path, ref_prefix, journal_account=None, ye_date=None):
    """Full AJE pipeline: detect format, extract entries, map accounts, post.

    For IIF/Venice files: uses parse_csw_aje() directly.
    For PDF/image/Excel/CSV: calls aje_extract.py for AI extraction.
    No year-end rollforward: the perpetual RE model re-derives opening/closing RE
    automatically, so prior-period AJEs are just ordinary dated postings.

    Args:
        file_path: Path to the AJE document (any format)
        ref_prefix: Reference prefix for posted entries (e.g. '25AJE')
        journal_account: Journal account name (defaults to ref_prefix)
        ye_date: Accepted for backward compatibility; ignored (no rollforward).

    Returns:
        dict with: extraction_result, mapping, posting_result, reroll_result (if applicable)
    """
    import subprocess, json

    if journal_account is None:
        journal_account = ref_prefix

    ext = os.path.splitext(file_path)[1].lower()
    native_formats = {'.iif', '.txt', '.ven'}

    # Step 1: Extract entries
    if ext in native_formats:
        parsed = parse_csw_aje(file_path)
        entries = parsed['entries']
        csw_accounts = parsed['csw_accounts']
        extraction_result = {'format': parsed['format'], 'entry_count': len(entries)}
    else:
        # Call aje_extract.py for AI extraction
        script_dir = os.path.dirname(os.path.abspath(__file__))
        script_path = os.path.join(script_dir, 'aje_extract.py')
        if not os.path.exists(script_path):
            raise ValueError(f"aje_extract.py not found at {script_path}")

        result = subprocess.run(
            [sys.executable, script_path, file_path],
            capture_output=True, text=True, timeout=120
        )
        if result.returncode != 0:
            raise ValueError(f"aje_extract.py failed: {result.stderr.strip()}")

        parsed = json.loads(result.stdout)
        entries = parsed['entries']
        csw_accounts = parsed['csw_accounts']
        extraction_result = {'format': 'extracted', 'entry_count': len(entries)}

    if not entries:
        return {'extraction_result': extraction_result, 'mapping': {},
                'posting_result': {'posted': 0, 'skipped': 0, 'errors': []},
                'reroll_result': None}

    # Step 2: Map accounts
    mapping = auto_match_accounts(csw_accounts)

    # Build account_map: csw_name -> grid_account_id
    account_map = {}
    unmatched = []
    for csw_name, match in mapping.items():
        if match:
            account_map[csw_name] = match['id']
        else:
            unmatched.append(csw_name)

    # Step 3: Post AJEs. No year-end rollforward in the perpetual model — prior-period
    # adjustments are ordinary dated postings that auto-re-derive opening/closing RE.
    posting_result = import_aje_entries(entries, account_map, ref_prefix, journal_account)
    reroll_result = None

    return {
        'extraction_result': extraction_result,
        'mapping': {k: v for k, v in mapping.items()},
        'unmatched_accounts': unmatched,
        'posting_result': posting_result,
        'reroll_result': reroll_result,
    }


