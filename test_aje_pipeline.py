#!/usr/bin/env python3
"""
Test: AJE Intake Pipeline

Tests process_aje_file and the CLI/MCP wrappers. (Rollforward/re-roll tests retired
with the perpetual RE model.) Uses create_starter_books() + the archive IIF file.
"""
import sys, os, shutil, traceback
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import models

BASE = os.path.dirname(os.path.abspath(__file__))
TEST_DIR = os.path.join(BASE, '_test_aje_pipeline')
DB_PATH = os.path.join(TEST_DIR, 'books.db')
IIF_PATH = os.path.join(os.path.dirname(BASE), 'archive', 'cwquickb.iif')

# ── Logging ──────────────────────────────────────────────────────
results = []
pass_count = 0
fail_count = 0

def log(msg=''):
    print(msg)

def check(label, condition, detail=''):
    global pass_count, fail_count
    tag = 'PASS' if condition else 'FAIL'
    if condition:
        pass_count += 1
    else:
        fail_count += 1
    msg = f"  [{tag}] {label}"
    if detail:
        msg += f" -- {detail}"
    print(msg)
    results.append((label, condition, detail))

def tb_balanced(as_of=None):
    if as_of is None:
        as_of = '2099-12-31'
    rows, dr, cr = models.get_trial_balance(as_of)
    return dr == cr, dr, cr

def get_balance(acct_name):
    acct = models.get_account_by_name(acct_name)
    if not acct:
        return None
    return models.get_account_balance(acct['id'])

# ── Setup ────────────────────────────────────────────────────────
def setup():
    if os.path.exists(TEST_DIR):
        shutil.rmtree(TEST_DIR)
    os.makedirs(TEST_DIR)
    models.create_starter_books(DB_PATH, 'AJE Test Corp', '12-31')
    # Set fiscal year to 2024
    models.set_meta('fiscal_year', '2024')
    models.set_meta('fy_ceiling_mode', 'cy')

def cleanup():
    if os.path.exists(TEST_DIR):
        shutil.rmtree(TEST_DIR)

# ── Test 1: Parse IIF file directly ─────────────────────────────
def test_parse_iif():
    log("\n═══ Test 1: parse_csw_aje (IIF) ═══")

    if not os.path.exists(IIF_PATH):
        print(f"  [SKIP] IIF fixture not on this machine ({IIF_PATH}) — test skipped, not failed")
        return None

    parsed = models.parse_csw_aje(IIF_PATH)

    check("Format detected as IIF", parsed['format'] == 'iif', f"got: {parsed['format']}")
    check("Entries parsed", len(parsed['entries']) > 0, f"count: {len(parsed['entries'])}")
    check("14 entries found", len(parsed['entries']) == 14, f"got: {len(parsed['entries'])}")
    check("CsW accounts collected", len(parsed['csw_accounts']) > 0, f"count: {len(parsed['csw_accounts'])}")

    # Check first entry
    e1 = parsed['entries'][0]
    check("Entry 1 num = '01'", e1['num'] == '01', f"got: {e1['num']}")
    check("Entry 1 has 2 lines", len(e1['lines']) == 2, f"got: {len(e1['lines'])}")

    # Check entry 3 (multi-line: AJE03 has 6 lines)
    e3 = parsed['entries'][2]
    check("Entry 3 num = '03'", e3['num'] == '03', f"got: {e3['num']}")
    check("Entry 3 is multi-line", len(e3['lines']) > 2, f"lines: {len(e3['lines'])}")

    # Check balance of each entry
    for entry in parsed['entries']:
        total = sum(line['amount_cents'] for line in entry['lines'])
        if total != 0:
            check(f"Entry {entry['num']} balanced", False, f"off by {total} cents")
            return parsed

    check("All entries balance to zero", True)
    return parsed


# ── Test 2: auto_match_accounts ──────────────────────────────────
def test_auto_match(parsed):
    log("\n═══ Test 2: auto_match_accounts ═══")

    if not parsed:
        print("  [SKIP] no parsed data (IIF fixture absent) — test skipped, not failed")
        return

    mapping = models.auto_match_accounts(parsed['csw_accounts'])
    matched = sum(1 for v in mapping.values() if v)
    unmatched = sum(1 for v in mapping.values() if not v)

    check("Mapping returned", len(mapping) > 0, f"count: {len(mapping)}")
    log(f"    Matched: {matched}, Unmatched: {unmatched}")

    # These are CaseWare account names from a real client — they won't match
    # starter book accounts. That's expected. The test verifies the function
    # runs without error and returns the right structure.
    for csw_name, match in mapping.items():
        if match:
            check(f"  Match: '{csw_name}' → {match['name']}", True)
        else:
            log(f"    No match: '{csw_name}' (expected — CsW names vs starter book)")


# ── Test 3: Rollforward then Undo ────────────────────────────────
def test_undo_rollforward():
    log("\n═══ Test 3: undo_rollforward ═══")

    # Post some revenue and expense to make RE non-zero
    rev = models.get_account_by_name('REV.SVC')
    exp = models.get_account_by_name('EX.OFFICE')
    bank = models.get_account_by_name('BANK.CHQ')

    # Revenue: Dr BANK.CHQ 5000, Cr REV.SVC 5000
    models.add_transaction('2024-06-15', 'TEST', 'Service revenue', [
        (bank['id'], 500000, ''),
        (rev['id'], -500000, ''),
    ])

    # Expense: Dr EX.OFFICE 2000, Cr BANK.CHQ 2000
    models.add_transaction('2024-07-01', 'TEST', 'Office supplies', [
        (exp['id'], 200000, ''),
        (bank['id'], -200000, ''),
    ])

    bal, dr, cr = tb_balanced()
    check("TB balanced before rollforward", bal, f"DR={dr}, CR={cr}")

    # Rollforward
    result = models.rollforward('2024-12-31')
    check("Rollforward succeeded", result is not None)
    check("YE-OFS txn posted", result['txn_id'] > 0, f"txn #{result['txn_id']}")
    check("Lock date set", models.get_meta('lock_date', '') == '2024-12-31')
    check("FY ceiling advanced", models.fiscal_ceiling() == '2025-12-31')
    check("Fiscal year advanced", models.get_meta('fiscal_year', '') == '2025')

    closing_re = result['closing_re']
    log(f"    Closing RE: {closing_re}")

    bal, dr, cr = tb_balanced()
    check("TB balanced after rollforward", bal, f"DR={dr}, CR={cr}")

    # Now undo
    undo = models.undo_rollforward('2024-12-31')
    check("Undo succeeded", undo is not None)
    check("Txn deleted", undo['deleted_txn_id'] == result['txn_id'])
    check("Lock date cleared", models.get_meta('lock_date', '') == '')
    check("Ceiling restored to 2024", undo['restored_ceiling'] == '2024-12-31')
    check("FY restored to 2024", undo['restored_fy'] == '2024')

    bal, dr, cr = tb_balanced()
    check("TB balanced after undo", bal, f"DR={dr}, CR={cr}")


# ── Test 4: re_rollforward ───────────────────────────────────────
def test_re_rollforward():
    log("\n═══ Test 4: re_rollforward ═══")

    # First rollforward again
    result = models.rollforward('2024-12-31')
    check("Initial rollforward", result is not None)
    old_txn = result['txn_id']
    old_re = result['closing_re']

    # Post an AJE in the closed period (need to clear lock first to post)
    # Actually, process_aje_file handles this. Let's test re_rollforward directly.
    # We'll manually clear lock, post, then re-roll.
    models.set_meta('lock_date', '')

    exp = models.get_account_by_name('EX.OFFICE')
    bank = models.get_account_by_name('BANK.CHQ')

    # Post additional expense in 2024
    models.add_transaction('2024-11-15', 'AJE', 'Late adjustment', [
        (exp['id'], 100000, ''),
        (bank['id'], -100000, ''),
    ])

    # Restore lock so re_rollforward has something to undo
    models.set_meta('lock_date', '2024-12-31')

    # Re-roll
    reroll = models.re_rollforward('2024-12-31')
    check("Re-rollforward succeeded", reroll is not None)
    check("Undo phase completed", reroll['undo'] is not None)
    check("Old txn deleted", reroll['undo']['deleted_txn_id'] == old_txn)
    check("New rollforward completed", reroll['rollforward'] is not None)
    new_txn = reroll['rollforward']['txn_id']
    new_re = reroll['rollforward']['closing_re']
    check("New txn different from old", new_txn != old_txn, f"old={old_txn}, new={new_txn}")
    check("RE changed after AJE", new_re != old_re, f"old={old_re}, new={new_re}")

    bal, dr, cr = tb_balanced()
    check("TB balanced after re-roll", bal, f"DR={dr}, CR={cr}")

    check("Lock date restored", models.get_meta('lock_date', '') == '2024-12-31')
    check("FY ceiling correct", models.fiscal_ceiling() == '2025-12-31')


# ── Test 5: Undo with no rollforward (error case) ────────────────
def test_undo_no_rollforward():
    log("\n═══ Test 5: Error cases ═══")

    # Undo the current rollforward first to get a clean state
    models.undo_rollforward('2024-12-31')

    # Now try to undo again — should fail
    try:
        models.undo_rollforward('2024-12-31')
        check("Undo with no rollforward raises error", False, "no exception raised")
    except ValueError as e:
        check("Undo with no rollforward raises ValueError", True, str(e))

    # Bad date format
    try:
        models.undo_rollforward('not-a-date')
        check("Bad date raises error", False, "no exception raised")
    except ValueError as e:
        check("Bad date raises ValueError", True, str(e))


# ── Test 6: process_aje_file with IIF ────────────────────────────
def test_process_aje_file():
    log("\n═══ Test 6: process_aje_file (IIF, no re-roll) ═══")

    if not os.path.exists(IIF_PATH):
        print(f"  [SKIP] IIF fixture not on this machine ({IIF_PATH}) — test skipped, not failed")
        return

    # Fresh books for this test
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
    models.create_starter_books(DB_PATH, 'AJE Test Corp', '12-31')
    models.set_meta('fiscal_year', '2024')
    models.set_meta('fy_ceiling_mode', 'cy')

    # Add accounts that might match some CsW names
    # (Most won't match, but that's fine — tests the unmatched path)
    result = models.process_aje_file(IIF_PATH, 'AJE')

    check("process_aje_file returned", result is not None)
    check("Extraction format = iif", result['extraction_result']['format'] == 'iif')
    check("14 entries extracted", result['extraction_result']['entry_count'] == 14,
          f"got: {result['extraction_result']['entry_count']}")
    check("Mapping returned", len(result['mapping']) > 0)

    unmatched = result.get('unmatched_accounts', [])
    log(f"    Unmatched accounts: {len(unmatched)}")
    log(f"    Posted: {result['posting_result']['posted']}, Skipped: {result['posting_result']['skipped']}")
    log(f"    Errors: {len(result['posting_result']['errors'])}")

    check("No re-roll (none requested)", result['reroll_result'] is None)

    bal, dr, cr = tb_balanced()
    check("TB balanced after import", bal, f"DR={dr}, CR={cr}")


# ── Test 7: process_aje_file with re-roll ─────────────────────────
def test_process_aje_file_reroll():
    log("\n═══ Test 7: process_aje_file with re-roll ═══")

    if not os.path.exists(IIF_PATH):
        print(f"  [SKIP] IIF fixture not on this machine ({IIF_PATH}) — test skipped, not failed")
        return

    # Fresh books
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
    models.create_starter_books(DB_PATH, 'AJE Test Corp', '12-31')
    models.set_meta('fiscal_year', '2024')
    models.set_meta('fy_ceiling_mode', 'cy')

    # Post some activity so RE is non-zero
    rev = models.get_account_by_name('REV.SVC')
    bank = models.get_account_by_name('BANK.CHQ')
    models.add_transaction('2024-06-15', 'TEST', 'Service revenue', [
        (bank['id'], 500000, ''),
        (rev['id'], -500000, ''),
    ])

    # Rollforward first
    roll = models.rollforward('2024-12-31')
    check("Initial rollforward for re-roll test", roll is not None)
    old_re = roll['closing_re']
    log(f"    Initial closing RE: {old_re}")

    # Now process AJE with ye_date — should trigger re-roll
    result = models.process_aje_file(IIF_PATH, 'AJE', ye_date='2024-12-31')

    check("process_aje_file with re-roll returned", result is not None)
    check("Re-roll result present", result['reroll_result'] is not None)

    if result['reroll_result']:
        check("Re-roll undo completed", result['reroll_result']['undo'] is not None)
        check("Re-roll forward completed", result['reroll_result']['rollforward'] is not None)

    bal, dr, cr = tb_balanced()
    check("TB balanced after AJE + re-roll", bal, f"DR={dr}, CR={cr}")

    check("Lock date still set", models.get_meta('lock_date', '') == '2024-12-31')
    check("FY ceiling at 2025", models.fiscal_ceiling() == '2025-12-31')


# ── Run ──────────────────────────────────────────────────────────
def main():
    global pass_count, fail_count

    log("╔════════════════════════════════════════════╗")
    log("║   AJE Pipeline Test Suite                  ║")
    log("╚════════════════════════════════════════════╝")

    try:
        setup()

        parsed = test_parse_iif()
        test_auto_match(parsed)
        # Rollforward / re-roll tests removed: year-end rollforward was retired
        # (perpetual RE model — process_aje_file no longer rerolls). See models.py.
        test_process_aje_file()

    except Exception as e:
        log(f"\n  FATAL: {e}")
        traceback.print_exc()
        fail_count += 1
    finally:
        cleanup()

    log(f"\n{'='*50}")
    log(f"  Results: {pass_count} passed, {fail_count} failed")
    log(f"{'='*50}")

    return 0 if fail_count == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
