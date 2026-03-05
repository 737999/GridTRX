"""
pdf_reports — shared PDF generators for GridTRX.

Zero Flask dependencies.  All functions operate on models.py directly
and return raw bytes (or helper data).  Used by both app.py (web UI)
and mcp_server.py (MCP tools).
"""
import io
import os
from datetime import datetime, timedelta
from collections import OrderedDict

import models


# ═══════════════════════════════════════════════════════════════════
# SHARED HELPERS
# ═══════════════════════════════════════════════════════════════════

def _setup_fonts():
    """Register the best available monospace TTF font.
    Returns (font_name, font_bold_name) for use with reportlab Canvas.
    Falls back to Courier / Courier-Bold if nothing else found.
    """
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    font = 'Courier'
    font_b = 'Courier-Bold'
    candidates = [
        ('LiberationMono', '/usr/share/fonts/truetype/liberation/LiberationMono-Regular.ttf',
                           '/usr/share/fonts/truetype/liberation/LiberationMono-Bold.ttf'),
        ('DejaVuMono',     '/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf',
                           '/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf'),
        ('Consolas',       'C:/Windows/Fonts/consola.ttf',
                           'C:/Windows/Fonts/consolab.ttf'),
        ('CourierNew',     'C:/Windows/Fonts/cour.ttf',
                           'C:/Windows/Fonts/courbd.ttf'),
        ('Menlo',          '/System/Library/Fonts/Menlo.ttc',
                           '/System/Library/Fonts/Menlo.ttc'),
    ]
    for name, regular, bold in candidates:
        if os.path.exists(regular) and os.path.exists(bold):
            try:
                try:
                    pdfmetrics.getFont(name)
                except KeyError:
                    pdfmetrics.registerFont(TTFont(name, regular))
                    pdfmetrics.registerFont(TTFont(name + '-Bold', bold))
                font = name
                font_b = name + '-Bold'
                break
            except Exception:
                continue
    return font, font_b


def _fmt_money(cents):
    """Format cents as dollar string. No normal-balance. Debits positive."""
    if cents == 0:
        return '\u2014'
    neg = cents < 0
    val = abs(cents) / 100.0
    s = f'{val:,.2f}'
    return f'({s})' if neg else s


def _short_date(d):
    """Format YYYY-MM-DD as dd-Mon-yy."""
    if not d:
        return ''
    try:
        dt = datetime.strptime(d[:10], '%Y-%m-%d')
        return dt.strftime('%d-%b-%y')
    except Exception:
        return d[:10]


def _get_bs_account_ids():
    """Return set of account_ids that appear in the BS report."""
    reports = models.get_reports()
    bs = next((r for r in reports if r['name'] == 'BS'), None)
    if not bs:
        return set()
    items = models.get_report_items(bs['id'])
    return {i['account_id'] for i in items if i['account_id']}


def _get_report_account_order(report_name):
    """Return ordered list of (account_id, acct_name, acct_desc) from a report."""
    reports = models.get_reports()
    rpt = next((r for r in reports if r['name'] == report_name), None)
    if not rpt:
        return []
    items = models.get_report_items(rpt['id'])
    seen = set()
    result = []
    for item in items:
        aid = item['account_id']
        atype = item['account_type'] or ''
        if aid and aid not in seen and atype == 'posting':
            seen.add(aid)
            result.append((aid, item['acct_name'] or '', item['acct_desc'] or ''))
    return result


def _build_account_detail(account_id, acct_name, acct_desc, begin, end, is_bs, dr_cr_filter='all'):
    """Build GL detail rows for one account.
    Returns (opening, rows, closing).
    Debits positive, credits negative. No normal-balance flipping.
    """
    rows = []

    # Opening balance
    if is_bs:
        if begin:
            d = datetime.strptime(begin, '%Y-%m-%d') - timedelta(days=1)
            opening = models.get_account_balance(account_id, date_to=d.strftime('%Y-%m-%d'))
        else:
            opening = 0
    else:
        opening = 0

    # Get transactions in period
    with models.get_db() as db:
        sql = """
            SELECT t.id as txn_id, t.date, t.reference, t.description as txn_desc,
                   l.amount, l.description as line_desc, l.id as line_id,
                   GROUP_CONCAT(DISTINCT a2.name) as cross_accounts
            FROM lines l
            JOIN transactions t ON l.transaction_id = t.id
            LEFT JOIN lines l2 ON l2.transaction_id = t.id AND l2.account_id != ?
            LEFT JOIN accounts a2 ON l2.account_id = a2.id
            WHERE l.account_id = ?"""
        params = [account_id, account_id]
        if begin:
            sql += " AND t.date >= ?"
            params.append(begin)
        if end:
            sql += " AND t.date <= ?"
            params.append(end)
        sql += " GROUP BY l.id ORDER BY t.date, t.id, l.sort_order"
        txns = db.execute(sql, params).fetchall()

    balance = opening

    for txn in txns:
        amt = txn['amount']
        debit = amt if amt > 0 else 0
        credit = -amt if amt < 0 else 0

        if dr_cr_filter == 'debit' and amt <= 0:
            continue
        if dr_cr_filter == 'credit' and amt >= 0:
            continue

        balance += amt
        cross_raw = txn['cross_accounts'] or ''
        cross = '-split-' if ',' in cross_raw else cross_raw

        rows.append({
            'type': 'txn',
            'date': txn['date'],
            'ref': txn['reference'] or '',
            'desc': txn['line_desc'] or txn['txn_desc'] or '',
            'debit': debit,
            'credit': credit,
            'balance': balance,
            'cross': cross,
        })

    closing = balance
    return opening, rows, closing


# ═══════════════════════════════════════════════════════════════════
# PDF GENERATORS — all return bytes
# ═══════════════════════════════════════════════════════════════════

def report_pdf(company, report_name, col_labels, col_types, rows, hide_zero=False):
    """Generate a portrait PDF for a single-column or comparative report (2-6 columns).

    Args:
        company: Company name string.
        report_name: Report name (e.g. "BS", "IS").
        col_labels: List of column header labels.
        col_types: List of 'actual'|'change'|'pct_change'|'spacer' per column.
        rows: List of (item_dict, values_list). values_list is a list of ints
              (cents for amounts, basis-points for pct).
        hide_zero: If True, skip account rows where all values are zero.

    Returns:
        bytes — the PDF content.
    """
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas

    font, font_b = _setup_fonts()

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    pw, ph = letter  # 612, 792
    margin = 36
    right_edge = pw - margin
    usable_w = pw - 2 * margin

    fs = 8
    line_h = 10
    ncols = len(col_labels)

    # Adapt column width based on column count
    if ncols <= 3:
        col_w = 90
    elif ncols <= 5:
        col_w = 75
    else:
        col_w = 65

    desc_w = usable_w - ncols * col_w
    if desc_w < 120:
        desc_w = 120
        col_w = max(50, (usable_w - desc_w) // ncols)

    max_desc_chars = int(desc_w / (fs * 0.52))

    # Column right-edge positions (amounts are right-aligned)
    col_rights = []
    x = margin + desc_w
    for i in range(ncols):
        col_rights.append(x + col_w - 4)
        x += col_w

    y = ph - margin
    page_num = 1

    def fmt_val(v, ctype):
        if v is None:
            return ''
        if ctype == 'pct_change':
            if v == 0:
                return '\u2014'
            return f'{v / 100:.1f}%'
        if ctype == 'spacer':
            return ''
        return _fmt_money(v)

    def header():
        nonlocal y
        c.setFont(font_b, 10)
        c.drawCentredString(pw / 2, ph - margin + 5, company)
        c.setFont(font_b, 8)
        c.drawCentredString(pw / 2, ph - margin - 7, report_name)
        c.setFont(font, 6)
        c.drawRightString(right_edge, ph - margin + 5, f'Page {page_num}')
        y = ph - margin - 18

    def col_header():
        nonlocal y
        c.setFont(font_b, fs - 1)
        c.drawString(margin, y, 'Description')
        for i, label in enumerate(col_labels):
            ct = col_types[i] if i < len(col_types) else 'actual'
            if ct != 'spacer':
                c.drawRightString(col_rights[i], y, label)
        y -= 2
        c.setLineWidth(0.5)
        c.line(margin, y, right_edge, y)
        y -= line_h

    def check_page(need=2):
        nonlocal y, page_num
        if y < margin + need * line_h:
            c.showPage()
            page_num += 1
            header()
            col_header()

    header()
    col_header()

    for item, bals in rows:
        itype = item.get('item_type', 'account')
        indent = item.get('indent', 0) or 0

        if itype == 'separator':
            check_page()
            style = item.get('sep_style', 'single')
            if style == 'double':
                c.setLineWidth(0.5)
                line_y = y + line_h * 0.4
                c.line(margin + desc_w, line_y, right_edge, line_y)
                c.line(margin + desc_w, line_y - 2.5, right_edge, line_y - 2.5)
            elif style == 'blank':
                pass
            else:
                c.setLineWidth(0.3)
                c.line(margin + desc_w, y + line_h * 0.4, right_edge, y + line_h * 0.4)
            y -= line_h
            continue

        if itype == 'label':
            check_page()
            desc = item.get('description') or ''
            if desc:
                c.setFont(font_b, fs)
                display = '  ' * indent + desc
                c.drawString(margin, y, display[:max_desc_chars])
            y -= line_h
            continue

        # account or total
        if itype in ('account', 'total'):
            if hide_zero and itype == 'account':
                if isinstance(bals, list) and all((b is None or b == 0) for b in bals):
                    continue

            check_page()
            # Description
            if item.get('acct_desc'):
                desc = item['acct_desc']
            else:
                desc = item.get('description') or item.get('acct_desc') or item.get('acct_name') or ''
            is_total = itype == 'total'
            fn = font_b if is_total else font
            c.setFont(fn, fs)
            display = '  ' * indent + desc
            c.drawString(margin, y, display[:max_desc_chars])

            # Values
            if isinstance(bals, list):
                for i, v in enumerate(bals):
                    ct = col_types[i] if i < len(col_types) else 'actual'
                    if ct == 'spacer':
                        continue
                    c.setFont(fn, fs)
                    c.drawRightString(col_rights[i], y, fmt_val(v, ct))
            else:
                # Single value (backward compat)
                c.setFont(fn, fs)
                c.drawRightString(col_rights[0], y, fmt_val(bals, 'actual'))

            y -= line_h

    c.save()
    buf.seek(0)
    return buf.read()


def gl_pdf(company, accounts, bs_ids, begin, end, dr_cr_filter='all'):
    """Generate General Ledger as monospaced PDF.

    Args:
        company: Company name.
        accounts: List of (account_id, acct_name, acct_desc).
        bs_ids: Set of account_ids on the Balance Sheet.
        begin: Start date (YYYY-MM-DD) or ''.
        end: End date (YYYY-MM-DD) or ''.
        dr_cr_filter: 'all', 'debit', or 'credit'.

    Returns:
        bytes — the PDF content.
    """
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas

    font, font_b = _setup_fonts()

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    pw, ph = letter
    margin = 36
    right_edge = pw - margin

    fs = 6.5
    line_h = 8.5

    col_date = margin
    col_ref = margin + 40
    col_desc = margin + 66
    col_debit = margin + 310
    col_credit = margin + 375
    col_balance = margin + 440
    col_cross = margin + 510
    desc_max = 62

    y = ph - margin
    page_num = 1

    begin_s = _short_date(begin) if begin else 'Start'
    end_s = _short_date(end) if end else 'Current'

    def header():
        nonlocal y
        c.setFont(font_b, 8)
        c.drawString(margin, ph - margin + 5, f'{company} \u2014 General Ledger')
        c.setFont(font, 6)
        c.drawString(margin, ph - margin - 4, f'{begin_s} to {end_s}')
        c.drawRightString(right_edge, ph - margin + 5, f'Page {page_num}')
        y = ph - margin - 12

    def col_header():
        nonlocal y
        c.setFont(font_b, fs)
        c.drawString(col_date, y, 'Date')
        c.drawString(col_ref, y, 'Ref')
        c.drawString(col_desc, y, 'Description')
        c.drawRightString(col_debit + 58, y, 'Debit')
        c.drawRightString(col_credit + 58, y, 'Credit')
        c.drawRightString(col_balance + 58, y, 'Balance')
        c.drawString(col_cross, y, 'Acct')
        y -= 2
        c.setLineWidth(0.4)
        c.line(margin, y, right_edge, y)
        y -= line_h

    def check_page(need=2):
        nonlocal y, page_num
        if y < margin + need * line_h:
            c.showPage()
            page_num += 1
            header()
            col_header()

    def draw_row(date_s, ref_s, desc_s, debit, credit, balance, cross='', bold=False):
        nonlocal y
        check_page()
        fn = font_b if bold else font
        c.setFont(fn, fs)
        c.drawString(col_date, y, date_s)
        c.drawString(col_ref, y, (ref_s or '')[:6])
        c.drawString(col_desc, y, (desc_s or '')[:desc_max])
        if debit:
            c.drawRightString(col_debit + 58, y, _fmt_money(debit))
        if credit:
            c.drawRightString(col_credit + 58, y, _fmt_money(-credit))
        if balance is not None:
            c.drawRightString(col_balance + 58, y, _fmt_money(balance))
        if cross:
            c.drawString(col_cross, y, cross[:12])
        y -= line_h

    header()

    for idx, (aid, aname, adesc) in enumerate(accounts):
        is_bs = aid in bs_ids
        opening, rows, closing = _build_account_detail(aid, aname, adesc, begin, end, is_bs, dr_cr_filter)

        if not rows and opening == 0:
            continue

        check_page(5)

        c.setFont(font_b, 8)
        c.drawString(margin, y, f'{aname}  {adesc}')
        y -= line_h
        col_header()

        draw_row(begin_s, '', 'Opening Balance', 0, 0, opening, bold=True)

        total_dr, total_cr = 0, 0
        for r in rows:
            total_dr += r['debit']
            total_cr += r['credit']
            draw_row(_short_date(r['date']), r['ref'], r['desc'],
                     r['debit'], r['credit'], r['balance'], r['cross'])

        check_page()
        c.setLineWidth(0.3)
        c.line(col_debit, y + line_h - 2, col_balance + 68, y + line_h - 2)

        draw_row(end_s, '', 'Closing Balance', total_dr, total_cr, closing, bold=True)

        c.setLineWidth(0.4)
        c.line(col_balance, y + line_h - 2, col_balance + 68, y + line_h - 2)
        c.line(col_balance, y + line_h - 5, col_balance + 68, y + line_h - 5)
        y -= line_h * 0.5

    c.save()
    buf.seek(0)
    return buf.read()


def account_ledger_pdf(company, account_id, acct_name, acct_desc, begin, end, is_bs=False):
    """Generate a single-account ledger PDF with Dr/Cr columns and running balance.

    Same layout as the GL report but for one account only.
    """
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas

    font, font_b = _setup_fonts()

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    pw, ph = letter
    margin = 36
    right_edge = pw - margin

    fs = 6.5
    line_h = 8.5

    col_date = margin
    col_ref = margin + 40
    col_desc = margin + 66
    col_debit = margin + 310
    col_credit = margin + 375
    col_balance = margin + 440
    col_cross = margin + 510
    desc_max = 62

    y = ph - margin
    page_num = 1

    begin_s = _short_date(begin) if begin else 'Start'
    end_s = _short_date(end) if end else 'Current'

    def header():
        nonlocal y
        c.setFont(font_b, 10)
        c.drawCentredString(pw / 2, ph - margin + 5, company)
        c.setFont(font_b, 8)
        c.drawCentredString(pw / 2, ph - margin - 7, f'{acct_name}  {acct_desc}')
        c.setFont(font, 6)
        c.drawString(margin, ph - margin - 16, f'{begin_s} to {end_s}')
        c.drawRightString(right_edge, ph - margin + 5, f'Page {page_num}')
        y = ph - margin - 24

    def col_header():
        nonlocal y
        c.setFont(font_b, fs)
        c.drawString(col_date, y, 'Date')
        c.drawString(col_ref, y, 'Ref')
        c.drawString(col_desc, y, 'Description')
        c.drawRightString(col_debit + 58, y, 'Debit')
        c.drawRightString(col_credit + 58, y, 'Credit')
        c.drawRightString(col_balance + 58, y, 'Balance')
        c.drawString(col_cross, y, 'Acct')
        y -= 2
        c.setLineWidth(0.4)
        c.line(margin, y, right_edge, y)
        y -= line_h

    def check_page(need=2):
        nonlocal y, page_num
        if y < margin + need * line_h:
            c.showPage()
            page_num += 1
            header()
            col_header()

    def draw_row(date_s, ref_s, desc_s, debit, credit, balance, cross='', bold=False):
        nonlocal y
        check_page()
        fn = font_b if bold else font
        c.setFont(fn, fs)
        c.drawString(col_date, y, date_s)
        c.drawString(col_ref, y, (ref_s or '')[:6])
        c.drawString(col_desc, y, (desc_s or '')[:desc_max])
        if debit:
            c.drawRightString(col_debit + 58, y, _fmt_money(debit))
        if credit:
            c.drawRightString(col_credit + 58, y, _fmt_money(-credit))
        if balance is not None:
            c.drawRightString(col_balance + 58, y, _fmt_money(balance))
        if cross:
            c.drawString(col_cross, y, cross[:12])
        y -= line_h

    header()
    col_header()

    opening, rows, closing = _build_account_detail(account_id, acct_name, acct_desc, begin, end, is_bs)

    draw_row(begin_s, '', 'Opening Balance', 0, 0, opening, bold=True)

    total_dr, total_cr = 0, 0
    for r in rows:
        total_dr += r['debit']
        total_cr += r['credit']
        draw_row(_short_date(r['date']), r['ref'], r['desc'],
                 r['debit'], r['credit'], r['balance'], r['cross'])

    check_page()
    c.setLineWidth(0.3)
    c.line(col_debit, y + line_h - 2, col_balance + 68, y + line_h - 2)

    draw_row(end_s, '', 'Closing Balance', total_dr, total_cr, closing, bold=True)

    c.setLineWidth(0.4)
    c.line(col_balance, y + line_h - 2, col_balance + 68, y + line_h - 2)
    c.line(col_balance, y + line_h - 5, col_balance + 68, y + line_h - 5)

    c.save()
    buf.seek(0)
    return buf.read()


def aje_pdf(company, account_id, acct_name, acct_desc, begin, end):
    """Generate AJE report for one account as PDF, grouped by reference.

    Args:
        company: Company name.
        account_id: The account's database ID.
        acct_name: Account name (e.g. 'EX.OFFICE').
        acct_desc: Account description.
        begin: Start date (YYYY-MM-DD) or ''.
        end: End date (YYYY-MM-DD) or ''.

    Returns:
        bytes — the PDF content.
    """
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas

    font, font_b = _setup_fonts()

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    pw, ph = letter
    margin = 36
    right_edge = pw - margin

    fs = 7
    line_h = 9

    col_ref = margin
    col_date = margin + 68
    col_desc = margin + 130
    amt_r = right_edge - 68
    col_acct = right_edge - 60

    y = ph - margin
    page_num = 1

    begin_s = _short_date(begin) if begin else 'Start'
    end_s = _short_date(end) if end else 'Current'

    def page_header():
        nonlocal y
        c.setFont(font_b, 9)
        c.drawString(margin, ph - margin + 5, company)
        c.setFont(font_b, 8)
        c.drawString(margin, ph - margin - 6, f'{acct_name} \u2014 {acct_desc}')
        c.setFont(font, 6.5)
        c.drawString(margin, ph - margin - 15, f'{begin_s} to {end_s}')
        c.drawRightString(right_edge, ph - margin + 5, f'Page {page_num}')
        y = ph - margin - 24

    def col_headers():
        nonlocal y
        c.setFont(font_b, fs)
        c.drawString(col_ref, y, 'Ref')
        c.drawString(col_date, y, 'Date')
        c.drawString(col_desc, y, 'Description')
        c.drawRightString(amt_r, y, 'Amount')
        c.drawString(col_acct, y, 'Account')
        y -= 2
        c.setLineWidth(0.4)
        c.line(margin, y, right_edge, y)
        y -= line_h

    def check_page(need=2):
        nonlocal y, page_num
        if y < margin + need * line_h:
            c.showPage()
            page_num += 1
            page_header()
            col_headers()

    # Query: all transaction lines touching this account
    with models.get_db() as db:
        sql = """
            SELECT t.id as txn_id, t.date, t.reference, t.description as txn_desc,
                   l.amount, l.description as line_desc,
                   GROUP_CONCAT(DISTINCT a2.name) as cross_accounts
            FROM lines l
            JOIN transactions t ON l.transaction_id = t.id
            LEFT JOIN lines l2 ON l2.transaction_id = t.id AND l2.account_id != ?
            LEFT JOIN accounts a2 ON l2.account_id = a2.id
            WHERE l.account_id = ?
        """
        params = [account_id, account_id]
        if begin:
            sql += " AND t.date >= ?"
            params.append(begin)
        if end:
            sql += " AND t.date <= ?"
            params.append(end)
        sql += " GROUP BY l.id ORDER BY t.reference, t.date, t.id, l.sort_order"
        txn_rows = db.execute(sql, params).fetchall()

    # Group by reference
    groups = OrderedDict()
    for row in txn_rows:
        ref = row['reference'] or '(no ref)'
        if ref not in groups:
            groups[ref] = []
        groups[ref].append(row)

    # Draw PDF
    page_header()
    col_headers()

    group_keys = list(groups.keys())
    for gi, ref in enumerate(group_keys):
        lines = groups[ref]
        check_page(min(len(lines) + 1, 5))

        for row in lines:
            check_page()
            c.setFont(font, fs)
            c.drawString(col_ref, y, (row['reference'] or '')[:10])
            c.drawString(col_date, y, _short_date(row['date']))
            c.drawString(col_desc, y, (row['line_desc'] or row['txn_desc'] or '')[:42])
            c.drawRightString(amt_r, y, _fmt_money(row['amount']))
            cross = row['cross_accounts'] or ''
            cross = '-split-' if ',' in cross else cross
            c.drawString(col_acct, y, cross[:10])
            y -= line_h

        if gi < len(group_keys) - 1:
            y -= line_h * 2

    c.save()
    buf.seek(0)
    return buf.read()
