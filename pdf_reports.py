"""
pdf_reports — shared PDF generators for GridTRX.

Zero Flask dependencies.  All functions operate on models.py directly
and return raw bytes (or helper data).  Used by both app.py (web UI)
and mcp_server.py (MCP tools).
"""
import io
import os
import hashlib
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


def engagement_scorecard_pdf(company, period, ytd_data, period_summary):
    """Generate engagement scorecard PDF (page 3 of monthly packet).

    Args:
        company: Company name string.
        period: Current period 'YYYY-MM'.
        ytd_data: List from models.get_ytd_engagement() —
                  [{period, events: {type: count}, total}]
        period_summary: Dict from models.get_engagement_summary() —
                        {event_type: count}

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

    y = ph - margin

    # Header
    c.setFont(font_b, 10)
    c.drawCentredString(pw / 2, y + 5, company)
    y -= 14
    c.setFont(font_b, 8)
    c.drawCentredString(pw / 2, y, f'Engagement Scorecard — {period}')
    y -= 20

    # Current month summary
    c.setFont(font_b, 8)
    c.drawString(margin, y, 'This Month')
    y -= 12

    event_labels = {
        'files_received': 'Files Received',
        'import_completed': 'Import Completed',
        'reminder_first': 'Reminder Sent (1st)',
        'reminder_second': 'Reminder Sent (2nd)',
        'client_responded': 'Client Responded',
        'suspense_cleared': 'Suspense Cleared',
        'report_generated': 'Report Generated',
        'data_reminder': 'Data Reminder Sent',
        'data_deadline': 'Deadline Reminder Sent',
        'packet_sent': 'Packet Delivered',
    }

    c.setFont(font, 7)
    for etype, label in event_labels.items():
        count = period_summary.get(etype, 0)
        if count > 0:
            mark = 'Yes' if count == 1 else f'{count}x'
            c.drawString(margin + 10, y, f'{label}: {mark}')
            y -= 10
    if not period_summary:
        c.drawString(margin + 10, y, '(no events recorded)')
        y -= 10

    y -= 10

    # YTD grid
    c.setFont(font_b, 8)
    c.drawString(margin, y, 'Year-to-Date Activity')
    y -= 14

    # Column layout for YTD grid
    col_period = margin
    col_files = margin + 60
    col_import = margin + 120
    col_remind = margin + 180
    col_susp = margin + 240
    col_report = margin + 300
    col_packet = margin + 360

    c.setFont(font_b, 6.5)
    c.drawString(col_period, y, 'Period')
    c.drawString(col_files, y, 'Files')
    c.drawString(col_import, y, 'Import')
    c.drawString(col_remind, y, 'Reminders')
    c.drawString(col_susp, y, 'Suspense')
    c.drawString(col_report, y, 'Report')
    c.drawString(col_packet, y, 'Packet')
    y -= 2
    c.setLineWidth(0.4)
    c.line(margin, y, right_edge, y)
    y -= 10

    c.setFont(font, 6.5)
    for entry in ytd_data:
        ev = entry['events']
        c.drawString(col_period, y, entry['period'])
        c.drawString(col_files, y, str(ev.get('files_received', 0)))
        c.drawString(col_import, y, str(ev.get('import_completed', 0)))
        remind_count = ev.get('reminder_first', 0) + ev.get('reminder_second', 0) + \
                       ev.get('data_reminder', 0) + ev.get('data_deadline', 0)
        c.drawString(col_remind, y, str(remind_count))
        c.drawString(col_susp, y, str(ev.get('suspense_cleared', 0)))
        c.drawString(col_report, y, str(ev.get('report_generated', 0)))
        c.drawString(col_packet, y, str(ev.get('packet_sent', 0)))
        y -= 10

    c.save()
    buf.seek(0)
    return buf.read()


def _monthly_is_pdf(company, report_id, year, through_month):
    """Generate a 13-column IS: one column per month (Jan-Dec) + YTD total.

    Landscape orientation. Months beyond through_month are blank.
    """
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.pdfgen import canvas
    from calendar import monthrange

    font, font_b = _setup_fonts()

    buf = io.BytesIO()
    pw, ph = landscape(letter)  # 792 x 612
    c = canvas.Canvas(buf, pagesize=landscape(letter))
    margin = 28
    right_edge = pw - margin
    usable_w = pw - 2 * margin

    fs = 5.5
    line_h = 8
    ncols = 13  # 12 months + YTD

    # Layout: description column + 13 amount columns
    desc_w = 110
    col_w = (usable_w - desc_w) // ncols
    max_desc_chars = int(desc_w / (fs * 0.5))

    # Column right-edge positions
    col_rights = []
    x = margin + desc_w
    for i in range(ncols):
        col_rights.append(x + col_w - 2)
        x += col_w

    month_names = ['Jan','Feb','Mar','Apr','May','Jun',
                   'Jul','Aug','Sep','Oct','Nov','Dec','YTD']

    # Compute each month's data
    month_columns = []  # list of 12 column data lists
    display_items = models.get_report_items(report_id)
    all_items = models.get_all_report_items()

    for m in range(1, 13):
        if m <= through_month:
            d_from = f'{year}-{m:02d}-01'
            d_to = f'{year}-{m:02d}-{monthrange(year, m)[1]:02d}'
            col = models.compute_report_column(
                report_id, d_from, d_to,
                _display_items=display_items, _all_items=all_items)
            month_columns.append(col)
        else:
            month_columns.append(None)

    # YTD column
    ytd_from = f'{year}-01-01'
    ytd_to = f'{year}-{through_month:02d}-{monthrange(year, through_month)[1]:02d}'
    ytd_col = models.compute_report_column(
        report_id, ytd_from, ytd_to,
        _display_items=display_items, _all_items=all_items)
    month_columns.append(ytd_col)

    # Build unified row list from YTD column (has all items)
    rows = []
    for idx, (item, ytd_amt) in enumerate(ytd_col):
        month_vals = []
        for m_idx in range(12):
            mc = month_columns[m_idx]
            if mc is None:
                month_vals.append(None)
            else:
                month_vals.append(mc[idx][1] if idx < len(mc) else 0)
        month_vals.append(ytd_amt)
        rows.append((item, month_vals))

    y = ph - margin
    page_num = 1

    def header():
        nonlocal y
        c.setFont(font_b, 8)
        c.drawCentredString(pw / 2, ph - margin + 5, company)
        c.setFont(font_b, 6)
        c.drawCentredString(pw / 2, ph - margin - 5,
            f'Income Statement — Monthly Detail — {year}')
        c.setFont(font, 5)
        c.drawRightString(right_edge, ph - margin + 5, f'Page {page_num}')
        y = ph - margin - 12

    def col_header():
        nonlocal y
        c.setFont(font_b, fs - 0.5)
        c.drawString(margin, y, 'Description')
        for i, label in enumerate(month_names):
            c.drawRightString(col_rights[i], y, label)
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

    header()
    col_header()

    for item, vals in rows:
        itype = item.get('item_type', 'account')
        indent = item.get('indent', 0) or 0

        if itype == 'separator':
            check_page()
            style = item.get('sep_style', 'single')
            if style == 'double':
                c.setLineWidth(0.4)
                ly = y + line_h * 0.4
                c.line(margin + desc_w, ly, right_edge, ly)
                c.line(margin + desc_w, ly - 2, right_edge, ly - 2)
            elif style != 'blank':
                c.setLineWidth(0.2)
                c.line(margin + desc_w, y + line_h * 0.4, right_edge, y + line_h * 0.4)
            y -= line_h
            continue

        if itype == 'label':
            check_page()
            desc = item.get('description') or ''
            if desc:
                c.setFont(font_b, fs)
                c.drawString(margin, y, ('  ' * indent + desc)[:max_desc_chars])
            y -= line_h
            continue

        if itype in ('account', 'total'):
            # Skip zero rows for accounts
            if itype == 'account' and all((v is None or v == 0) for v in vals):
                continue

            check_page()
            if item.get('acct_desc'):
                desc = item['acct_desc']
            else:
                desc = item.get('description') or item.get('acct_desc') or item.get('acct_name') or ''
            is_total = itype == 'total'
            fn = font_b if is_total else font
            c.setFont(fn, fs)
            c.drawString(margin, y, ('  ' * indent + desc)[:max_desc_chars])

            for i, v in enumerate(vals):
                if v is None:
                    continue
                c.setFont(fn, fs)
                c.drawRightString(col_rights[i], y, _fmt_money(v))
            y -= line_h

    c.save()
    buf.seek(0)
    return buf.read()


def monthly_packet_pdf(company, db_path, period_end, period_start=None):
    """Generate monthly client packet: BS + YTD IS + 13-col Monthly IS + Scorecard.

    Merges separate PDFs using PyPDF2. Applies StandardEncryption
    (no password to open, owner password 'gridbk', no modifications allowed).

    Args:
        company: Company name string.
        db_path: Path to the client's books.db.
        period_end: End date 'YYYY-MM-DD'.
        period_start: Start date 'YYYY-MM-DD' (defaults to Jan 1 of the year).

    Returns:
        dict with keys:
            pdf_bytes: bytes — the merged PDF content
            sha256: str — hex digest of the PDF
            pages: int — number of pages
    """
    from PyPDF2 import PdfMerger, PdfReader, PdfWriter

    models.init_db(db_path)

    # Derive period
    year = int(period_end[:4])
    month = int(period_end[5:7])
    period = period_end[:7]  # YYYY-MM
    fy_start = f'{year}-01-01'

    # Page 1: Balance Sheet (cumulative as-of period end)
    bs_report = models.find_report_by_name('BS')
    bs_bytes = b''
    if bs_report:
        col_data = models.compute_report_column(bs_report['id'], None, period_end)
        bs_bytes = report_pdf(
            company, f'Balance Sheet — As at {_short_date(period_end)}',
            [_short_date(period_end)], ['actual'],
            [(item, [amt]) for item, amt in col_data]
        )

    # Page 2: Income Statement YTD (Jan 1 to period end)
    is_report = models.find_report_by_name('IS')
    is_bytes = b''
    if is_report:
        col_data = models.compute_report_column(is_report['id'], fy_start, period_end)
        is_bytes = report_pdf(
            company, f'Income Statement — Year to Date',
            [f'{_short_date(fy_start)} to {_short_date(period_end)}'], ['actual'],
            [(item, [amt]) for item, amt in col_data]
        )

    # Page 3: 13-column Monthly IS (landscape)
    monthly_is_bytes = b''
    if is_report:
        monthly_is_bytes = _monthly_is_pdf(company, is_report['id'], year, month)

    # Page 4: Engagement Scorecard
    ytd_data = models.get_ytd_engagement(year)
    period_summary = models.get_engagement_summary(period)
    scorecard_bytes = engagement_scorecard_pdf(company, period, ytd_data, period_summary)

    # Merge
    merger = PdfMerger()
    for pdf_bytes in [bs_bytes, is_bytes, monthly_is_bytes, scorecard_bytes]:
        if pdf_bytes:
            merger.append(io.BytesIO(pdf_bytes))
    merged_buf = io.BytesIO()
    merger.write(merged_buf)
    merger.close()

    merged_bytes = merged_buf.getvalue()

    # Apply encryption: no password to open, owner password restricts editing
    reader = PdfReader(io.BytesIO(merged_bytes))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.encrypt(user_password='', owner_password='gridbk',
                   permissions_flag=0b0100)  # allow printing only

    encrypted_buf = io.BytesIO()
    writer.write(encrypted_buf)
    final_bytes = encrypted_buf.getvalue()

    # Hash
    sha = hashlib.sha256(final_bytes).hexdigest()
    page_count = len(reader.pages)

    return {
        'pdf_bytes': final_bytes,
        'sha256': sha,
        'pages': page_count,
    }
