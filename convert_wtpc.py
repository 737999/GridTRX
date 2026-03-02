"""
WTPC Converter — Converts NV DATAE text dump to Grid books.db
Uses the WTPC chart of accounts from the NV setup screenshots.
"""
import re, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import models

MONTHS = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
          'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
NOISE = set(MONTHS.keys()) | {'C', 'A', 'etsf'}

def parse_date(s):
    m = re.match(r'(\w{3})\s+(\d{1,2}),(\d{2})', s.strip())
    if not m: return None
    mon, day, yr = m.group(1), int(m.group(2)), int(m.group(3))
    if mon not in MONTHS: return None
    return f"{2000+yr:04d}-{MONTHS[mon]:02d}-{day:02d}"

def parse_amt(s):
    s = s.strip().replace(',','')
    if not s or s == '0.00': return 0
    neg = s.endswith('-')
    if neg: s = s[:-1]
    if s.startswith('('): s = s[1:]; neg = True
    if s.endswith(')'): s = s[:-1]
    parts = s.split('.')
    cents = int(parts[0]) * 100 + int(parts[1][:2].ljust(2,'0')) if '.' in s else int(s) * 100
    return -cents if neg else cents

# ─── Full WTPC Chart of Accounts ─────────────────────────────────
# From the NV setup screenshots (BS + IS)
# Format: (name, normal_bal, description, account_type)
WTPC_ACCOUNTS = [
    # BS - Current Assets - Bank
    ('CASH','D','Cash box','posting'),
    ('BANK.CHQ','D','Bank - Chequing TD 8735','posting'),
    ('BANK.SAV','D','Bank - ATB svgs','posting'),
    ('BANK.TAX','D','Bank - Tax reserve TD 7734','posting'),
    ('BANK.ATB','D','Bank - ATB chqing 3600','posting'),
    ('BANK.ROY','D','Bank - Royal','posting'),
    ('TOTBANK','D','Total Bank Accounts','total'),
    # Deposit clearing
    ('DC','D','Clearing - general','posting'),
    ('DC.E','D','Clearing - eTransfer','posting'),
    ('DC.SQ','D','Clearing - Square','posting'),
    ('DC.TOT','D','Total Deposit Clearing','total'),
    # AR
    ('AR','D','AR - General','posting'),
    ('AR.T1','D','AR - T1/US personal','posting'),
    ('AR.MISC','D','AR - misc','posting'),
    ('AR.OLD','D','AR - Old 2021','posting'),
    ('AR.NEW','D','AR - New 2023','posting'),
    ('AR.TOT','D','Total Accounts Receivable','total'),
    # Other current
    ('AR.CLEAR','D','AR Clearing','posting'),
    ('AR.WIP','D','WIP - New','posting'),
    ('WIP','D','Work in Progress','posting'),
    ('AR.RANCH','D','AR - Ranch','posting'),
    ('BRO','D','Brokerage Account','posting'),
    ('SEC','D','Securities','posting'),
    ('COINS','D','Coins/Collectibles','posting'),
    ('DEP','D','Deposits','posting'),
    ('PREPAIDS','D','Prepaid Expenses','posting'),
    ('CA','D','Total Current Assets','total'),
    # Capital assets
    ('EQUIP','D','Equipment','posting'),
    ('LEASEHOLDS','D','Leaseholds','posting'),
    ('FURN','D','Furniture','posting'),
    ('COMP','D','Computer Equipment','posting'),
    ('TOTFA','D','Total Capital Assets','total'),
    # Accumulated amortization
    ('BUILD.DEP','C','Accum Amort - Building','posting'),
    ('LEASE.DEP','C','Accum Amort - Leaseholds','posting'),
    ('EQUIP.DEP','C','Accum Amort - Equipment','posting'),
    ('FURN.DEP','C','Accum Amort - Furniture','posting'),
    ('COMP.DEP','C','Accum Amort - Computer','posting'),
    ('TOTDEP','C','Total Accum Amortization','total'),
    ('NETFA','D','Net Capital Assets','total'),
    # Goodwill
    ('GOODWILL','D','Goodwill','posting'),
    ('TA','D','TOTAL ASSETS','total'),
    # Current liabilities
    ('AP','C','Accounts Payable','posting'),
    ('AP.ACC','C','AP - Accruals','posting'),
    ('AP.TRADE','C','AP - Trade','posting'),
    ('AP.AMEX','C','AP - Amex','posting'),
    ('AP.OXFPRO','C','AP - Oxford Properties','posting'),
    ('AP.RE','C','AP - Real Estate','posting'),
    ('AP.DEP','C','AP - Deposits held','posting'),
    ('LOC.ATB','C','Line of Credit - ATB','posting'),
    # GST
    ('GST.OUT','C','GST Collected','posting'),
    ('GST.IN','D','GST Paid (ITCs)','posting'),
    ('GST.REMIT','C','GST Remittance','posting'),
    ('GST.PAY','C','GST Payable','posting'),
    ('TOTGST','C','Total GST','total'),
    # Tax
    ('FEDTAX','C','Federal Tax Payable','posting'),
    ('PROTAX','C','Provincial Tax Payable','posting'),
    ('TOT.TAX','C','Total Tax Payable','total'),
    ('CL','C','Total Current Liabilities','total'),
    # Long-term
    ('LOAN.ATB','C','Loan - ATB','posting'),
    ('LOAN.CEBA','C','Loan - CEBA','posting'),
    ('AP.SHARE','C','Shareholder Loans','posting'),
    ('TOTTERM','C','Total Term Debt','total'),
    ('LTL','C','Total Long-Term Liabilities','total'),
    ('TL','C','TOTAL LIABILITIES','total'),
    # Equity
    ('CAPITAL','C','Share Capital','posting'),
    ('RE','C','Retained Earnings','posting'),
    ('EQ','C','Total Equity','total'),
    # IS - Revenue
    ('REV.CLEAR','C','Revenue - Clearing','posting'),
    ('REV.SQ','C','Revenue - Square','posting'),
    ('REV.FEES','C','Revenue - Fees','posting'),
    ('REV.T1','C','Revenue - T1/Personal','posting'),
    ('REV.US','C','Revenue - US Returns','posting'),
    ('REV.UN','C','Revenue - UN Returns','posting'),
    ('REV.WIP','C','Revenue - WIP Adjust','posting'),
    ('REV.RECOV','C','Revenue - Recoveries','posting'),
    ('REV.SUND','C','Revenue - Sundry','posting'),
    ('REV.25','C','Revenue - 2025','posting'),
    ('TOTREV','C','Total Revenue','total'),
    # Cost of sales
    ('CS.SAL','D','Salary - Production','posting'),
    ('CS.PRTAX','D','Payroll Tax - Production','posting'),
    ('CS.SUB','D','Subcontractors','posting'),
    ('CS.PROF','D','Professional Fees','posting'),
    ('CS.TAXCAN','D','Tax Software - Canada','posting'),
    ('CS.TAXUS','D','Tax Software - US','posting'),
    ('CS.ADMIN','D','Admin - COS','posting'),
    ('CS.SHIP','D','Shipping - COS','posting'),
    ('CS.DISB','D','Disbursements - COS','posting'),
    ('GROSS','C','Gross Profit','total'),
    # Operating expenses
    ('EX.ADV','D','Advertising & Marketing','posting'),
    ('EX.POD','D','Podcast Expenses','posting'),
    ('EX.BAD','D','Bad Debts','posting'),
    ('EX.PEN','D','Penalties & Interest','posting'),
    ('EX.MEALS','D','Meals & Entertainment','posting'),
    ('EX.COMP','D','Computer Expenses','posting'),
    ('EX.DEP','D','Depreciation','posting'),
    ('EX.DUES','D','Dues & Memberships','posting'),
    ('EX.EDU','D','Education & Training','posting'),
    ('EX.INS','D','Insurance','posting'),
    ('EX.LIC','D','Licenses & Permits','posting'),
    ('EX.LIB','D','Library & Resources','posting'),
    ('EX.MAINT','D','Maintenance & Repairs','posting'),
    ('EX.OFFICE','D','Office & General','posting'),
    ('EX.PDUES','D','Professional Dues','posting'),
    ('EX.PROF','D','Professional Fees','posting'),
    ('EX.RENT','D','Rent','posting'),
    ('EX.SC','D','Service Charges','posting'),
    ('EX.SHIP','D','Shipping & Postage','posting'),
    ('EX.SUSP','D','Suspense','posting'),
    ('EX.TAXS','D','Tax Software','posting'),
    ('EX.TAXUS','D','Tax Software - US Exp','posting'),
    ('EX.TEL','D','Telephone & Internet','posting'),
    ('EX.TRAVEL','D','Travel','posting'),
    ('EX.CC','D','Credit Card Charges','posting'),
    ('EX.ADMIN','D','Admin Expenses','posting'),
    ('EX.AUTO','D','Auto Expenses','posting'),
    ('TOTEX','D','Total Operating Expenses','total'),
    ('OPINC','C','Operating Income','total'),
    # Other items
    ('EX.LIFE','D','Life Insurance','posting'),
    ('EX.INVEST','D','Investment Expenses','posting'),
    ('EX.LTINT','D','Long-term Interest','posting'),
    ('REV.INT','C','Interest Income','posting'),
    ('REV.RENTAL','C','Rental Income','posting'),
    ('REV.GAIN','C','Capital Gains','posting'),
    ('REV.ASSET','C','Asset Disposal','posting'),
    ('EX.FOREX','D','Foreign Exchange','posting'),
    ('EX.AMORT','D','Amortization','posting'),
    ('EX.OWN','D','Owner Benefits','posting'),
    ('TAXINC','C','Income Before Taxes','total'),
    ('EX.INTAX','D','Income Tax','posting'),
    ('NETINC','C','Net Income (Loss)','total'),
    ('NI','C','Net Income → RE','total'),
    # RE section
    ('RE.OPEN','C','Retained Earnings - Opening','posting'),
    ('DIVPAID','D','Dividends Paid','posting'),
    ('DIVPAID2','D','Dividends Paid 2','posting'),
    ('RE.CLOSE','C','Retained Earnings - Closing','total'),
    # AP sub-report accounts
    ('P.BUNJOA','C','Bund Joanne','posting'),
    ('P.WARTAX','C','Warman Tax','posting'),
]

# ─── BS Report Items ─────────────────────────────────────────────
# (item_type, description, account_name, indent, total_to_1, sep_style)
BS_ITEMS = [
    ('label','CURRENT ASSETS',None,0,'',''),
    ('separator','',None,0,'','blank'),
    ('label','Bank Accounts:',None,1,'',''),
    ('account','',  'CASH',2,'TOTBANK',''),
    ('account','',  'BANK.CHQ',2,'TOTBANK',''),
    ('account','',  'BANK.SAV',2,'TOTBANK',''),
    ('account','',  'BANK.TAX',2,'TOTBANK',''),
    ('account','',  'BANK.ATB',2,'TOTBANK',''),
    ('account','',  'BANK.ROY',2,'TOTBANK',''),
    ('separator','',None,0,'','single'),
    ('total','',    'TOTBANK',3,'CA',''),
    ('separator','',None,0,'','blank'),
    ('label','Deposit Clearing:',None,1,'',''),
    ('account','',  'DC',2,'DC.TOT',''),
    ('account','',  'DC.E',2,'DC.TOT',''),
    ('account','',  'DC.SQ',2,'DC.TOT',''),
    ('separator','',None,0,'','single'),
    ('total','',    'DC.TOT',3,'CA',''),
    ('separator','',None,0,'','blank'),
    ('label','Accounts Receivable:',None,1,'',''),
    ('account','',  'AR',2,'AR.TOT',''),
    ('account','',  'AR.T1',2,'AR.TOT',''),
    ('account','',  'AR.MISC',2,'AR.TOT',''),
    ('account','',  'AR.OLD',2,'AR.TOT',''),
    ('account','',  'AR.NEW',2,'AR.TOT',''),
    ('separator','',None,0,'','single'),
    ('total','',    'AR.TOT',3,'CA',''),
    ('separator','',None,0,'','blank'),
    ('label','Other Current Assets:',None,1,'',''),
    ('account','',  'AR.CLEAR',2,'CA',''),
    ('account','',  'AR.WIP',2,'CA',''),
    ('account','',  'WIP',2,'CA',''),
    ('account','',  'AR.RANCH',2,'CA',''),
    ('account','',  'BRO',2,'CA',''),
    ('account','',  'SEC',2,'CA',''),
    ('account','',  'COINS',2,'CA',''),
    ('account','',  'DEP',2,'CA',''),
    ('account','',  'PREPAIDS',2,'CA',''),
    ('separator','',None,0,'','single'),
    ('total','Total Current Assets','CA',1,'TA',''),
    ('separator','',None,0,'','blank'),
    ('label','CAPITAL ASSETS',None,0,'',''),
    ('account','',  'EQUIP',2,'TOTFA',''),
    ('account','',  'LEASEHOLDS',2,'TOTFA',''),
    ('account','',  'FURN',2,'TOTFA',''),
    ('account','',  'COMP',2,'TOTFA',''),
    ('separator','',None,0,'','single'),
    ('total','',    'TOTFA',3,'NETFA',''),
    ('separator','',None,0,'','blank'),
    ('label','Accumulated Amortization:',None,1,'',''),
    ('account','',  'BUILD.DEP',2,'TOTDEP',''),
    ('account','',  'LEASE.DEP',2,'TOTDEP',''),
    ('account','',  'EQUIP.DEP',2,'TOTDEP',''),
    ('account','',  'FURN.DEP',2,'TOTDEP',''),
    ('account','',  'COMP.DEP',2,'TOTDEP',''),
    ('separator','',None,0,'','single'),
    ('total','',    'TOTDEP',3,'NETFA',''),
    ('separator','',None,0,'','single'),
    ('total','Net Capital Assets','NETFA',1,'TA',''),
    ('separator','',None,0,'','blank'),
    ('account','',  'GOODWILL',2,'TA',''),
    ('separator','',None,0,'','double'),
    ('total','TOTAL ASSETS','TA',0,'',''),
    ('separator','',None,0,'','blank'),
    ('separator','',None,0,'','blank'),
    ('label','CURRENT LIABILITIES',None,0,'',''),
    ('account','',  'AP',2,'CL',''),
    ('account','',  'AP.ACC',2,'CL',''),
    ('account','',  'AP.TRADE',2,'CL',''),
    ('account','',  'AP.AMEX',2,'CL',''),
    ('account','',  'AP.OXFPRO',2,'CL',''),
    ('account','',  'AP.RE',2,'CL',''),
    ('account','',  'AP.DEP',2,'CL',''),
    ('account','',  'LOC.ATB',2,'CL',''),
    ('separator','',None,0,'','blank'),
    ('label','GST:',None,1,'',''),
    ('account','',  'GST.OUT',2,'TOTGST',''),
    ('account','',  'GST.IN',2,'TOTGST',''),
    ('account','',  'GST.REMIT',2,'TOTGST',''),
    ('account','',  'GST.PAY',2,'TOTGST',''),
    ('separator','',None,0,'','single'),
    ('total','',    'TOTGST',3,'CL',''),
    ('separator','',None,0,'','blank'),
    ('label','Tax:',None,1,'',''),
    ('account','',  'FEDTAX',2,'TOT.TAX',''),
    ('account','',  'PROTAX',2,'TOT.TAX',''),
    ('separator','',None,0,'','single'),
    ('total','',    'TOT.TAX',3,'CL',''),
    ('separator','',None,0,'','single'),
    ('total','Total Current Liabilities','CL',1,'TL',''),
    ('separator','',None,0,'','blank'),
    ('label','LONG-TERM LIABILITIES',None,0,'',''),
    ('account','',  'LOAN.ATB',2,'TOTTERM',''),
    ('account','',  'LOAN.CEBA',2,'TOTTERM',''),
    ('account','',  'AP.SHARE',2,'TOTTERM',''),
    ('separator','',None,0,'','single'),
    ('total','',    'TOTTERM',3,'LTL',''),
    ('separator','',None,0,'','single'),
    ('total','Total Long-Term Liabilities','LTL',1,'TL',''),
    ('separator','',None,0,'','blank'),
    ('label','EQUITY',None,0,'',''),
    ('account','',  'CAPITAL',2,'EQ',''),
    ('account','',  'RE',2,'EQ',''),
    ('separator','',None,0,'','single'),
    ('total','Total Equity','EQ',1,'TL',''),
    ('separator','',None,0,'','double'),
    ('total','TOTAL LIABILITIES & EQUITY','TL',0,'',''),
]

# ─── IS Report Items ─────────────────────────────────────────────
IS_ITEMS = [
    ('label','REVENUE',None,0,'',''),
    ('account','',  'REV.CLEAR',2,'TOTREV',''),
    ('account','',  'REV.SQ',2,'TOTREV',''),
    ('account','',  'REV.FEES',2,'TOTREV',''),
    ('account','',  'REV.T1',2,'TOTREV',''),
    ('account','',  'REV.US',2,'TOTREV',''),
    ('account','',  'REV.UN',2,'TOTREV',''),
    ('account','',  'REV.WIP',2,'TOTREV',''),
    ('account','',  'REV.RECOV',2,'TOTREV',''),
    ('account','',  'REV.SUND',2,'TOTREV',''),
    ('account','',  'REV.25',2,'TOTREV',''),
    ('separator','',None,0,'','single'),
    ('total','Total Revenue','TOTREV',1,'GROSS',''),
    ('separator','',None,0,'','blank'),
    ('label','COST OF SALES',None,0,'',''),
    ('account','',  'CS.SAL',2,'GROSS',''),
    ('account','',  'CS.PRTAX',2,'GROSS',''),
    ('account','',  'CS.SUB',2,'GROSS',''),
    ('account','',  'CS.PROF',2,'GROSS',''),
    ('account','',  'CS.TAXCAN',2,'GROSS',''),
    ('account','',  'CS.TAXUS',2,'GROSS',''),
    ('account','',  'CS.ADMIN',2,'GROSS',''),
    ('account','',  'CS.SHIP',2,'GROSS',''),
    ('account','',  'CS.DISB',2,'GROSS',''),
    ('separator','',None,0,'','single'),
    ('total','Gross Profit','GROSS',1,'OPINC',''),
    ('separator','',None,0,'','blank'),
    ('label','EXPENSES',None,0,'',''),
    ('account','',  'EX.ADV',2,'TOTEX',''),
    ('account','',  'EX.POD',2,'TOTEX',''),
    ('account','',  'EX.BAD',2,'TOTEX',''),
    ('account','',  'EX.PEN',2,'TOTEX',''),
    ('account','',  'EX.MEALS',2,'TOTEX',''),
    ('account','',  'EX.COMP',2,'TOTEX',''),
    ('account','',  'EX.DEP',2,'TOTEX',''),
    ('account','',  'EX.DUES',2,'TOTEX',''),
    ('account','',  'EX.EDU',2,'TOTEX',''),
    ('account','',  'EX.INS',2,'TOTEX',''),
    ('account','',  'EX.LIC',2,'TOTEX',''),
    ('account','',  'EX.LIB',2,'TOTEX',''),
    ('account','',  'EX.MAINT',2,'TOTEX',''),
    ('account','',  'EX.OFFICE',2,'TOTEX',''),
    ('account','',  'EX.PDUES',2,'TOTEX',''),
    ('account','',  'EX.PROF',2,'TOTEX',''),
    ('account','',  'EX.RENT',2,'TOTEX',''),
    ('account','',  'EX.SC',2,'TOTEX',''),
    ('account','',  'EX.SHIP',2,'TOTEX',''),
    ('account','',  'EX.SUSP',2,'TOTEX',''),
    ('account','',  'EX.TAXS',2,'TOTEX',''),
    ('account','',  'EX.TAXUS',2,'TOTEX',''),
    ('account','',  'EX.TEL',2,'TOTEX',''),
    ('account','',  'EX.TRAVEL',2,'TOTEX',''),
    ('account','',  'EX.CC',2,'TOTEX',''),
    ('account','',  'EX.ADMIN',2,'TOTEX',''),
    ('account','',  'EX.AUTO',2,'TOTEX',''),
    ('separator','',None,0,'','single'),
    ('total','Total Operating Expenses','TOTEX',1,'OPINC',''),
    ('separator','',None,0,'','single'),
    ('total','Operating Income','OPINC',1,'TAXINC',''),
    ('separator','',None,0,'','blank'),
    ('label','OTHER ITEMS',None,0,'',''),
    ('account','',  'EX.LIFE',2,'TAXINC',''),
    ('account','',  'EX.INVEST',2,'TAXINC',''),
    ('account','',  'EX.LTINT',2,'TAXINC',''),
    ('account','',  'REV.INT',2,'TAXINC',''),
    ('account','',  'REV.RENTAL',2,'TAXINC',''),
    ('account','',  'REV.GAIN',2,'TAXINC',''),
    ('account','',  'REV.ASSET',2,'TAXINC',''),
    ('account','',  'EX.FOREX',2,'TAXINC',''),
    ('account','',  'EX.AMORT',2,'TAXINC',''),
    ('account','',  'EX.OWN',2,'TAXINC',''),
    ('separator','',None,0,'','single'),
    ('total','Income Before Taxes','TAXINC',1,'NETINC',''),
    ('separator','',None,0,'','blank'),
    ('account','',  'EX.INTAX',2,'NETINC',''),
    ('separator','',None,0,'','double'),
    ('total','Net Income (Loss)','NETINC',0,'NI',''),
    ('separator','',None,0,'','blank'),
    ('label','RETAINED EARNINGS',None,0,'',''),
    ('account','',  'RE.OPEN',2,'RE.CLOSE',''),
    ('total','',    'NI',2,'RE.CLOSE',''),
    ('account','',  'DIVPAID',2,'RE.CLOSE',''),
    ('account','',  'DIVPAID2',2,'RE.CLOSE',''),
    ('separator','',None,0,'','double'),
    ('total','Retained Earnings - Closing','RE.CLOSE',0,'',''),
]

def create_wtpc_books(output_path, datae_path=None):
    """Create WTPC books with full chart of accounts and optionally import transactions."""
    if os.path.exists(output_path):
        os.remove(output_path)
    
    models.init_db(output_path)
    models.set_meta('company_name', 'WTPC Professional Corporation')
    models.set_meta('fiscal_year_end', '05-31')
    
    # Create accounts
    acct_map = {}  # name -> id
    for name, nb, desc, atype in WTPC_ACCOUNTS:
        acct_map[name] = models.add_account(name, nb, desc, atype)
    
    # Create reports
    bs_id = models.add_report('BS', 'Balance Sheet - WTPC', 10)
    is_id = models.add_report('IS', 'Income Statement - WTPC', 20)
    aje_id = models.add_report('AJE', 'Adjusting Entries', 30)
    trx_id = models.add_report('TRX', 'Transaction Journal', 40)
    
    # Build BS items
    pos = 10
    for item_type, desc, acct_name, indent, tt1, sep in BS_ITEMS:
        aid = acct_map.get(acct_name) if acct_name else None
        itype = item_type
        if itype == 'account' and acct_name and acct_name in acct_map:
            acct_row = [a for a in WTPC_ACCOUNTS if a[0] == acct_name]
            if acct_row and acct_row[0][3] == 'total':
                itype = 'total'
        models.add_report_item(bs_id, itype, desc, aid, indent, pos, tt1, sep_style=sep)
        pos += 10
    
    # Build IS items
    pos = 10
    for item_type, desc, acct_name, indent, tt1, sep in IS_ITEMS:
        aid = acct_map.get(acct_name) if acct_name else None
        itype = item_type
        if itype == 'account' and acct_name and acct_name in acct_map:
            acct_row = [a for a in WTPC_ACCOUNTS if a[0] == acct_name]
            if acct_row and acct_row[0][3] == 'total':
                itype = 'total'
        models.add_report_item(is_id, itype, desc, aid, indent, pos, tt1, sep_style=sep)
        pos += 10
    
    print(f"Created {len(acct_map)} accounts, 4 reports")
    print(f"  BS: {len(BS_ITEMS)} items")
    print(f"  IS: {len(IS_ITEMS)} items")
    
    if datae_path:
        import_datae(datae_path, acct_map)
    
    return acct_map

def import_datae(datae_path, acct_map):
    """Import transactions from DATAE text dump."""
    with open(datae_path, 'rb') as f:
        raw = f.read()
    text = raw.decode('latin-1')
    lines = text.split('\n')
    
    # Parse all transaction lines
    # Strategy: group lines by txn number, take the SECOND row (final version)
    # for simple 2-account txns
    txn_groups = {}  # txn_num -> list of lines
    current_num = None
    
    noise_accts = set(MONTHS.keys()) | {'C', 'A'}
    
    for line in lines:
        line = line.rstrip('\r')
        if not line.strip():
            continue
        
        # Header line (with txn number)
        m = re.match(r'(\d{10})\s+A\s+C\s+(\S+)\s+(\S*)\s+(\w{3}\s+\d{1,2},\d{2})\s*(.*)', line)
        if m:
            current_num = int(m.group(1))
            if current_num not in txn_groups:
                txn_groups[current_num] = {'lines': [], 'date': None, 'main_acct': None}
            acct1 = m.group(2)
            acct2 = m.group(3)
            date_str = parse_date(m.group(4))
            rest = m.group(5).strip()
            
            if acct1 not in noise_accts:
                txn_groups[current_num]['lines'].append(('header', acct1, acct2, date_str, rest))
                txn_groups[current_num]['date'] = date_str
                txn_groups[current_num]['main_acct'] = acct1
            continue
        
        # Continuation line (no txn number, same pattern)
        m2 = re.match(r'\s{10,11}A\s+C\s+(\S+)\s+(\S*)\s+(\w{3}\s+\d{1,2},\d{2})\s*(.*)', line)
        if m2 and current_num is not None:
            acct1 = m2.group(1)
            acct2 = m2.group(2)
            date_str = parse_date(m2.group(3))
            rest = m2.group(4).strip()
            if acct1 not in noise_accts:
                txn_groups[current_num]['lines'].append(('cont', acct1, acct2, date_str, rest))
                if date_str:
                    txn_groups[current_num]['date'] = date_str
            continue
        
        # Distribution line: "             A            EX.SUSP     desc    amount"
        m3 = re.match(r'\s{10,15}([AC])\s+(\S+)\s+(.*)', line)
        if m3 and current_num is not None:
            side = m3.group(1)  # A=debit context, C=credit context  
            acct = m3.group(2)
            rest = m3.group(3).strip()
            if acct not in noise_accts:
                txn_groups[current_num]['lines'].append(('dist', side, acct, rest))
            continue
    
    print(f"Parsed {len(txn_groups)} transaction groups")
    
    # Auto-create any accounts that appear in transactions but not in our chart
    all_txn_accts = set()
    for num, grp in txn_groups.items():
        for entry in grp['lines']:
            if entry[0] == 'header' or entry[0] == 'cont':
                if entry[1] not in noise_accts: all_txn_accts.add(entry[1])
                if entry[2] and entry[2] not in noise_accts: all_txn_accts.add(entry[2])
            elif entry[0] == 'dist':
                if entry[2] not in noise_accts: all_txn_accts.add(entry[2])
    
    missing = all_txn_accts - set(acct_map.keys())
    for name in sorted(missing):
        # Guess normal balance from name prefix
        if name.startswith('R.'):
            nb, desc = 'C', f'Revenue Client - {name}'
        elif name.startswith('W.'):
            nb, desc = 'D', f'WIP - {name}'
        elif name.startswith('P.'):
            nb, desc = 'C', f'Payable - {name}'
        elif name.startswith('REV'):
            nb, desc = 'C', name
        elif name.startswith(('EX','CS')):
            nb, desc = 'D', name
        else:
            nb, desc = 'D', name
        acct_map[name] = models.add_account(name, nb, desc)
    
    if missing:
        print(f"Auto-created {len(missing)} missing accounts: {sorted(missing)[:10]}...")
    
    # Now convert to Grid transactions
    # For each NV txn group, the SECOND main line (cont) has the final account assignment
    imported = 0
    skipped = 0
    errors = []
    
    for num in sorted(txn_groups.keys()):
        grp = txn_groups[num]
        if not grp['date']:
            skipped += 1
            continue
        
        headers = [e for e in grp['lines'] if e[0] == 'header']
        conts = [e for e in grp['lines'] if e[0] == 'cont']
        dists = [e for e in grp['lines'] if e[0] == 'dist']
        
        if not headers:
            skipped += 1
            continue
        
        # Use the continuation line (row 2) as the final version
        final = conts[0] if conts else headers[0]
        _, acct1, acct2, _, rest = final
        
        # Extract amount from rest
        amt_match = re.search(r'([\d,]+\.\d{2}-?)\s*$', rest)
        desc = rest
        amount = 0
        if amt_match:
            amount = parse_amt(amt_match.group(1))
            desc = rest[:amt_match.start()].strip()
        
        # Clean up description prefixes
        desc = re.sub(r'^etsf\s*', '', desc).strip()
        if not desc:
            desc = headers[0][4] if headers else ''
            amt2 = re.search(r'([\d,]+\.\d{2}-?)\s*$', desc)
            if amt2: desc = desc[:amt2.start()].strip()
            desc = re.sub(r'^etsf\s*', '', desc).strip()
        
        if amount == 0 and not dists:
            skipped += 1
            continue
        
        date_str = grp['date']
        
        # Resolve account IDs
        a1_id = acct_map.get(acct1)
        a2_id = acct_map.get(acct2) if acct2 else None
        
        if not a1_id:
            skipped += 1
            continue
        
        try:
            if acct2 and a2_id and not dists:
                # Simple 2-account transaction
                # Amount convention: negative means money leaving acct1
                if amount < 0:
                    # Acct1 credited, acct2 debited
                    models.add_transaction(date_str, '', desc, [
                        (a1_id, amount, desc),
                        (a2_id, -amount, desc)])
                elif amount > 0:
                    models.add_transaction(date_str, '', desc, [
                        (a1_id, amount, desc),
                        (a2_id, -amount, desc)])
                else:
                    skipped += 1
                    continue
            elif dists:
                # Distribution transaction - use dist lines
                txn_lines = [(a1_id, amount, desc)]
                remaining = -amount
                
                for d in dists:
                    _, d_side, d_acct, d_rest = d
                    d_amt_m = re.search(r'([\d,]+\.\d{2}-?)\s*$', d_rest)
                    if not d_amt_m: continue
                    d_amt = parse_amt(d_amt_m.group(1))
                    d_desc = d_rest[:d_amt_m.start()].strip()
                    d_desc = re.sub(r'^\*', '', d_desc).strip()
                    
                    d_acct_id = acct_map.get(d_acct)
                    if not d_acct_id: continue
                    
                    txn_lines.append((d_acct_id, d_amt, d_desc or desc))
                
                # Check balance
                total = sum(l[1] for l in txn_lines)
                if total != 0:
                    # Try to fix by adjusting
                    if abs(total) < 200 and len(txn_lines) >= 2:
                        # Small rounding - adjust last line
                        acct_id, amt, d = txn_lines[-1]
                        txn_lines[-1] = (acct_id, amt - total, d)
                    else:
                        skipped += 1
                        continue
                
                if len(txn_lines) >= 2:
                    models.add_transaction(date_str, '', desc, txn_lines)
                else:
                    skipped += 1
                    continue
            else:
                skipped += 1
                continue
            imported += 1
        except Exception as e:
            errors.append(f"Txn {num}: {e}")
            skipped += 1
    
    print(f"Imported {imported} transactions, skipped {skipped}")
    if errors:
        print(f"Errors ({len(errors)}):")
        for e in errors[:10]:
            print(f"  {e}")
    
    # Verify trial balance
    tb, dr, cr = models.get_trial_balance()
    diff = abs(dr - cr)
    print(f"Trial balance: DR={models.fmt_amount(dr)} CR={models.fmt_amount(cr)}")
    if diff == 0:
        print("✓ BALANCED!")
    else:
        print(f"✗ Out of balance by {models.fmt_amount(diff)}")

if __name__ == '__main__':
    datae = '/mnt/user-data/uploads/DATAE_NV.txt'
    output = '/home/claude/grid/WTPC/books.db'
    os.makedirs('/home/claude/grid/WTPC', exist_ok=True)
    create_wtpc_books(output, datae)
