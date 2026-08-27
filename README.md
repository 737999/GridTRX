# GridTRX

**[Watch the demo: Full accounting cycle in 15 minutes](https://youtu.be/9mmHbgEB3PQ)**

**Built for agent use as a double-entry, full-cycle accounting suite and reporting framework. You prompt in plain English -> the agent completes the books correctly.**

GridTRX is a bookkeeping system built for AI agents first and human accountants second. Feed it a bank statement, and the system produces a full set of auditable books: balance sheet, income statement, general ledger, trial balance. No subscriptions, no cloud, full privacy. Your books are one SQLite file on your own disk.

**The whole install story:** unzip this, then tell your AI agent — *"Read SKILL.md and do my books."* The agent handles the rest.

## Built by a CPA

GridTRX was built by a cross-border tax CPA practitioner who does dual-citizen and international tax returns for a living. Not a developer who read an accounting textbook. A real accountant who got tired of watching the industry bolt chatbots onto QuickBooks and call it AI. The chart of accounts, the import rules, the report structures, and the suspense workflow all come from doing real client work, not from a product spec.

## What You Get

**Financial Statements**
- Balance Sheet with opening balances, closing balances, $ change, % change
- Income Statement with period comparisons, up to 13 comparative columns
- General Ledger with running balances and cross-account references
- Trial Balance — always ties, debits equal credits
- Adjusting Journal Entry report
- Statement of Retained Earnings — computed, always current
- Sales tax reporting
- All reports export to CSV and PDF, for any time period. Just ask your agent.

**Full Accounting Cycle**
- Opening balances module for bringing an existing business across
- Import bank data from CSV, OFX, and QBO files with auto-categorization
- Customizable import rules match vendors to accounts automatically
- Unrecognized transactions land in Suspense (EX.SUSP) for review and clearing — tell the agent what they are, or clear them yourself in the browser
- Adjusting journal entries as a first-class module, including CaseWare-style AJE import
- **Perpetual retained earnings — there is no year-end close ritual.** Retained earnings are always correct by construction; you move the working year forward and the statements follow. Nothing to run, nothing to undo.
- Editable lock-date enforcement for closed periods
- Bank reconciliation with statement-by-statement continuity checks
- Sales tax code support (GST/HST with automatic net + tax splitting)

**Architecture**
- Each client is one SQLite file. Copy it, back it up. Treat it as what it is — confidential financial records. Don't transmit it over channels the client hasn't approved.
- Amounts stored as integers (cents). No floating-point rounding.
- Every transaction balances. Sum of all lines = 0. Always.
- One data layer (`models.py`) — CLI, MCP server, and browser UI all call the same functions.

**Agent safety:** destructive commands respect the lock date, every import is tagged as a single batch (undoable whole in the browser UI), and each book is snapshotted daily on first open. Agents are instructed (SKILL.md) to confirm with their human before deleting, re-importing, or changing lock dates.

## How It Works

An AI agent operates GridTRX on behalf of a human. The human never touches the software.

```
Human: "Here's my bank statement. Can you do my books?"
  ↓
Agent: Creates books → Imports bank data → Applies rules
  ↓
Agent: "I couldn't categorize 3 transactions. What's the Amazon charge for?"
  ↓
Human: "Office supplies."
  ↓
Agent: Adds rule → Deletes old entry → Re-imports → Trial balance ties
  ↓
Agent: "Books are done. Here's your balance sheet and income statement."
```

No clicking. No menus. No login. Just a conversation.

## Three Interfaces, One Engine

**MCP Server (preferred for agents)** — 58 structured JSON tools for AI agents, wrapping `models.py` directly. No text parsing, typed parameters, deterministic output.

**CLI (fallback for agents, power users)** — Interactive shell and one-shot commands. Zero dependencies beyond the Python standard library. Any terminal-based agent can drive it via subprocess.

**Browser UI (for humans)** — Flask web interface at `localhost:5000`. Ledger browsing with inline editing and keyboard navigation, report viewer with drill-down, comparative reports up to 13 columns, bank import with rule preview, reconciliation marking, print-ready PDF output, dark mode. Same database, same data.

All three hit the same `models.py` data layer. Nothing is out of sync. MCP and CLI enforce a workspace boundary when `GRIDTRX_WORKSPACE` is set.

## Quick Start

### Agents

Point your agent at this folder and say: **"Read SKILL.md and do my books."** SKILL.md and ai.txt are the complete operating manual — commands, conventions, and the rules of the house. They are maintained with the engine and are always current.

### MCP Server (Claude Desktop / Claude Code)

```json
{
  "mcpServers": {
    "gridtrx": {
      "command": "python",
      "args": ["/path/to/mcp_server.py"],
      "env": {"GRIDTRX_WORKSPACE": "/path/to/clients"}
    }
  }
}
```

Set `GRIDTRX_WORKSPACE` to restrict database access to one directory. The MCP server rejects any path outside it.

### CLI

```bash
# Create new books, then work in them
python cli.py
Grid> new ~/clients/acme "Acme Corp"
Grid/Acme Corp> post 2025-01-15 "Office supplies" 84.00 EX.OFFICE BANK.CHQ
Grid/Acme Corp> importcsv ~/downloads/jan2025.csv BANK.CHQ
Grid/Acme Corp> tb
Grid/Acme Corp> report BS

# One-shot mode for agents
python cli.py ~/clients/acme tb
python cli.py ~/clients/acme report IS 2025-01-01 2025-12-31
python cli.py ~/clients/acme importcsv ~/downloads/jan2025.csv BANK.CHQ
```

The full command set (accounts, ledgers, AJEs, imports, rules, reconciliation, opening balances, exports) is documented in [SKILL.md](SKILL.md) — the same document the agent reads.

### Browser UI

```bash
python run.py
```

Opens at `http://localhost:5000`.

## Display Format

```
 1,500.00      ← $1,500 debit
(1,500.00)     ← $1,500 credit
    —          ← zero
```

Positive = debit. Parentheses = credit. No sign-flipping. What you see is what's stored.

## Account Naming

GridTRX uses descriptive account names, not numeric codes. If source data has numeric codes (1010, 5800, etc.), ignore them and map by description.

| Prefix | Type | Examples |
|--------|------|----------|
| `BANK.xxx` | Bank accounts | `BANK.CDN`, `BANK.US`, `BANK.CHQ` |
| `REV.xxx` | Revenue | `REV`, `REV.SVC`, `REV.FOREIGN` |
| `EX.xxx` | Expenses | `EX.PHONE`, `EX.OFFICE`, `EX.WAGES` |
| `GST.xxx` | Tax accounts | `GST.IN`, `GST.OUT`, `GST.CLR` |
| `AR.xxx` | Accounts receivable | `AR` |
| `AP.xxx` | Accounts payable | `AP`, `AP.CC` |
| `SHL.xxx` | Shareholder loans | `SHL.DANA` |
| `RE.xxx` | Retained earnings | `RE`, `RE.OB` |

If no existing account matches, create one using the `EX.` or `REV.` prefix convention.

## OpenClaw Skill

GridTRX is available as an [OpenClaw skill on ClawHub](https://clawhub.ai). Any OpenClaw agent can install it and immediately handle bookkeeping tasks.

See [SKILL.md](SKILL.md) for the full skill definition.

## Requirements

**CLI only:** Python 3.9+ (standard library only — no pip install needed)

**MCP Server:** Python 3.9+ and `pip install mcp`

**Browser UI:** Python 3.9+ and `pip install flask`

All dependencies are declared in `requirements.txt`. Install before first use — nothing is installed at runtime.

## License

AGPLv3 — see [LICENSE](LICENSE)
