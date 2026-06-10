# data-analysis

A skill for Claude Code and Cowork that turns raw data files into interactive HTML reports with charts, statistics, and plain-language insights - readable by anyone, not just data experts.

## Installation

### Claude Code

```bash
mkdir -p ~/.claude/skills
git clone <repo-url> ~/.claude/skills/data-analysis
```

Or copy manually:
```bash
mkdir -p ~/.claude/skills/data-analysis
cp SKILL.md ~/.claude/skills/data-analysis/
```

### Cowork

Install via **Settings > Capabilities > Install skill** and drop the `.skill` file.

## Usage

Just ask naturally - the skill triggers automatically:

```
analyze my sales.csv and tell me which products perform best
```

```
look at employes.xlsx - are the salaries consistent with seniority?
```

```
which items in my inventory.json are about to run out of stock?
```

### Voice and language

The skill auto-detects the user's language and writes insights in the same language (French prompt → French report).

## What it produces

A self-contained **HTML file** you can open in any browser:

| Section | Content |
|---------|---------|
| Header | Dataset name, row/column count, date |
| Summary | 3–5 plain-language insights answering your actual question |
| KPI cards | Key totals, averages, and maximums |
| Charts | Bar, line, histogram - powered by Chart.js |
| Correlations | Strongest relationships between numeric columns |
| Data quality | Missing values and anomaly alerts |
| Sample table | First 5 rows |

## Supported input formats

| Format | Extension |
|--------|-----------|
| CSV / TSV | `.csv`, `.tsv` |
| Excel | `.xlsx`, `.xls` |
| JSON | `.json` (array of objects, columnar object, or JSON Lines) |

Large files (>100k rows) are sampled at 10,000 rows - noted in the report header. Excel workbooks contribute their first sheet; the other sheet names are included in the analysis output. CSV and TSV are read as UTF-8 (with or without BOM), Windows-1252, or Latin-1, with the delimiter detected automatically. A semicolon file that uses a decimal comma is parsed as a number.

## Key features

- **Plain-language insights** - no statistical jargon; "sales peaked in November (+68%)" not "the November mean exceeds μ + 2σ"
- **Anomaly detection** - flags outliers using IQR method
- **Time series** - auto-detects date columns and plots trends over time
- **Correlation analysis** - identifies relationships between numeric columns
- **Two-pass audit** - drafts insights, then audits them for vagueness before writing the final report

## Author

Concepteur Donatien Byakombe · 2026

## License

MIT
