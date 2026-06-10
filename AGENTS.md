# AGENTS.md

Guidance for AI coding agents (Claude Code, Cowork, etc.) working in this repository.

## What this repo is

A Claude Code / Cowork skill implemented as Markdown + a bundled Python script pattern. The runtime artifact is `SKILL.md`: the agent reads its YAML frontmatter (metadata + allowed tools) followed by the instructions. There is no build step.

## Key files

- `SKILL.md` - the skill itself. YAML frontmatter (`name`, `version`, `description`, `allowed-tools`) followed by the analysis workflow, Python script template, HTML template, and edge case guidance. **This is the source of truth.**
- `README.md` - for humans: installation, usage, feature table, supported formats.
- `AGENTS.md` - this file.

## The maintenance contract

`SKILL.md` and `README.md` must stay in sync. When you change behavior or content:

- **Python script:** if you change the analysis logic (new stat, new chart type, new format support), update both the script in SKILL.md and the "Supported input formats" or "Key features" table in README.md.
- **HTML template:** if you change the template structure or color palette, note it in the version history at the bottom of SKILL.md.
- **Version:** bump `version:` in SKILL.md frontmatter and add a line to the "Version history" section at the bottom of SKILL.md. Bump README if features changed.
- **Edge cases:** if you add or remove an edge case, update the "Edge cases and what NOT to do" section in SKILL.md.

## Editing SKILL.md

- Preserve valid YAML frontmatter (formatting and indentation matter).
- The Python script inside the skill is a **template** - it should be general enough to handle any CSV/Excel/JSON, not hardcoded to specific column names.
- The HTML template must remain self-contained (no external dependencies except Chart.js from cdnjs.cloudflare.com).
- Keep the two-pass process (analyze → audit → report) intact - it is the core quality mechanism.

## Adding a new output format

If you add support for a new input format (e.g. Parquet, SQLite):
1. Add the loading logic to the Python script in SKILL.md
2. Add the extension to the "Supported input formats" table in README.md
3. Add an edge case note if the format has quirks
4. Bump version (minor bump, e.g. 1.0.0 → 1.1.0)
