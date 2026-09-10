# CLAUDE.md — Digital Products Project (Avangrid)

## Project Context

- **Project**: Digital Products (formerly "SAP CoE") at Avangrid
- **Client**: Avangrid (energy utility, Iberdrola subsidiary)
- **Consultant**: NTT DATA
- **Working directory**: `/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/`
- **Python**: Always use `.venv/bin/python` for all commands
- **Key deliverables**: Prior Art Review PPTX, Project Kickoff PPTX, Discovery Sessions Workbook (Excel), Operating Model Observatory (web app)

---

## MANDATORY RULES — Knowledge Base (KB) System

**These rules are non-negotiable and apply to every conversation, every task, every query.**

### Rule 1: NEVER Read Source Files Directly for Content

**DO NOT** use `markitdown`, `python-pptx`, `pdfplumber`, `openpyxl`, `Read`, `cat`, or any other tool to read `.pptx`, `.pdf`, or `.xlsx` files for the purpose of answering questions, analyzing content, or extracting information.

**ALL content queries MUST go through the KB system.**

This applies to every file format:
- `.pptx` — NEVER open with python-pptx or markitdown to read content
- `.pdf` — NEVER open with pdfplumber, PyPDF2, or the Read tool to read content
- `.xlsx` — NEVER open with openpyxl or pandas to read content

### Rule 2: ALL Content Queries Go Through the KB CLI

```bash
# Hybrid search (vector + BM25) — the default and most common command
PYTHONPATH=src .venv/bin/python -m sapkb.cli search "query"

# Filter by project
PYTHONPATH=src .venv/bin/python -m sapkb.cli search "query" --project "P&S / P&T"

# Show full chunk text (use when you need detailed content)
PYTHONPATH=src .venv/bin/python -m sapkb.cli search "query" --text

# Check for contradictions across sources
PYTHONPATH=src .venv/bin/python -m sapkb.cli search "query" --divergences

# KB statistics
PYTHONPATH=src .venv/bin/python -m sapkb.cli stats

# Ingest a single file
PYTHONPATH=src .venv/bin/python -m sapkb.cli ingest path/to/file

# Ingest with visual pipeline (Claude Haiku for text, Sonnet for complex visuals)
PYTHONPATH=src .venv/bin/python -m sapkb.cli ingest path/to/file --visual

# Batch ingest a directory
PYTHONPATH=src .venv/bin/python -m sapkb.cli ingest path/to/dir --glob "*.pptx"
```

### Rule 3: The ONLY Exception — Editing Files

The **only** time you may open a source file directly is when the task is to **edit or modify** it:
- `python-pptx` to **modify** slides (add, rearrange, change content)
- `openpyxl` to **modify** Excel workbooks
- Other libraries to **write/modify** files

**Reading for content or analysis is NEVER an exception.** If you need to understand what is in a file before editing it, search the KB first.

### Rule 4: If the KB Does Not Have the Answer, Say So

When the user asks about content from any project document:

1. **ALWAYS** search the KB first
2. If the KB returns relevant results, use them
3. If the KB does **not** have enough information, **tell the user** — do NOT silently fall back to reading the file directly
4. Suggest ingesting the file if it has not been ingested yet: `PYTHONPATH=src .venv/bin/python -m sapkb.cli ingest path/to/file --visual`

### Rule 5: Projects in the KB

The KB contains documents from these projects:
- **P&S / P&T** (Planning & Scheduling / Planning & Tracking)
- **Digital CoE** (Digital Center of Excellence)
- **Infrastructure**
- **Renewables DTO** (Digital Transformation Office)
- **SAP Digital Products**

Use the `--project` flag to filter searches when relevant.

### Rule 6: NEVER Call the Anthropic API for Chat or Analysis

All chat, analysis, and reasoning happens inside Claude Code (included in the plan). **Do NOT** call the Anthropic API (`anthropic.Anthropic()`, `client.messages.create()`, etc.) for conversation or analysis tasks.

The Anthropic API is used **exclusively** during the ingest pipeline:
- **Claude Haiku** — for text-heavy slides (cheaper, faster)
- **Claude Sonnet** — for visually complex slides (diagrams, charts, Venn diagrams, flows)

This is handled automatically by `PYTHONPATH=src .venv/bin/python -m sapkb.cli ingest --visual`.

### Rule 7: Visual Ingest Pipeline for New Files

When ingesting new `.pptx` files, prefer the visual pipeline:

```bash
PYTHONPATH=src .venv/bin/python -m sapkb.cli ingest path/to/file.pptx --visual
```

This renders each slide as an image and uses Claude vision to describe complex visuals that text extraction would miss (org charts, process flows, Venn diagrams, architecture diagrams).

---

## Quick Reference

| Task | Command |
|------|---------|
| Search the KB | `PYTHONPATH=src .venv/bin/python -m sapkb.cli search "query"` |
| Search with full text | `PYTHONPATH=src .venv/bin/python -m sapkb.cli search "query" --text` |
| Search by project | `PYTHONPATH=src .venv/bin/python -m sapkb.cli search "query" --project "Digital CoE"` |
| Find contradictions | `PYTHONPATH=src .venv/bin/python -m sapkb.cli search "query" --divergences` |
| KB stats | `PYTHONPATH=src .venv/bin/python -m sapkb.cli stats` |
| Ingest file (visual) | `PYTHONPATH=src .venv/bin/python -m sapkb.cli ingest file.pptx --visual` |
| Ingest directory | `PYTHONPATH=src .venv/bin/python -m sapkb.cli ingest dir/ --glob "*.pptx"` |
