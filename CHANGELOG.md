# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2026-05-16

### Added

- MCP server (Python, stdio transport) that wraps think-cell's `.ppttc` JSON
  automation: builds `.ppttc` files from plain data and converts them into
  PowerPoint decks (`.pptx`) via think-cell's `ppttc.exe`.
- 8 MCP tools: `create_chart`, `build_presentation`, `create_auto_deck`,
  `set_deck_branding`, `convert_to_pptx`, `validate_ppttc`, `list_chart_types`,
  and `diagnose_thinkcell`.
- 8 chart types: `waterfall`, `bar`, `stacked_bar`, `line`, `scatter`,
  `mekko`, `area`, and `combo` — each with per-type input validation and data
  shaping.
- `create_auto_deck`: zero-setup multi-slide decks built on think-cell's
  official automation template, requiring no PowerPoint template preparation
  or element naming. Each slide carries a fixed layout (slide title plus a
  left and right chart, both on a date axis); each slide becomes one `.ppttc`
  entry.
- `set_deck_branding`: re-themes the auto-deck template — removes the
  think-cell logo and applies a configurable accent palette (a
  Big-Four-inspired palette by default). Only standard PowerPoint parts are
  touched, so named think-cell elements stay intact.
- `build_presentation`: combines multiple charts and slides into a single
  `.ppttc` file, with per-chart names emitted as one `.ppttc` entry per slide.
- Structural `.ppttc` validator (`validate_ppttc`) that checks JSON structure,
  required keys, cell types, and row consistency.
- Read-only think-cell environment diagnostics (`diagnose_thinkcell`):
  Windows registry and filesystem checks for the install, PowerPoint, the COM
  add-in registration, disabled add-ins, licence, and template content — no
  processes are launched.
- `list_chart_types`: catalog of every chart type with its required/optional
  parameters, expected data shape, and a ready-to-use example.
- Self-contained test suite (`python tests.py`) and GitHub Actions CI.

[0.1.0]: https://github.com/Hectirry/thinkcell-mcp/releases/tag/v0.1.0
