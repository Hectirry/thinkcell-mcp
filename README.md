# think-cell MCP Server

An [MCP](https://modelcontextprotocol.io) server that wraps think-cell's JSON
data automation. It lets Claude build think-cell automation files (`.ppttc`)
from plain data and convert them into PowerPoint decks (`.pptx`) using
think-cell's `ppttc.exe`.

## Tools

| Tool | Purpose |
| --- | --- |
| `create_chart` | Build a `.ppttc` file for a single chart. |
| `build_presentation` | Combine many charts/slides into one `.ppttc` file. |
| `create_auto_deck` | Build a multi-slide `.ppttc` deck with **no template setup**. |
| `set_deck_branding` | Recolour the auto-deck template and drop the think-cell logo. |
| `convert_to_pptx` | Run `ppttc.exe` to turn a `.ppttc` into a `.pptx`. |
| `validate_ppttc` | Structurally validate a `.ppttc` document. |
| `list_chart_types` | Describe every supported chart type. |
| `diagnose_thinkcell` | Check why think-cell automation isn't working. |

Supported chart types: `waterfall`, `bar`, `stacked_bar`, `line`, `scatter`,
`mekko`, `area`, `combo`.

## Requirements

- **Windows** with [think-cell](https://www.think-cell.com/) installed — a
  separately licensed product. It provides `ppttc.exe`, by default under
  `C:\Program Files (x86)\think-cell`. Set the `THINKCELL_DIR` environment
  variable if think-cell is installed elsewhere.
- **Python 3.10+**.
- **Microsoft PowerPoint** — `ppttc.exe` drives it during conversion.

> Building and validating `.ppttc` files works on any OS. Only
> `convert_to_pptx` and `diagnose_thinkcell` require Windows + think-cell +
> PowerPoint.

## Installation

1. Clone the repository and enter it:

   ```powershell
   git clone https://github.com/Hectirry/thinkcell-mcp.git
   cd thinkcell-mcp
   ```

2. Install the dependencies:

   ```powershell
   python -m pip install -r requirements.txt
   ```

3. (Optional) Confirm the server starts — it waits on stdio; press `Ctrl+C`
   to stop:

   ```powershell
   python server.py
   ```

## Register with Claude

The server speaks MCP over **stdio**. Add the block below to the `mcpServers`
section of your Claude configuration, adjusting the paths for your machine.

- **Claude Desktop:** edit `%APPDATA%\Claude\claude_desktop_config.json`.
- **Claude Code / VS Code:** copy `.mcp.json.example` to `.mcp.json`, or run
  `claude mcp add`.

```json
{
  "mcpServers": {
    "thinkcell": {
      "command": "python",
      "args": ["/absolute/path/to/thinkcell-mcp/server.py"]
    }
  }
}
```

`command` must resolve to a Python 3.10+ interpreter — use the full path to
`python.exe` if `python` is not on your `PATH`. If you already have other
servers, add only the inner `"thinkcell": { ... }` entry. Restart Claude
after editing the config.

## Quickstart

The fastest end-to-end flow needs **no PowerPoint template setup**:
`create_auto_deck` builds the `.ppttc` against think-cell's bundled official
template, and `convert_to_pptx` renders the finished `.pptx`.

**1. Build a `.ppttc` with `create_auto_deck`** — one slide with a title, a
left chart and a right chart, each with its own title:

```json
{
  "output_name": "Quickstart",
  "slides": [
    {
      "slide_title": "Quarterly Review",
      "left_title": "Market share",
      "right_title": "Orders",
      "left_chart": {
        "categories": ["Q1", "Q2", "Q3"],
        "series": [
          {"name": "Our brand", "values": [52, 54, 57]},
          {"name": "Competitor", "values": [21, 20, 18]}
        ]
      },
      "right_chart": {
        "categories": ["Q1", "Q2", "Q3"],
        "series": [
          {"name": "Delivered", "values": [740, 760, 805]}
        ]
      }
    }
  ]
}
```

Chart `categories` accept **either** plain string labels like `"Q1"` (a
category axis) **or** ISO dates `"YYYY-MM-DD"` (a date axis) — the axis kind
is auto-detected. This call returns a `ppttc_path` in the `output/` folder.

**2. Render the `.pptx` with `convert_to_pptx`** — pass the `ppttc_path` from
step 1:

```json
{
  "ppttc_path": "C:\\Users\\you\\thinkcell-mcp\\output\\Quickstart.ppttc"
}
```

**Result:** a finished PowerPoint deck (`Quickstart.pptx`) with real
think-cell charts — no template preparation and no manual element naming.

See *Tool reference and examples* below for every option of each tool.

## How think-cell automation works

think-cell's automation does **not** create charts from nothing. The flow is:

1. You prepare a PowerPoint **template** (`.pptx`) that already contains
   think-cell charts. Each chart and text field is given a **name** in the
   think-cell *Data links* / element naming UI.
2. This server produces a `.ppttc` file — JSON that maps each **name** to a
   data table.
3. `ppttc.exe` opens the template, pushes the data into the matching named
   elements, and writes the finished `.pptx`.

**Key point:** the `chart_name` in a `.ppttc` must match the name assigned to
a think-cell element in the template. To name an element: select the chart in
PowerPoint and type a unique name in the **"Name" field of its think-cell mini
toolbar**, then press Enter. `create_chart` takes a `chart_name` (default
`Chart1`); `build_presentation` lets you set names per chart (or auto-assigns
`Chart1`, `Chart2`, ...). **If the names do not match, `ppttc.exe` runs but
updates nothing.** The chart *type* (waterfall, line, ...) is decided by the
chart you placed in the template — this server validates and shapes the data
to match, and the per-type rules guard against malformed input.

**There is no general API to "control" the think-cell add-in.** Unlike some
Office add-ins, think-cell exposes no COM/scripting interface for creating
charts or driving its ribbon — the `.ppttc` / `ppttc.exe` pipeline is the
*only* supported automation, and it just pushes data into a prepared
template. If a conversion seems to do nothing, the cause is almost always
environmental (the add-in isn't registered, PowerPoint disabled it, the
licence lapsed, or the template has no think-cell charts). Run
`diagnose_thinkcell` to pinpoint it.

Generated `.ppttc` files are written to the `output/` folder next to
`server.py`.

## Zero-setup decks (no PowerPoint, no naming)

Naming think-cell elements inside PowerPoint is the one manual step the
`.ppttc` pipeline normally requires. The **`create_auto_deck`** tool removes
it. think-cell ships an official automation template whose elements are
already named. This project does **not** redistribute that file; instead
`autodeck.py` copies it from your local think-cell install
(`<THINKCELL_DIR>/ppttc/template.pptx`) into `templates/thinkcell_auto.pptx`
the first time `create_auto_deck` runs. Its five named elements are fixed:

| Element | Role |
| --- | --- |
| `SlideTitle` | Slide title text field. |
| `LeftChartTitle` | Title above the left chart. |
| `RightChartTitle` | Title above the right chart. |
| `LeftChart` | Left chart — date or category axis, `percentage` series. |
| `RightChart` | Right chart — date or category axis, `number` series. |

The user never opens PowerPoint or names anything — they just supply data.
Each slide becomes one `.ppttc` entry (one slide). The resulting `.ppttc` can
be handed straight to `convert_to_pptx`.

### Branding

think-cell's template carries think-cell's own logo and a green/blue accent
palette. On first use the bundled copy is **automatically de-branded** — the
logo is removed and a professional, blue-anchored palette (drawn from the
Big Four firms' brand colours) is applied. Both live in standard PowerPoint
parts (the theme and the slide master), not in the named think-cell elements,
so re-skinning never affects automation. Use **`set_deck_branding`** to apply
your own accent colours or to toggle the logo. Per-series `fill` values in
`create_auto_deck` still override the palette for individual series.

## Tool reference and examples

### 1. `create_chart`

Build a `.ppttc` file for one chart.

```json
{
  "chart_type": "bar",
  "title": "Revenue by Region",
  "data": [
    {"category": "North", "Revenue": 120},
    {"category": "South", "Revenue": 95},
    {"category": "West", "Revenue": 140}
  ],
  "template_path": "C:\\decks\\regions.pptx",
  "chart_name": "RevenueChart"
}
```

Returns, e.g.:

```json
{
  "success": true,
  "ppttc_path": "C:\\Users\\you\\thinkcell-mcp\\output\\Revenue_by_Region.ppttc",
  "chart_type": "bar",
  "chart_name": "RevenueChart",
  "template": "C:\\decks\\regions.pptx",
  "errors": []
}
```

`template_path` is optional; when omitted the placeholder `template.pptx` is
written and must be corrected before conversion. `chart_name` (default
`Chart1`) **must match the name given to the chart inside think-cell** — select
the chart in PowerPoint and set the "Name" field in its think-cell mini
toolbar. If the names differ, `ppttc.exe` runs but updates nothing.

`options` is an optional object for chart-type-specific settings:
`category_key` (any categorical chart) forces which row key is the category
column; `line_series` (combo only) lists which series render as lines. An
unrecognized option key is rejected with a clear error.

### 2. `build_presentation`

Combine several charts/slides into a single `.ppttc` file. Each chart entry
may set `chart_name` to match a named element in the template.

```json
{
  "template_path": "C:\\decks\\quarterly.pptx",
  "output_name": "Q3_Review",
  "slides": [
    {
      "title": "Revenue Bridge",
      "charts": [
        {
          "chart_type": "waterfall",
          "chart_name": "BridgeChart",
          "data": [
            {"category": "Opening", "Value": 100},
            {"category": "New", "Value": 40},
            {"category": "Churn", "Value": -15},
            {"category": "Closing", "Value": 125}
          ]
        }
      ]
    },
    {
      "title": "Revenue Trend",
      "charts": [
        {
          "chart_type": "line",
          "chart_name": "TrendChart",
          "data": [
            {"category": "Jul", "2025": 100},
            {"category": "Aug", "2025": 112},
            {"category": "Sep", "2025": 125}
          ]
        }
      ]
    }
  ]
}
```

Returns the `.ppttc` path plus `slide_count` and `chart_count`.

### 3. `create_auto_deck`

Build a multi-slide `.ppttc` deck **without preparing or naming any
template** — it always uses the bundled `templates/thinkcell_auto.pptx`
(see *Zero-setup decks* above). Every slide has the same fixed layout: a
slide title, plus a left chart and a right chart, each with its own title.
All fields are optional, but every slide must carry at least one. Each chart's
`categories` may be **either** ISO dates `"YYYY-MM-DD"` (producing a **date
axis**) **or** plain string labels like `"Q1"` (producing a **category
axis**) — the axis kind is auto-detected per chart. The left chart renders
`percentage` values, the right chart renders plain `number` values. Each slide
becomes one `.ppttc` entry (one slide).

Input schema — `slides` is a non-empty list of slide objects:

| Slide field | Type | Notes |
| --- | --- | --- |
| `slide_title` | string | optional — slide heading |
| `left_title` | string | optional — title above the left chart |
| `right_title` | string | optional — title above the right chart |
| `left_chart` | object | optional — left chart (`percentage` values) |
| `right_chart` | object | optional — right chart (`number` values) |

A chart object has `categories` (a non-empty list of either ISO date strings
`"YYYY-MM-DD"` for a date axis, or plain string labels for a category axis)
and `series` (non-empty list of series objects). Each series object has `name`
(string), `values` (list of numbers, one per category) and an optional
`fill` (hex colour like `"#ff0000"` applied to every point of that series).

```json
{
  "output_name": "Germany_Review",
  "slides": [
    {
      "slide_title": "Competition: Germany",
      "left_title": "Market share",
      "right_title": "Our orders (10K)",
      "left_chart": {
        "categories": ["2024-01-01", "2025-01-01"],
        "series": [
          {"name": "Our brand", "values": [55, 55]},
          {"name": "Competitor 1", "values": [4, 4.5]}
        ]
      },
      "right_chart": {
        "categories": ["2024-01-01", "2025-01-01"],
        "series": [
          {"name": "Pending", "values": [0, 14], "fill": "#ff0000"},
          {"name": "Delivered", "values": [760, 747]}
        ]
      }
    }
  ]
}
```

Returns `{"success": true, "ppttc_path": "...", "slide_count": 1,
"template": "...", "errors": []}`, or `success: false` with a list of
structured validation errors. `output_name` is optional (default
`auto_deck`); the `.ppttc` is written to the `output/` folder.

### 4. `set_deck_branding`

Re-brand the template that `create_auto_deck` uses. Rewrites the bundled
template's theme accent palette and removes the think-cell logo — once — so
every future deck inherits it. Only standard PowerPoint parts are touched;
named think-cell elements are left intact.

```json
{
  "accent_colors": ["#00338D", "#0091DA", "#86BC25", "#E88D14", "#797878", "#FFDB00"],
  "remove_logo": true
}
```

`accent_colors` takes up to six `#RRGGBB` colours (theme accent1..accent6);
omit it to apply the default Big-Four-inspired palette. Returns
`{"success": true, "template": "...", "accents": [...], "logo_removed": bool,
"errors": []}`.

### 5. `convert_to_pptx`

Run `ppttc.exe` to produce a `.pptx`. Internally runs:
`ppttc.exe <input.ppttc> -o <output.pptx>`.

```json
{
  "ppttc_path": "C:\\Users\\you\\thinkcell-mcp\\output\\Q3_Review.ppttc",
  "output_path": "C:\\decks\\Q3_Review_final.pptx"
}
```

Returns `{"success": true, "pptx_path": "...", "stdout": "...", "stderr": "...", "error": null}`,
or a structured error (missing input, `ppttc.exe` not found, non-zero exit,
timeout). `output_path` is optional and defaults to the input path with a
`.pptx` extension.

### 6. `validate_ppttc`

Structurally validate a `.ppttc` document. Pass either `ppttc_path` or
`ppttc_json`.

```json
{
  "ppttc_path": "C:\\Users\\you\\thinkcell-mcp\\output\\Q3_Review.ppttc"
}
```

Returns `{"valid": true, "errors": []}`, or `valid: false` with a list of
problems (bad JSON, missing `template`/`data`/`name`/`table`, malformed
cells, inconsistent row widths, ...).

### 7. `list_chart_types`

No input. Returns every supported chart type with its required/optional
parameters, the expected data shape, and a ready-to-use example:

```json
{
  "count": 8,
  "chart_types": {
    "bar": {
      "required_params": ["chart_type", "data", "title"],
      "optional_params": ["template_path", "chart_name", "options"],
      "data_shape": "List of row objects. Each row has one category key ...",
      "example": { "chart_type": "bar", "title": "Revenue by Region", "data": [ ... ] }
    }
  }
}
```

### 8. `diagnose_thinkcell`

Explain why think-cell automation isn't working. Runs read-only Windows
registry and filesystem checks — no processes are launched. `template_path`
is optional.

```json
{
  "template_path": "C:\\decks\\quarterly.pptx"
}
```

Returns a report whose `summary` gives an overall `status` (`ok` / `warning`
/ `error`) and actionable `recommendations`, plus per-area `checks`:

```json
{
  "platform_supported": true,
  "summary": {
    "status": "error",
    "headline": "think-cell automation cannot work until the errors below are resolved.",
    "recommendations": ["In PowerPoint: File > Options > Add-ins > ..."]
  },
  "checks": {
    "thinkcell_installation": { "status": "ok", "summary": "...", "details": { ... } },
    "powerpoint": { "status": "ok", "summary": "..." },
    "com_addin": { "status": "error", "summary": "The think-cell COM add-in is ...", "recommendation": "..." },
    "office_disabled_items": { "status": "ok", "summary": "..." },
    "license": { "status": "info", "summary": "..." },
    "template": { "status": "warning", "summary": "...", "details": { ... } }
  }
}
```

Checks performed: think-cell installation and `ppttc.exe`; PowerPoint
presence; the think-cell **COM add-in registration and `LoadBehavior`**;
PowerPoint's **disabled / crashing add-in list**; think-cell registry/licence
presence; and (with `template_path`) whether the `.pptx` is valid and
actually contains think-cell content.

## Chart data shapes

Every chart type takes `data` as a **list of row objects**.

- **Categorical** charts (`bar`, `stacked_bar`, `line`, `area`, `waterfall`,
  `mekko`, `combo`): each row has one category key (`category`, `label` or
  `name`) plus one or more numeric series keys. **Every row must define the
  same set of series** — a row that omits a series is rejected, because a
  missing value would otherwise be charted as `0` rather than as a gap.

  ```json
  [{"category": "Q1", "Plan": 100, "Actual": 92},
   {"category": "Q2", "Plan": 110, "Actual": 118}]
  ```

- **`scatter`**: each row is a point with numeric `x` and `y`, plus an
  optional numeric `size` (bubble chart) and optional `label`.

  ```json
  [{"label": "A", "x": 9.99, "y": 1200},
   {"label": "B", "x": 14.99, "y": 850}]
  ```

Call `list_chart_types` for the exact rules and an example per type.

## Project structure

```
thinkcell-mcp/
  server.py            MCP server entry point (FastMCP, stdio transport)
  autodeck.py          zero-setup .ppttc builder (powers create_auto_deck)
  branding.py          re-themes the auto-deck template (set_deck_branding)
  converter.py         ppttc.exe wrapper with structured error handling
  validator.py         .ppttc JSON structural validation
  diagnostics.py       think-cell environment diagnostics (registry/files)
  tests.py             self-contained test suite (run: python tests.py)
  charts/
    __init__.py        chart registry + catalog helpers
    base.py            shared ChartBuilder base class + .ppttc writer
    waterfall.py       one module per chart type
    bar.py
    stacked_bar.py
    line.py
    scatter.py
    mekko.py
    area.py
    combo.py
  templates/           thinkcell_auto.pptx is copied here from your think-cell
                       install on first use (not shipped with the repo)
  output/              generated .ppttc / .pptx files (git-ignored)
  pyproject.toml       packaging metadata
  requirements.txt     Python dependencies
  .mcp.json.example    sample Claude MCP registration
  LICENSE              MIT licence
  README.md            this file
```

## Notes and limitations

- All tools return structured objects. On failure they return
  `success: false` (or `valid: false`) with an `errors` list — never a bare
  string.
- This server builds and validates `.ppttc` files on any OS, but
  `convert_to_pptx` requires Windows with think-cell and PowerPoint
  installed.
- The data table layout follows the `thinkcell` library's `.ppttc` format.
  Exact rendering (totals on a waterfall, which series is a line in a combo,
  Mekko column widths) is governed by the chart you set up in the template.

## License

Released under the [MIT License](LICENSE).

This is an independent, unofficial integration. "think-cell" is a trademark
of think-cell Sales GmbH & Co. KG; this project is not affiliated with or
endorsed by think-cell. It requires a separately licensed think-cell
installation and does not redistribute any think-cell software or template
files.
