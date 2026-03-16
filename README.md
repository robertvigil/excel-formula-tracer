# excel-formula-tracer

Trace and validate cell references across Excel spreadsheets.

An [Office Script](https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel) that analyzes any Excel cell and recursively traces its entire dependency chain — producing a color-coded, hierarchical breakdown of every precedent cell, formula, and value.

## Why

Excel's built-in **Trace Precedents** draws arrows for one level at a time and doesn't work well across sheets. This script builds the full dependency tree in one shot and writes it to a dedicated "Formula Trace" sheet you can scroll, filter, and share.

## Features

- **Whitelist-based tracing** — only cells listed in the "Trace Config" tab are followed; everything else appears as `[not in scope]`
- **Recursive tracing** — follows dependencies up to a configurable depth (`MAX_TRACE_DEPTH`)
- **Cross-sheet support** — resolves references like `Sheet2!A1` and `'OH Rates Calc'!E75`
- **Range expansion** — expands `P17:P19` into individual cells (P17, P18, P19)
- **Range limit** — large ranges (configurable via `RANGE_EXPANSION_LIMIT`) are shown as collapsed notation instead of blowing up the trace
- **Circular reference detection** — marks already-visited cells with `[already traced above]`
- **Comments integration** — Column B of Trace Config provides optional annotations per cell (single cell references only, not ranges)
- **Preserves formatting** — numeric formats from source cells carry over to the trace output
- **Color-coded output** — each depth level gets a distinct background color; special colors for circular refs, max depth, not-in-scope cells, and oversized ranges

## Output

The script creates a **Formula Trace** sheet with these columns:

| Column | Description |
|--------|-------------|
| Level | Depth in the dependency tree (0 = root cell) |
| Sheet | Sheet name where the cell lives |
| Cell | Cell address with indentation showing hierarchy |
| Status | Indicators: `[already traced above]`, `[max depth reached]`, `[not in scope]`, `[range limit exceeded]` |
| Formula | Formula text (displayed as-is, not executed) |
| Value | Current cell value (preserves source number format) |
| Comments | Lookup formula fetching comments from the Trace Config tab |

## Usage

1. Open your workbook in Excel (desktop or web)
2. Go to **Automate → New Script** (or **Automate → Scripts**)
3. Paste the contents of `formula-trace.osts`
4. Select any cell with a formula and run the script
5. A **Trace Config** tab is created with your selected cell as the root (A2)
6. Populate Column A with additional cells/ranges to include in the trace scope
7. Optionally add comments in Column B for single cell references
8. Run the script again to generate the **Formula Trace** sheet

## Configuration

Edit the constants at the top of the script:

```typescript
const MAX_TRACE_DEPTH = 5;        // How many levels deep to trace
const RANGE_EXPANSION_LIMIT = 15; // Max cells before a range is collapsed
const CONFIG_TAB = "Trace Config"; // Name of the configuration tab
```

## Trace Config Sheet

The script uses a "Trace Config" tab with this layout:

| A (Cell) | B (Comment) |
|----------|-------------|
| `'Sheet1'!A1` | Rate source |
| `'Sheet1'!B5:B20` | |
| `'Sheet2'!C3` | Updated quarterly |

- **Column A (Cell)** — whitelist of cells/ranges to trace. Supports individual cells (`'Sheet1'!A1`) and ranges (`'Sheet Name'!E73:E85`). **A2 is always the root cell** (starting point for the trace).
- **Column B (Comment)** — optional annotations. Comments only apply to single cell references; comments on range entries are disregarded since they can't map to a single row in the output.

If the Trace Config tab doesn't exist when you run the script, it's created automatically with the active cell as A2.

## Limitations

- Uses regex-based formula parsing — may miss some edge cases in complex formulas
- Does not handle named ranges or structured references (`Table1[Column]`)
- Does not handle external workbook references (`[Book1]Sheet1!A1`)

## Credits

This project was vibe coded with [Claude Code](https://docs.anthropic.com/en/docs/claude-code) (Claude Opus), Anthropic's AI-powered coding agent. From architecture to implementation — every line of code, this README, and even the commit messages were authored by Claude.

## License

MIT
