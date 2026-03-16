# excel-formula-tracer

Trace and validate cell references across Excel spreadsheets.

An [Office Script](https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel) that analyzes any Excel cell and recursively traces its entire dependency chain — producing a color-coded, hierarchical breakdown of every precedent cell, formula, and value.

## Why

Excel's built-in **Trace Precedents** draws arrows for one level at a time and doesn't work well across sheets. This script builds the full dependency tree in one shot and writes it to a dedicated "Formula Trace" sheet you can scroll, filter, and share.

## Features

- **Recursive tracing** — follows dependencies up to a configurable depth (`MAX_TRACE_DEPTH`)
- **Cross-sheet support** — resolves references like `Sheet2!A1` and `'OH Rates Calc'!E75`
- **Range expansion** — expands `P17:P19` into individual cells (P17, P18, P19)
- **Range limit** — large ranges (configurable via `RANGE_EXPANSION_LIMIT`) are shown as collapsed notation instead of blowing up the trace
- **Circular reference detection** — marks already-visited cells with `[already traced above]`
- **Cell filtering** — optional Include/Exclude lists (via a Comments sheet) to focus or prune the trace
- **Comments integration** — optional lookup column that pulls annotations from a Comments sheet
- **Preserves formatting** — numeric formats from source cells carry over to the trace output
- **Color-coded output** — each depth level gets a distinct background color; special colors for circular refs, max depth, excluded cells, and oversized ranges

## Output

The script creates a **Formula Trace** sheet with these columns:

| Column | Description |
|--------|-------------|
| Level | Depth in the dependency tree (0 = selected cell) |
| Sheet | Sheet name where the cell lives |
| Cell | Cell address with indentation showing hierarchy |
| Status | Indicators: `[already traced above]`, `[max depth reached]`, `[not followed]`, `[range limit exceeded]` |
| Formula | Formula text (displayed as-is, not executed) |
| Value | Current cell value (preserves source number format) |
| Comments | *(Optional)* Lookup formula fetching comments from the Comments sheet |

## Usage

1. Open your workbook in Excel (desktop or web)
2. Go to **Automate → New Script** (or **Automate → Scripts**)
3. Paste the contents of `formula-trace.osts`
4. Select any cell with a formula
5. Run the script

A new **Formula Trace** sheet appears with the full dependency tree.

## Configuration

Edit the constants at the top of the script:

```typescript
const MAX_TRACE_DEPTH = 5;        // How many levels deep to trace
const RANGE_EXPANSION_LIMIT = 15; // Max cells before a range is collapsed
const HIDE_EXCLUDED_CELLS = true; // Omit excluded cells from output entirely
const COMMENTS_TAB = "Comments";  // Set to "" to disable comments integration
```

## Comments Sheet

When `COMMENTS_TAB` is set, the script looks for (or creates) a sheet with this layout:

| A (Cell) | B (Comment) | C (Include In Trace) | D (Exclude From Trace) |
|----------|-------------|----------------------|------------------------|
| `'Sheet1'!A1` | Rate source | `'Sheet1'!A1:A10` | `'Sheet2'!Z1` |

- **Include In Trace** — whitelist; only these cells are traced (all others skipped)
- **Exclude From Trace** — blacklist; these cells appear but aren't followed further
- Both support individual cells and ranges (e.g., `'Sheet Name'!E73:E85`)
- Include takes precedence: a cell must be in Include *and* not in Exclude
- The root cell (level 0) is always traced regardless of filters

## Limitations

- Uses regex-based formula parsing — may miss some edge cases in complex formulas
- Does not handle named ranges or structured references (`Table1[Column]`)
- Does not handle external workbook references (`[Book1]Sheet1!A1`)

## License

MIT
