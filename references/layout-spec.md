# Layout specification

These numbers are the template. Change them only if you intend to change the look,
and change them in `scripts/build.py → SPEC` so both the DOCX and the PDF move together.

## Page

| Item | Value |
|---|---|
| Paper | A4 — 21.0 × 29.7 cm (595.28 × 841.89 pt) |
| Top margin | 1.70 cm |
| Bottom margin | 1.60 cm |
| Left / right margin | 1.90 cm each |
| Text column width | 17.20 cm |

## Type

| Element | Face | Size | Weight | Notes |
|---|---|---|---|---|
| Name | Times New Roman | 16 pt | Bold | Centred |
| Contact line | Times New Roman | body | Regular | Centred, fields separated by `  \|  ` |
| Availability line | Times New Roman | body | Bold | Centred — this is the line screeners act on |
| Section heading | Times New Roman | 11.5 pt | Bold | ALL CAPS, 0.75 pt rule underneath |
| Body / bullet text | Times New Roman | body (starts at 10 pt) | Regular | |
| Entry line | Times New Roman | body | Bold | date · org · role · location |
| Entry role on second line | Times New Roman | body | Italic | Only when the row cannot fit on one line |
| Bullet dot | Times New Roman | body × 0.80 | Regular | `●` (U+25CF) |

## Spacing

| Item | Value |
|---|---|
| After name | 2.0 pt |
| After contact / availability | 2.0 pt |
| Before section heading | 8.0 pt |
| After section heading rule | 3.0 pt |
| Before entry | 2.0 pt |
| After entry row | 1.0 pt |
| After bullet | 1.0 pt |
| PDF line advance | 1.16 × body size |

## Columns and indents

| Item | Value |
|---|---|
| Date column width | 3.35 cm (left tab) |
| Org column ends / role starts | 8.55 cm (left tab) |
| Location | right-aligned at 17.20 cm |
| Bullet left indent | 0.55 cm |
| Bullet hanging indent | −0.45 cm |
| Bullet dot x-offset | 0.10 cm from margin |

## Row layout logic

An entry row is `date · org · role · location`. It stays on one line only when:

1. `org` finishes before the 8.55 cm role column (with a 6 pt gap), **and**
2. `role + location` fit in the space from 8.55 cm to the right margin.

Otherwise the row degrades in this order:

- **org too long** → date and location on line 1, org wraps underneath at the 3.35 cm column.
- **role + location too long** → location stays right-aligned on line 1, role drops to line 2 in italic.

## Auto-fit

`body` starts at 10 pt and drops in 0.1 pt steps (floor 8.0 pt) until the document fits
`--max-pages`. The builder prints the chosen size. Below 9.0 pt it prints a hint, because
past that point the right fix is to cut content, not to shrink type.

## Fonts

Windows: `C:\Windows\Fonts\times.ttf`, `timesbd.ttf`, `timesi.ttf`, `timesbi.ttf`.
Elsewhere the builder falls back to the PDF base-14 Times equivalents (`tiro` / `tibo` /
`tiit` / `tibi`), which have identical metrics, so measurement and wrapping behave the same.
