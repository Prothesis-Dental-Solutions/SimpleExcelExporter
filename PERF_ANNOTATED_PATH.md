# Performance — annotated-object path

TODO list to speed up `new SpreadsheetWriter(stream, annotatedObject)` + `.Write()` — the path used in production by PCLOUD.

## Context

The library exposes two constructors:

- `SpreadsheetWriter(Stream, WorkbookDfn)` — caller passes a fully built workbook definition. Fast.
- `SpreadsheetWriter(Stream, object)` — the **annotated-object path**, used by PCLOUD. The constructor walks the object graph with reflection to build the `WorkbookDfn` internally. Slow on large inputs.

## Baseline (measured on the ConsoleApp, 1 million `Player` rows)

| Phase | Duration |
|---|---|
| `Preparing the SpreadsheetWriter` (≈ `BuildWorkbook` via reflection) | **55 s** |
| `Writing the Excel file` (`.Write()`) | **120 s** |
| Total for one sheet | **~175 s** |

Comparison on the same workbook shape via the `WorkbookDfn` path (no reflection): `Preparing` = 11 s, `Writing` = 36 s. The 4× slowdown on the annotated path is entirely in the reflection and in the post-reflection object graph it produces.

## Hot spots identified in `SpreadSheetWriter.cs`

### 1. `Type.GetProperties()` called per-player instead of per-type — BuildWorkbook

```
foreach (player) CalculateMaxNumberOfElement(player);   // calls GetProperties() internally
foreach (player) { type.GetProperties(); AddHeaderCells...; }   // 1M calls, same result every time
foreach (player) { type.GetProperties(); AddCells...;        }   // 1M calls, same result every time
```

Every call hits the CLR reflection machinery. `PropertyInfo[]` is immutable per `Type`.

**Fix** — a `Dictionary<Type, PropertyInfo[]>` cache (field on `SpreadsheetWriter`). Get-or-add pattern.

**Estimated gain: 10–15 s on BuildWorkbook.**

### 2. String key in `_cachedAttributes` causes ~100M allocations — BuildWorkbook

```csharp
private T? GetAttributeFrom<T>(PropertyInfo propertyInfo)
{
    var key = $"{propertyInfo.Module.MetadataToken}_{propertyInfo.MetadataToken}_{typeof(T).Name}";
    if (_cachedAttributes.TryGetValue(key, out var cachedAttribute)) { ... }
    ...
}
```

For 1M players × ~30 properties × 4–5 attribute types queried per property, the interpolation allocates on the order of **120M short strings** just for dictionary keys.

**Fix** — switch to `Dictionary<(PropertyInfo, Type), Attribute?>` (the tuple is a `ValueTuple` struct, no heap allocation) or to `ConditionalWeakTable<PropertyInfo, ...>`.

**Estimated gain: 5–10 s on BuildWorkbook + large GC pressure reduction.**

### 3. `CalculateMaxNumberOfElement` reflects per-player

Same story as #1 but buried inside the MultiColumn width calculator. It queues objects and re-runs `GetProperties()` for each. Once #1 is in place, `CalculateMaxNumberOfElement` can use the same cache — zero duplicate work.

**Gain folded into #1.**

### 4. `ColumnReferenceHelper.ToLetters(columnIndex)` uncached — Write

```csharp
CellReference = $"{ColumnReferenceHelper.ToLetters(columnIndex)}{rowIndex}",
```

1M rows × ~30 columns = **30M calls**. Each one creates a `StringBuilder`, loops, calls `ToString()`. Columns 1–100 produce exactly 100 distinct strings, each returned hundreds of thousands of times.

**Fix** — pre-populate a static `string[] ColumnLettersCache = new string[N]` at class initialization, indexed by columnIndex. Fallback to the current algorithm for columnIndex > N (> 100 covers the vast majority of real-world sheets; Excel's max is 16384).

**Estimated gain: 5–10 s on Write.**

### 5. `CellReference` string concatenation per cell — Write

Even after #4, each cell still does `$"{letters}{rowIndex}"` which allocates a new string. 30M allocations remaining.

**Fix** — either precompute `letters + row` on the fly inside the worksheet writer (bypassing the intermediate `CellReference` property), or write directly to the `XmlWriter` without building the string. More invasive, needs to touch `CreateCell` / `WriteWorksheets`.

**Estimated gain: 2–5 s on Write. Lower priority.**

## Priority plan

### Round 1 — three quick wins (~30–40 min)

- [ ] #1 Cache `PropertyInfo[]` by `Type`
- [ ] #2 Tuple key in `_cachedAttributes`
- [ ] #4 Static lookup table for column letters

All three are:
- Non-invasive (internal caches, no API surface change)
- Semantics-preserving (same output XML byte-for-byte)
- Covered by the existing 31 tests — no new tests strictly needed, but adding a perf-oriented benchmark (see below) documents the gain

**Target after Round 1: 100–120 s total instead of 175 s (~35% speed-up).**

### Round 2 — structural (later)

- [ ] #5 Bypass intermediate `CellReference` string, write directly to `XmlWriter`
- [ ] Stream `SheetData` row-by-row instead of materializing the whole sheet in memory before writing (point #3 of the original PR review, deferred)

## Validation protocol

1. **Before**: run `src/ConsoleApp` on the current `fix/xlsx-ooxml-compliance-numbers` tip, record the `Total execution time`.
2. Apply one optimization per commit.
3. **After each commit**: re-run the console app, record the delta.
4. **After each commit**: run `dotnet test` — all 31 tests must stay green.
5. **After each commit**: diff two generated .xlsx files (before/after) byte-for-byte. They must be identical — any difference in output means the refactor broke semantics.

For a more rigorous benchmark, introduce BenchmarkDotNet in a dedicated test project later. The console app is enough for a first pass.

## Non-goals

- Changing the public API (the two constructor signatures stay as-is).
- Altering the XML output format.
- Introducing parallelism (speculative; would complicate the shared `_cachedAttributes` and `_multiColumnAttribute` state).

## Process

- Best merged **after** PR #9 lands on `master`, as a dedicated PR titled `perf: speed up annotated-object path`.
- Version bump: patch (`1.5.1`) — no API change, no new feature, pure performance.
