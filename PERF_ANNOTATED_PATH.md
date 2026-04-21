# Performance — annotated-object path

Tracks the work done on branch `perf/annotated-object-path` to speed up the annotated-object path (`new SpreadsheetWriter(stream, annotatedObject)` + `.Write()`), the code path used by PCLOUD.

## Context

The library exposes two constructors:

- `SpreadsheetWriter(Stream, WorkbookDfn)` — caller passes a fully built workbook definition. Fast.
- `SpreadsheetWriter(Stream, object)` — the **annotated-object path**, used by PCLOUD. The constructor walks the object graph with reflection to build the `WorkbookDfn` internally. Slow on large inputs.

## Hot spots identified during the PR #9 review

### 1. `Type.GetProperties()` called per-player instead of per-type

```
foreach (player) CalculateMaxNumberOfElement(player);   // calls GetProperties() internally
foreach (player) { type.GetProperties(); AddHeaderCells...; }   // 1M calls, same result every time
foreach (player) { type.GetProperties(); AddCells...;        }   // 1M calls, same result every time
```

**Status: tested and reverted (commit `63156ab`, dropped during rebase).**

Added a `Dictionary<Type, PropertyInfo[]>` cache and a `GetTypeProperties(Type)` helper, replacing the hot-loop call sites. Controlled 3-run benchmark showed **no measurable gain** (~1 s slower vs baseline, within noise). The CLR already caches `PropertyInfo[]` internally, so the application-level cache was redundant. A second benchmark pass that rebuilt the branch without `#1` confirmed the final runtime was within 0.4 s of the previous result — dropping `#1` did not regress anything.

Decision: dropped from the PR to avoid dead optimisation code.

### 2. String key in `_cachedAttributes` — ~100 M allocations

```csharp
private T? GetAttributeFrom<T>(PropertyInfo propertyInfo)
{
    var key = $"{propertyInfo.Module.MetadataToken}_{propertyInfo.MetadataToken}_{typeof(T).Name}";
    if (_cachedAttributes.TryGetValue(key, out var cachedAttribute)) { ... }
    ...
}
```

For 1 M players × ~30 properties × 4-5 attribute types queried per property, the interpolation allocates on the order of **120 M short strings** just for dictionary keys.

**Status: applied (commit `9a2a229`). Main gain of the PR.**

Switched to `Dictionary<(PropertyInfo, Type), Attribute?>` — the tuple is a `ValueTuple` struct, no heap allocation. Tuple equality compares by reference on both components, which is exactly what the string key was simulating.

**Measured gain: −8.9 s (−7.0 %)** on the 1 M-row fixture (126.29 s → 117.43 s).

### 3. `ColumnReferenceHelper.ToLetters(columnIndex)` uncached

```csharp
CellReference = $"{ColumnReferenceHelper.ToLetters(columnIndex)}{rowIndex}",
```

1 M rows × ~30 columns = **30 M calls**. Each one creates a `StringBuilder`, loops, calls `ToString()`. The result for a given `columnIndex` is immutable and tiny.

**Status: applied (commit `b9d1bae`).**

Added a static `string[] Cache` of size 16 384 (Excel's hard column cap). Pre-populated at class initialization with the computed letters for every valid column (A..XFD). `ToLetters()` hits the array directly for the common case and keeps the original algorithm as a fallback past the cache size.

The static footprint is roughly 450 KB of interned strings — negligible next to the multi-MB workbooks the library produces.

**Measured gain: −6.5 s (−5.5 %)** on the 1 M-row fixture (117.43 s → 110.97 s).

### 4. `CalculateMaxNumberOfElement` reflects per-player

Same story as #1 but buried inside the MultiColumn width calculator. Would have folded into #1's gain if #1 had produced one. Since #1 didn't move the needle, this was not revisited — the CLR cache is good enough here too.

**Status: not applied.**

### 5. `CellReference` string concatenation per cell

Even after #3, each cell still does `$"{letters}{rowIndex}"` which allocates a new string. 30 M allocations remaining.

**Status: not applied (Round 2 candidate).**

## Measured results

Controlled benchmark (`scripts/benchmark.sh`, 3 runs per state, seeded `Random(42)`, machine idle during runs).

| State | Runtime (median) | Δ vs baseline | File size (`TestWithData3.xlsx`) |
|---|---:|---:|---:|
| baseline `1.5.0-alpha.1` | 126.29 s | — | 95.0 MB |
| + `#2` tuple key | **117.43 s** | −8.9 s (−7.0 %) | 95.0 MB (± 1 B) |
| + `#3` 16 384-entry letters cache | **110.97 s** | **−15.3 s (−12.1 %)** | 95.0 MB (± 1 B) |

Numbers come from the rebased branch (without `#1`), so the gains attributed to `#2` and `#3` are what actually ship in this PR. An earlier benchmark pass that included `#1` as a stepping stone reached 110.58 s — essentially the same final runtime (within 0.4 s), which is why dropping `#1` was safe.

File sizes are identical across all states (variance ≤ 2 bytes from ZIP compression timing noise) — optims are behaviour-preserving.

## Non-goals

- Changing the public API (the two constructor signatures stay as-is).
- Altering the XML output format.
- Introducing parallelism (speculative; would complicate the shared `_cachedAttributes` and `_multiColumnAttribute` state).

## Process followed

- One optimisation per commit, each verified against the 48-test suite and exercised through `scripts/benchmark.sh`.
- The rejected attempt (#1) was dropped via `git rebase --onto` rather than landed as a revert commit — the branch was never pushed with #1 in it.
- Version bumped to `1.5.0-alpha.3`: the `1.5.0` stable has not shipped yet, so this pre-release continues the `1.5.0-alpha.*` series rather than jumping ahead to `1.5.1`. `1.5.0-alpha.2` is already tagged on master (the pre-release that carried only the initial post-`alpha.1` bump, no perf work), so this branch lands on top as `alpha.3`. PCLOUD will validate this pre-release the same way as `alpha.1`, and a successful run tags the stable `1.5.0`.

## Round 2 (deferred)

- `#5` — skip the intermediate `CellReference` string, write directly to `XmlWriter`. Roughly 30 M string concatenations would go away. Needs touching `CreateCell` and `WriteWorksheets`.
- Stream `SheetData` row-by-row instead of materialising the entire sheet in memory before writing. Addresses the memory footprint rather than CPU time; more invasive.
- Consider BenchmarkDotNet in a dedicated test project for sub-second-level micro-benchmarks (GC allocation counts, per-operation time).

## Outputs retained

Per-state `.xlsx` files from this benchmark live at `~/SimpleExcelExporter-bench-outputs-perf/<label>/` for manual inspection.
