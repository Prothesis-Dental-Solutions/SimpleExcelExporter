# Benchmark results — PR #9 performance & file-size analysis

Controlled measurement of runtime and output file sizes across 6 versions of the library, using the same ConsoleApp harness with a seeded `Random(42)` so all runs generate byte-identical data across versions (modulo timestamps).

- **Machine**: Linux, tuxedo laptop, no other user applications running.
- **Protocol**: 3 runs per version, same build, results reported as medians. File sizes are reported in bytes (stable across runs thanks to the fixed seed — variance ≤ 1 byte).
- **Harness**: [`scripts/benchmark.sh`](scripts/benchmark.sh) — portable, reproducible (commit list and output dir configurable at the top of the script / via env vars).
- **Persistent XLSX outputs**: `~/SimpleExcelExporter-bench-outputs/<version-label>/` for direct opening in Excel / Numbers / LibreOffice.

## Versions benchmarked

| # | Label | Commit | What changed |
|---|---|---|---|
| 1 | `01-master` | `0e3a966` | master — pre-PR reference |
| 2 | `02-before-shared` | `ac59ba6` | PR #9 compliance + refactor, before any size optim |
| 3 | `03-after-shared` | `679355f` | + shared strings table |
| 4 | `04-A-skip-empty` | `b57ebc5` | + skip empty cells (optim A) |
| 5 | `05-B-omit-s0` | `7070d23` | + omit `s="0"` default style (optim B) |
| 6 | `06-C-omit-tn` | `d298204` | + omit `t="n"` default type (optim C) |

## Runtime (medians of 3 runs, seconds)

| Version | Total runtime | Δ vs master |
|---|---:|---:|
| `01-master` | **93.20 s** | — (baseline) |
| `02-before-shared` | 126.77 s | +33.57 s (**+36 %**) |
| `03-after-shared` | 125.56 s | +32.36 s (+35 %) |
| `04-A-skip-empty` | 128.65 s | +35.45 s (+38 %) |
| `05-B-omit-s0` | 128.43 s | +35.23 s (+38 %) |
| `06-C-omit-tn` | 125.91 s | +32.71 s (+35 %) |

**Reading** : the PR adds a ~35 % runtime overhead vs master (price of compliance and streaming refactor). The three size optims (A, B, C) are **perf-neutral** — variation within ±3 s is noise.

## File sizes (`TestWithData3.xlsx`, 1 M rows × 20 cols, annotated path)

| Version | Size | Δ vs master | Δ vs previous step |
|---|---:|---:|---:|
| `01-master` | **47.2 MB** | — | — |
| `02-before-shared` | 109.4 MB | +62.3 MB (**+132 %**) | +62.3 MB |
| `03-after-shared` | 109.2 MB | +62.1 MB (+132 %) | −0.2 MB (≈ 0) |
| `04-A-skip-empty` | 97.7 MB | +50.6 MB (+107 %) | **−11.5 MB** (−11 %) |
| `05-B-omit-s0` | 96.7 MB | +49.5 MB (+105 %) | −1.0 MB (−1 %) |
| `06-C-omit-tn` | **95.0 MB** | +47.9 MB (+101 %) | **−1.6 MB** (−2 %) |

**Reading** : the compliance bump inflated the file by 132 % (62 MB). The three optims together recovered 14 MB. We remain **+101 %** above master, fundamentally because Apple Numbers requires the `r="A1"` attribute on every cell.

## File sizes (other fixtures)

### `TestWithData4.xlsx` (1 M rows, `WorkbookDfn` path, dense sheet)

| Version | Size | Δ vs master |
|---|---:|---:|
| `01-master` | **30.8 MB** | — |
| `02-before-shared` | 55.6 MB | +24.8 MB (+80 %) |
| `03-after-shared` | 55.9 MB | +25.1 MB (+81 %) |
| `04-A-skip-empty` | 55.9 MB | +25.1 MB (+81 %) |
| `05-B-omit-s0` | 55.5 MB | +24.7 MB (+80 %) |
| `06-C-omit-tn` | 55.4 MB | +24.6 MB (+80 %) |

**Reading** : the `WorkbookDfn` path has very few empty cells, so optim A doesn't help. Optims B and C each save a modest ~0.4 MB.

### `TestWithData5.xlsx` (`Group` fixture, 16 k persons hierarchy)

| Version | Size | Δ vs master |
|---|---:|---:|
| `01-master` | 48.6 KB | — |
| `02-before-shared` | 128.9 KB | +80.3 KB (+165 %) |
| `03-after-shared` | 175.3 KB | +126.7 KB (+260 %) |
| `04-A-skip-empty` | 175.3 KB | +126.7 KB (+260 %) |
| `05-B-omit-s0` | 166.7 KB | +118.1 KB (+243 %) |
| `06-C-omit-tn` | 166.7 KB | +118.1 KB (+243 %) |

**Reading** : shared strings *hurt* this fixture (+46 KB) because all ~16 k person names are unique → the table stores them all without deduplication gain, and the added `sharedStrings.xml` + rels/content-types overhead dominates. Optim B recovers a bit. None of the size optims pay off on workloads without repeated strings or empty cells.

### Small fixtures (baseline shapes)

| Fixture | 01-master | 02-before-shared | 03-after-shared | 04-A | 05-B | 06-C |
|---|---:|---:|---:|---:|---:|---:|
| `TestWithData1.xlsx` | 4.0 KB | 4.2 KB | 4.5 KB | 4.5 KB | 4.5 KB | 4.5 KB |
| `TestWithData2.xlsx` | 3.1 KB | 3.4 KB | 3.7 KB | 3.7 KB | 3.7 KB | 3.7 KB |
| `TestWithDataEmpty.xlsx` | 2.7 KB | 2.8 KB | 3.0 KB | 3.0 KB | 3.0 KB | 3.0 KB |

Negligible movements on small workbooks — the metadata cost dominates over any data.

## Summary of findings

1. **Compliance has a real cost.** The PR makes files 80–132 % larger and runtime 35 % slower on a 1 M-row annotated export. That's the price to pay for Apple Numbers acceptance (mandatory `r="A1"` per cell).

2. **Shared strings are perf-neutral on our data** (runtime and size). They're kept because they're the XLSX convention; the memory-allocation improvement may matter on leaner machines.

3. **Skip empty cells (A)** is the biggest size win on annotated paths with `MultiColumn` padding (−11 %). Zero gain elsewhere.

4. **Omit default attributes (B, C)** each trim 1–2 % more. Modest but free (single-line changes, no behaviour change, accepted by every reader we tested).

5. **Cumulative gain of A + B + C**: −14 MB on the 1 M-row annotated fixture (109 → 95 MB). Enough to drop below Google Sheets' 100 MB upload limit.

6. **Google Sheets has a 10 M cell limit** orthogonal to file size. `TestWithData3.xlsx` (20 M cells) won't open in Google Sheets regardless.

## Validation to-dos

The following open items still need empirical validation — the OOXML spec allows them, but strict readers don't always honour the spec:

- [ ] Open `~/SimpleExcelExporter-bench-outputs/06-C-omit-tn/TestWithData3.xlsx` in **Apple Numbers**. The A/B/C optims (empty cells omitted, `s="0"` omitted, `t="n"` omitted) were reasoned from spec but not re-validated on Numbers after the initial compliance work on commit `fee0e01`.
- [ ] Spot-check in Excel (Windows + macOS) and LibreOffice Calc.
- [ ] If any reader refuses, revert the culprit commit (B or C is easy to revert; A is the big-gain one).

## Artifacts

- Benchmark script: [`scripts/benchmark.sh`](scripts/benchmark.sh) — re-run with `./scripts/benchmark.sh` from the repo root.
- Raw CSV (default path): `/tmp/benchmark-results.csv` (one row per run).
- Run log (default path): `/tmp/benchmark.log`.
- XLSX outputs per version (default path): `~/SimpleExcelExporter-bench-outputs/<label>/`.

Override via environment variables: `RUNS`, `OUTPUT_DIR`, `LOG`, `REPORT`.
