# SimpleExcelExporter

[![NuGet](https://img.shields.io/nuget/v/SimpleExcelExporter.svg?style=flat-square&label=nuget)](https://www.nuget.org/packages/SimpleExcelExporter/)
[![NuGet pre-releases](https://img.shields.io/nuget/vpre/SimpleExcelExporter.svg?style=flat-square&label=nuget%20pre)](https://www.nuget.org/packages/SimpleExcelExporter/)
[![CI](https://github.com/Prothesis-Dental-Solutions/SimpleExcelExporter/actions/workflows/dotnet.yml/badge.svg)](https://github.com/Prothesis-Dental-Solutions/SimpleExcelExporter/actions/workflows/dotnet.yml)
[![License](https://img.shields.io/badge/license-LGPL--3.0--or--later-blue.svg?style=flat-square)](LICENSE)

Small focused C# library for exporting .NET objects to `.xlsx` files. Built on top of [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK), it trades the SDK's full flexibility for a much simpler API aimed at one specific use case: **turning a list of annotated POCOs into a spreadsheet**.

The generated files open in:
- Microsoft Excel (Windows + macOS)
- LibreOffice Calc
- Google Sheets (up to its 10 M-cell quota)
- **Apple Numbers** — since `1.5.0` the output is strict ECMA-376 compliant

## Installation

```shell
dotnet add package SimpleExcelExporter
```

Target framework: **.NET 8**.

## Two ways to use it

### 1. Annotated POCOs (recommended for most use cases)

Annotate your model classes with attributes from `SimpleExcelExporter.Annotations`, then let the exporter discover the structure via reflection.

```csharp
using SimpleExcelExporter;
using SimpleExcelExporter.Annotations;
using SimpleExcelExporter.Definitions;

public class Player
{
    [CellDefinition(CellDataType.String)]
    [Header(typeof(TeamRes), "PlayerNameColumnName")]
    [Index(1)]
    public string? PlayerName { get; set; }

    [CellDefinition(CellDataType.Date)]
    [Header(typeof(TeamRes), "DateOfBirthColumnName")]
    [Index(2)]
    public DateTime? DateOfBirth { get; set; }

    [CellDefinition(CellDataType.Number)]
    [Header(typeof(TeamRes), "NumberOfVictoryColumnName")]
    [Index(3)]
    public int? NumberOfVictory { get; set; }
}

public class Team
{
    private ICollection<Player>? _players;

    [SheetName(typeof(TeamRes), "SheetName")]
    [EmptyResultMessage(typeof(TeamRes), "EmptyResultMessage")]
    public ICollection<Player> Players => _players ??= new HashSet<Player>();
}

// …and then:
var team = new Team { Players = { /* … */ } };
using var stream = new FileStream("team.xlsx", FileMode.Create);
new SpreadsheetWriter(stream, team).Write();
```

Header labels come from a resource file (`TeamRes.resx`) so the same workbook can be exported in several languages.

See the full worked example at [SimpleExcelExporterExample](https://github.com/Prothesis-Dental-Solutions/SimpleExcelExporterExample).

### 2. `WorkbookDfn` — explicit, attribute-free

If annotating your domain classes is not an option, build a `WorkbookDfn` by hand and pass that to the writer. You keep full control of sheet names, column ordering, cell types, and data.

```csharp
var workbook = new WorkbookDfn();
var sheet = new WorksheetDfn("Players");
workbook.Worksheets.Add(sheet);

sheet.ColumnHeadings.Cells.Add(new CellDfn("Name"));
sheet.ColumnHeadings.Cells.Add(new CellDfn("Date of birth"));

var row = new RowDfn();
row.Cells.Add(new CellDfn("Alexandre", cellDataType: CellDataType.String));
row.Cells.Add(new CellDfn(new DateTime(1974, 2, 1), cellDataType: CellDataType.Date));
sheet.Rows.Add(row);

using var stream = new FileStream("players.xlsx", FileMode.Create);
new SpreadsheetWriter(stream, workbook).Write();
```

A more complete example is in [`src/ConsoleApp/Program.cs`](src/ConsoleApp/Program.cs).

## Supported attributes

| Attribute | Applied on | Purpose |
|---|---|---|
| `[SheetName]` | property returning `IEnumerable<T>` | Localised sheet name from a resource file |
| `[EmptyResultMessage]` | idem | Message shown in the sheet when the collection is empty |
| `[Header]` | property | Localised column header |
| `[Index]` | property | Column order within a sheet (1-indexed) |
| `[CellDefinition]` | property | Cell data type (Date, Number, Boolean, Percentage, Time, String) |
| `[MultiColumn]` | property of collection type | Expands a sub-collection into repeating columns |
| `[IgnoreFromSpreadSheet]` | property | Skip this property during export |

## What's in the generated XLSX?

Output conforms to ECMA-376 strict mode since `1.5.0`:

- Every `<c>` carries `r="A1"`-style references (Apple Numbers requires it).
- Non-empty string cells route through `xl/sharedStrings.xml` (deduplicated).
- Empty cells are omitted from the sheet XML — readers infer the position (matches Excel's native output).
- Default style attributes (`s="0"`, `t="n"`) are omitted to keep files compact.
- `docProps/core.xml` is populated with `dc:creator`, `dcterms:created`, `dcterms:modified`.
- Deterministic writes when running under CI (`ContinuousIntegrationBuild=true`), plus a companion `.snupkg` with [Source Link](https://github.com/dotnet/sourcelink) for step-into-source debugging.

## Performance and file-size characteristics

On a 1 M-row annotated fixture (see [`BENCHMARK_RESULTS.md`](BENCHMARK_RESULTS.md) for the full table and methodology):

- Runtime is dominated by the reflection walk of annotated objects (~60 % of total) rather than the XML write.
- Files sit ~100 % larger than the pre-`1.5.0` output — the price of the mandatory `r="A1"` attribute on every cell. Three spec-conformant optimisations ([skip empty cells, omit default `s="0"`, omit default `t="n"`](BENCHMARK_RESULTS.md)) partially offset this and keep a 1 M × 20 fixture under Google Sheets' 100 MB upload limit.

A reproducible benchmark harness is at [`scripts/benchmark.sh`](scripts/benchmark.sh).

## Repository layout

```
src/
  SimpleExcelExporter/         the library (published as the NuGet package)
  ConsoleApp/                  runnable harness used for manual + perf checks
test/
  SimpleExcelExporterTests/    48 NUnit tests (unit + OOXML compliance + behaviour)
scripts/
  benchmark.sh                 perf & file-size benchmark across library versions
.github/
  workflows/                   CI (on push/PR) and Release (on tag)
  dependabot.yml               weekly NuGet + Actions dependency scans
```

## Releasing

Version bumps, NuGet publication, and GitHub Releases are driven entirely by pushing a `v*` tag that matches `<Version>` in the csproj. See [RELEASING.md](RELEASING.md) for the step-by-step flow, SemVer guidance, and troubleshooting.

## License

[LGPL-3.0-or-later](LICENSE).
