# CLAUDE.md

## Companion repo: SimpleExcelExporterExample

`../SimpleExcelExporterExample/` (https://github.com/Prothesis-Dental-Solutions/SimpleExcelExporterExample) is the public "getting started" sample, linked from this repo's `README.md`. It consumes `SimpleExcelExporter` via NuGet, not via `ProjectReference`, so it does **not** track `master` automatically.

When bumping the library — `<Version>` or `<TargetFramework>` in `src/SimpleExcelExporter/SimpleExcelExporter.csproj` — update the example repo's `SimpleExcelExporterExample.csproj` in the same change set:

- `<PackageReference Include="SimpleExcelExporter" Version="..." />` → new lib version
- `<TargetFramework>` → match the lib

If the public API changes, or the snippets in this repo's `README.md` change, also refresh `Program.cs` / `WorkbookDfnExample.cs` and the example README so the public-facing examples don't go stale.
