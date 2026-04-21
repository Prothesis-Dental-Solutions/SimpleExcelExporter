# Releasing

This document describes how to publish a new version of `SimpleExcelExporter` to NuGet.

## TL;DR

1. Open a PR that bumps `<Version>` in `src/SimpleExcelExporter/SimpleExcelExporter.csproj`.
2. Merge the PR to `master`.
3. Create a git tag that matches the version (prefixed with `v`) and push it.
4. The `Release` workflow publishes to NuGet and creates a GitHub Release.

No package is ever published without an explicit tag. There is no rush — tag only when you are ready.

## Step by step

### 1. Bump the version

Edit `src/SimpleExcelExporter/SimpleExcelExporter.csproj`:

```xml
<Version>1.5.0</Version>
```

Follow [Semantic Versioning](https://semver.org):

- **PATCH** (`1.4.4` → `1.4.5`): backwards-compatible bug fix.
- **MINOR** (`1.4.x` → `1.5.0`): new backwards-compatible capability. A change that produces different output for an existing call (new spreadsheet features, new compatibility) counts as MINOR even if no new public method is added.
- **MAJOR** (`1.x.x` → `2.0.0`): breaking change of the public API.

Open a PR, get it reviewed, merge to `master`.

### 2. Create and push the tag

Once merged:

```bash
git checkout master
git pull
git tag -a v1.5.0 -m "Release 1.5.0"
git push origin v1.5.0
```

The tag format is always `v` followed by the exact version string in the csproj.

### 3. The workflow does the rest

`.github/workflows/release.yml` triggers on any tag matching `v*` and:

1. Verifies that the tag matches `<Version>` in the csproj. If not, the release aborts before anything is published.
2. Restores, builds, and tests the solution in `Release` configuration.
3. Packs `SimpleExcelExporter.<version>.nupkg`.
4. Pushes it to [nuget.org](https://www.nuget.org/packages/SimpleExcelExporter) using the `NUGET_API_KEY` repository secret and `--skip-duplicate` (safe to re-run).
5. Creates a GitHub Release on the tag with auto-generated notes.

## Pre-releases

For testing a version in a downstream project before committing to a stable release, use a SemVer pre-release suffix:

```xml
<Version>1.5.0-alpha.1</Version>
```

```bash
git tag -a v1.5.0-alpha.1 -m "Pre-release 1.5.0-alpha.1"
git push origin v1.5.0-alpha.1
```

Iterate with `-alpha.2`, `-alpha.3`, then `-beta.1`, `-rc.1` as the version stabilizes. Consumers must reference the exact version (or use a wildcard like `1.5.0-alpha.*`) because NuGet hides pre-releases by default.

When the pre-release campaign is over:

```xml
<Version>1.5.0</Version>
```

Tag `v1.5.0`. The stable release supersedes the pre-releases (which stay published but can be unlisted from the nuget.org UI if desired).

## Troubleshooting

### The release workflow failed on "Verify tag matches csproj version"

The tag and the csproj `<Version>` disagree. Either:

- Delete the mismatched tag and re-tag with the correct version:
  ```bash
  git tag -d v1.5.0
  git push origin :refs/tags/v1.5.0
  git tag -a v1.5.0 -m "Release 1.5.0"
  git push origin v1.5.0
  ```
- Or, if the csproj is wrong, open a quick PR to fix it, then re-tag.

### `dotnet nuget push` says the version already exists

`--skip-duplicate` turns this into a success; no action needed. If you genuinely need to change a published package, increment the version — NuGet does not allow overwriting a published package.

### I forgot to tag and merged a release-bumping PR

Nothing was published. Create the tag on `master` any time — the workflow will run at that moment.
