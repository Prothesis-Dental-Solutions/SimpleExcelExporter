<!-- Describe what the PR does and why. Link any related issues. -->

## Checklist

- [ ] Tests added or updated
- [ ] Breaking change? If yes, describe the migration in the PR description and plan a MAJOR version bump
- [ ] If this PR ships user-visible changes: `<Version>` in `src/SimpleExcelExporter/SimpleExcelExporter.csproj` bumped accordingly (see [RELEASING.md](../RELEASING.md))
- [ ] After merge: create and push a tag `v<version>` on `master` to trigger the NuGet publish (the release is not automatic on merge)
