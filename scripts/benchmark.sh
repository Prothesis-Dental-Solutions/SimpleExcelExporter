#!/bin/bash
# Perf & file-size benchmark harness used to produce BENCHMARK_RESULTS.md.
#
# What it does:
#   1. Snapshots the current src/ConsoleApp/ (so every version is benchmarked with the
#      same seeded data generator).
#   2. For each version in STATES below: git checkout, overlay the snapshot on ConsoleApp,
#      build in Release, run the ConsoleApp RUNS times, capture total runtime and the file
#      sizes of every .xlsx produced.
#   3. On the last run of each version, copies the produced .xlsx files to
#      $OUTPUT_DIR/<label>/ so they can be opened in Excel / Numbers / LibreOffice.
#   4. Returns to the starting branch.
#
# Usage:
#   ./scripts/benchmark.sh
#
# Environment variables (all optional):
#   OUTPUT_DIR   Destination for the generated .xlsx files per version.
#                Default: $HOME/SimpleExcelExporter-bench-outputs
#   RUNS         Runs per state (median reported in BENCHMARK_RESULTS.md).
#                Default: 3
#   LOG          Run log path. Default: /tmp/benchmark.log
#   REPORT       CSV output path. Default: /tmp/benchmark-results.csv
#
# Expected duration: ~40 min for 6 states × 3 runs on the default fixture
# (TestWithData3 = 1 M rows × 20 cols).
#
# Notes:
#   - The STATES array references specific commit SHAs. Update it if you want to
#     benchmark a different slice of history.
#   - The script restores the starting branch on completion, but force-checkout
#     (-f) discards any uncommitted changes in tracked files. Commit or stash
#     before running.
#   - Requires a seeded Random in ConsoleApp (committed in 487f810) for
#     reproducible file sizes between runs.

set -e

# Resolve the repo root from this script's location, so callers can invoke it
# from anywhere.
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$SCRIPT_DIR/.." && pwd)"

OUTPUT_DIR="${OUTPUT_DIR:-$HOME/SimpleExcelExporter-bench-outputs}"
RUNS="${RUNS:-3}"
LOG="${LOG:-/tmp/benchmark.log}"
REPORT="${REPORT:-/tmp/benchmark-results.csv}"
SNAPSHOT="/tmp/benchmark-consoleapp-snapshot"

# Versions to benchmark: "label sha".
# Add or remove entries to change the coverage.
STATES=(
  "01-master          0e3a966"
  "02-before-shared   ac59ba6"
  "03-after-shared    679355f"
  "04-A-skip-empty    b57ebc5"
  "05-B-omit-s0       7070d23"
  "06-C-omit-tn       d298204"
)

STARTING_BRANCH="$(cd "$REPO" && git rev-parse --abbrev-ref HEAD)"
cd "$REPO"

echo "state,run,total_s,size_Data1,size_Data2,size_Data3,size_Data4,size_Data5,size_DataEmpty" > "$REPORT"
: > "$LOG"

# Snapshot the current (seeded) ConsoleApp. This is what gets layered onto every
# checked-out version, so every run exercises the same data generator — only the
# library under test (src/SimpleExcelExporter/) varies between states.
rm -rf "$SNAPSHOT"
mkdir -p "$SNAPSHOT"
cp -r "$REPO/src/ConsoleApp/." "$SNAPSHOT/"
echo "=== $(date +%H:%M:%S) snapshot saved to $SNAPSHOT ===" >> "$LOG"

mkdir -p "$OUTPUT_DIR"

run_state() {
  local label=$1 sha=$2
  local version_dir="$OUTPUT_DIR/$label"
  mkdir -p "$version_dir"
  {
    echo "=== $(date +%H:%M:%S) $label ($sha) ==="
    git checkout -f "$sha" 2>&1
    cp -r "$SNAPSHOT/." "$REPO/src/ConsoleApp/" 2>&1
    dotnet restore --nologo --verbosity quiet 2>&1 || true
  } >> "$LOG" 2>&1

  if ! dotnet build --no-restore --configuration Release --nologo --verbosity quiet >> "$LOG" 2>&1; then
    echo "=== $(date +%H:%M:%S) BUILD FAILED on $label ($sha) — skipping ===" >> "$LOG"
    echo "$label,BUILD-FAILED,0,0,0,0,0,0,0" >> "$REPORT"
    return 0
  fi

  for i in $(seq 1 "$RUNS"); do
    rm -rf "$REPO/src/ConsoleApp/bin/release/net8.0/ExampleOutput-"* 2>/dev/null || true
    echo "  $(date +%H:%M:%S) $label run $i/$RUNS..." >> "$LOG"
    local out total
    out=$(cd "$REPO/src/ConsoleApp/bin/release/net8.0" && ./ConsoleApp 2>&1)
    total=$(echo "$out" | grep "Total execution time" | awk '{print $4}')
    local dir="$REPO/src/ConsoleApp/bin/release/net8.0"
    local outputdir
    outputdir=$(ls -d "$dir"/ExampleOutput-* 2>/dev/null | head -1)
    local s1 s2 s3 s4 s5 se
    s1=$(stat -c%s "$outputdir/TestWithData1.xlsx" 2>/dev/null || echo "0")
    s2=$(stat -c%s "$outputdir/TestWithData2.xlsx" 2>/dev/null || echo "0")
    s3=$(stat -c%s "$outputdir/TestWithData3.xlsx" 2>/dev/null || echo "0")
    s4=$(stat -c%s "$outputdir/TestWithData4.xlsx" 2>/dev/null || echo "0")
    s5=$(stat -c%s "$outputdir/TestWithData5.xlsx" 2>/dev/null || echo "0")
    se=$(stat -c%s "$outputdir/TestWithDataEmpty.xlsx" 2>/dev/null || echo "0")
    echo "$label,$i,$total,$s1,$s2,$s3,$s4,$s5,$se" >> "$REPORT"
    echo "    $(date +%H:%M:%S) total=${total}s data3=${s3}B" >> "$LOG"

    if [ "$i" = "$RUNS" ]; then
      cp "$outputdir/"*.xlsx "$version_dir/" 2>/dev/null || true
      echo "    $(date +%H:%M:%S) xlsx copied to $version_dir" >> "$LOG"
    fi
  done
}

for entry in "${STATES[@]}"; do
  # Split the entry on whitespace into label + sha.
  # shellcheck disable=SC2086
  set -- $entry
  run_state "$1" "$2"
done

git checkout -f "$STARTING_BRANCH" >> "$LOG" 2>&1
echo "=== $(date +%H:%M:%S) DONE ===" >> "$LOG"
echo "Results:  $REPORT"
echo "Log:      $LOG"
echo "Outputs:  $OUTPUT_DIR"
