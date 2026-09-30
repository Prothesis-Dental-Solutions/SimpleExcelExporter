#!/usr/bin/env bash
# Enforce the .editorconfig rules for line endings on every tracked text file:
#   - LF only: no CR character (CRLF or lone CR)
#   - a final newline (except NuGet-generated packages.lock.json, see .editorconfig)
# Checks the Git index, so it sees exactly what is committed.
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"

status=0

if cr_files=$(git grep --cached -I -l $'\r'); then
  echo "::error::CR characters found. The repository uses LF line endings (see .editorconfig and .gitattributes):"
  echo "$cr_files"
  status=1
fi

missing_newline=""
while IFS= read -r -d '' file; do
  blob=$(git show ":$file")
  # Skip empty and binary files.
  [ -n "$blob" ] || continue
  git show ":$file" | grep -Iq . || continue
  if [ -n "$(git show ":$file" | tail -c1)" ]; then
    missing_newline+="$file"$'\n'
  fi
done < <(git ls-files -z -- ':!:**/packages.lock.json')

if [ -n "$missing_newline" ]; then
  echo "::error::Files without a final newline (insert_final_newline = true in .editorconfig):"
  printf '%s' "$missing_newline"
  status=1
fi

if [ "$status" -eq 0 ]; then
  echo "Line endings OK: LF only, final newline present."
fi
exit "$status"
