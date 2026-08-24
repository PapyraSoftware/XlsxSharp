#!/usr/bin/env bash
# Regenerates the CodeGen-produced parser files (*.g.cs) from the OOXML XSD in Schemas/.
#
# Usage: XlsxSharp.IO.CodeGen/regenerate.sh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
SCHEMA="$SCRIPT_DIR/Schemas/sml.xsd"

STYLES_TARGET="$REPO_ROOT/XlsxSharp/Excel/IO/StylesReader.g.cs"
CACHE_RECORDS_TARGET="$REPO_ROOT/XlsxSharp/Excel/IO/PivotCacheRecordsReader.g.cs"

echo "Generating $STYLES_TARGET"
dotnet run --project "$SCRIPT_DIR" -c Release -- styles "$SCHEMA" "$STYLES_TARGET" >/dev/null

echo "Generating $CACHE_RECORDS_TARGET"
dotnet run --project "$SCRIPT_DIR" -c Release -- cache-records "$SCHEMA" "$CACHE_RECORDS_TARGET" >/dev/null

echo
echo "Done. Review 'git diff' before committing -- schema changes can add or"
echo "remove elements/attributes, which changes the signature of the hand-coded"
echo "hooks in XlsxSharp/Excel/IO/StylesReader.cs and PivotCacheRecordsReader.cs."
