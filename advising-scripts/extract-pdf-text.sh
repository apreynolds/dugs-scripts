#!/usr/bin/env bash
# extract-pdf-text.sh — extract text from a text-based PDF using pdftotext
# Usage:
#   ./extract-pdf-text.sh input.pdf              # prints to stdout
#   ./extract-pdf-text.sh input.pdf output.txt   # writes to file

set -euo pipefail

if [[ $# -lt 1 ]]; then
    echo "Usage: $0 <input.pdf> [output.txt]" >&2
    exit 1
fi

input="$1"

if [[ ! -f "$input" ]]; then
    echo "Error: file not found: $input" >&2
    exit 1
fi

if [[ $# -ge 2 ]]; then
    pdftotext "$input" "$2"
    echo "Text written to $2"
else
    pdftotext "$input" -
fi
