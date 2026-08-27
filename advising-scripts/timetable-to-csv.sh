#!/usr/bin/env bash
# timetable-to-csv.sh — convert UNB course timetable text dump to wide CSV
# Usage:
#   ./timetable-to-csv.sh input.html              # prints to stdout
#   ./timetable-to-csv.sh input.html output.csv   # writes to file
#   ./timetable-to-csv.sh -h | --help             # show this help
#
# Reads the saved-as-text version of UNB's course timetables page (a
# tab/marker layout where each section header is a single tab-separated line
# and any schedule body is a stack of "*Crs day(s)*", "*Lab day(s)*", etc.
# markers, each followed by its value on the next line). Emits one CSV row
# per section with a fixed wide schema:
#
#   course,section,title,lec_days,lec_time,lab_days,lab_time,tut_days,tut_time
#
# Sections without any schedule (e.g. FRE1A online sections) get all six
# day/time columns left blank.

set -euo pipefail

# --- help -------------------------------------------------------------------

if [[ "${1:-}" == "-h" || "${1:-}" == "--help" ]]; then
    # Reprint the leading comment block as help text.
    awk 'NR==1 && /^#!/ { next }
         /^#/            { sub(/^# ?/, ""); print; next }
         { exit }' "$0"
    exit 0
fi

# --- arg handling -----------------------------------------------------------

if [[ $# -lt 1 ]]; then
    echo "Usage: $0 <input.html> [output.csv]" >&2
    exit 1
fi

input="$1"

if [[ ! -f "$input" ]]; then
    echo "Error: file not found: $input" >&2
    exit 1
fi

# --- parsing ----------------------------------------------------------------
#
# All real work happens in awk. Logic:
#
#   * A "header line" starts with a course code (e.g. MATH*1003, CS*1073) and
#     is tab-separated:    course \t section \t title \t [marker]
#     The 4th field, if present, is the first marker (always "*Crs day(s)*").
#
#   * Standalone marker lines look like *Crs day(s)*, *Lab day(s)*,
#     *Tut day(s)*, *Crs time*, *Lab time*, *Tut time*. The first time
#     marker often has a leading tab (it lived in column 5 of the original
#     table); the regex below is permissive about leading whitespace.
#
#   * The line immediately after a marker is the value for that marker.
#     Markers stack in order: all day(s) markers first, then all time markers.
#
# We carry a `pending` slot name across lines: when set, the next non-marker
# line goes into that slot and `pending` is cleared.

awk_script='
function trim(s) {
    sub(/^[[:space:]]+/, "", s)
    sub(/[[:space:]]+$/, "", s)
    return s
}

# Quote a CSV field if it contains a comma, double quote, or newline.
function csv(s) {
    if (s ~ /[",]/ || s ~ /\n/) {
        gsub(/"/, "\"\"", s)
        return "\"" s "\""
    }
    return s
}

# Map a marker string like "*Crs day(s)*" to its slot name.
function marker_slot(m,    s) {
    s = m
    sub(/^[[:space:]]*\*/, "", s)
    sub(/\*[[:space:]]*$/, "", s)
    if (s == "Crs day(s)") return "lec_days"
    if (s == "Crs time")   return "lec_time"
    if (s == "Lab day(s)") return "lab_days"
    if (s == "Lab time")   return "lab_time"
    if (s == "Tut day(s)") return "tut_days"
    if (s == "Tut time")   return "tut_time"
    return ""
}

# Emit the current record (if any) and reset state.
function flush() {
    if (course != "") {
        print csv(course) "," csv(section) "," csv(title) ","         \
              csv(slot["lec_days"]) "," csv(slot["lec_time"]) ","     \
              csv(slot["lab_days"]) "," csv(slot["lab_time"]) ","     \
              csv(slot["tut_days"]) "," csv(slot["tut_time"])
    }
    course = ""; section = ""; title = ""; pending = ""
    delete slot
}

BEGIN {
    FS = "\t"
    print "course,section,title,lec_days,lec_time,lab_days,lab_time,tut_days,tut_time"
}

# Header line: starts with DEPT*CODE (codes are followed by " \t" — a trailing
# space before the tab — so we allow optional whitespace before the tab).
/^[A-Za-z]+\*[A-Za-z0-9]+[[:space:]]*\t/ {
    flush()
    course  = trim($1)
    section = trim($2)
    title   = trim($3)
    if (NF >= 4 && $4 ~ /\*[A-Za-z]/) {
        pending = marker_slot($4)
    }
    next
}

# Standalone marker line (sometimes prefixed with a tab).
/^[[:space:]]*\*(Crs|Lab|Tut)[[:space:]](day\(s\)|time)\*[[:space:]]*$/ {
    pending = marker_slot($0)
    next
}

# Anything else: if a marker is pending, this line is its value.
{
    if (pending != "") {
        slot[pending] = trim($0)
        pending = ""
    }
}

END { flush() }
'

if [[ $# -ge 2 ]]; then
    awk "$awk_script" "$input" > "$2"
    echo "CSV written to $2" >&2
else
    awk "$awk_script" "$input"
fi
