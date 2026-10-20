# paths.py — where the DUGS data lives. The one file to change if a folder moves.
# Usage: imported by the scripts; not run directly.
#
# Every script's default paths come from here, and every one can still be
# overridden on its command line (see each script's -h).
#
# The data is in the Teams folders on OneDrive, never in this repository.
# Each team's folder is synced to this machine under ONEDRIVE; on Windows
# the same "OneDrive - University of New Brunswick" folder sits in the home
# folder too. If yours has a different name, or the team folders are synced
# somewhere else, change the lines below.

from pathlib import Path

ONEDRIVE = Path.home() / "OneDrive - University of New Brunswick"

# One folder per team. Curriculum and Timetable keep everything under the
# channel's General folder; Advising and DUGS keep it at the top.
CURRICULUM = ONEDRIVE / "TEAM-CURRICULUM" / "General"
TIMETABLE = ONEDRIVE / "TEAM-TIMETABLE" / "General"
ADVISING = ONEDRIVE / "TEAM-ADVISING"
DUGS = ONEDRIVE / "DUGS"

# Curriculum
MASTER_LIST = CURRICULUM / "COURSES-math-stat.xlsx"
OTHER_UNITS = CURRICULUM / "COURSES-required-by-other-units.md"
COURSE_SUMMARIES = CURRICULUM / "informer5-course-summaries"
CALENDAR = CURRICULUM / "academic-calendar"  # one folder per calendar year

# Timetable
TIMETABLE_EXPORTS = TIMETABLE / "timetable-exports-by-year"

# Advising
REQUIREMENTS = ADVISING / "degree-requirements"
ADVISING_TIMETABLE = ADVISING / "resources" / "advising-timetable.xlsx"
