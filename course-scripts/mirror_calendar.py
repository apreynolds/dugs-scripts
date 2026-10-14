#!/usr/bin/env -S uv run --script
# /// script
# dependencies = []
# ///
# mirror_calendar.py — save a copy of UNB's online undergraduate calendar (Fredericton).
# Usage:
#   ./mirror_calendar.py OUTDIR
#       e.g. ./mirror_calendar.py ~/Documents/3work/WIP-DUGS-related/curriculum/academic-calendar/2027-28
#   ./mirror_calendar.py OUTDIR --base URL  mirror an archived year instead of "current"
#   ./mirror_calendar.py -h | --help
#
#   OUTDIR        folder to write into. Two subfolders are made:
#                   html/   each page exactly as the server sent it
#                   text/   the same page's main content as plain text
#                 plus MIRROR-LOG.txt, listing every page and the date.
#       --base    the calendar's root URL (default: the "current" calendar)
#       --delay   seconds to wait between pages (default 0.2)
#       --retext  download nothing; rebuild text/ from the pages already in
#                 html/ (for when the text conversion itself is changed)
#
# WHAT THIS IS FOR
# From 2026-27 the calendar is published only as a website, and "current" is
# replaced every year. This keeps a dated copy that can be searched like the
# PDF-derived text files of earlier years.
#
# WHAT IS MIRRORED
#   frederictonprograms/   every page: the program regulations
#   frederictoncourses/    one page per department, which holds every course
#                          entry in full. The separate page each course also
#                          has (math-1003.html) repeats the same text, so
#                          those thousands of pages are skipped.
# Saint John's sections are left out on purpose.

import argparse
import datetime
import re
import subprocess
import sys
import time
import urllib.parse
from html.parser import HTMLParser
from pathlib import Path

DEFAULT_BASE = "https://www.unb.ca/academics/calendar/undergraduate/current/"
SECTIONS = ["frederictonprograms/", "frederictoncourses/"]

# A single course's own page: 'math-1003.html', 'cs-1073.html', 'me-3522.html'.
COURSE_PAGE = re.compile(r"/[a-z]{2,5}-\d{3,4}[a-z]?\.html$")

# Tags that end a line of text when they open or close.
BLOCK_TAGS = {"p", "div", "li", "ul", "ol", "tr", "table", "h1", "h2", "h3", "h4",
              "h5", "h6", "br", "section", "dt", "dd", "blockquote"}


class LinkFinder(HTMLParser):
    """Collects the href of every <a> on a page."""

    def __init__(self):
        super().__init__()
        self.links = []

    def handle_starttag(self, tag, attrs):
        if tag == "a":
            href = dict(attrs).get("href")
            if href:
                self.links.append(href)


class TextMaker(HTMLParser):
    """Turns a piece of HTML into plain text, one block per line."""

    def __init__(self):
        super().__init__()
        self.parts = []
        self.skipping = 0  # inside <script> or <style>

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self.skipping += 1
        elif tag in BLOCK_TAGS:
            self.parts.append("\n")
            # Mark headings the markdown way ('## Honours Program'), so a
            # search can tell which part of a long page a line belongs to.
            if re.fullmatch(r"h[1-6]", tag):
                self.parts.append("#" * int(tag[1]) + " ")
        elif tag in ("td", "th"):
            self.parts.append(" | ")  # keeps the cells of a table row apart

    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self.skipping = max(0, self.skipping - 1)
        elif tag in BLOCK_TAGS:
            self.parts.append("\n")

    def handle_data(self, data):
        if not self.skipping:
            self.parts.append(data)

    def text(self):
        joined = "".join(self.parts).replace("\xa0", " ")
        lines = [re.sub(r"\s+", " ", line).strip() for line in joined.split("\n")]
        # Drop blank lines, and the empty table cells left at the start of a row.
        lines = [re.sub(r"^(\| )+", "", line) for line in lines]
        return "\n".join(line for line in lines if line and line != "|") + "\n"


def main_content(html):
    """The part of a page between the main-content marker and the footer."""
    start = html.find('id="mcontent"')
    if start == -1:
        return html
    end = html.find("<footer", start)
    return html[start - 200 if start > 200 else 0: end if end != -1 else len(html)]


def fetch(url):
    """Return (final URL, page text), or (url, None) if the page cannot be had."""
    # curl does the downloading. Python's own urllib was tried first and got
    # every page cut off at about 24 kB when run behind a filtering proxy
    # (IncompleteRead), where curl reads the same pages whole.
    #   -s silent, -L follow redirects, -f fail on 404 instead of saving the
    #   error page, -w prints the final URL and content type after the body.
    # A few of the site's own links have a space in them ('registrars warrants').
    url = urllib.parse.quote(url, safe=":/%")
    marker = "\n@@MIRROR@@"
    command = ["curl", "-sLf", "--max-time", "60", "--retry", "3",
               "-A", "mirror_calendar.py (UNB Math & Stats)",
               "-w", marker + "%{url_effective}\t%{content_type}", url]
    result = subprocess.run(command, capture_output=True)
    if result.returncode != 0:
        print(f"  FAILED {url}: curl exit {result.returncode}", file=sys.stderr)
        return url, None
    body, _, tail = result.stdout.decode("utf-8", errors="replace").rpartition(marker)
    final, _, kind = tail.partition("\t")
    if "html" not in kind:
        return final, None
    return final, body


def wanted(url, base):
    """Is this a page the mirror should hold?"""
    if not url.startswith(base):
        return False
    rest = url[len(base):]
    if not any(rest.startswith(section) for section in SECTIONS):
        return False
    if not (rest.endswith(".html") or rest.endswith("/")):
        return False  # a PDF, an image, a search form
    if rest.startswith("frederictoncourses/") and COURSE_PAGE.search(url):
        return False  # one course's own page; its department page has the text
    return True


def main():
    parser = argparse.ArgumentParser(description="Mirror UNB's online undergraduate calendar (Fredericton).")
    parser.add_argument("outdir", help="folder to write into, e.g. 2026-27")
    parser.add_argument("--base", default=DEFAULT_BASE, help="the calendar's root URL")
    parser.add_argument("--delay", type=float, default=0.2, help="seconds between pages (default 0.2)")
    parser.add_argument("--retext", action="store_true",
                        help="download nothing: rebuild text/ from the pages already in html/")
    args = parser.parse_args()

    base = args.base if args.base.endswith("/") else args.base + "/"
    out = Path(args.outdir).expanduser()
    queue = [base + section + "index.html" for section in SECTIONS]
    seen = set(queue)
    log = []

    while queue:
        url = queue.pop(0)
        relative = url[len(base):]
        if relative.endswith("/"):
            relative += "index.html"
        page = out / "html" / relative
        if args.retext:
            # Rebuild text/ from the pages already saved; nothing is downloaded.
            if not page.exists():
                continue
            final, html = url, page.read_text(encoding="utf-8")
        else:
            final, html = fetch(url)
            if html is None:
                log.append(f"FAILED\t{url}")
                continue
            page.parent.mkdir(parents=True, exist_ok=True)
            page.write_text(html, encoding="utf-8")

        maker = TextMaker()
        maker.feed(main_content(html))
        text = out / "text" / (relative[:-5] + ".txt")
        text.parent.mkdir(parents=True, exist_ok=True)
        text.write_text(f"SOURCE: {url}\n\n" + maker.text(), encoding="utf-8")
        log.append(f"ok\t{url}")
        print(f"{len(log):4d}  {relative}")

        finder = LinkFinder()
        finder.feed(html)
        for href in finder.links:
            # Resolve '../physics/index.html' against this page, drop any '#part'.
            target = urllib.parse.urldefrag(urllib.parse.urljoin(final, href.strip()))[0]
            if target.endswith("/"):
                target += "index.html"
            if target not in seen and wanted(target, base):
                seen.add(target)
                queue.append(target)
        if not args.retext:
            time.sleep(args.delay)

    if args.retext:
        print(f"\nRebuilt the text of {len(log)} pages in {out}")
        return
    stamp = datetime.date.today().isoformat()
    header = f"Mirrored {stamp} from {base}\n{len(log)} pages\n\n"
    (out / "MIRROR-LOG.txt").write_text(header + "\n".join(log) + "\n", encoding="utf-8")
    failed = sum(1 for line in log if line.startswith("FAILED"))
    print(f"\n{len(log)} pages into {out} ({failed} failed)")


if __name__ == "__main__":
    main()
