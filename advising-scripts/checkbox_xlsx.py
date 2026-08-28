# checkbox_xlsx.py — turn boolean cells into native Excel checkboxes.
# Usage: imported by tracker_writer.py; not run directly.
#
# Excel 365 renders a TRUE/FALSE cell as a tickable checkbox when its CELL
# FORMAT carries a "FeaturePropertyBag" extension. openpyxl cannot write that,
# so this post-processes the saved .xlsx: it reopens the archive, injects four
# pieces, and writes it back. Only stdlib zipfile is needed.
#
# ---------------------------------------------------------------------------
# WHY POST-PROCESSING, AND WHY NOT A TEMPLATE
# ---------------------------------------------------------------------------
#
# openpyxl 3.1.5 cannot be coaxed into this, and loading a hand-built template
# does not work either — it fails three separate ways:
#
#   * save() rebuilds the archive from the object model, so an unknown part
#     like xl/featurePropertyBag/ is simply dropped;
#   * keep_vba=True does not rescue it. Its ARC_VBA regex matches only
#     xl/vba, vmlDrawing, ctrlProps, customUI, activeX and media/*.emf, and
#     packaging/manifest.py further restricts carried-over content types to
#     (ACTIVEX, CTRL, VBA);
#   * styles/cell_style.py declares CellStyle.extLst and accepts it in
#     __init__ — then never assigns it, and leaves it out of __elements__ and
#     __attrs__, so it is never serialised. descriptors/excel.py's Extension
#     holds only a `uri` anyway, so the child element cannot be represented.
#
# XlsxWriter implements this feature properly, but it is write-only: adopting
# it would mean rewriting the whole generator — conditional formatting, data
# validation, comments, tables, the hidden sheet — for a cosmetic change to
# two columns.
#
# ---------------------------------------------------------------------------
# THE FOUR PIECES
# ---------------------------------------------------------------------------
#
#   1. a new part, xl/featurePropertyBag/featurePropertyBag.xml, declaring a
#      Checkbox bag and the chain of bags that maps a cell format to it;
#   2. a content-type override for that part;
#   3. a workbook relationship pointing at it;
#   4. an <extLst> inside the <xf> cell formats that should render as
#      checkboxes, referencing bag 0 of the mapped list.
#
# The CELL is untouched: it stays an ordinary boolean. That is the whole
# reason this degrades gracefully — Office 2019 and LibreOffice ignore the
# extension and show TRUE/FALSE, which is still correct and still readable by
# tracker_reader.
#
# All four were VERIFIED against a workbook Excel itself wrote — a scratch
# file with one checkbox in it — not merely built from the documentation. The
# bag XML, the content type and the relationship came back byte-identical.
# Excel's own file had the extension on the default cell format, so it never
# showed whether <extLst> may follow <alignment>; ours does, and Excel accepts
# it. If this part ever needs changing, make another such file and diff again
# rather than reasoning about the spec.
#
# ---------------------------------------------------------------------------
# THE TRAP: CLONE THE <xf>, NEVER MUTATE IT
# ---------------------------------------------------------------------------
#
# openpyxl DEDUPES cell formats. In the tracker the gpa1/gpa2 cells share one
# <xf> with Term, Course, CH, Gr., Pass and Ovr. — every cell that got
# BODY_FONT plus centred alignment. Adding the extension to that <xf> in place
# would turn all six of those columns into checkboxes, and the generated XML
# would look perfectly reasonable while doing it.
#
# So each distinct format in use by a target cell is CLONED to the end of
# <cellXfs> with the extension attached, and only the target cells are
# repointed at the clone. Nothing else in the workbook changes appearance.
#
# Element order matters to Excel: <extLst> must be the LAST child of <xf>,
# after <alignment> and <protection>. Get it wrong, or collide an rId, or
# declare a content type without its part, and Excel offers to "repair" the
# file — which silently discards things.

import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree

FEATURE_BAG_PART = "xl/featurePropertyBag/featurePropertyBag.xml"
FEATURE_BAG_TYPE = "application/vnd.ms-excel.featurepropertybag+xml"
FEATURE_BAG_REL = ("http://schemas.microsoft.com/office/2022/11/"
                   "relationships/FeaturePropertyBag")
FEATURE_BAG_NS = ("http://schemas.microsoft.com/office/spreadsheetml/2022/"
                  "featurepropertybag")

# The uri is a fixed GUID Excel looks for; it is not ours to choose.
XF_EXT_URI = "{C7286773-470A-42A8-94C5-96B5CB345126}"

# Bags are referenced by their position in this document, so the order is
# load-bearing: Checkbox is 0, XFControls 1, XFComplement 2. The final
# XFComplements bag lists which complements are mapped, and the i="0" in the
# <xf> extension indexes into THAT list.
FEATURE_BAG_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<FeaturePropertyBags xmlns="' + FEATURE_BAG_NS + '">'
    '<bag type="Checkbox"/>'
    '<bag type="XFControls"><bagId k="CellControl">0</bagId></bag>'
    '<bag type="XFComplement"><bagId k="XFControls">1</bagId></bag>'
    '<bag type="XFComplements" extRef="XFComplementsMapperExtRef">'
    '<a k="MappedFeaturePropertyBags"><bagId>2</bagId></a>'
    '</bag>'
    '</FeaturePropertyBags>'
)

XF_EXTENSION = (
    '<extLst><ext uri="' + XF_EXT_URI + '" xmlns:xfpb="' + FEATURE_BAG_NS +
    '"><xfpb:xfComplement i="0"/></ext></extLst>'
)

# One XML attribute, value quoted. Matching attributes explicitly rather than
# with [^>]* matters: [^>]* happily eats the "/" of a self-closing tag, and the
# alternation then falls through to the <xf>...</xf> branch and swallows the
# NEXT element whole. That silently merged two cell formats into one here, and
# the resulting file looked entirely reasonable. Quoted values may legitimately
# contain "/" and ">" (a numFmtId's format code, for one).
ATTRS = r'(?:\s+[\w:.-]+="[^"]*")*\s*'

# One <xf ... />, or one <xf ...> ... </xf>. Cell formats do not nest, so the
# non-greedy close is unambiguous.
XF_RE = re.compile(r"<xf" + ATTRS + r"(?:/>|>.*?</xf>)", re.S)
CELL_XFS_RE = re.compile(r"(<cellXfs\b[^>]*>)(.*?)(</cellXfs>)", re.S)


def _cell_re(coordinate):
    """Match the opening tag of one cell, self-closing or not."""
    return re.compile(
        r'<c(?:\s+[\w:.-]+="[^"]*")*?\s+r="' + re.escape(coordinate) +
        r'"' + ATTRS + r"(?:/>|>)")


class CheckboxError(Exception):
    """Raised when the workbook is not shaped the way this expects."""


SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _sheet_part(archive, sheet_name):
    """Which xl/worksheets/sheetN.xml holds `sheet_name`?

    Resolved through workbook.xml and its rels rather than assuming the
    tracker is sheet1. The hidden _TrackData sheet means there is more than
    one, and quietly editing the wrong one would be hard to spot.

    Parsed rather than pattern-matched: openpyxl writes these attributes in
    the order Type, Target, Id, and nothing guarantees that stays true. A
    regex expecting one order fails loudly here, but a regex expecting one
    order somewhere less careful would not.
    """
    workbook = ElementTree.fromstring(archive["xl/workbook.xml"])
    rel_id = None
    for sheet in workbook.iter("{" + SPREADSHEET_NS + "}sheet"):
        if sheet.get("name") == sheet_name:
            rel_id = sheet.get("{" + RELS_NS + "}id")
            break
    if rel_id is None:
        raise CheckboxError("no sheet named " + repr(sheet_name) +
                            " in xl/workbook.xml")

    rels = ElementTree.fromstring(archive["xl/_rels/workbook.xml.rels"])
    for relationship in rels.iter("{" + PACKAGE_RELS_NS + "}Relationship"):
        if relationship.get("Id") == rel_id:
            path = relationship.get("Target", "")
            if path.startswith("/"):
                return path.lstrip("/")
            return path if path.startswith("xl/") else "xl/" + path
    raise CheckboxError("no relationship " + rel_id +
                        " in xl/_rels/workbook.xml.rels")


def _split_cell_xfs(styles):
    """Return (prefix, [xf strings], suffix) around the <cellXfs> children."""
    match = CELL_XFS_RE.search(styles)
    if match is None:
        raise CheckboxError("xl/styles.xml has no <cellXfs> block")
    formats = XF_RE.findall(match.group(2))
    if not formats:
        raise CheckboxError("<cellXfs> contains no <xf> elements")

    # The split must account for every <xf> and lose nothing between them.
    # Without this guard a pattern that merges two elements produces a shorter
    # list, a wrong count attribute, and a clone containing two formats — none
    # of which is visible in the output. That is not hypothetical; it is what
    # the first version of XF_RE did.
    if len(formats) != match.group(2).count("<xf"):
        raise CheckboxError(
            "parsed " + str(len(formats)) + " <xf> elements but the block "
            "contains " + str(match.group(2).count("<xf")) +
            " — the pattern is not splitting cell formats correctly")
    if "".join(formats) != match.group(2).strip():
        raise CheckboxError(
            "<cellXfs> holds content between its <xf> elements that this "
            "would discard")
    return match, formats


def _with_extension(xf):
    """Add the checkbox extension as the LAST child of one <xf>."""
    if XF_EXT_URI in xf:
        return xf
    if xf.endswith("/>"):
        # A childless <xf .../> has to grow a body to hold the extension.
        return xf[:-2] + ">" + XF_EXTENSION + "</xf>"
    return xf[:-len("</xf>")] + XF_EXTENSION + "</xf>"


def _cell_style_indexes(sheet_xml, coordinates):
    """Map each target coordinate to the style index its cell uses.

    A cell with no s= attribute uses format 0, which is why the default is
    not simply "skip it".
    """
    found = {}
    for coordinate in coordinates:
        match = _cell_re(coordinate).search(sheet_xml)
        if match is None:
            continue  # cell never written; nothing to convert
        style = re.search(r'\bs="(\d+)"', match.group(0))
        found[coordinate] = int(style.group(1)) if style else 0
    if not found:
        raise CheckboxError("none of the requested cells exist in the sheet")
    return found


def _repoint(sheet_xml, coordinates, mapping, cell_styles):
    """Rewrite the s= attribute of each target cell to its cloned format."""
    for coordinate in coordinates:
        if coordinate not in cell_styles:
            continue
        new_index = str(mapping[cell_styles[coordinate]])

        def replace(match, new_index=new_index):
            tag = match.group(0)
            if re.search(r'\bs="\d+"', tag):
                return re.sub(r'\bs="\d+"', 's="' + new_index + '"', tag)
            # No style yet: insert one straight after the reference.
            return re.sub(r'(\br="[^"]+")', r'\1 s="' + new_index + '"',
                          tag, count=1)

        sheet_xml = _cell_re(coordinate).sub(replace, sheet_xml, count=1)
    return sheet_xml


def _next_rel_id(rels):
    """An rId that is not already taken. A collision means a repair prompt."""
    used = {int(n) for n in re.findall(r'\bId="rId(\d+)"', rels)}
    candidate = 1
    while candidate in used:
        candidate += 1
    return "rId" + str(candidate)


def add_checkboxes(path, sheet_name, coordinates):
    """Render the given cells as native Excel checkboxes.

    path        the .xlsx openpyxl has already saved
    sheet_name  the worksheet holding them, e.g. "Tracker"
    coordinates iterable of cell references, e.g. ["J6", "K6", ...]

    The cells must already hold booleans. Returns the number of distinct cell
    formats that were cloned.
    """
    path = Path(path)
    coordinates = list(coordinates)
    if not coordinates:
        return 0

    with zipfile.ZipFile(path) as source:
        order = source.namelist()
        archive = {name: source.read(name) for name in order}

    sheet_part = _sheet_part(archive, sheet_name)
    if sheet_part not in archive:
        raise CheckboxError("missing sheet part " + sheet_part)

    sheet_xml = archive[sheet_part].decode("utf-8")
    styles = archive["xl/styles.xml"].decode("utf-8")

    cell_styles = _cell_style_indexes(sheet_xml, coordinates)
    match, formats = _split_cell_xfs(styles)

    # Clone, never mutate — see the header. One clone per distinct format in
    # use, appended in a stable order so the output is reproducible.
    mapping = {}
    for index in sorted(set(cell_styles.values())):
        if index >= len(formats):
            raise CheckboxError(
                "cell format " + str(index) + " is out of range; <cellXfs> "
                "holds " + str(len(formats)))
        mapping[index] = len(formats)
        formats.append(_with_extension(formats[index]))

    rebuilt = (match.group(1) + "".join(formats) + match.group(3))
    rebuilt = re.sub(r'\bcount="\d+"', 'count="' + str(len(formats)) + '"',
                     rebuilt, count=1)
    styles = styles[:match.start()] + rebuilt + styles[match.end():]

    archive["xl/styles.xml"] = styles.encode("utf-8")
    archive[sheet_part] = _repoint(
        sheet_xml, coordinates, mapping, cell_styles).encode("utf-8")

    # The part itself.
    archive[FEATURE_BAG_PART] = FEATURE_BAG_XML.encode("utf-8")
    if FEATURE_BAG_PART not in order:
        order.append(FEATURE_BAG_PART)

    # Its content type. A part without one, or one without its part, is a
    # repair prompt either way.
    content_types = archive["[Content_Types].xml"].decode("utf-8")
    if FEATURE_BAG_TYPE not in content_types:
        content_types = content_types.replace(
            "</Types>",
            '<Override PartName="/' + FEATURE_BAG_PART + '" ContentType="' +
            FEATURE_BAG_TYPE + '"/></Types>')
        archive["[Content_Types].xml"] = content_types.encode("utf-8")

    # And the relationship that ties it to the workbook.
    rels = archive["xl/_rels/workbook.xml.rels"].decode("utf-8")
    if FEATURE_BAG_REL not in rels:
        rels = rels.replace(
            "</Relationships>",
            '<Relationship Id="' + _next_rel_id(rels) + '" Type="' +
            FEATURE_BAG_REL + '" Target="featurePropertyBag/'
            'featurePropertyBag.xml"/></Relationships>')
        archive["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    _assert_well_formed(archive)

    # Write beside the original and move into place, so a failure part way
    # through cannot leave a half-written workbook where the good one was.
    temporary = path.with_suffix(path.suffix + ".checkbox")
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as out:
        for name in order:
            out.writestr(name, archive[name])
    shutil.move(str(temporary), str(path))
    return len(mapping)


def _assert_well_formed(archive):
    """Parse the parts we edited, so a broken edit fails here and not in Excel.

    Parse only — reserialising through ElementTree would rewrite namespace
    prefixes across the whole file and churn parts we never touched.
    """
    for name, data in archive.items():
        if not name.endswith(".xml") and not name.endswith(".rels"):
            continue
        try:
            ElementTree.fromstring(data)
        except ElementTree.ParseError as error:
            raise CheckboxError(name + " is not well-formed after the edit: "
                                + str(error))
