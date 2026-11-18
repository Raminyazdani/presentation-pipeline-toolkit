"""Reorder a slide deck without rebuilding its slide and media parts."""

import argparse
from copy import copy
import os
from pathlib import Path
import shutil
import sys
import tempfile
from zipfile import BadZipFile, ZipFile

from lxml import etree


P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
MANIFEST = "ppt/presentation.xml"
PPTX_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
)


def read_xml(archive, name):
    try:
        contents = archive.read(name)
    except KeyError:
        raise ValueError(f"missing PPTX part: {name}") from None
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    try:
        root = etree.fromstring(contents, parser=parser)
    except etree.XMLSyntaxError as exc:
        raise ValueError(f"invalid XML in {name}: {exc.msg}") from None
    if root.getroottree().docinfo.doctype:
        raise ValueError(f"unexpected document type declaration in {name}")
    return root


def ordered_manifest(archive, order):
    names = archive.namelist()
    if len(names) != len(set(names)):
        raise ValueError("ambiguous PPTX package with duplicate part names")
    types = read_xml(archive, "[Content_Types].xml")
    if any("digital-signature" in t.get("ContentType", "") for t in types):
        raise ValueError("digitally signed presentations cannot be reordered")
    if not any(t.get("PartName") == f"/{MANIFEST}" and
               t.get("ContentType") == PPTX_CONTENT_TYPE for t in types):
        raise ValueError("expected a standard .pptx presentation")
    root = read_xml(archive, MANIFEST)
    if root.xpath("boolean(.//*[local-name()='sectionLst'])"):
        raise ValueError("sectioned decks are unsupported; remove sections in a copy first")
    slide_list = root.find(P + "sldIdLst")
    if root.tag != P + "presentation" or slide_list is None or len(slide_list) == 0:
        raise ValueError("expected a Transitional .pptx with at least one slide")
    slides = list(slide_list)
    if any(slide.tag != P + "sldId" for slide in slides):
        raise ValueError("unexpected element in the slide list")
    for key in ("id", R + "id"):
        ids = [slide.get(key) for slide in slides]
        if None in ids or len(set(ids)) != len(ids):
            raise ValueError("ambiguous or missing slide identifiers")
    if sorted(order) != list(range(1, len(slides) + 1)):
        raise ValueError(f"order must include every position from 1 to {len(slides)} exactly once")
    for position in order:
        slide_list.append(slides[position - 1])
    return etree.tostring(root.getroottree(), encoding="UTF-8", xml_declaration=True)


def reorder(source, output, order):
    """Publish a new file only after successfully copying the entire package."""
    if source.suffix.lower() != ".pptx" or output.suffix.lower() != ".pptx":
        raise ValueError("use .pptx file names for the input and output")
    if os.path.lexists(output):
        raise FileExistsError(f"output already exists: {output}")
    with ZipFile(source) as original:
        manifest = ordered_manifest(original, order)
        with tempfile.TemporaryDirectory(dir=output.parent, prefix=".pptx-order-") as work:
            staged = Path(work) / "staged.pptx"
            with ZipFile(staged, "w") as result:
                result.comment = original.comment
                for info in original.infolist():
                    if info.filename == MANIFEST:
                        result.writestr(copy(info), manifest)
                    else:
                        with original.open(info) as reader, result.open(copy(info), "w") as writer:
                            shutil.copyfileobj(reader, writer)
            # Link creation is exclusive even if another process just wrote the target.
            try:
                os.link(staged, output)
            except FileExistsError:
                raise FileExistsError(f"output already exists: {output}") from None


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="existing .pptx file")
    parser.add_argument("--order", nargs="+", type=int, required=True,
                        help="all slide positions, in the new order (1-based)")
    parser.add_argument("-o", "--output", type=Path, required=True, help="new .pptx file")
    args = parser.parse_args(argv)
    try:
        reorder(args.input, args.output, args.order)
    except BadZipFile:
        parser.exit(2, "pptx-order: input is damaged or is not a readable PPTX package\n")
    except (OSError, ValueError, RuntimeError) as exc:
        parser.exit(2, f"pptx-order: {exc}\n")
    print(f"Wrote {args.output} ({len(args.order)} slides in order {' '.join(map(str, args.order))})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
