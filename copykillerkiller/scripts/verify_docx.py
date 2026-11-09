#!/usr/bin/env python3
"""Compare two DOCX packages without modifying them; Python 3.9+ stdlib only."""

import argparse
import hashlib
import json
import os
from pathlib import Path
import posixpath
import sys
import xml.etree.ElementTree as ET
from zipfile import BadZipFile, ZipFile


WORD_NS = (
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/wordprocessingml/main",
)
MATH_NS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "http://purl.oclc.org/ooxml/officeDocument/math",
)
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOCUMENT_PART = "word/document.xml"
REQUIRED_PARTS = ("[Content_Types].xml", "_rels/.rels", DOCUMENT_PART)
MAIN_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)
PROTECTED = {
    "tables": (WORD_NS, "tbl"),
    "math_objects": (MATH_NS, "oMath"),
    "math_display": (MATH_NS, "oMathPara"),
    "section_properties": (WORD_NS, "sectPr"),
    "hyperlinks": (WORD_NS, "hyperlink"),
    "footnote_references": (WORD_NS, "footnoteReference"),
    "endnote_references": (WORD_NS, "endnoteReference"),
    "simple_fields": (WORD_NS, "fldSimple"),
    "complex_field_markers": (WORD_NS, "fldChar"),
    "field_instructions": (WORD_NS, "instrText"),
}
TEXT_TAGS = {
    "{%s}%s" % (namespace, local)
    for namespaces, names in ((WORD_NS, ("t", "instrText", "delText")), (MATH_NS, ("t",)))
    for namespace in namespaces
    for local in names
}


def sha256_stream(stream):
    digest = hashlib.sha256()
    for block in iter(lambda: stream.read(1024 * 1024), b""):
        digest.update(block)
    return digest.hexdigest()


def parse_xml(data, label):
    # DOCX XML does not require DTDs. Reject them rather than expand entities.
    if b"<!DOCTYPE" in data.replace(b"\x00", b"").upper():
        raise ValueError("%s: DTD declarations are not supported" % label)
    try:
        return ET.fromstring(data)
    except ET.ParseError as error:
        raise ValueError("%s: invalid XML (%s)" % (label, error)) from error


def read_docx(path):
    with path.open("rb") as source:
        file_hash = sha256_stream(source)
        source.seek(0)
        with ZipFile(source) as package:
            members = [item for item in package.infolist() if not item.is_dir()]
            names = [item.filename for item in members]
            if len(names) != len(set(names)):
                raise ValueError("%s: duplicate ZIP part names" % path)
            missing = sorted(set(REQUIRED_PARTS) - set(names))
            if missing:
                raise ValueError("%s: missing DOCX parts: %s" % (path, ", ".join(missing)))
            hashes = {}
            for item in members:
                with package.open(item) as part:
                    hashes[item.filename] = sha256_stream(part)
            roots = {
                name: parse_xml(package.read(name), "%s:%s" % (path, name))
                for name in REQUIRED_PARTS
            }
    content_types = roots["[Content_Types].xml"]
    if content_types.tag != "{%s}Types" % CONTENT_TYPES_NS:
        raise ValueError("%s: invalid content-types root" % path)
    if not any(
        item.get("PartName") == "/" + DOCUMENT_PART
        and item.get("ContentType") == MAIN_CONTENT_TYPE
        for item in content_types.findall("{%s}Override" % CONTENT_TYPES_NS)
    ):
        raise ValueError("%s: missing DOCX main-document content type" % path)
    if roots["_rels/.rels"].tag != "{%s}Relationships" % RELATIONSHIPS_NS:
        raise ValueError("%s: invalid package-relationships root" % path)
    office_document_types = {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
    }
    if not any(
        item.get("Type") in office_document_types
        and item.get("TargetMode", "Internal") == "Internal"
        and posixpath.normpath(item.get("Target", "").lstrip("/")) == DOCUMENT_PART
        for item in roots["_rels/.rels"].findall("{%s}Relationship" % RELATIONSHIPS_NS)
    ):
        raise ValueError("%s: missing main-document package relationship" % path)
    document = roots[DOCUMENT_PART]
    if document.tag not in {"{%s}document" % namespace for namespace in WORD_NS}:
        raise ValueError("%s: invalid Word document root" % path)
    body_tag = document.tag.rsplit("}", 1)[0] + "}body"
    if document.find(body_tag) is None:
        raise ValueError("%s: missing Word document body" % path)
    return {"sha256": file_hash, "zip_part_count": len(hashes)}, hashes, document


def meaningful_text(text, preserve=False):
    return text if text and (preserve or text.strip()) else ""


def xml_structure(element):
    """Ignore prefix/attribute order and indentation, but preserve Word text spaces."""
    return (
        element.tag,
        sorted(element.attrib.items()),
        meaningful_text(element.text, element.tag in TEXT_TAGS),
        [(xml_structure(child), meaningful_text(child.tail)) for child in element],
    )


def protected_fingerprints(document, namespaces, local):
    tags = {"{%s}%s" % (namespace, local) for namespace in namespaces}
    return [
        hashlib.sha256(
            json.dumps(xml_structure(element), ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        ).hexdigest()
        for element in document.iter()
        if element.tag in tags
    ]


def compare(source_path, revision_path):
    if os.path.samefile(source_path, revision_path):
        raise ValueError("source and revision refer to the same file")
    source, before, source_xml = read_docx(source_path)
    revision, after, revision_xml = read_docx(revision_path)
    added = sorted(after.keys() - before.keys())
    removed = sorted(before.keys() - after.keys())
    common = before.keys() & after.keys()
    changed = sorted(name for name in common if before[name] != after[name])
    elements = {}
    for label, (namespaces, local) in PROTECTED.items():
        old = protected_fingerprints(source_xml, namespaces, local)
        new = protected_fingerprints(revision_xml, namespaces, local)
        elements[label] = {"source_count": len(old), "revision_count": len(new), "changed": old != new}
    return {
        "schema_version": 1,
        "source": source,
        "revision": revision,
        "zip_parts": {
            "added": added,
            "removed": removed,
            "changed": changed,
            "counts": {
                "added": len(added), "removed": len(removed), "changed": len(changed),
                "unchanged": len(common) - len(changed),
            },
        },
        "protected_document_xml": {
            "part": DOCUMENT_PART,
            "any_changed": any(item["changed"] for item in elements.values()),
            "elements": elements,
            "comparison": "Ordered element subtrees with expanded XML names; ignores indentation "
                          "and attribute order. ZIP part comparison ignores metadata and compression.",
        },
        "limitations": [
            "This is a structural preflight, not a pass/fail assessment or full OOXML validation.",
            "File SHA-256 includes ZIP metadata; ZIP part changes compare uncompressed bytes only.",
            "Protected checks cover selected elements in word/document.xml, not every document feature. "
            "Element positions relative to unprotected prose and relationship targets are not compared.",
            "Complex fields are checked through fldChar markers and instrText instructions, not entire "
            "field spans, cached results, field pairing, or evaluation correctness.",
            "Render both documents and inspect every page for layout, fonts, clipping, and glyph errors.",
            "Manually verify prose, meaning, numbers, citations, references, and requested changes.",
        ],
    }


def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", metavar="SOURCE.docx", type=Path)
    parser.add_argument("revision", metavar="REVISION.docx", type=Path)
    args = parser.parse_args()
    try:
        result = compare(args.source.resolve(strict=True), args.revision.resolve(strict=True))
    except (OSError, ValueError, BadZipFile, RuntimeError, NotImplementedError) as error:
        print("verify_docx: %s" % error, file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
