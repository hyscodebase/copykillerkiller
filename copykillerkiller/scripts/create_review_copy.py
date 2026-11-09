#!/usr/bin/env python3
"""Create a separate DOCX with Word review comments; Python 3.9+ stdlib only."""

from collections import defaultdict
from copy import copy
import argparse
import hashlib
from io import BytesIO
import json
import os
from pathlib import Path
import posixpath
import re
import sys
from xml.dom import Node, minidom
from xml.parsers.expat import ExpatError
from zipfile import BadZipFile, ZipFile, ZIP_DEFLATED


WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
STRICT_WORD = "http://purl.oclc.org/ooxml/wordprocessingml/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
STRICT_REL = "http://purl.oclc.org/ooxml/officeDocument/relationships/"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
XMLNS = "http://www.w3.org/2000/xmlns/"
XML = "http://www.w3.org/XML/1998/namespace"
DOC = "word/document.xml"
DOC_RELS = "word/_rels/document.xml.rels"
COMMENTS = "word/comments.xml"
MAIN_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
COMMENTS_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
REL_TYPE = "application/vnd.openxmlformats-package.relationships+xml"


class ReviewError(Exception):
    pass


class SafeParser(argparse.ArgumentParser):
    def error(self, message):
        self.exit(2, "create_review_copy: invalid command line; use --help\n")


def parse_xml(data, label):
    if b"<!DOCTYPE" in data.replace(b"\x00", b"").upper():
        raise ReviewError(label + ": DTD declarations are not supported")
    try:
        return minidom.parseString(data)
    except ExpatError:
        raise ReviewError(label + ": invalid XML") from None


def children(element, namespace, local):
    return [node for node in element.childNodes if node.nodeType == Node.ELEMENT_NODE
            and node.namespaceURI == namespace and node.localName == local]


def new_element(document, namespace, prefix, local, attributes=None):
    name = (prefix + ":" if prefix else "") + local
    element = document.createElementNS(namespace, name)
    for key, value in (attributes or {}).items():
        element.setAttribute(key, value)
    return element


def stored_text(paragraph, word_namespace):
    # Deliberately no whitespace normalization, field instructions, or math m:t.
    return "".join(
        child.data
        for element in paragraph.getElementsByTagNameNS(word_namespace, "t")
        for child in element.childNodes
        if child.nodeType in (Node.TEXT_NODE, Node.CDATA_SECTION_NODE)
    )


def validate_notes(notes, paragraphs, word_namespace):
    if not isinstance(notes, list) or not notes:
        raise ReviewError("notes must be a nonempty JSON list")
    for number, note in enumerate(notes, 1):
        if not isinstance(note, dict) or set(note) != {"paragraph", "expected_text", "comment"}:
            raise ReviewError("note %d must contain paragraph, expected_text, and comment only" % number)
        index = note["paragraph"]
        if type(index) is not int or not 1 <= index <= len(paragraphs):
            raise ReviewError("note %d has an invalid direct body paragraph number" % number)
        for key in ("expected_text", "comment"):
            if not isinstance(note[key], str) or not note[key].strip():
                raise ReviewError("note %d has empty or invalid %s" % (number, key))
        if note["expected_text"] != stored_text(paragraphs[index - 1], word_namespace):
            raise ReviewError("note %d expected_text does not match paragraph %d" % (number, index))
        if any(not (c in "\t\n\r" or 0x20 <= ord(c) <= 0xD7FF
                    or 0xE000 <= ord(c) <= 0xFFFD or 0x10000 <= ord(c) <= 0x10FFFF)
               for c in note["comment"]):
            raise ReviewError("note %d contains characters not permitted in XML" % number)


def reject_existing_comments(parts, content_types, document):
    known_comment_part = re.compile(
        r"word/comments(?:extended|extensible|ids|userdata)?[0-9]*\.xml", re.IGNORECASE
    )
    if any(known_comment_part.fullmatch(name) for name in parts):
        raise ReviewError("existing comment-related parts are not supported")
    for node in content_types.documentElement.childNodes:
        if node.nodeType == Node.ELEMENT_NODE and "comment" in node.getAttribute("ContentType").lower():
            raise ReviewError("existing comment content types are not supported")
    for name, data in parts.items():
        if name.endswith(".rels"):
            rels = parse_xml(data, "package relationships")
            for item in rels.getElementsByTagNameNS(REL, "Relationship"):
                if "comment" in item.getAttribute("Type").lower():
                    raise ReviewError("existing comment relationships are not supported")
        elif name.startswith("word/") and name.endswith(".xml"):
            if b"comment" not in data.replace(b"\x00", b"").lower():
                continue
            root = document if name == DOC else parse_xml(data, "Word XML")
            for namespace in (WORD, STRICT_WORD):
                for local in ("comment", "commentRangeStart", "commentRangeEnd", "commentReference"):
                    if root.getElementsByTagNameNS(namespace, local):
                        raise ReviewError("existing Word comment anchors are not supported")


def build_review(source_bytes, notes):
    with ZipFile(BytesIO(source_bytes)) as archive:
        infos = archive.infolist()
        if len({item.filename for item in infos}) != len(infos):
            raise ReviewError("DOCX has duplicate ZIP part names")
        parts = {item.filename: archive.read(item) for item in infos}
        archive_comment = archive.comment
    if not {DOC, "[Content_Types].xml", "_rels/.rels"}.issubset(parts):
        raise ReviewError("DOCX is missing required package parts")
    document = parse_xml(parts[DOC], "document")
    root = document.documentElement
    word_namespace = root.namespaceURI
    if root.localName != "document" or word_namespace not in (WORD, STRICT_WORD):
        raise ReviewError("invalid Word document root")
    bodies = children(root, word_namespace, "body")
    if len(bodies) != 1:
        raise ReviewError("DOCX must have one direct document body")
    paragraphs = children(bodies[0], word_namespace, "p")
    validate_notes(notes, paragraphs, word_namespace)
    content_types = parse_xml(parts["[Content_Types].xml"], "content types")
    ct_root = content_types.documentElement
    if ct_root.namespaceURI != CT or ct_root.localName != "Types":
        raise ReviewError("invalid content-types root")
    if not any(item.getAttribute("PartName") == "/" + DOC
               and item.getAttribute("ContentType") == MAIN_TYPE
               for item in children(ct_root, CT, "Override")):
        raise ReviewError("missing DOCX main-document content type")
    package_rels = parse_xml(parts["_rels/.rels"], "package relationships").documentElement
    if package_rels.namespaceURI != REL or package_rels.localName != "Relationships":
        raise ReviewError("invalid package-relationships root")
    if not any(item.getAttribute("Type") in {OFFICE_REL + "officeDocument", STRICT_REL + "officeDocument"}
               and item.getAttribute("TargetMode") in ("", "Internal")
               and posixpath.normpath(item.getAttribute("Target").lstrip("/")) == DOC
               for item in children(package_rels, REL, "Relationship")):
        raise ReviewError("missing main-document package relationship")
    reject_existing_comments(parts, content_types, document)

    if DOC_RELS in parts:
        relationships = parse_xml(parts[DOC_RELS], "document relationships")
    else:
        relationships = minidom.parseString('<Relationships xmlns="%s"/>' % REL)
    rel_root = relationships.documentElement
    if rel_root.namespaceURI != REL or rel_root.localName != "Relationships":
        raise ReviewError("invalid document-relationships root")
    ids = [item.getAttribute("Id") for item in children(rel_root, REL, "Relationship")]
    if "" in ids or len(ids) != len(set(ids)):
        raise ReviewError("invalid or duplicate document relationship IDs")
    relation_id = "rIdReviewComments"
    while relation_id in ids:
        relation_id += "_"
    rel_root.appendChild(new_element(relationships, REL, rel_root.prefix, "Relationship", {
        "Id": relation_id, "Type": (STRICT_REL if word_namespace == STRICT_WORD else OFFICE_REL) + "comments",
        "Target": "comments.xml",
    }))
    ct_root.appendChild(new_element(content_types, CT, ct_root.prefix, "Override", {
        "PartName": "/" + COMMENTS, "ContentType": COMMENTS_TYPE,
    }))
    if not any(item.getAttribute("Extension") == "rels" for item in children(ct_root, CT, "Default")):
        default = new_element(content_types, CT, ct_root.prefix, "Default", {
            "Extension": "rels", "ContentType": REL_TYPE,
        })
        ct_root.insertBefore(default, ct_root.firstChild)

    prefix = root.prefix
    if not prefix:
        prefix = "w"
        while root.hasAttribute("xmlns:" + prefix) and root.getAttribute("xmlns:" + prefix) != word_namespace:
            prefix += "_"
        root.setAttributeNS(XMLNS, "xmlns:" + prefix, word_namespace)
    comments = minidom.parseString('<w:comments xmlns:w="%s"/>' % word_namespace)
    groups = defaultdict(list)
    for comment_id, note in enumerate(notes):
        groups[note["paragraph"]].append(str(comment_id))
        comment = comments.createElementNS(word_namespace, "w:comment")
        comment.setAttributeNS(word_namespace, "w:id", str(comment_id))
        comment.setAttributeNS(word_namespace, "w:author", "Review tool")
        for line in note["comment"].replace("\r\n", "\n").replace("\r", "\n").split("\n"):
            paragraph = comments.createElementNS(word_namespace, "w:p")
            run = comments.createElementNS(word_namespace, "w:r")
            text = comments.createElementNS(word_namespace, "w:t")
            text.setAttributeNS(XML, "xml:space", "preserve")
            text.appendChild(comments.createTextNode(line))
            run.appendChild(text)
            paragraph.appendChild(run)
            comment.appendChild(paragraph)
        comments.documentElement.appendChild(comment)
    for number, comment_ids in groups.items():
        paragraph = paragraphs[number - 1]
        first_content = next((node for node in paragraph.childNodes
                              if node.nodeType == Node.ELEMENT_NODE
                              and not (node.namespaceURI == word_namespace and node.localName == "pPr")), None)
        for comment_id in comment_ids:
            start = new_element(document, word_namespace, prefix, "commentRangeStart")
            start.setAttributeNS(word_namespace, prefix + ":id", comment_id)
            paragraph.insertBefore(start, first_content)
        for comment_id in reversed(comment_ids):
            end = new_element(document, word_namespace, prefix, "commentRangeEnd")
            end.setAttributeNS(word_namespace, prefix + ":id", comment_id)
            paragraph.appendChild(end)
        for comment_id in comment_ids:
            run = new_element(document, word_namespace, prefix, "r")
            reference = new_element(document, word_namespace, prefix, "commentReference")
            reference.setAttributeNS(word_namespace, prefix + ":id", comment_id)
            run.appendChild(reference)
            paragraph.appendChild(run)
    replacements = {
        DOC: document.toxml(encoding="UTF-8"),
        DOC_RELS: relationships.toxml(encoding="UTF-8"),
        "[Content_Types].xml": content_types.toxml(encoding="UTF-8"),
        COMMENTS: comments.toxml(encoding="UTF-8"),
    }
    # Complete ZIP serialization before the output path is ever created.
    buffer = BytesIO()
    with ZipFile(buffer, "w", compression=ZIP_DEFLATED) as output:
        output.comment = archive_comment
        for info in infos:
            output.writestr(copy(info), replacements.get(info.filename, parts[info.filename]))
        for name in sorted(replacements.keys() - parts.keys()):
            output.writestr(name, replacements[name])
    return buffer.getvalue(), sorted(replacements.keys() - parts.keys()), sorted(replacements.keys() & parts.keys())


def write_exclusive(path, data):
    descriptor = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    identity = os.fstat(descriptor)
    try:
        with os.fdopen(descriptor, "wb") as output:
            output.write(data)
            output.flush()
            os.fsync(output.fileno())
    except BaseException:
        # Remove only the file we created, never a replacement at the same path.
        try:
            current = path.lstat()
            if (current.st_dev, current.st_ino) == (identity.st_dev, identity.st_ino):
                path.unlink()
        except OSError:
            pass
        raise


def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = SafeParser(description=__doc__)
    parser.add_argument("source", metavar="SOURCE.docx", type=Path)
    parser.add_argument("--expected-sha256", required=True)
    parser.add_argument("--notes", required=True, type=Path)
    parser.add_argument("--out", required=True, type=Path)
    args = parser.parse_args()
    try:
        if not re.fullmatch(r"[0-9a-fA-F]{64}", args.expected_sha256):
            raise ReviewError("expected SHA-256 must contain exactly 64 hexadecimal characters")
        if os.path.lexists(args.out):
            raise ReviewError("output already exists; refusing to overwrite it")
        source_bytes = args.source.read_bytes()
        source_hash = hashlib.sha256(source_bytes).hexdigest()
        if source_hash != args.expected_sha256.lower():
            raise ReviewError("source SHA-256 does not match the expected value")
        try:
            notes = json.loads(args.notes.read_text(encoding="utf-8"))
        except (UnicodeError, json.JSONDecodeError):
            raise ReviewError("notes must contain valid UTF-8 JSON") from None
        review, added, changed = build_review(source_bytes, notes)
        write_exclusive(args.out, review)
    except ReviewError as error:
        print("create_review_copy: %s" % error, file=sys.stderr)
        return 2
    except FileExistsError:
        print("create_review_copy: output already exists; refusing to overwrite it", file=sys.stderr)
        return 2
    except OSError as error:
        print("create_review_copy: file operation failed (errno %s)" % error.errno, file=sys.stderr)
        return 2
    except (BadZipFile, RuntimeError, ValueError, NotImplementedError):
        print("create_review_copy: invalid or unsupported DOCX package", file=sys.stderr)
        return 2
    print(json.dumps({
        "schema_version": 1, "source_sha256": source_hash,
        "review_sha256": hashlib.sha256(review).hexdigest(),
        "comments_added": len(notes), "paragraphs_annotated": sorted({note["paragraph"] for note in notes}),
        "added_parts": added, "changed_parts": changed,
        "limitations": [
            "Paragraph numbers count direct body paragraphs, including empty paragraphs; tables are not targets.",
            "expected_text compares concatenated stored w:t text exactly, not rendered text or math m:t.",
            "Existing comments are rejected. This tool does not edit prose, verify citations, or predict scan scores.",
            "Verify comment visibility and document layout in Word; headless PDF rendering may omit comments.",
        ],
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
