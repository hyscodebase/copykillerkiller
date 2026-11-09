#!/usr/bin/env python3
"""Find evidence linking a Copykiller layout report to candidate DOCX drafts."""

import argparse
from bisect import bisect_right
from collections import Counter
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
import unicodedata
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile


WORD_NAMESPACES = {
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/wordprocessingml/main",
}
CONTAINERS = {"tbl", "tr", "tc", "sdt", "sdtContent", "customXml", "ins", "moveTo"}
MARKER = re.compile(r"^[ \t]*(낮음|중간|높음)(?:(?:\t[ \t]*| {2,})(.*)|[ \t]*)$")
SUMMARY_HEADER = re.compile(r"AI\s*작성률\s+전체\s*영역\s+AI\s*의심\s*영역\s+높음\s+중간\s+낮음")
SUMMARY_ROW = re.compile(r"(\d{1,3})\s*%\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)")
ANCHOR_MIN = 24
ANCHOR_MAX = 56
MATCH_LIMIT = 5


def compact(text):
    return unicodedata.normalize("NFC", "".join(text.split()))


def word_name(element):
    namespace, _, name = element.tag.lstrip("{").partition("}")
    return name if namespace in WORD_NAMESPACES else ""


def inline_text(paragraph):
    for child in paragraph:
        name = word_name(child)
        if name == "t":
            yield child.text or ""
        elif name not in {"p", "del", "moveFrom", "txbxContent"}:
            yield from inline_text(child)


def docx_blocks(parent, prefix):
    counts = Counter()
    for child in parent:
        name = word_name(child)
        if name not in CONTAINERS and name != "p":
            continue
        counts[name] += 1
        location = f"{prefix}/{name}[{counts[name]}]"
        if name == "p":
            normalized = compact("".join(inline_text(child)))
            if normalized:
                yield location, normalized
        else:
            yield from docx_blocks(child, location)


def load_draft(path):
    if path.suffix.lower() != ".docx":
        raise ValueError("원고는 .docx 파일이어야 합니다.")
    with path.open("rb") as source:
        digest = hashlib.sha256()
        for chunk in iter(lambda: source.read(1 << 20), b""):
            digest.update(chunk)
        source.seek(0)
        with ZipFile(source) as package:
            document = package.read("word/document.xml")
            if b"<!DOCTYPE" in document.replace(b"\x00", b"").upper():
                raise ValueError("DOCX 본문에 지원하지 않는 DTD 선언이 있습니다.")
            root = ET.fromstring(document)
    body = next((child for child in root if word_name(child) == "body"), None)
    if body is None:
        raise ValueError("DOCX에서 본문을 찾지 못했습니다.")
    return digest.hexdigest(), list(docx_blocks(body, "body"))


def load_report(path):
    suffix = path.suffix.lower()
    if suffix == ".txt":
        return path.read_text(encoding="utf-8-sig")
    if suffix != ".pdf":
        raise ValueError("보고서는 .pdf 또는 UTF-8 .txt 파일이어야 합니다.")
    try:
        result = subprocess.run(
            ["pdftotext", "-layout", "-enc", "UTF-8", str(path.resolve()), "-"],
            capture_output=True, encoding="utf-8", check=True, timeout=90,
        )
    except FileNotFoundError as exc:
        raise ValueError("pdftotext를 찾지 못했습니다. Poppler를 설치하거나 배치가 보존된 UTF-8 텍스트를 사용하세요.") from exc
    except subprocess.CalledProcessError as exc:
        raise ValueError("PDF 텍스트를 읽지 못했습니다. 파일 상태와 열기 권한을 확인하세요.") from exc
    except subprocess.TimeoutExpired as exc:
        raise ValueError("PDF 텍스트 추출 시간이 초과됐습니다.") from exc
    return result.stdout


def report_pages(text):
    pages = text.split("\f")
    if len(pages) > 1 and not pages[-1].strip():
        pages.pop()
    return pages


def find_summary(pages):
    tables = []
    for page_number, page in enumerate(pages, 1):
        lines = [line.strip() for line in page.splitlines() if line.strip()]
        for header, row in zip(lines, lines[1:]):
            if not SUMMARY_HEADER.fullmatch(header):
                continue
            match = SUMMARY_ROW.fullmatch(row)
            if match is None:
                return {"status": "unverified_layout"}
            tables.append((page_number, tuple(map(int, match.groups()))))
    if not tables:
        return {"status": "not_found"}
    if len(tables) > 1:
        return {"status": "multiple_tables"}
    page, (score, total, suspected, high, medium, low) = tables[0]
    if not (0 <= score <= 100 and high + medium + low == suspected <= total):
        return {"status": "inconsistent_counts", "page": page}
    return {
        "status": "extracted", "page": page, "score_percent": score,
        "total_regions": total, "suspected_regions": suspected,
        "severity_counts": {"높음": high, "중간": medium, "낮음": low},
    }


def marker_anchors(pages):
    for page_number, page in enumerate(pages, 1):
        lines = page.splitlines()
        for index, line in enumerate(lines):
            match = MARKER.fullmatch(line)
            if match is None:
                continue
            # A label on its own line is a visual cue without a reliable anchor.
            anchor = compact(match.group(2) or "")
            if anchor:
                for next_line in lines[index + 1:index + 3]:
                    if len(anchor) >= ANCHOR_MIN or not next_line.strip() or MARKER.fullmatch(next_line):
                        break
                    anchor += compact(next_line)
            yield {"page": page_number, "line": index + 1, "label": match.group(1)}, anchor[:ANCHOR_MAX]


def find_anchor(anchor, manuscript, offsets, locations):
    if len(anchor) < ANCHOR_MIN:
        return {"status": "insufficient_text", "locations": []}
    found = []
    start = manuscript.find(anchor)
    while start != -1 and len(found) <= MATCH_LIMIT:
        first = bisect_right(offsets, start) - 1
        last = bisect_right(offsets, start + len(anchor) - 1) - 1
        found.append({"start": locations[first], "end": locations[last]})
        start = manuscript.find(anchor, start + 1)
    status = ("unmatched", "unique", "ambiguous")[min(len(found), 2)]
    return {"status": status, "locations": found[:MATCH_LIMIT], "more_locations": len(found) > MATCH_LIMIT}


def compare_draft(index, path, pages, markers):
    sha256, blocks = load_draft(path)
    offsets = []
    length = 0
    for _, normalized in blocks:
        offsets.append(length)
        length += len(normalized)
    manuscript = "".join(normalized for _, normalized in blocks)
    locations = [location for location, _ in blocks]
    repeats = Counter(normalized for _, normalized in blocks)
    exact = []
    for location, normalized in blocks:
        if len(normalized) < ANCHOR_MAX or repeats[normalized] != 1:
            continue
        matched_pages = [number for number, page in enumerate(pages, 1) if normalized in page]
        if matched_pages:
            exact.append({"location": location, "pages": matched_pages})
    return {
        "draft_index": index, "sha256": sha256, "text_blocks": len(blocks),
        "full_block_match_count": len(exact), "full_block_matches": exact[:12],
        "marker_matches": [
            {**marker, **find_anchor(anchor, manuscript, offsets, locations)}
            for marker, anchor in markers
        ],
    }


def prepare(report_path, draft_paths):
    pages = report_pages(load_report(report_path))
    summary = find_summary(pages)
    markers = list(marker_anchors(pages))
    normalized_pages = [compact(page) for page in pages]
    return {
        "schema_version": 1, "report_pages": len(pages), "report_summary": summary,
        "blank_text_pages": [number for number, page in enumerate(normalized_pages, 1) if not page],
        "marker_candidates": [marker for marker, _ in markers],
        "drafts": [compare_draft(index, path, normalized_pages, markers) for index, path in enumerate(draft_paths, 1)],
        "visual_review_required": True,
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("report", type=Path, help="Copykiller PDF or UTF-8 layout text")
    parser.add_argument("drafts", type=Path, nargs="+", metavar="draft.docx", help="one or more DOCX candidates")
    args = parser.parse_args()
    try:
        result = prepare(args.report, args.drafts)
    except OSError as exc:
        parser.exit(1, f"prepare_review: 입력을 읽지 못했습니다 ({exc.strerror or type(exc).__name__}).\n")
    except (BadZipFile, ET.ParseError, KeyError, RuntimeError, ValueError) as exc:
        parser.exit(1, f"prepare_review: {exc}\n")
    sys.stdout.buffer.write((json.dumps(result, ensure_ascii=False, indent=2) + "\n").encode("utf-8"))


if __name__ == "__main__":
    main()
