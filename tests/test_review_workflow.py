import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from test_preflight import make_docx, run_script, table, text_run


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CONTENT_TYPE = "{http://schemas.openxmlformats.org/package/2006/content-types}"
PAST_TEXT = "The first draft drew its conclusions from one uninterrupted observation period."
CURRENT_TEXT = "The current draft compares two observation periods and reports their separate results."
TABLE_TEXT = "The table records the sensor settings shared by both observation periods."
NOTE = "Which interval in the observation log supports the comparison?"
EQUATION = '<w:p><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x = 5</m:t></m:r></m:oMath></w:p>'
WORD_EXTENSIONS = (
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
    ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)
DOCUMENT_RELS = "word/_rels/document.xml.rels"
HYPERLINK_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
HYPERLINK_TARGET = "https://example.invalid/observation-log"
LINK_TEXT = "Synthetic observation log"
RELS_XML = (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId7" Type="{HYPERLINK_TYPE}" Target="{HYPERLINK_TARGET}" TargetMode="External"/>'
    '</Relationships>'
)


def manuscript(text):
    styled_paragraph = (
        '<w:p w14:paraId="14589A10"><w:pPr><w:spacing w:after="240"/></w:pPr>'
        + text_run(text) + '</w:p>'
    )
    page_break = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    link = '<w:p><w:hyperlink r:id="rId7">' + text_run(LINK_TEXT) + '</w:hyperlink></w:p>'
    return styled_paragraph + table(TABLE_TEXT) + EQUATION + page_break + link


def comment_free_body(document):
    body = ET.fromstring(document).find(W + "body")
    for para in body.iter(W + "p"):
        for child in list(para):
            if child.tag in {W + "commentRangeStart", W + "commentRangeEnd"}:
                para.remove(child)
            elif child.tag == W + "r" and child.find(W + "commentReference") is not None:
                if child.find(W + "t") is not None:
                    raise AssertionError("The comment reference shares a run with body text")
                para.remove(child)
    return ET.tostring(body)


def raster_pages(docx, directory):
    directory.mkdir()
    profile = (directory / "profile").resolve().as_uri()
    subprocess.run(
        ["soffice", f"-env:UserInstallation={profile}", "--headless", "--convert-to", "pdf:writer_pdf_Export",
         "--outdir", str(directory), str(docx)],
        check=True, capture_output=True, timeout=90,
    )
    pdf = directory / f"{docx.stem}.pdf"
    text = subprocess.check_output(["pdftotext", "-enc", "UTF-8", str(pdf), "-"], timeout=30).decode("utf-8")
    rendered_text = " ".join(text.split())
    if CURRENT_TEXT not in rendered_text or LINK_TEXT not in rendered_text:
        raise AssertionError("Rendered PDF is missing expected text from the current manuscript")
    subprocess.run(["pdftoppm", "-r", "96", str(pdf), str(directory / "page")], check=True, timeout=30)
    return [path.read_bytes() for path in sorted(directory.glob("page-*.ppm"))]


class ReviewWorkflowTest(unittest.TestCase):
    def setUp(self):
        workspace = tempfile.TemporaryDirectory(prefix="copykillerkiller-e2e-")
        self.addCleanup(workspace.cleanup)
        self.root = Path(workspace.name)
        self.past = self.root / "past" / "manuscript.docx"
        self.current = self.root / "current" / "manuscript.docx"
        self.review = self.root / "review.docx"
        self.notes = self.root / "notes.json"
        self.report = self.root / "synthetic-past-report.txt"
        self.past.parent.mkdir()
        self.current.parent.mkdir()
        extra_parts = {DOCUMENT_RELS: RELS_XML}
        make_docx(self.past, manuscript(PAST_TEXT), root_attributes=WORD_EXTENSIONS, extra_parts=extra_parts)
        make_docx(self.current, manuscript(CURRENT_TEXT), root_attributes=WORD_EXTENSIONS, extra_parts=extra_parts)
        self.report.write_text(
            "AI작성률  전체 영역  AI 의심 영역  높음  중간  낮음\n40%  3  2  0  1  1\f"
            f"낮음   {PAST_TEXT}\n\n중간   {TABLE_TEXT}\f",
            encoding="utf-8",
        )
        self.write_note(CURRENT_TEXT)

    def write_note(self, expected_text):
        self.notes.write_text(json.dumps([{
            "paragraph": 1, "expected_text": expected_text, "comment": NOTE,
        }]), encoding="utf-8")

    def create_review(self, source, digest, destination=None):
        return run_script("create_review_copy.py", source, "--expected-sha256", digest,
                          "--notes", self.notes, "--out", destination or self.review)

    def read_json(self, result):
        self.assertEqual(result.returncode, 0, result.stderr)
        return json.loads(result.stdout)

    def test_past_report_to_current_review_and_format_verification(self):
        inputs = {path: path.read_bytes() for path in (self.report, self.past, self.current)}
        drafts = [self.past, self.current]
        preflight = self.read_json(run_script("prepare_review.py", "--target-index", 2, self.report, *drafts))
        self.assertEqual(preflight["report_summary"]["score_percent"], 40)
        assessment = preflight["target_assessment"]
        self.assertEqual(assessment["status"], "other_draft_aligned")
        self.assertEqual(assessment["aligned_draft_indices"], [1])
        self.assertEqual(assessment["draft_index"], 2)
        index = assessment["draft_index"] - 1
        working = drafts[index]
        digest = preflight["drafts"][index]["sha256"]
        created = self.read_json(self.create_review(working, digest))
        self.assertEqual(created["source_sha256"], digest)
        self.assertEqual(created["review_sha256"], hashlib.sha256(self.review.read_bytes()).hexdigest())
        self.assertEqual(created["comments_added"], 1)

        verification = self.read_json(run_script("verify_docx.py", working, self.review))
        self.assertEqual(verification["source"]["sha256"], hashlib.sha256(inputs[self.current]).hexdigest())
        self.assertFalse(verification["main_body_text_changed"])
        self.assertFalse(verification["protected_document_xml"]["any_changed"])
        for key in ("tables", "math_objects", "section_properties", "hyperlinks"):
            self.assertEqual(verification["protected_document_xml"]["elements"][key]["source_count"], 1)
        with ZipFile(self.current) as source, ZipFile(self.review) as review:
            self.assertIsNone(review.testzip())
            changed_parts = {"[Content_Types].xml", "word/document.xml", DOCUMENT_RELS}
            self.assertEqual(set(review.namelist()), set(source.namelist()) | {"word/comments.xml"})
            for part in set(source.namelist()) - changed_parts:
                self.assertEqual(source.read(part), review.read(part), part)
            source_xml = source.read("word/document.xml")
            review_xml = review.read("word/document.xml")
            self.assertEqual(comment_free_body(source_xml), comment_free_body(review_xml))
            self.assertIn(b"xmlns:w14=", review_xml)
            self.assertIn(b'mc:Ignorable="w14"', review_xml)
            self.assertIn(CURRENT_TEXT.encode(), review_xml)
            self.assertNotIn(PAST_TEXT.encode(), review_xml)
            doc = ET.fromstring(review_xml)
            para = doc.find(f"{W}body/{W}p")
            self.assertEqual(para[0].tag, W + "pPr")
            anchors = [list(para.iter(W + name)) for name in ("commentRangeStart", "commentRangeEnd", "commentReference")]
            self.assertEqual([len(nodes) for nodes in anchors], [1, 1, 1])
            comments = ET.fromstring(review.read("word/comments.xml")).findall(W + "comment")
            self.assertEqual(len(comments), 1)
            self.assertEqual("".join(comments[0].itertext()).strip(), NOTE)
            self.assertEqual([nodes[0].get(W + "id") for nodes in anchors], [comments[0].get(W + "id")] * 3)
            rels = ET.fromstring(review.read(DOCUMENT_RELS)).findall(REL + "Relationship")
            comment_rels = [rel for rel in rels if
                rel.get("Type", "").endswith("/comments") and rel.get("Target") == "comments.xml"
            ]
            self.assertEqual(len(comment_rels), 1)
            source_rels = ET.fromstring(source.read(DOCUMENT_RELS)).findall(REL + "Relationship")
            self.assertEqual([rel.attrib for rel in rels if rel not in comment_rels], [rel.attrib for rel in source_rels])
            types = ET.fromstring(review.read("[Content_Types].xml"))
            source_types = ET.fromstring(source.read("[Content_Types].xml"))
            self.assertTrue(any(
                item.get("PartName") == "/word/comments.xml"
                and item.get("ContentType") == "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
                for item in types.findall(CONTENT_TYPE + "Override")
            ))
            self.assertEqual(
                [(item.tag, item.attrib) for item in types if item.get("PartName") != "/word/comments.xml"],
                [(item.tag, item.attrib) for item in source_types],
            )

        if os.environ.get("COPYKILLER_RENDER_TEST") == "1":
            missing = [tool for tool in ("soffice", "pdftoppm", "pdftotext") if shutil.which(tool) is None]
            self.assertFalse(missing, f"Missing rendering tools: {', '.join(missing)}")
            source_pages = raster_pages(self.current, self.root / "current-render")
            review_pages = raster_pages(self.review, self.root / "review-render")
            self.assertEqual(len(source_pages), 2)
            self.assertEqual(len(source_pages), len(review_pages))
            for page, (source_image, review_image) in enumerate(zip(source_pages, review_pages), 1):
                self.assertTrue(source_image == review_image, f"Rendered page {page} differs")

        self.assertEqual({path: path.read_bytes() for path in inputs}, inputs)

    def test_stale_hash_and_comment_anchor_do_not_create_a_copy(self):
        current_digest = hashlib.sha256(self.current.read_bytes()).hexdigest()
        past_digest = hashlib.sha256(self.past.read_bytes()).hexdigest()
        for source, digest in ((self.current, past_digest), (self.past, current_digest), (self.past, past_digest)):
            with self.subTest(source=source.parent.name, digest_matches_current=digest == current_digest):
                result = self.create_review(source, digest)
                self.assertNotEqual(result.returncode, 0)
                self.assertFalse(self.review.exists())

        self.write_note(PAST_TEXT)
        result = self.create_review(self.current, current_digest)
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(self.review.exists())

    def test_existing_output_is_preserved(self):
        self.review.write_bytes(b"existing review")
        current_digest = hashlib.sha256(self.current.read_bytes()).hexdigest()
        result = self.create_review(self.current, current_digest)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(self.review.read_bytes(), b"existing review")

    def test_existing_comments_require_a_preserving_editor(self):
        current_digest = hashlib.sha256(self.current.read_bytes()).hexdigest()
        self.read_json(self.create_review(self.current, current_digest))
        annotated = self.review.read_bytes()
        destination = self.root / "second-review.docx"
        result = self.create_review(self.review, hashlib.sha256(annotated).hexdigest(), destination)
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(destination.exists())
        self.assertEqual(self.review.read_bytes(), annotated)
