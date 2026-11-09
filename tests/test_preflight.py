import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from xml.sax.saxutils import escape
from zipfile import ZipFile


SCRIPTS = Path(__file__).resolve().parents[1] / "copykillerkiller" / "scripts"
SENTENCE = "가상 실험에서는 조사자가 빛의 세기를 바꾸고 같은 간격으로 온도를 측정하였다."
TABLE_SENTENCE = "비교에 사용한 센서는 모두 동일한 눈금과 단위를 표시하도록 설정하였다."
SUMMARY = "AI작성률   전체 영역   AI 의심 영역   높음   중간   낮음\n\n52%   12   6   0   2   4"


def text_run(text):
    return f"<w:r><w:t>{escape(text)}</w:t></w:r>"


def paragraph(text):
    return f"<w:p>{text_run(text)}</w:p>"


def table(text):
    return f"<w:tbl><w:tr><w:tc>{paragraph(text)}</w:tc></w:tr></w:tbl>"


def make_docx(path, body, *, root_attributes="", extra_parts=None):
    with ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>""")
        package.writestr("_rels/.rels", """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>""")
        package.writestr("word/document.xml", f"""<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"{root_attributes}>
<w:body>{body}<w:sectPr/></w:body></w:document>""")
        for name, data in (extra_parts or {}).items():
            package.writestr(name, data)


def run_script(script, *paths):
    return subprocess.run(
        [sys.executable, str(SCRIPTS / script), *map(str, paths)],
        capture_output=True, encoding="utf-8",
    )


class PreflightTest(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.root = Path(folder.name)
        self.report = self.root / "가상 보고서.txt"
        self.original = self.root / "원고 A.docx"
        self.revised = self.root / "원고 B.docx"

    def test_candidates_keep_report_score_and_map_table(self):
        self.report.write_text(f"{SUMMARY}\f낮음   {SENTENCE}\n\n중간   {TABLE_SENTENCE}\f아래 쪽은 표시 경계 확인이 필요하다.\f", encoding="utf-8")
        make_docx(self.original, paragraph(SENTENCE) + table(TABLE_SENTENCE))
        make_docx(self.revised, paragraph(SENTENCE.replace("온도", "습도")) + table(TABLE_SENTENCE))
        before = self.original.read_bytes()
        result = run_script("prepare_review.py", self.report, self.original, self.revised)
        self.assertEqual(result.returncode, 0, result.stderr)
        data = json.loads(result.stdout)
        self.assertEqual(data["report_summary"]["score_percent"], 52)
        self.assertEqual(data["report_summary"]["suspected_regions"], 6)
        self.assertEqual(data["report_pages"], 3)
        self.assertTrue(data["visual_review_required"])
        self.assertEqual([mark["page"] for mark in data["marker_candidates"]], [2, 2])
        first, second = data["drafts"]
        self.assertEqual(first["sha256"], hashlib.sha256(before).hexdigest())
        self.assertEqual(first["marker_matches"][0]["status"], "unique")
        self.assertEqual(second["marker_matches"][0]["status"], "unmatched")
        self.assertEqual(first["marker_matches"][1]["locations"][0]["start"], "body/tbl[1]/tr[1]/tc[1]/p[1]")
        self.assertNotIn(SENTENCE, result.stdout)
        self.assertNotIn(self.original.name, result.stdout)
        self.assertEqual(self.original.read_bytes(), before)

    def test_repeated_and_label_only_matches_stay_unresolved(self):
        self.report.write_text(f"낮음   {SENTENCE}\n\n높음\n    {TABLE_SENTENCE}\f\f", encoding="utf-8")
        make_docx(self.original, paragraph(SENTENCE) * 2)
        result = run_script("prepare_review.py", self.report, self.original)
        self.assertEqual(result.returncode, 0, result.stderr)
        data = json.loads(result.stdout)
        self.assertEqual(data["blank_text_pages"], [2])
        self.assertEqual(data["report_summary"]["status"], "not_found")
        matches = data["drafts"][0]["marker_matches"]
        self.assertEqual([match["status"] for match in matches], ["ambiguous", "insufficient_text"])

    def test_report_alignment_does_not_replace_the_working_target(self):
        self.report.write_text(f"{SUMMARY}\f낮음   {SENTENCE}\n\n중간   {TABLE_SENTENCE}", encoding="utf-8")
        body = paragraph(SENTENCE) + table(TABLE_SENTENCE)
        make_docx(self.original, body)
        make_docx(self.revised, paragraph("이 문단은 새로운 자료로 수정하였다.") + table(TABLE_SENTENCE))
        annotated = self.root / "검토 사본.docx"
        make_docx(annotated, body.replace("<w:p>", '<w:p><w:commentRangeStart w:id="0"/>', 1))

        inputs = (self.report, self.original, self.revised, annotated)
        for target, expected in ((2, "other_draft_aligned"), (3, "shared_alignment")):
            with self.subTest(target=target):
                result = run_script("prepare_review.py", "--target-index", target, *inputs)
                self.assertEqual(result.returncode, 0, result.stderr)
                data = json.loads(result.stdout)
                self.assertEqual(data["target_assessment"], {
                    "draft_index": target, "status": expected,
                    "aligned_draft_indices": [1, 3], "distinct_anchors_checked": 2,
                })
                self.assertEqual(data["report_summary"]["score_percent"], 52)
        result = run_script("prepare_review.py", "--target-index", 1, self.report, self.original, self.revised)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)["target_assessment"]["status"], "target_aligned")
        result = run_script("prepare_review.py", "--target-index", 1, self.report, self.revised)
        self.assertEqual(result.returncode, 0, result.stderr)
        data = json.loads(result.stdout)
        self.assertEqual(data["target_assessment"]["status"], "unresolved")

    def test_single_repeated_report_anchor_is_not_enough_to_flag_target(self):
        self.report.write_text(f"낮음   {SENTENCE}\f중간   {SENTENCE}", encoding="utf-8")
        make_docx(self.original, paragraph(SENTENCE))
        make_docx(self.revised, paragraph("다른 문단을 작성했다."))
        result = run_script("prepare_review.py", "--target-index", 2, self.report, self.original, self.revised)
        self.assertEqual(result.returncode, 0, result.stderr)
        assessment = json.loads(result.stdout)["target_assessment"]
        self.assertEqual(assessment["distinct_anchors_checked"], 1)
        self.assertEqual(assessment["status"], "unresolved")

    def test_invalid_docx_exits_without_results(self):
        self.report.write_text(SUMMARY, encoding="utf-8")
        self.original.write_bytes(b"not a package")
        result = run_script("prepare_review.py", self.report, self.original)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(result.stdout, "")
        self.assertIn("prepare_review", result.stderr)

    def test_revision_preflight_separates_prose_from_table_changes(self):
        make_docx(self.original, paragraph(SENTENCE) + table(TABLE_SENTENCE))
        before = self.original.read_bytes()
        changed_prose = paragraph(SENTENCE.replace("온도", "습도"))
        make_docx(self.revised, changed_prose + table(TABLE_SENTENCE))
        result = run_script("verify_docx.py", self.original, self.revised)
        self.assertEqual(result.returncode, 0, result.stderr)
        data = json.loads(result.stdout)
        self.assertEqual(data["source"]["sha256"], hashlib.sha256(before).hexdigest())
        self.assertEqual(data["zip_parts"]["changed"], ["word/document.xml"])
        self.assertTrue(data["main_body_text_changed"])
        self.assertFalse(data["protected_document_xml"]["any_changed"])
        self.assertNotIn(self.original.name, result.stdout)

        make_docx(self.revised, changed_prose + table("표 내용이 달라졌다."))
        result = run_script("verify_docx.py", self.original, self.revised)
        self.assertEqual(result.returncode, 0, result.stderr)
        data = json.loads(result.stdout)
        self.assertTrue(data["protected_document_xml"]["elements"]["tables"]["changed"])
        self.assertEqual(self.original.read_bytes(), before)

    def test_annotations_and_run_splits_keep_main_body_text(self):
        make_docx(self.original, paragraph(SENTENCE))
        midpoint = len(SENTENCE) // 2
        annotated = (
            '<w:p><w:commentRangeStart w:id="0"/>'
            + text_run(SENTENCE[:midpoint]) + text_run(SENTENCE[midpoint:])
            + '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>'
        )
        make_docx(self.revised, annotated)
        result = run_script("verify_docx.py", self.original, self.revised)
        self.assertEqual(result.returncode, 0, result.stderr)
        data = json.loads(result.stdout)
        self.assertNotEqual(data["source"]["sha256"], data["revision"]["sha256"])
        self.assertFalse(data["main_body_text_changed"])

    def test_revision_preflight_rejects_same_file(self):
        make_docx(self.original, paragraph(SENTENCE))
        result = run_script("verify_docx.py", self.original, self.original)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(result.stdout, "")

    def test_missing_file_errors_omit_personal_paths(self):
        self.report.write_text(SUMMARY, encoding="utf-8")
        make_docx(self.original, paragraph(SENTENCE))
        for script, first in (("prepare_review.py", self.report), ("verify_docx.py", self.original)):
            with self.subTest(script=script):
                result = run_script(script, first, self.revised)
                self.assertNotEqual(result.returncode, 0)
                self.assertEqual(result.stdout, "")
                self.assertNotIn(self.revised.name, result.stderr)
                self.assertNotIn(str(self.root), result.stderr)


if __name__ == "__main__":
    unittest.main()
