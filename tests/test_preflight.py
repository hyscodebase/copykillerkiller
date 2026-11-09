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


def paragraph(text):
    return f"<w:p><w:r><w:t>{escape(text)}</w:t></w:r></w:p>"


def table(text):
    return f"<w:tbl><w:tr><w:tc>{paragraph(text)}</w:tc></w:tr></w:tbl>"


def make_docx(path, body):
    with ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>""")
        package.writestr("_rels/.rels", """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>""")
        package.writestr("word/document.xml", f"""<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>{body}<w:sectPr/></w:body></w:document>""")


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
        self.assertFalse(data["protected_document_xml"]["any_changed"])
        self.assertNotIn(self.original.name, result.stdout)

        make_docx(self.revised, changed_prose + table("표 내용이 달라졌다."))
        result = run_script("verify_docx.py", self.original, self.revised)
        self.assertEqual(result.returncode, 0, result.stderr)
        data = json.loads(result.stdout)
        self.assertTrue(data["protected_document_xml"]["elements"]["tables"]["changed"])
        self.assertEqual(self.original.read_bytes(), before)

    def test_revision_preflight_rejects_same_file(self):
        make_docx(self.original, paragraph(SENTENCE))
        result = run_script("verify_docx.py", self.original, self.original)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(result.stdout, "")


if __name__ == "__main__":
    unittest.main()
