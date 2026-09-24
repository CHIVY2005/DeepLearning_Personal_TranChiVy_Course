import re
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


SUBMISSION = Path(__file__).with_name(
    "BT01_Regression_TranChiVy_3123580065.docx"
)
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def document_paragraphs():
    if not SUBMISSION.exists() or not zipfile.is_zipfile(SUBMISSION):
        return []
    with zipfile.ZipFile(SUBMISSION) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    return [
        "".join(node.text or "" for node in paragraph.iter(f"{W_NS}t")).strip()
        for paragraph in root.iter(f"{W_NS}p")
    ]


class SubmissionDocumentTests(unittest.TestCase):
    def test_submission_is_a_valid_docx_with_student_identity(self):
        self.assertTrue(SUBMISSION.exists(), f"Missing submission: {SUBMISSION.name}")
        self.assertTrue(zipfile.is_zipfile(SUBMISSION), "Submission is not a valid DOCX")
        text = "\n".join(document_paragraphs())
        self.assertIn("Trần Chí Vỹ", text)
        self.assertIn("3123580065", text)

    def test_slide_example_contains_all_three_solution_methods(self):
        text = "\n".join(document_paragraphs()).lower()
        self.assertIn("giải hệ phương trình chuẩn", text)
        self.assertIn("pseudo-inverse", text)
        self.assertIn("gradient descent", text)
        self.assertRegex(text, r"a\s*=\s*0[,.]9767")
        self.assertRegex(text, r"b\s*=\s*[−-]102[,.]2093")

    def test_both_exercise_sets_have_all_25_numbered_solutions(self):
        paragraphs = document_paragraphs()
        self.assertIn("PHẦN II. BÀI TẬP 1", paragraphs)
        self.assertIn("PHẦN III. BÀI TẬP 2", paragraphs)
        set_1 = paragraphs.index("PHẦN II. BÀI TẬP 1")
        set_2 = paragraphs.index("PHẦN III. BÀI TẬP 2")
        first_count = sum(bool(re.fullmatch(r"Câu (?:[1-9]|10)\..+", p)) for p in paragraphs[set_1:set_2])
        second_count = sum(bool(re.fullmatch(r"Câu (?:[1-9]|1[0-5])\..+", p)) for p in paragraphs[set_2:])
        self.assertEqual(first_count, 10)
        self.assertEqual(second_count, 15)

    def test_document_has_no_unresolved_placeholders(self):
        text = "\n".join(document_paragraphs())
        self.assertTrue(text, "Document text is missing")
        for placeholder in ("TODO", "TBD", "HoTen", "MaSV", "□"):
            self.assertNotIn(placeholder, text)


if __name__ == "__main__":
    unittest.main()
