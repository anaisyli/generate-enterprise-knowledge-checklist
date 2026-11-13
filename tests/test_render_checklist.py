import copy
import importlib.util
import json
import os
import subprocess
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "render_checklist.py"
SUBPROCESS_ENV = {**os.environ, "PYTHONIOENCODING": "utf-8"}
spec = importlib.util.spec_from_file_location("render_checklist", SCRIPT)
renderer = importlib.util.module_from_spec(spec)
spec.loader.exec_module(renderer)


def fixture(sub_count=6):
    categories = ["分拣机 (Sorters)", "输送机 (Conveyors)", "测量设备 (Measurement Equipment)"]
    subs = [f"小类{i} (Type {i})" for i in range(1, sub_count + 1)]
    paragraphs = []

    def add(kind, text):
        paragraphs.append({"kind": kind, "text": text})

    def heading(text):
        add("heading", text)

    add("title", "物流设备行业知识库清单")
    add("blue", renderer.BLUE_INTRO)
    for text in ("1. 公司介绍 (Company Profile)", "1.1 公司简介 (Company Overview)",
                 "1.2 公司能力 (Company Capabilities)", "1.3 发展历程 (Company Milestones)",
                 "2. 产品信息 (Product Information)", f"2.1 产品大类1：{categories[0]}"):
        heading(text)
    for i, sub in enumerate(subs[:5], 1):
        heading(f"2.1.{i} {sub}")
        for label in renderer.PRODUCT_LABELS:
            add("label", label)
            add("text", "请提供相关资料。")
    add("blue", renderer.BLUE_FIRST)
    for i, category in enumerate(categories[1:], 2):
        heading(f"2.{i} 产品大类{i}：{category}")
        add("blue", renderer.BLUE_OTHER)
    for text in ("3. 解决方案 (Solutions)", "3.1 ODM/OEM解决方案 (ODM/OEM Solutions)",
                 "3.2 行业解决方案 (Industry Solutions)", "4. 客户成功案例 (Customer Success Stories)",
                 "4.1 行业案例 (Industry Cases)", "4.2 其他案例 (Other Cases)",
                 "5. 行业趋势与未来展望 (Industry Trends & Outlook)"):
        heading(text)
    add("blue", renderer.BLUE_TRENDS)
    for text in ("5.1 技术趋势 (Technology Trends)", "5.2 市场趋势 (Market Trends)",
                 "6. 常见问题 (FAQ)", "7. 其他相关文件 (Other Documents)"):
        heading(text)
    return {"product_categories": categories, "first_category_subcategories": subs,
            "paragraphs": paragraphs}


class ChecklistTests(unittest.TestCase):
    def test_reference_templates_are_anonymized_and_keep_blue_notes(self):
        skill_root = SCRIPT.parents[1]
        templates = list((skill_root / "assets" / "templates").glob("*.docx"))
        self.assertEqual(len(templates), 5)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        for path in templates:
            with self.subTest(template=path.name), zipfile.ZipFile(path) as z:
                self.assertFalse(any(name.startswith("docProps/") for name in z.namelist()))
                xml = ET.fromstring(z.read("word/document.xml"))
                self.assertGreater(len(xml.findall(".//w:color[@w:val='0000FF']", ns)), 0)
                self.assertEqual(xml.findall(".//w:hyperlink", ns), [])
                body = "".join(t.text or "" for t in xml.findall(".//w:t", ns))
                self.assertNotRegex(body, r"https?://|www\.")
                if path.name.startswith("体外诊断"):
                    self.assertIn("示例型号甲", body)

    def test_documentation_does_not_contain_real_urls(self):
        skill_root = SCRIPT.parents[1]
        for path in (skill_root / "README.md", skill_root / "SKILL.md",
                     skill_root / "references" / "workflow.md"):
            with self.subTest(path=path.name):
                text = path.read_text(encoding="utf-8")
                self.assertNotRegex(text, r"https?://|www\.")
                self.assertNotRegex(text, r"(?i)(?:[A-Z]:\\|\\\\[^\\\s]+\\[^\\\s]+)")

    def test_limit_to_five_and_handle_fewer(self):
        for count in (1, 5, 6):
            with self.subTest(count=count):
                paragraphs = renderer.validate(fixture(count))
                heads = [p for p in paragraphs if p["kind"] == "heading"
                         and p["text"].startswith("2.1.")]
                self.assertEqual(len(heads), min(count, 5))

    def test_rejects_missing_product_subcategories(self):
        data = fixture()
        data["first_category_subcategories"] = []
        with self.assertRaisesRegex(ValueError, "非空"):
            renderer.validate(data)

    def test_rejects_number_gap_and_wrong_product_name(self):
        data = fixture()
        for p in data["paragraphs"]:
            if p["text"].startswith("2.2 "):
                p["text"] = p["text"].replace("2.2 ", "2.4 ", 1)
                break
        with self.assertRaisesRegex(ValueError, "不连续"):
            renderer.validate(data)
        data = fixture()
        data["first_category_subcategories"][0] = "其他 (Other)"
        with self.assertRaisesRegex(ValueError, "名称或顺序"):
            renderer.validate(data)

    def test_rejects_missing_or_moved_blue_and_extra_product_prose(self):
        for mutation in ("intro", "first", "other", "trend"):
            data = fixture()
            text = {"intro": renderer.BLUE_INTRO, "first": renderer.BLUE_FIRST,
                    "other": renderer.BLUE_OTHER, "trend": renderer.BLUE_TRENDS}[mutation]
            i = next(i for i, p in enumerate(data["paragraphs"]) if p["text"] == text)
            data["paragraphs"].pop(i)
            with self.subTest(mutation=mutation), self.assertRaises(ValueError):
                renderer.validate(data)
        data = fixture()
        i = next(i for i, p in enumerate(data["paragraphs"]) if p["text"] == renderer.BLUE_OTHER)
        data["paragraphs"].insert(i + 1, {"kind": "text", "text": "不应展开。"})
        with self.assertRaisesRegex(ValueError, "只能包含"):
            renderer.validate(data)

    def test_rejects_missing_label_and_extra_subcategory(self):
        data = fixture()
        data["paragraphs"].remove({"kind": "label", "text": renderer.PRODUCT_LABELS[0]})
        with self.assertRaisesRegex(ValueError, "四个标准"):
            renderer.validate(data)
        data = fixture()
        i = next(i for i, p in enumerate(data["paragraphs"]) if p["text"] == renderer.BLUE_FIRST)
        data["paragraphs"].insert(i, {"kind": "heading", "text": "2.1.6 小类6 (Type 6)"})
        with self.assertRaisesRegex(ValueError, "前五个"):
            renderer.validate(data)

    def test_output_is_word_xml_with_blue_text_and_no_tables(self):
        data = fixture()
        paragraphs = renderer.validate(data)
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source, out = root / "input.json", root / "result.docx"
            source.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
            run = subprocess.run([sys.executable, str(SCRIPT), str(source), str(out)],
                                 capture_output=True, text=True, encoding="utf-8", env=SUBPROCESS_ENV)
            self.assertEqual(run.returncode, 0, run.stderr)
            with zipfile.ZipFile(out) as z:
                for name in z.namelist():
                    ET.fromstring(z.read(name))
                xml = ET.fromstring(z.read("word/document.xml"))
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            self.assertEqual(xml.findall(".//w:tbl", ns), [])
            result = []
            blue = []
            for p in xml.findall(".//w:body/w:p", ns):
                text = "".join(t.text or "" for t in p.findall(".//w:t", ns))
                result.append(text)
                if p.find(".//w:color[@w:val='0000FF']", ns) is not None:
                    blue.append(text)
            self.assertEqual(result, [p["text"] for p in paragraphs])
            self.assertEqual(blue, [p["text"] for p in paragraphs if p["kind"] == "blue"])
            before = out.read_bytes()
            rerun = subprocess.run([sys.executable, str(SCRIPT), str(source), str(out)],
                                   capture_output=True, text=True, encoding="utf-8", env=SUBPROCESS_ENV)
            self.assertNotEqual(rerun.returncode, 0)
            self.assertEqual(out.read_bytes(), before)

    def test_invalid_input_creates_no_output(self):
        data = copy.deepcopy(fixture())
        data["paragraphs"][1]["text"] = "不正确的提示"
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source, out = root / "input.json", root / "result.docx"
            source.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
            run = subprocess.run([sys.executable, str(SCRIPT), str(source), str(out)],
                                 capture_output=True, text=True, encoding="utf-8", env=SUBPROCESS_ENV)
            self.assertNotEqual(run.returncode, 0)
            self.assertFalse(out.exists())


if __name__ == "__main__":
    unittest.main()
