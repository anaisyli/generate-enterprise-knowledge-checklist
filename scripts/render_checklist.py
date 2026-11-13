#!/usr/bin/env python3
"""Validate a checklist manuscript and render a simple, table-free DOCX."""

import argparse
import io
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

BLUE_INTRO = "*建议优先提供英文资料，中文资料也可。"
BLUE_FIRST = "*其余产品小类请按实际情况继续补充，可参照以上框架提供资料。"
BLUE_OTHER = "*请根据实际产品类目补充该大类下的产品小类，并参照以上框架提供资料。"
BLUE_TRENDS = "*此部分可提供已有行业文章、技术文章、白皮书、展会资料、公众号文章、行业报告或内部培训资料。"
PRODUCT_LABELS = (
    "产品概述 (Product Overview)：",
    "产品特性 (Product Features)：",
    "应用场景 (Applications)：",
    "产品规格参数 (Product Specifications)：",
)
HEADING = re.compile(r"^(\d+(?:\.\d+)*)(?:\.\s+|\s+)(\S.*)$")
INVALID_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff\ud800-\udfff]")


def _strings(value, name):
    if not isinstance(value, list) or not value:
        raise ValueError(f"{name} 必须是非空字符串数组。")
    if any(not isinstance(s, str) or not s.strip() or s != s.strip() for s in value):
        raise ValueError(f"{name} 必须只包含非空、首尾无空白的字符串。")
    if len(set(value)) != len(value):
        raise ValueError(f"{name} 中有重复名称。")
    return value


def validate(data):
    if not isinstance(data, dict):
        raise ValueError("JSON 顶层必须是对象。")
    categories = _strings(data.get("product_categories"), "product_categories")
    subcategories = _strings(data.get("first_category_subcategories"), "first_category_subcategories")
    paragraphs = data.get("paragraphs")
    if not isinstance(paragraphs, list) or not paragraphs:
        raise ValueError("paragraphs 必须是非空数组。")
    for i, p in enumerate(paragraphs):
        if not isinstance(p, dict) or p.get("kind") not in {"title", "heading", "label", "text", "blue"}:
            raise ValueError(f"第 {i + 1} 段的 kind 无效。")
        text = p.get("text")
        if not isinstance(text, str) or not text.strip() or text != text.strip():
            raise ValueError(f"第 {i + 1} 段的 text 必须为非空且首尾无空白的字符串。")
        if INVALID_XML.search(text) or "\n" in text or "\r" in text:
            raise ValueError(f"第 {i + 1} 段包含 XML 无效字符或换行，请拆成独立段落。")
    if paragraphs[0]["kind"] != "title" or sum(p["kind"] == "title" for p in paragraphs) != 1:
        raise ValueError("必须仅有一个 title，且位于第一段。")
    if len(paragraphs) < 2 or paragraphs[1] != {"kind": "blue", "text": BLUE_INTRO}:
        raise ValueError("标题后必须紧接英文资料蓝色提示。")

    headings = {}
    last_child = {}
    current = None
    for i, p in enumerate(paragraphs):
        if p["kind"] != "heading":
            continue
        match = HEADING.fullmatch(p["text"])
        if not match:
            raise ValueError(f"第 {i + 1} 段标题编号格式无效：{p['text']}")
        parts = tuple(int(s) for s in match.group(1).split("."))
        if len(parts) > 3 or any(n < 1 for n in parts):
            raise ValueError(f"第 {i + 1} 段标题层级无效。")
        parent = parts[:-1]
        if parent and (parent not in headings or current is None or current[:len(parent)] != parent):
            raise ValueError(f"标题 {match.group(1)} 的父标题缺失或不在当前章节。")
        if parts[-1] != last_child.get(parent, 0) + 1:
            raise ValueError(f"标题 {match.group(1)} 的编号不连续。")
        last_child[parent] = parts[-1]
        headings[parts] = (i, match.group(2))
        current = parts
    for number in [(1,), (2,), (3,), (4,), (5,), (6,), (7,),
                   (1, 1), (1, 2), (1, 3), (3, 1), (3, 2),
                   (4, 1), (4, 2), (5, 1), (5, 2)]:
        if number not in headings:
            raise ValueError(f"缺少必要标题 {'.'.join(map(str, number))}。")
    if any(key[0] > 7 for key in headings):
        raise ValueError("只能包含 1 至 7 七个一级章节。")

    for n, category in enumerate(categories, 1):
        key = (2, n)
        expected = f"产品大类{n}：{category}"
        if key not in headings or headings[key][1] != expected:
            raise ValueError(f"产品大类 {n} 必须是：2.{n} {expected}")
    actual_categories = [k for k in headings if len(k) == 2 and k[0] == 2]
    if len(actual_categories) != len(categories):
        raise ValueError("产品大类标题数量与目录不一致。")
    expected_subs = subcategories[:5]
    actual_subs = [k for k in headings if len(k) == 3 and k[:2] == (2, 1)]
    if len(actual_subs) != len(expected_subs):
        raise ValueError("第一大类只应展开实际存在的前五个小类。")
    for n, name in enumerate(expected_subs, 1):
        key = (2, 1, n)
        if key not in headings or headings[key][1] != name:
            raise ValueError(f"小类 2.1.{n} 与目录名称或顺序不一致。")
        start = headings[key][0] + 1
        next_heading = next((i for i in range(start, len(paragraphs))
                             if paragraphs[i]["kind"] == "heading"), len(paragraphs))
        block = paragraphs[start:next_heading]
        labels = [p["text"] for p in block if p["kind"] == "label"]
        if labels != list(PRODUCT_LABELS):
            raise ValueError(f"小类 2.1.{n} 必须依次包含四个标准产品资料要点。")
        for label in PRODUCT_LABELS:
            idx = next(i for i, p in enumerate(block) if p["kind"] == "label" and p["text"] == label)
            if idx + 1 >= len(block) or block[idx + 1]["kind"] != "text":
                raise ValueError(f"小类 2.1.{n} 的 {label} 后缺少正文。")
    if any(k[0] == 2 and len(k) == 3 and k[1] != 1 for k in headings):
        raise ValueError("其他产品大类不能展开小类。")

    expected_blue_positions = {1: BLUE_INTRO}
    for n in range(1, len(categories) + 1):
        start = headings[(2, n)][0]
        end = headings[(2, n + 1)][0] if n < len(categories) else headings[(3,)][0]
        block = paragraphs[start + 1:end]
        if n == 1:
            if not block or block[-1] != {"kind": "blue", "text": BLUE_FIRST}:
                raise ValueError("第一产品大类结尾缺少蓝色补充提示。")
            expected_blue_positions[end - 1] = BLUE_FIRST
        else:
            if block != [{"kind": "blue", "text": BLUE_OTHER}]:
                raise ValueError(f"产品大类 {n} 只能包含大类标题和蓝色补充提示。")
            expected_blue_positions[start + 1] = BLUE_OTHER
    trends = headings[(5,)][0] + 1
    expected_blue_positions[trends] = BLUE_TRENDS
    actual_blue_positions = {i: p["text"] for i, p in enumerate(paragraphs) if p["kind"] == "blue"}
    if actual_blue_positions != expected_blue_positions:
        raise ValueError("蓝色提示的内容、位置或数量有误。")
    return paragraphs


def word_count(paragraphs):
    text = "\n".join(p["text"] for p in paragraphs)
    return len(re.findall(r"[\u3400-\u9fff]", text)) + len(re.findall(r"[A-Za-z]+(?:[-'][A-Za-z]+)*", text))


def _paragraph(p):
    styles = {"title": "Title", "heading": "Heading", "label": "Label"}
    kind, text = p["kind"], p["text"]
    style = styles.get(kind)
    if kind == "heading":
        match = HEADING.fullmatch(text)
        style += str(len(match.group(1).split(".")))
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    rpr = '<w:rPr><w:color w:val="0000FF"/></w:rPr>' if kind == "blue" else ""
    return f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def render(paragraphs):
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:document {ns}><w:body>'
                + "".join(_paragraph(p) for p in paragraphs)
                + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
                  '</w:sectPr></w:body></w:document>')
    styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              f'<w:styles {ns}>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
              '<w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Microsoft YaHei"/>'
              '<w:sz w:val="22"/></w:rPr><w:pPr><w:spacing w:after="120"/></w:pPr></w:style>'
              + "".join(
                  f'<w:style w:type="paragraph" w:styleId="{name}"><w:name w:val="{name}"/>'
                  f'<w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="{size}"/></w:rPr></w:style>'
                  for name, size in (("Title", 34), ("Heading1", 30), ("Heading2", 26),
                                     ("Heading3", 24), ("Label", 22))
              ) + '</w:styles>')
    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                     '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                     '</Types>')
    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                 '</Relationships>')
    doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                '</Relationships>')
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in (("[Content_Types].xml", content_types), ("_rels/.rels", root_rels),
                              ("word/document.xml", document), ("word/styles.xml", styles),
                              ("word/_rels/document.xml.rels", doc_rels)):
            z.writestr(name, content.encode("utf-8"))
    return buffer.getvalue()


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="UTF-8 JSON 稿件")
    parser.add_argument("output", type=Path, help="新建 DOCX 文件")
    args = parser.parse_args(argv)
    try:
        if args.output.suffix.lower() != ".docx":
            raise ValueError("输出文件必须使用 .docx 后缀。")
        with args.input.open(encoding="utf-8-sig") as f:
            data = json.load(f)
        paragraphs = validate(data)
        document = render(paragraphs)
        with args.output.open("xb") as f:
            f.write(document)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"导出失败：{exc}", file=sys.stderr)
        return 1
    count = word_count(paragraphs)
    print(f"已生成：{args.output}")
    print(f"近似字数（中文字符 + 英文单词，含标题和提示）：{count}")
    if not 5000 <= count <= 7000:
        print("警告：篇幅偏离约 6000 字目标；请结合实际产品数量和参考模板人工检查，不要虚构内容凑字。")
    print("请人工复核行业准确性、英文、公司名、蓝色提示及页面排版。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
