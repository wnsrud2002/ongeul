#!/usr/bin/env python3
"""mdmaker 테스트. 표준 unittest만 쓰고, 샘플은 여기서 직접 만든다(사용자 문서 사용 금지)."""
import argparse
import base64
import io
import datetime
import json
import os
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

import mdmaker as M
import pdf
import pdfwrite
import xls

PNG = base64.b64decode(
    b"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==")

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
V = 'xmlns:v="urn:schemas-microsoft-com:vml"'

DOCUMENT = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document {W} {R} {V}><w:body>
 <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>계약서 제목</w:t></w:r></w:p>
 <w:p><w:r><w:t xml:space="preserve">본문 앞 문단 </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>굵게</w:t></w:r>
      <w:r><w:rPr><w:strike/><w:color w:val="FF0000"/></w:rPr><w:t>취소선빨강</w:t></w:r>
      <w:r><w:t>。漢字와 特殊文字 ±§</w:t></w:r>
      <w:r><w:footnoteReference w:id="2"/></w:r></w:p>
 <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>첫째 항목</w:t></w:r></w:p>
 <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr></w:pPr>
      <w:r><w:t>둘째 항목</w:t></w:r></w:p>
 <w:p><w:hyperlink r:id="rId1"><w:r><w:t>링크문구</w:t></w:r></w:hyperlink></w:p>
 <w:p><w:ins w:id="10" w:author="x"><w:r><w:t>삽입된문장</w:t></w:r></w:ins>
      <w:del w:id="11" w:author="x"><w:r><w:delText>삭제된문장</w:delText></w:r></w:del></w:p>
 <w:tbl>
  <w:tr><w:tc><w:p><w:r><w:t>표셀A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>표셀B</w:t></w:r></w:p></w:tc></w:tr>
  <w:tr><w:tc><w:p><w:r><w:t>표셀C</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>표셀D</w:t></w:r></w:p></w:tc></w:tr>
 </w:tbl>
 <w:tbl>
  <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
        <w:p><w:r><w:t>병합머리</w:t></w:r></w:p></w:tc></w:tr>
  <w:tr><w:tc><w:p><w:r><w:t>병합좌</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>병합우</w:t></w:r></w:p></w:tc></w:tr>
 </w:tbl>
 <w:p><w:r><w:t>본문 뒤 문단</w:t></w:r></w:p>
 <w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent>
      <w:p><w:r><w:t>텍스트상자내용</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>
      <v:shape><v:imagedata r:id="rId2"/></v:shape></w:pict></w:r></w:p>
</w:body></w:document>"""

RELS = f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/a" TargetMode="External"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
 <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>"""

STYLES = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:styles {W}><w:style w:styleId="Heading1"><w:name w:val="heading 1"/></w:style></w:styles>"""

NUMBERING = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:numbering {W}>
 <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>
 <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>
 <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
 <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>"""

FOOTNOTES = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:footnotes {W}><w:footnote w:id="2"><w:p><w:r><w:t>각주본문</w:t></w:r></w:p></w:footnote></w:footnotes>"""

HEADER = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:hdr {W}><w:p><w:r><w:t>머리글문구</w:t></w:r></w:p></w:hdr>"""


def make_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        z.writestr("word/document.xml", DOCUMENT)
        z.writestr("word/_rels/document.xml.rels", RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/numbering.xml", NUMBERING)
        z.writestr("word/footnotes.xml", FOOTNOTES)
        z.writestr("word/header1.xml", HEADER)
        z.writestr("word/media/image1.png", PNG)


def make_xlsx(path: Path) -> None:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "매출"
    ws["A1"] = "항목"
    ws["B1"] = "금액"
    ws["A2"] = "영"
    ws["B2"] = 0
    ws["A3"] = "거짓"
    ws["B3"] = False
    ws["A4"] = "빈문자열"          # B4는 아래에서 진짜 빈 문자열 셀로 심는다
    ws["A5"] = "날짜"
    ws["B5"] = datetime.datetime(2026, 9, 20)
    ws["B5"].number_format = "yyyy-mm-dd"
    ws["A6"] = "코드"
    ws["B6"] = "007"
    ws["A7"] = "합계"
    ws["B7"] = "=SUM(B2:B2)"           # 계산값 캐시 없음
    ws["A8"] = "비율"
    ws["B8"] = 0.125
    ws["B8"].number_format = "0.0%"
    ws["A9"] = "줄바꿈|파이프"
    ws["B9"] = "a\nb\tc|d"
    ws.merge_cells("A11:B11")
    ws["A11"] = "병합머리"
    ws.column_dimensions["C"].hidden = True
    hid = wb.create_sheet("숨김시트")
    hid["A1"] = "숨겨진값"
    hid.sheet_state = "hidden"
    wb.save(path)
    _inject_empty_string_cell(path)


def _inject_empty_string_cell(path: Path) -> None:
    """openpyxl은 빈 문자열을 저장하지 못하므로 시트 XML에 직접 넣는다."""
    import re
    z = zipfile.ZipFile(path)
    parts = {n: z.read(n) for n in z.namelist()}
    z.close()
    name = [n for n in parts if n.endswith("worksheets/sheet1.xml")][0]
    xml = re.sub(r'(<row r="4".*?)</row>', r'\1<c r="B4" t="str"><v></v></c></row>',
                 parts[name].decode(), count=1, flags=re.S)
    parts[name] = xml.encode()
    with zipfile.ZipFile(path, "w") as o:
        for n, d in parts.items():
            o.writestr(n, d)


class Tmp(unittest.TestCase):
    def setUp(self):
        self.d = Path(tempfile.mkdtemp(prefix="mdmaker-test-"))
        self.inp = self.d / "in"
        self.out = self.d / "out"
        self.inp.mkdir()

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)

    def run_cli(self, *args):
        return M.main([str(a) for a in args])

    def md_of(self, name, incomplete=False):
        base = self.out / "_incomplete" if incomplete else self.out
        return (base / (name + ".md")).read_text(encoding="utf-8")


class TestText(Tmp):
    def test_cp949_and_exact_preservation(self):
        p = self.inp / "메모 파일.txt"
        body = "가나다 漢字\n\n\n연속 빈 줄과   반복   공백\n# 제목처럼 보이는 줄\n"
        p.write_bytes(body.encode("cp949"))
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)
        md = self.md_of("메모 파일.txt")
        self.assertIn(body, md)
        self.assertIn("cp949", md)

    def test_unknown_encoding_fails(self):
        p = self.inp / "깨진.txt"
        p.write_bytes(b"\xff\xfe\x00abc\x81\x40\xff")
        self.assertEqual(self.run_cli(p, "--out", self.out), 1)
        self.assertIn("failed", self.md_of("깨진.txt", incomplete=True))

    def test_markdown_passthrough(self):
        p = self.inp / "a.md"
        p.write_text("# 원래 제목\n\n- 항목\n", encoding="utf-8")
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)
        self.assertIn("# 원래 제목", self.md_of("a.md"))


class TestCsv(Tmp):
    def test_quotes_delimiters_newlines(self):
        p = self.inp / "표.csv"
        p.write_text('a,"b,콤마","c\n줄바꿈"\n1,,"따옴표""안"\n', encoding="utf-8")
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)
        md = self.md_of("표.csv")
        for want in ("b,콤마", "c<br>줄바꿈", '따옴표"안'):
            self.assertIn(want, md)

    def test_ragged_rows_reported(self):
        p = self.inp / "들쭉.csv"
        p.write_text("a,b,c\nd\n", encoding="utf-8")
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)
        meta = json.loads((self.out / "들쭉.csv.assets" / "meta.json").read_text(encoding="utf-8"))
        self.assertEqual(meta["structure"]["row_widths"], [3, 1])


class TestDocx(Tmp):
    def setUp(self):
        super().setUp()
        self.p = self.inp / "계약서.docx"
        make_docx(self.p)

    def test_full_conversion(self):
        self.assertEqual(self.run_cli(self.p, "--out", self.out), 0)
        md = self.md_of("계약서.docx")
        self.assertIn("# 계약서 제목", md)
        self.assertIn("**굵게**", md)
        self.assertIn("~~", md)
        self.assertIn("color:#FF0000", md)
        self.assertIn("。漢字와 特殊文字 ±§", md)
        self.assertIn("- 첫째 항목", md)
        self.assertIn("  1. 둘째 항목", md)
        self.assertIn("[링크문구](https://example.invalid/a)", md)
        self.assertIn("{+삽입:삽입된문장+}", md)
        self.assertIn("{-삭제:삭제된문장-}", md)
        self.assertIn("각주본문", md)
        self.assertIn("머리글문구", md)
        self.assertIn("텍스트상자내용", md)
        self.assertIn("media/image1.png", md)

    def test_body_table_body_order(self):
        self.run_cli(self.p, "--out", self.out)
        md = self.md_of("계약서.docx")
        self.assertLess(md.index("본문 앞 문단"), md.index("표셀A"))
        self.assertLess(md.index("표셀A"), md.index("본문 뒤 문단"))

    def test_simple_table_is_markdown_merged_is_rowwise(self):
        self.run_cli(self.p, "--out", self.out)
        md = self.md_of("계약서.docx")
        self.assertIn("| 표셀A | 표셀B |", md)
        self.assertIn("원본에 헤더 행 표시가 없어", md)
        self.assertIn("**R1C1** (가로병합 2칸): 병합머리", md)

    def test_asset_bytes_preserved(self):
        self.run_cli(self.p, "--out", self.out)
        a = self.out / "계약서.docx.assets"
        self.assertEqual((a / "media" / "image1.png").read_bytes(), PNG)
        meta = json.loads((a / "meta.json").read_text(encoding="utf-8"))
        self.assertEqual(meta["status"], "success")
        self.assertEqual(meta["assets"][0]["sha256"], M.sha256(PNG))

    def test_missing_text_is_detected_as_partial(self):
        """검사기가 실제로 누락을 잡는지 확인한다(성공만 확인하는 검사는 증명이 아니다)."""
        res = M.Res(self.p, "docx")
        res.markdown = "# 계약서 제목\n"      # 대부분 빠진 결과
        with M.open_zip(self.p, M.Limits()) as zf:
            M.verify_ooxml_text(res, zf, M.Limits(), "docx")
        self.assertEqual(res.status, M.PARTIAL)
        self.assertTrue(any(c["item"] == "텍스트 노드" and not c["ok"] for c in res.checks))

    def test_bad_signature_fails(self):
        bad = self.inp / "가짜.docx"
        bad.write_bytes(b"not a zip at all")
        self.assertEqual(self.run_cli(bad, "--out", self.out), 1)
        self.assertIn("시그니처", self.md_of("가짜.docx", incomplete=True))


class TestXlsx(Tmp):
    def setUp(self):
        super().setUp()
        self.p = self.inp / "매출.xlsx"
        make_xlsx(self.p)

    def test_values_and_structure(self):
        self.assertEqual(self.run_cli(self.p, "--out", self.out, "--xlsx-table", "md"), 0)
        md = self.md_of("매출.xlsx")
        self.assertIn("## 시트: 매출", md)
        self.assertIn("## 시트: 숨김시트 (숨김: hidden)", md)
        self.assertIn("숨겨진값", md)
        self.assertIn("=SUM(B2:B2) → [원본에 계산값 없음]", md)
        self.assertIn("[빈 문자열]", md)
        self.assertIn("FALSE", md)
        self.assertIn("2026-09-20T00:00:00", md)
        self.assertIn("007", md)
        self.assertIn("0.125", md)          # 표시(12.5%)가 아니라 원시 값
        self.assertIn("`A11:B11`", md)
        self.assertIn("숨김 열: C", md)
        self.assertIn("a<br>b\tc\\|d", md)

    def test_number_format_kept_in_meta(self):
        self.run_cli(self.p, "--out", self.out)
        meta = json.loads((self.out / "매출.xlsx.assets" / "meta.json").read_text(encoding="utf-8"))
        fmts = {(c["sheet"], c["coord"]): c for c in meta["structure"]["cells"]}
        self.assertEqual(fmts[("매출", "B8")]["number_format"], "0.0%")
        self.assertEqual(fmts[("매출", "B5")]["number_format"], "yyyy-mm-dd")
        self.assertAlmostEqual(fmts[("매출", "B5")]["serial"], 46285.0)
        self.assertEqual(fmts[("매출", "B7")]["formula"], "=SUM(B2:B2)")

    def test_tsv_mode(self):
        self.assertEqual(self.run_cli(self.p, "--out", self.out, "--xlsx-table", "tsv"), 0)
        md = self.md_of("매출.xlsx")
        self.assertIn("```tsv", md)
        self.assertIn("a\\nb\\tc|d", md)

    def test_cell_limit_reports_partial(self):
        self.assertEqual(self.run_cli(self.p, "--out", self.out, "--max-cells", "3"), 1)
        md = self.md_of("매출.xlsx", incomplete=True)
        self.assertIn("셀 한도", md)

    def test_dropped_cell_is_detected(self):
        res = M.Res(self.p, "xlsx")
        res.markdown = "# 매출.xlsx\n"
        M.verify_xlsx(res, self.p, {}, M.Limits())
        self.assertEqual(res.status, M.PARTIAL)


class TestPipeline(Tmp):
    def test_batch_recursive_and_mixed_status(self):
        (self.inp / "재무").mkdir()
        (self.inp / "재무" / "메모.txt").write_text("안녕", encoding="utf-8")
        make_xlsx(self.inp / "재무" / "매출.xlsx")
        (self.inp / "그림.bmp").write_bytes(b"BM..")
        code = self.run_cli(self.inp, "--out", self.out, "--recursive")
        self.assertEqual(code, 1)                       # 미지원 파일이 하나 있다
        self.assertTrue((self.out / "재무" / "메모.txt.md").exists())
        self.assertTrue((self.out / "재무" / "매출.xlsx.md").exists())
        self.assertFalse((self.out / "그림.bmp.md").exists())

    def test_no_overwrite_then_overwrite(self):
        p = self.inp / "a.txt"
        p.write_text("처음", encoding="utf-8")
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)
        p.write_text("나중", encoding="utf-8")
        self.assertEqual(self.run_cli(p, "--out", self.out), 1)
        self.assertIn("처음", self.md_of("a.txt"))       # 기존 결과가 보호된다
        self.assertEqual(self.run_cli(p, "--out", self.out, "--overwrite"), 0)
        self.assertIn("나중", self.md_of("a.txt"))

    def test_output_inside_input_not_rescanned(self):
        (self.inp / "a.txt").write_text("내용", encoding="utf-8")
        out = self.inp / "out"
        self.assertEqual(self.run_cli(self.inp, "--out", out, "--recursive"), 0)
        self.assertEqual(self.run_cli(self.inp, "--out", out, "--recursive", "--overwrite"), 0)
        self.assertFalse((out / "out").exists())

    def test_url_rejected(self):
        self.assertEqual(self.run_cli("https://example.invalid/a.docx", "--out", self.out), 2)

    def test_wrong_option_for_format_is_error(self):
        make_xlsx(self.inp / "b.xlsx")
        self.assertEqual(self.run_cli(self.inp / "b.xlsx", "--out", self.out,
                                      "--encoding", "utf-8"), 2)

    def test_temp_dirs_cleaned(self):
        p = self.inp / "a.txt"
        p.write_text("x", encoding="utf-8")
        self.run_cli(p, "--out", self.out)
        self.assertEqual([q for q in self.out.iterdir() if q.name.startswith(".mdmaker-")], [])


class TestSafety(Tmp):
    def test_dtd_rejected(self):
        with self.assertRaises(ValueError):
            M.parse_xml(b'<?xml version="1.0"?><!DOCTYPE a [<!ENTITY x "y">]><a/>', M.Limits())

    def test_zip_path_escape_rejected(self):
        p = self.inp / "악성.docx"
        with zipfile.ZipFile(p, "w") as z:
            z.writestr("../../evil.xml", "<a/>")
        with self.assertRaises(ValueError):
            M.open_zip(p, M.Limits())

    def test_zip_escape_file_is_failed_not_crash(self):
        p = self.inp / "악성.docx"
        with zipfile.ZipFile(p, "w") as z:
            z.writestr("../../evil.xml", "<a/>")
        self.assertEqual(self.run_cli(p, "--out", self.out), 1)
        self.assertIn("경로 탈출", self.md_of("악성.docx", incomplete=True))


if __name__ == "__main__":
    unittest.main()


# ------------------------------------------------------------------ HWPX 샘플
HP = 'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"'
HH = 'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"'
HS = 'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"'
HC = 'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"'

CONTAINER = """<?xml version="1.0" encoding="UTF-8"?>
<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container">
 <ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf"
   media-type="application/hwpml-package+xml"/></ocf:rootfiles></ocf:container>"""

# spine 순서를 파일명 순서와 반대로 둔다: 이름 정렬로 읽으면 틀리는 샘플이다.
CONTENT_HPF = """<?xml version="1.0" encoding="UTF-8"?>
<opf:package xmlns:opf="http://www.idpf.org/2007/opf" version="">
 <opf:manifest>
  <opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>
  <opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>
  <opf:item id="section1" href="Contents/section1.xml" media-type="application/xml"/>
  <opf:item id="image1" href="BinData/image1.png" media-type="image/png"/>
 </opf:manifest>
 <opf:spine>
  <opf:itemref idref="section1" linear="yes"/>
  <opf:itemref idref="section0" linear="yes"/>
 </opf:spine></opf:package>"""

HEADER_XML = f"""<?xml version="1.0" encoding="UTF-8"?>
<hh:head {HH}><hh:refList>
 <hh:styles>
  <hh:style id="0" type="PARA" name="바탕글"/>
  <hh:style id="1" type="PARA" name="개요 1"/>
 </hh:styles>
 <hh:paraProperties>
  <hh:paraPr id="0"><hh:heading type="NONE" level="0"/></hh:paraPr>
  <hh:paraPr id="2"><hh:heading type="BULLET" level="0"/></hh:paraPr>
 </hh:paraProperties>
</hh:refList></hh:head>"""

SECTION0 = f"""<?xml version="1.0" encoding="UTF-8"?>
<hs:sec {HS} {HP} {HC}>
 <hp:p paraPrIDRef="0" styleIDRef="1"><hp:run charPrIDRef="0">
   <hp:t>한글 제목</hp:t></hp:run><hp:linesegarray><hp:lineseg/></hp:linesegarray></hp:p>
 <hp:p paraPrIDRef="0" styleIDRef="0"><hp:run><hp:t>앞<hp:tab/>뒤<hp:lineBreak/>다음<hp:fwSpace/>끝</hp:t></hp:run></hp:p>
 <hp:p paraPrIDRef="0" styleIDRef="0"><hp:run><hp:t>파이프|와 역슬래시\\ 그리고 # 우물정</hp:t></hp:run></hp:p>
 <hp:p paraPrIDRef="2" styleIDRef="0"><hp:run><hp:t>글머리 항목</hp:t></hp:run></hp:p>
 <hp:p paraPrIDRef="0" styleIDRef="0"><hp:run>
   <hp:fieldBegin type="HYPERLINK"><hp:parameters>
     <hp:stringParam name="Command">https://example.invalid/a;1</hp:stringParam>
   </hp:parameters></hp:fieldBegin><hp:t>링크글자</hp:t><hp:fieldEnd/></hp:run></hp:p>
 <hp:p paraPrIDRef="0" styleIDRef="0"><hp:run><hp:t>각주달린문장</hp:t>
   <hp:footNote><hp:subList><hp:p><hp:run><hp:t>각주내용</hp:t></hp:run></hp:p></hp:subList></hp:footNote>
   </hp:run></hp:p>
 <hp:p paraPrIDRef="0" styleIDRef="0"><hp:run><hp:equation><hp:script>E=mc^2</hp:script></hp:equation></hp:run></hp:p>
 <hp:p paraPrIDRef="0" styleIDRef="0"><hp:run><hp:pic><hp:imgRect/><hc:img binaryItemIDRef="image1"/></hp:pic></hp:run></hp:p>
 <hp:p paraPrIDRef="0" styleIDRef="0"><hp:run><hp:tbl rowCnt="2" colCnt="2">
  <hp:tr>
   <hp:tc><hp:cellAddr colAddr="0" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/>
    <hp:subList><hp:p><hp:run><hp:t>표셀가</hp:t></hp:run></hp:p></hp:subList></hp:tc>
   <hp:tc><hp:cellAddr colAddr="1" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/>
    <hp:subList><hp:p><hp:run><hp:t>표셀나</hp:t></hp:run></hp:p></hp:subList></hp:tc>
  </hp:tr>
  <hp:tr>
   <hp:tc><hp:cellAddr colAddr="0" rowAddr="1"/><hp:cellSpan colSpan="1" rowSpan="1"/>
    <hp:subList><hp:p><hp:run><hp:t>표셀다</hp:t></hp:run></hp:p></hp:subList></hp:tc>
   <hp:tc><hp:cellAddr colAddr="1" rowAddr="1"/><hp:cellSpan colSpan="1" rowSpan="1"/>
    <hp:subList><hp:p><hp:run><hp:t>표셀라</hp:t></hp:run></hp:p></hp:subList></hp:tc>
  </hp:tr></hp:tbl></hp:run></hp:p>
 <hp:p paraPrIDRef="0" styleIDRef="0"><hp:run><hp:tbl rowCnt="2" colCnt="2">
  <hp:tr><hp:tc><hp:cellAddr colAddr="0" rowAddr="0"/><hp:cellSpan colSpan="2" rowSpan="1"/>
    <hp:subList><hp:p><hp:run><hp:t>병합머리</hp:t></hp:run></hp:p>
     <hp:p><hp:run><hp:t>둘째문단</hp:t></hp:run></hp:p></hp:subList></hp:tc></hp:tr>
  <hp:tr><hp:tc><hp:cellAddr colAddr="0" rowAddr="1"/><hp:cellSpan colSpan="1" rowSpan="1"/>
    <hp:subList><hp:p><hp:run><hp:t>중첩바깥</hp:t>
     <hp:tbl rowCnt="1" colCnt="1"><hp:tr><hp:tc><hp:cellAddr colAddr="0" rowAddr="0"/>
      <hp:subList><hp:p><hp:run><hp:t>중첩안쪽</hp:t></hp:run></hp:p></hp:subList></hp:tc></hp:tr></hp:tbl>
     </hp:run></hp:p></hp:subList></hp:tc>
   <hp:tc><hp:cellAddr colAddr="1" rowAddr="1"/><hp:cellSpan colSpan="1" rowSpan="1"/>
    <hp:subList><hp:p><hp:run><hp:t>오른칸</hp:t></hp:run></hp:p></hp:subList></hp:tc></hp:tr>
 </hp:tbl></hp:run></hp:p>
</hs:sec>"""

SECTION1 = f"""<?xml version="1.0" encoding="UTF-8"?>
<hs:sec {HS} {HP} {HC}>
 <hp:p paraPrIDRef="0" styleIDRef="0"><hp:run><hp:t>둘째구역이먼저온다</hp:t></hp:run></hp:p>
</hs:sec>"""


def make_hwpx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("mimetype", "application/hwp+zip")
        z.writestr("version.xml", '<?xml version="1.0"?><hv:HCFVersion xmlns:hv="x"/>')
        z.writestr("META-INF/container.xml", CONTAINER)
        z.writestr("Contents/content.hpf", CONTENT_HPF)
        z.writestr("Contents/header.xml", HEADER_XML)
        z.writestr("Contents/section0.xml", SECTION0)
        z.writestr("Contents/section1.xml", SECTION1)
        z.writestr("BinData/image1.png", PNG)
        z.writestr("Preview/PrvText.txt", "미리보기전용문구")


class TestHwpx(Tmp):
    def setUp(self):
        super().setUp()
        self.p = self.inp / "보고서.hwpx"
        make_hwpx(self.p)

    def test_full_conversion(self):
        self.assertEqual(self.run_cli(self.p, "--out", self.out), 0)
        md = self.md_of("보고서.hwpx")
        self.assertIn("# 한글 제목", md)
        self.assertIn("앞\t뒤\n다음　끝", md)
        self.assertIn("- 글머리 항목", md)
        self.assertIn("[HYPERLINK: https://example.invalid/a;1]", md)
        self.assertIn("링크글자", md)
        self.assertIn("각주내용", md)
        self.assertIn("`[수식] E=mc^2`", md)
        self.assertIn("media/image1.png", md)

    def test_spine_order_beats_filename_order(self):
        self.run_cli(self.p, "--out", self.out)
        md = self.md_of("보고서.hwpx")
        self.assertLess(md.index("둘째구역이먼저온다"), md.index("한글 제목"))

    def test_table_paragraphs_not_duplicated(self):
        self.run_cli(self.p, "--out", self.out)
        md = self.md_of("보고서.hwpx")
        self.assertEqual(md.count("표셀가"), 1)
        self.assertEqual(md.count("중첩안쪽"), 1)
        self.assertIn("| 표셀가 | 표셀나 |", md)
        self.assertIn("**R1C1** (가로병합 2칸): 병합머리", md)

    def test_preview_text_not_used_as_body(self):
        self.run_cli(self.p, "--out", self.out)
        self.assertNotIn("미리보기전용문구", self.md_of("보고서.hwpx"))

    def test_escapes_round_trip(self):
        self.run_cli(self.p, "--out", self.out)
        md = self.md_of("보고서.hwpx")
        self.assertIn("파이프|와 역슬래시\\ 그리고 # 우물정", M.normalize_md(md))

    def test_asset_bytes_preserved(self):
        self.run_cli(self.p, "--out", self.out)
        a = self.out / "보고서.hwpx.assets"
        self.assertEqual((a / "media" / "image1.png").read_bytes(), PNG)
        meta = json.loads((a / "meta.json").read_text(encoding="utf-8"))
        self.assertEqual(meta["status"], "success")

    def test_missing_text_is_detected(self):
        res = M.Res(self.p, "hwpx")
        res.markdown = "# 한글 제목\n"
        with M.open_zip(self.p, M.Limits()) as zf:
            M.verify_hwpx(res, zf, M.Limits())
        self.assertEqual(res.status, M.PARTIAL)

    def test_missing_spine_falls_back_with_warning(self):
        p2 = self.inp / "spine없음.hwpx"
        z = zipfile.ZipFile(self.p)
        parts = {n: z.read(n) for n in z.namelist()}
        z.close()
        parts["Contents/content.hpf"] = CONTENT_HPF.replace(
            '<opf:itemref idref="section1" linear="yes"/>', "").replace(
            '<opf:itemref idref="section0" linear="yes"/>', "").encode()
        with zipfile.ZipFile(p2, "w") as o:
            for n, d in parts.items():
                o.writestr(n, d)
        self.assertEqual(self.run_cli(p2, "--out", self.out), 1)
        md = self.md_of("spine없음.hwpx", incomplete=True)
        self.assertIn("문서 순서 정보(spine)를 쓰지 못해", md)


# ------------------------------------------------------------------ HWP 5.x 샘플
SS, MS, CUTOFF = 512, 64, 4096
END, FREE, FATSECT = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD


def _cfb_write(streams: dict) -> bytes:
    """테스트용 최소 CFB 라이터. 형제는 right 포인터로만 잇는다(리더는 중위 순회)."""
    import struct
    sectors, fat = [], []

    def alloc(data: bytes, size: int) -> int:
        if not data:
            return END
        first = len(sectors)
        chunks = [data[i:i + size].ljust(size, b"\x00") for i in range(0, len(data), size)]
        for i, ch in enumerate(chunks):
            sectors.append(ch)
            fat.append(END if i == len(chunks) - 1 else len(sectors))
        return first

    # 디렉터리 항목: 루트 + 스토리지 + 스트림
    entries = [{"name": "Root Entry", "type": 5, "child": FREE, "right": FREE,
                "start": END, "size": 0}]
    tree: dict = {}
    for path in streams:
        parts = path.split("/")
        node = tree
        for p in parts[:-1]:
            node = node.setdefault(p, {})
        node[parts[-1]] = path

    mini_data = bytearray()
    minifat: list = []

    def add_mini(data: bytes) -> int:
        first = len(minifat)
        chunks = [data[i:i + MS].ljust(MS, b"\x00") for i in range(0, len(data), MS)]
        for i, ch in enumerate(chunks):
            mini_data.extend(ch)
            minifat.append(END if i == len(chunks) - 1 else len(minifat) + 1)
        return first

    big: list = []

    def build(node, parent_idx) -> int:
        idx_first = FREE
        prev = None
        for name, val in node.items():
            idx = len(entries)
            if isinstance(val, dict):
                entries.append({"name": name, "type": 1, "child": FREE, "right": FREE,
                                "start": END, "size": 0})
                entries[idx]["child"] = build(val, idx)
            else:
                data = streams[val]
                entries.append({"name": name, "type": 2, "child": FREE, "right": FREE,
                                "start": END, "size": len(data)})
                if len(data) < CUTOFF:
                    entries[idx]["start"] = add_mini(data) if data else END
                else:
                    big.append((idx, data))
            if prev is None:
                idx_first = idx
            else:
                entries[prev]["right"] = idx
            prev = idx
        return idx_first

    entries[0]["child"] = build(tree, 0)
    for idx, data in big:
        entries[idx]["start"] = alloc(data, SS)
    entries[0]["start"] = alloc(bytes(mini_data), SS)
    entries[0]["size"] = len(mini_data)
    mf = b"".join(struct.pack("<I", x) for x in minifat)
    mini_start = alloc(mf, SS) if mf else END
    n_mini = (len(mf) + SS - 1) // SS if mf else 0

    dir_raw = bytearray()
    for e in entries:
        nm = e["name"].encode("utf-16-le")
        b = bytearray(128)
        b[:len(nm)] = nm
        b[64:66] = struct.pack("<H", len(nm) + 2)
        b[66] = e["type"]
        b[67] = 1
        b[68:72] = struct.pack("<I", FREE)
        b[72:76] = struct.pack("<I", e["right"])
        b[76:80] = struct.pack("<I", e["child"])
        b[116:120] = struct.pack("<I", e["start"])
        b[120:128] = struct.pack("<Q", e["size"])
        dir_raw += b
    dir_start = alloc(bytes(dir_raw), SS)
    n_dir = (len(dir_raw) + SS - 1) // SS

    n_data = len(sectors)
    n_fat = 1
    while n_fat * (SS // 4) < n_data + n_fat:
        n_fat += 1
    fat_sec_nums = list(range(n_data, n_data + n_fat))
    fat = fat + [FATSECT] * n_fat
    fat += [FREE] * (n_fat * (SS // 4) - len(fat))
    fat_bytes = b"".join(struct.pack("<I", x) for x in fat)
    sectors.extend(fat_bytes[i:i + SS] for i in range(0, len(fat_bytes), SS))

    head = bytearray(SS)
    head[0:8] = M.hwp5.CFB_SIGNATURE
    head[24:26] = struct.pack("<H", 0x3E)
    head[26:28] = struct.pack("<H", 3)
    head[28:30] = struct.pack("<H", 0xFFFE)
    head[30:32] = struct.pack("<H", 9)
    head[32:34] = struct.pack("<H", 6)
    head[40:44] = struct.pack("<I", n_dir)
    head[44:48] = struct.pack("<I", n_fat)
    head[48:52] = struct.pack("<I", dir_start)
    head[56:60] = struct.pack("<I", CUTOFF)
    head[60:64] = struct.pack("<I", mini_start)
    head[64:68] = struct.pack("<I", n_mini)
    head[68:72] = struct.pack("<I", END)
    for i in range(109):
        v = fat_sec_nums[i] if i < len(fat_sec_nums) else FREE
        head[76 + i * 4:80 + i * 4] = struct.pack("<I", v)
    return bytes(head) + b"".join(sectors)


def _rec(tag: int, level: int, payload: bytes = b"") -> bytes:
    import struct
    if len(payload) >= 0xFFF:
        return struct.pack("<II", tag | (level << 10) | (0xFFF << 20), len(payload)) + payload
    return struct.pack("<I", tag | (level << 10) | (len(payload) << 20)) + payload


def _wstr(s: str) -> bytes:
    import struct
    return struct.pack("<H", len(s)) + s.encode("utf-16-le")


def _ctrl_char(code: int) -> bytes:
    import struct
    return struct.pack("<H", code) + b"\x00" * 12 + struct.pack("<H", code)


def _para(text: bytes, style: int = 0, level: int = 0, extra: bytes = b"") -> bytes:
    import struct
    head = struct.pack("<IIH", len(text) // 2, 0, 0) + bytes([style]) + b"\x00" * 5
    return (_rec(M.hwp5.TAG_PARA_HEADER, level, head)
            + _rec(M.hwp5.TAG_PARA_TEXT, level + 1, text) + extra)


def make_hwp(path: Path, compressed: bool = True, props: int = 0, version: int = 0x05000300,
             unknown_ctrl: bool = False) -> None:
    import struct
    T = lambda s: s.encode("utf-16-le")
    END_P = struct.pack("<H", 13)

    docinfo = (_rec(M.hwp5.TAG_STYLE, 0, _wstr("바탕글") + _wstr("Normal") + b"\x00" * 8)
               + _rec(M.hwp5.TAG_STYLE, 0, _wstr("개요 1") + _wstr("Outline 1") + b"\x00" * 8)
               + _rec(M.hwp5.TAG_BIN_DATA, 0,
                      struct.pack("<HH", 1, 1) + _wstr("png")))

    # 표: CTRL_HEADER('tbl ') + TABLE + 셀 4개
    def cell(col, row, cs, rs, text):
        # 실제 한글처럼 문단을 LIST_HEADER의 "자식"이 아니라 같은 수준의 형제로 둔다
        lh = struct.pack("<IIHHHHII", 1, 0, col, row, cs, rs, 1000, 500)
        return _rec(M.hwp5.TAG_LIST_HEADER, 2, lh) + _para(T(text) + END_P, level=2)

    table = (_rec(M.hwp5.TAG_CTRL_HEADER, 1, b" lbt" + b"\x00" * 20)
             + _rec(M.hwp5.TAG_TABLE, 2, struct.pack("<IHHH", 0, 2, 2, 0) + b"\x00" * 8)
             + cell(0, 0, 1, 1, "표셀가") + cell(1, 0, 1, 1, "표셀나")
             + cell(0, 1, 2, 1, "아래병합칸"))
    footnote = (_rec(M.hwp5.TAG_CTRL_HEADER, 1, b"  nf" + b"\x00" * 20)
                + _rec(M.hwp5.TAG_LIST_HEADER, 2, struct.pack("<II", 1, 0))
                + _para(T("각주내용") + END_P, level=2))
    pic_payload = bytearray(91)          # 실제 한글 문서와 같은 길이/위치
    pic_payload[71:73] = struct.pack("<H", 1)
    picture = (_rec(M.hwp5.TAG_CTRL_HEADER, 1, b" osg" + b"\x00" * 20)
               + _rec(M.hwp5.TAG_SHAPE_PICTURE, 2, bytes(pic_payload)))
    equation = (_rec(M.hwp5.TAG_CTRL_HEADER, 1, b" osg" + b"\x00" * 20)
                + _rec(M.hwp5.TAG_EQEDIT, 2, b"\x00" * 4 + _wstr("E=mc^2")))
    textbox = (_rec(M.hwp5.TAG_CTRL_HEADER, 1, b" osg" + b"\x00" * 20)
               + _rec(76, 2, b"\x00" * 16)                      # SHAPE_COMPONENT
               + _rec(M.hwp5.TAG_LIST_HEADER, 3, struct.pack("<II", 1, 0))
               + _para(T("글상자속문장") + END_P, level=3))
    mystery = (_rec(M.hwp5.TAG_CTRL_HEADER, 1, b"?zyx" + b"\x00" * 20)) if unknown_ctrl else b""

    section = (
        _para(T("한글 제목") + END_P, style=1)
        + _para(T("본문 가나다 漢字 그리고 |파이프") + END_P)
        + _para(T("앞") + _ctrl_char(9) + T("뒤") + struct.pack("<H", 10) + T("아랫줄") + END_P)
        + _para(_ctrl_char(11) + END_P, extra=table)
        + _para(T("각주달린문장") + _ctrl_char(17) + END_P, extra=footnote)
        + _para(_ctrl_char(11) + END_P, extra=picture)
        + _para(_ctrl_char(11) + END_P, extra=equation)
        + _para(_ctrl_char(11) + END_P, extra=textbox)
        + (_para(_ctrl_char(11) + END_P, extra=mystery) if unknown_ctrl else b""))

    def pack(data: bytes) -> bytes:
        import zlib
        if not compressed:
            return data
        co = zlib.compressobj(9, zlib.DEFLATED, -15)
        return co.compress(data) + co.flush()

    fh = bytearray(256)
    fh[0:17] = M.hwp5.HWP_SIGNATURE
    fh[32:36] = struct.pack("<I", version)
    fh[36:40] = struct.pack("<I", props | (1 if compressed else 0))
    path.write_bytes(_cfb_write({
        "FileHeader": bytes(fh),
        "DocInfo": pack(docinfo),
        "BodyText/Section0": pack(section),
        "BinData/BIN0001.png": pack(PNG),
        "PrvText": "미리보기전용문구".encode("utf-16-le"),
    }))


class TestHwp(Tmp):
    def setUp(self):
        super().setUp()
        self.p = self.inp / "옛문서.hwp"
        make_hwp(self.p)

    def test_full_conversion(self):
        self.assertEqual(self.run_cli(self.p, "--out", self.out), 0)
        md = self.md_of("옛문서.hwp")
        self.assertIn("# 한글 제목", md)
        self.assertIn("본문 가나다 漢字 그리고 |파이프", M.normalize_md(md))
        self.assertIn("앞\t뒤\n아랫줄", md)
        self.assertIn("- **R1C1**: 표셀가", md)      # 병합 셀이 있어 행별 표현으로 간다
        self.assertIn("- **R1C2**: 표셀나", md)
        self.assertIn("**R2C1** (가로병합 2칸): 아래병합칸", md)
        self.assertIn("각주내용", md)
        self.assertIn("`[수식] E=mc^2`", md)
        self.assertIn("media/BIN0001.png", md)
        self.assertIn("글상자속문장", md)        # 도형 안쪽 글상자 문단

    def test_uncompressed_file(self):
        p2 = self.inp / "무압축.hwp"
        make_hwp(p2, compressed=False)
        self.assertEqual(self.run_cli(p2, "--out", self.out), 0)
        self.assertIn("# 한글 제목", self.md_of("무압축.hwp"))

    def test_preview_not_used_as_body(self):
        self.run_cli(self.p, "--out", self.out)
        self.assertNotIn("미리보기전용문구", self.md_of("옛문서.hwp"))

    def test_image_bytes_preserved(self):
        self.run_cli(self.p, "--out", self.out)
        self.assertEqual((self.out / "옛문서.hwp.assets" / "media" / "BIN0001.png").read_bytes(), PNG)

    def test_protected_document_is_not_bypassed(self):
        for bit, label in ((2, "암호 설정"), (4, "배포용"), (0x10, "DRM")):
            p2 = self.inp / ("보호%d.hwp" % bit)
            make_hwp(p2, props=bit)
            self.assertEqual(self.run_cli(p2, "--out", self.out), 1)
            md = self.md_of("보호%d.hwp" % bit, incomplete=True)
            self.assertIn("unsupported", md)
            self.assertIn(label, md)

    def test_old_version_rejected(self):
        p2 = self.inp / "구버전.hwp"
        make_hwp(p2, version=0x03000000)
        self.assertEqual(self.run_cli(p2, "--out", self.out), 1)
        self.assertIn("HWP 3.0.0.0는 지원 대상이 아니다", self.md_of("구버전.hwp", incomplete=True))

    def test_unknown_control_blocks_success(self):
        p2 = self.inp / "모르는개체.hwp"
        make_hwp(p2, unknown_ctrl=True)
        self.assertEqual(self.run_cli(p2, "--out", self.out), 1)
        md = self.md_of("모르는개체.hwp", incomplete=True)
        self.assertIn("해석하지 못한 컨트롤", md)
        self.assertIn("unverified", md)

    def test_missing_text_is_detected(self):
        import zlib
        res = M.Res(self.p, "hwp")
        res.markdown = "# 한글 제목\n"
        streams = M.hwp5.Cfb(self.p.read_bytes()).streams()
        hdr = M.hwp5.Header(streams["FileHeader"])
        body = M.hwp5.decompress(streams["BodyText/Section0"], hdr.compressed)
        res.meta = {"images": []}
        M.verify_hwp(res, {"BodyText/Section0": body}, streams, hdr)
        self.assertEqual(res.status, M.PARTIAL)

    def test_non_cfb_hwp_fails(self):
        p = self.inp / "가짜.hwp"
        p.write_bytes("PK\x03\x04 hwpx인데 확장자만 hwp".encode())
        self.assertEqual(self.run_cli(p, "--out", self.out), 1)
        self.assertIn("HWP 5.x로 읽지 못했다", self.md_of("가짜.hwp", incomplete=True))


class TestRealWorldQuirks(Tmp):
    """실제 문서 239개를 돌려서 드러난 결함들. 합성 샘플만으로는 잡히지 않았다."""

    def _xlsx_patch(self, name, part, fn):
        src = self.inp / ("원본_%s" % name)
        make_xlsx(src)
        z = zipfile.ZipFile(src)
        parts = {n: z.read(n) for n in z.namelist()}
        z.close()
        parts[part] = fn(parts[part])
        p = self.inp / name
        with zipfile.ZipFile(p, "w") as o:
            for n, d in parts.items():
                o.writestr(n, d)
        src.unlink()
        return p

    def test_alternate_content_in_stylesheet(self):
        """<AlternateContent>로 감싼 xf 때문에 스타일 색인이 밀리는 파일."""
        def patch(d):
            s = d.decode()
            s = s.replace("</cellXfs>", '<AlternateContent><Fallback>'
                          '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
                          '</Fallback></AlternateContent><xf numFmtId="0" fontId="0" fillId="0" '
                          'borderId="0" xfId="0"/></cellXfs>')
            return s.encode()
        p = self._xlsx_patch("대체콘텐츠.xlsx", "xl/styles.xml", patch)
        # 마지막 xf를 참조하는 셀을 만든다
        z = zipfile.ZipFile(p)
        parts = {n: z.read(n) for n in z.namelist()}
        z.close()
        import re as _re
        idx = len(_re.findall(rb"<xf ", parts["xl/styles.xml"])) - 1
        sheet = [n for n in parts if n.endswith("sheet1.xml")][0]
        parts[sheet] = parts[sheet].replace(b'<c r="A1"', b'<c s="%d" r="A1"' % idx, 1)
        with zipfile.ZipFile(p, "w") as o:
            for n, d in parts.items():
                o.writestr(n, d)
        # 복구해서 읽고, 표시 형식은 원본 styles.xml 과 따로 대조해 통과한다
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)
        md = self.md_of("대체콘텐츠.xlsx")
        self.assertIn("항목", md)                                   # 내용은 살아 있다
        self.assertIn("styles.xml을 그대로는 읽지 못해", md)
        self.assertIn('"item": "표시 형식"', md)

    def test_value_hidden_in_merged_range(self):
        """병합 범위의 좌상단이 아닌 칸에 남은 값도 버리지 않는다."""
        def patch(d):
            ins = '<c r="B11" t="inlineStr"><is><t>숨은값</t></is></c><c r="A11"'.encode()
            return d.replace(b'<c r="A11"', ins, 1)
        p = self._xlsx_patch("병합잔여.xlsx", "xl/worksheets/sheet1.xml", patch)
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)
        md = self.md_of("병합잔여.xlsx")
        self.assertIn("숨은값", md)
        self.assertIn("병합 범위 안", md)

    def test_shared_formula_is_verified_not_skipped(self):
        def patch(d):
            return d.replace(b'<c r="B7"', b'<c r="C7"><f t="shared" si="9" ref="C7:C8">'
                                           b'SUM(B2:B2)</f><v>0</v></c>'
                                           b'<c r="D7"><f t="shared" si="9"/><v>0</v></c>'
                                           b'<c r="B7"', 1)
        p = self._xlsx_patch("공유수식.xlsx", "xl/worksheets/sheet1.xml", patch)
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)
        md = self.md_of("공유수식.xlsx")
        self.assertIn('"shared_formula": 1', md)
        self.assertNotIn("공유 수식", md.split("## 변환 상태")[1])   # 경고 없이 통과

    def test_extension_lies_about_format(self):
        p = self.inp / "사실은한글.hwpx"
        make_hwp(p)
        # 확장자가 틀렸다는 사실은 남기되, 내용 검사 결과를 바꾸지는 않는다
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)
        md = self.md_of("사실은한글.hwpx")
        self.assertIn("# 한글 제목", md)                             # 내용 기준으로 변환
        self.assertIn("확장자는 .hwpx지만 내용은 hwp다", md)
        self.assertIn("검사도 그 형식으로 했다", md)

    def test_literal_cr_in_docx_text(self):
        """XML 파서는 리터럴 CR을 LF로 바꾼다. 검사기도 같은 기준이어야 한다."""
        p = self.inp / "캐리지리턴.docx"
        make_docx(p)
        z = zipfile.ZipFile(p)
        parts = {n: z.read(n) for n in z.namelist()}
        z.close()
        parts["word/document.xml"] = parts["word/document.xml"].replace(
            "<w:t>본문 뒤 문단</w:t>".encode(), "<w:t>본문 뒤\r문단</w:t>".encode())
        with zipfile.ZipFile(p, "w") as o:
            for n, d in parts.items():
                o.writestr(n, d)
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)

    def test_source_text_containing_markdown_markers(self):
        """원문에 ** 나 백틱이 있어도 누락으로 오판하지 않는다."""
        p = self.inp / "별표.docx"
        make_docx(p)
        z = zipfile.ZipFile(p)
        parts = {n: z.read(n) for n in z.namelist()}
        z.close()
        parts["word/document.xml"] = parts["word/document.xml"].replace(
            "<w:t>본문 뒤 문단</w:t>".encode(),
            "<w:t>  **굵게 보이는 원문** 과 `백틱` 과 ~물결~</w:t>".encode())
        with zipfile.ZipFile(p, "w") as o:
            for n, d in parts.items():
                o.writestr(n, d)
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)   # 검사 통과가 핵심
        raw = self.md_of("별표.docx")
        self.assertIn("굵게 보이는 원문** 과 `백틱` 과 ~물결~", raw)   # 표식이 그대로 남는다

    def test_multi_paragraph_cell_keeps_text_contiguous(self):
        """행별 표 표현에서 들여쓰기를 끼워 넣어 셀 문장을 끊지 않는다."""
        self.assertEqual(M.cell_block("첫 줄\n둘째 줄"), "첫 줄<br>둘째 줄")
        self.assertIn("\n    | a |", M.cell_block("설명\n| a |"))   # 중첩 표는 들여쓴다

    def test_hwpx_field_paragraphs_kept(self):
        p = self.inp / "필드문단.hwpx"
        make_hwpx(p)
        z = zipfile.ZipFile(p)
        parts = {n: z.read(n) for n in z.namelist()}
        z.close()
        parts["Contents/section1.xml"] = SECTION1.replace(
            "<hp:t>둘째구역이먼저온다</hp:t>",
            '<hp:fieldBegin type="누름틀"><hp:subList><hp:p><hp:run>'
            '<hp:t>누름틀안쪽문장</hp:t></hp:run></hp:p></hp:subList></hp:fieldBegin>'
            "<hp:t>둘째구역이먼저온다</hp:t>").encode()
        with zipfile.ZipFile(p, "w") as o:
            for n, d in parts.items():
                o.writestr(n, d)
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)
        self.assertIn("누름틀안쪽문장", self.md_of("필드문단.hwpx"))


class TestEscapeRoundTrip(Tmp):
    def test_table_cell_is_not_escaped_twice(self):
        """표 칸 문단은 이미 이스케이프돼 있다. 또 이스케이프하면 `\\|` 가 `\\\\|` 가 된다."""
        p = self.inp / "파이프표.docx"
        make_docx(p)
        z = zipfile.ZipFile(p)
        parts = {n: z.read(n) for n in z.namelist()}
        z.close()
        parts["word/document.xml"] = parts["word/document.xml"].replace(
            "<w:t>표셀A</w:t>".encode(), "<w:t>|P(A)| = 2^4, A \\ B</w:t>".encode())
        with zipfile.ZipFile(p, "w") as o:
            for n, d in parts.items():
                o.writestr(n, d)
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)
        md = self.md_of("파이프표.docx")
        self.assertIn("|P(A)| = 2^4, A \\ B", M.normalize_md(md))
        self.assertEqual(M.cell_inline("이미\\|이스케이프됨"), "이미\\|이스케이프됨")


# ------------------------------------------------------------------ PPTX 샘플
A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'

PRESENTATION = f"""<?xml version="1.0"?>
<p:presentation {P} {R}><p:sldIdLst>
 <p:sldId id="257" r:id="rId2"/><p:sldId id="256" r:id="rId1"/>
</p:sldIdLst></p:presentation>"""

PRES_RELS = """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
</Relationships>"""

SLIDE1 = f"""<?xml version="1.0"?>
<p:sld {P} {A} {R}><p:cSld><p:spTree>
 <p:nvGrpSpPr/><p:grpSpPr/>
 <p:sp><p:nvSpPr><p:cNvPr id="2" name="제목" descr="제목 상자"/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="100" y="200"/></a:xfrm></p:spPr>
  <p:txBody><a:p><a:r><a:t>첫째 슬라이드 제목</a:t></a:r></a:p>
   <a:p><a:r><a:t>둘째 줄</a:t></a:r><a:br/><a:r><a:t>셋째 줄</a:t></a:r></a:p></p:txBody></p:sp>
 <p:grpSp><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="9" name="그룹안"/></p:nvSpPr>
   <p:txBody><a:p><a:r><a:t>그룹 안 글자</a:t></a:r></a:p></p:txBody></p:sp></p:grpSp>
 <p:pic><p:nvPicPr><p:cNvPr id="3" name="사진"/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="rId9"/></p:blipFill></p:pic>
 <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="표"/></p:nvGraphicFramePr>
  <a:graphic><a:graphicData><a:tbl>
   <a:tr><a:tc><a:txBody><a:p><a:r><a:t>표가</a:t></a:r></a:p></a:txBody></a:tc>
        <a:tc><a:txBody><a:p><a:r><a:t>표나</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
   <a:tr><a:tc gridSpan="2"><a:txBody><a:p><a:r><a:t>병합칸</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
  </a:tbl></a:graphicData></a:graphic></p:graphicFrame>
</p:spTree></p:cSld></p:sld>"""

SLIDE2 = f"""<?xml version="1.0"?>
<p:sld {P} {A} {R}><p:cSld><p:spTree>
 <p:sp><p:nvSpPr><p:cNvPr id="2" name="본문"/></p:nvSpPr>
  <p:txBody><a:p><a:r><a:t>두번째슬라이드지만먼저온다</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:sld>"""

NOTES1 = f"""<?xml version="1.0"?>
<p:notes {P} {A}><p:cSld><p:spTree>
 <p:sp><p:nvSpPr><p:cNvPr id="2" name="노트"/></p:nvSpPr>
  <p:txBody><a:p><a:r><a:t>발표자노트내용</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:notes>"""

SLIDE1_RELS = """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
 <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
</Relationships>"""


def make_pptx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        z.writestr("ppt/presentation.xml", PRESENTATION)
        z.writestr("ppt/_rels/presentation.xml.rels", PRES_RELS)
        z.writestr("ppt/slides/slide1.xml", SLIDE1)
        z.writestr("ppt/slides/slide2.xml", SLIDE2)
        z.writestr("ppt/slides/_rels/slide1.xml.rels", SLIDE1_RELS)
        z.writestr("ppt/notesSlides/notesSlide1.xml", NOTES1)
        z.writestr("ppt/media/image1.png", PNG)


class TestPptx(Tmp):
    def setUp(self):
        super().setUp()
        self.p = self.inp / "발표.pptx"
        make_pptx(self.p)

    def test_full_conversion(self):
        self.assertEqual(self.run_cli(self.p, "--out", self.out), 0)
        md = self.md_of("발표.pptx")
        self.assertIn("첫째 슬라이드 제목", md)
        self.assertIn("둘째 줄\n셋째 줄", md)
        self.assertIn("그룹 안 글자", md)
        self.assertIn("| 표가 | 표나 |", md) if "| 표가 | 표나 |" in md else \
            self.assertIn("**R1C1**: 표가", md)
        self.assertIn("병합칸", md)
        self.assertIn("media/image1.png", md)
        self.assertIn("발표자노트내용", md)

    def test_slide_order_follows_sldIdLst(self):
        self.run_cli(self.p, "--out", self.out)
        md = self.md_of("발표.pptx")
        self.assertLess(md.index("두번째슬라이드지만먼저온다"), md.index("첫째 슬라이드 제목"))

    def test_shape_positions_recorded(self):
        self.run_cli(self.p, "--out", self.out)
        meta = json.loads((self.out / "발표.pptx.assets" / "meta.json").read_text(encoding="utf-8"))
        shapes = {s["name"]: s for s in meta["structure"]["shapes"]}
        self.assertEqual((shapes["제목"]["x"], shapes["제목"]["y"]), ("100", "200"))

    def test_missing_text_detected(self):
        res = M.Res(self.p, "pptx")
        res.markdown = "# 발표.pptx\n"
        res.meta = {"images": []}
        with M.open_zip(self.p, M.Limits()) as zf:
            M.verify_pptx(res, zf, M.Limits())
        self.assertEqual(res.status, M.PARTIAL)


# ------------------------------------------------------------------ PDF 샘플
def _pdf_build(objs: dict, trailer_extra: str = "") -> bytes:
    out = bytearray(b"%PDF-1.7\n")
    for num in sorted(objs):
        body = objs[num]
        out += ("%d 0 obj\n" % num).encode()
        out += body if isinstance(body, bytes) else body.encode()
        out += b"\nendobj\n"
    out += ("trailer\n<< /Root 1 0 R /Size %d %s >>\n%%%%EOF\n"
            % (max(objs) + 1, trailer_extra)).encode()
    return bytes(out)


def _pdf_stream(d: str, data: bytes, compress=True) -> bytes:
    import zlib
    if compress:
        data = zlib.compress(data)
        d = d + " /Filter /FlateDecode"
    return ("<< %s /Length %d >>\nstream\n" % (d, len(data))).encode() + data + b"\nendstream"


CMAP = b"""/CIDInit /ProcSet findresource begin 12 dict begin begincmap
1 begincodespacerange <0000> <FFFF> endcodespacerange
2 beginbfchar
<0001> <AC00>
<0002> <B098>
endbfchar
endcmap CMapName currentdict /CMap defineresource pop end end"""

CONTENT1 = (b"BT /F1 12 Tf 72 720 Td (Hello PDF) Tj 0 -14 Td (second line) Tj ET\n"
            b"BT /F2 12 Tf 72 680 Td <00010002> Tj ET\n")
CONTENT2 = b"q 100 0 0 100 50 50 cm /Im1 Do Q\n"
FAKE_JPEG = b"\xff\xd8\xff\xe0" + b"mdmaker-test-image" * 4 + b"\xff\xd9"


def make_pdf(path: Path, encrypted: bool = False) -> None:
    objs = {
        1: "<< /Type /Catalog /Pages 2 0 R >>",
        2: "<< /Type /Pages /Kids [3 0 R 6 0 R] /Count 2 >>",
        3: ("<< /Type /Page /Parent 2 0 R /Contents 4 0 R /Annots [10 0 R]"
            " /Resources << /Font << /F1 5 0 R /F2 8 0 R >> >> >>"),
        4: _pdf_stream("", CONTENT1),
        5: "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
        6: ("<< /Type /Page /Parent 2 0 R /Contents 11 0 R"
            " /Resources << /XObject << /Im1 7 0 R >> >> >>"),
        7: _pdf_stream("/Type /XObject /Subtype /Image /Width 8 /Height 8 /ColorSpace /DeviceRGB"
                       " /BitsPerComponent 8 /Filter /DCTDecode", FAKE_JPEG, compress=False),
        8: ("<< /Type /Font /Subtype /Type0 /BaseFont /Batang /Encoding /Identity-H"
            " /ToUnicode 9 0 R /DescendantFonts [12 0 R] >>"),
        9: _pdf_stream("", CMAP),
        10: "<< /Type /Annot /Subtype /Link /Rect [0 0 10 10] /A << /URI (https://example.invalid/x) >> >>",
        11: _pdf_stream("", CONTENT2),
        12: "<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Batang >>",
    }
    extra = "/Encrypt 13 0 R" if encrypted else ""
    if encrypted:
        objs[13] = "<< /Filter /Standard /V 2 /R 3 /O <00> /U <00> /P -1 >>"
    path.write_bytes(_pdf_build(objs, extra))


class TestPdf(Tmp):
    def setUp(self):
        super().setUp()
        self.p = self.inp / "보고서.pdf"
        make_pdf(self.p)

    def test_text_and_korean(self):
        self.assertEqual(self.run_cli(self.p, "--out", self.out), 1)   # PDF는 항상 unverified
        md = self.md_of("보고서.pdf", incomplete=True)
        self.assertIn("Hello PDF", md)
        self.assertIn("second line", md)
        self.assertIn("가나", md)                       # ToUnicode CMap으로 복원
        self.assertIn("PDF는 문단·표·읽기 순서를 파일에 담지 않아", md)

    def test_line_break_from_position(self):
        self.run_cli(self.p, "--out", self.out)
        md = self.md_of("보고서.pdf", incomplete=True)
        self.assertIn("Hello PDF\nsecond line", md)

    def test_image_page_and_asset(self):
        self.run_cli(self.p, "--out", self.out)
        md = self.md_of("보고서.pdf", incomplete=True)
        self.assertIn("2쪽에 글자가 없고 이미지만 있다", md)
        assets = list((self.out / "_incomplete" / "보고서.pdf.assets" / "media").glob("*.jpg"))
        self.assertEqual(len(assets), 1)
        self.assertEqual(assets[0].read_bytes(), FAKE_JPEG)   # 원본 바이트 그대로

    def test_annotation_uri_kept(self):
        self.run_cli(self.p, "--out", self.out)
        self.assertIn("https://example.invalid/x", self.md_of("보고서.pdf", incomplete=True))

    def test_encrypted_pdf_not_bypassed(self):
        p2 = self.inp / "암호.pdf"
        make_pdf(p2, encrypted=True)
        self.assertEqual(self.run_cli(p2, "--out", self.out), 1)
        md = self.md_of("암호.pdf", incomplete=True)
        self.assertIn("unsupported", md)
        self.assertIn("보호 기능을 우회하지 않는다", md)

    def test_ocr_without_engine_is_reported(self):
        if M.ocr_available():
            self.skipTest("이 컴퓨터에는 tesseract가 있다")
        self.assertEqual(self.run_cli(self.p, "--out", self.out, "--ocr", "force"), 1)
        self.assertIn("OCR 엔진(tesseract)이 없어", self.md_of("보고서.pdf", incomplete=True))


class TestImage(Tmp):
    def test_image_preserved_without_interpretation(self):
        p = self.inp / "사진.png"
        p.write_bytes(PNG)
        self.assertEqual(self.run_cli(p, "--out", self.out), 1)
        md = self.md_of("사진.png", incomplete=True)
        self.assertIn("이미지 안의 글자는 해석하지 않았다", md)
        self.assertEqual((self.out / "_incomplete" / "사진.png.assets" / "media" / "사진.png")
                         .read_bytes(), PNG)


class TestLegacyOffice(Tmp):
    def test_without_libreoffice_is_unsupported(self):
        if M.soffice_path():
            self.skipTest("이 컴퓨터에는 LibreOffice가 있다")
        p = self.inp / "옛문서.doc"
        p.write_bytes(b"\xd0\xcf\x11\xe0" + b"\x00" * 100)
        self.assertEqual(self.run_cli(p, "--out", self.out), 1)
        md = self.md_of("옛문서.doc", incomplete=True)
        self.assertIn("unsupported", md)
        self.assertIn("LibreOffice가 필요하다", md)


class TestSplitAndReuse(Tmp):
    def _long_text(self):
        p = self.inp / "긴문서.txt"
        p.write_text("\n\n".join("문단 %d " % i + "가나다라" * 20 for i in range(40)),
                     encoding="utf-8")
        return p

    def test_split_parts_rejoin_exactly(self):
        p = self._long_text()
        self.assertEqual(self.run_cli(p, "--out", self.out, "--split-chars", "1200"), 0)
        parts = sorted(self.out.glob("긴문서.txt.part*.md"))
        self.assertGreater(len(parts), 1)
        bodies = []
        for f in parts:
            t = f.read_text(encoding="utf-8")
            bodies.append(t.split("-->\n", 1)[1])            # 조각 머리말 제거
        full = self.md_of("긴문서.txt").rsplit("\n---\n\n## 변환 상태", 1)[0]
        self.assertEqual(M.rejoin(bodies), full)
        self.assertIn('"item": "분할"', self.md_of("긴문서.txt"))

    def test_split_is_char_based_not_tokens(self):
        p = self._long_text()
        self.run_cli(p, "--out", self.out, "--split-chars", "1200")
        self.assertIn("문자 기준 분할(토큰 수와 다르다)",
                      sorted(self.out.glob("긴문서.txt.part*.md"))[0].read_text(encoding="utf-8"))

    def test_reuse_skips_only_when_everything_matches(self):
        p = self.inp / "a.txt"
        p.write_text("처음", encoding="utf-8")
        self.assertEqual(self.run_cli(p, "--out", self.out, "--reuse"), 0)
        self.assertEqual(self.run_cli(p, "--out", self.out, "--reuse"), 0)   # 두번째는 건너뜀
        cache = json.loads((self.out / ".mdmaker-cache.json").read_text(encoding="utf-8"))
        self.assertEqual(list(cache)[0], "a.txt")
        # 내용이 바뀌면 다시 변환한다(이름·시각이 아니라 해시로 판단)
        p.write_text("바뀜", encoding="utf-8")
        self.assertEqual(self.run_cli(p, "--out", self.out, "--reuse", "--overwrite"), 0)
        self.assertIn("바뀜", self.md_of("a.txt"))

    def test_reuse_refuses_when_result_missing(self):
        p = self.inp / "a.txt"
        p.write_text("처음", encoding="utf-8")
        self.run_cli(p, "--out", self.out, "--reuse")
        (self.out / "a.txt.md").unlink()
        self.assertEqual(self.run_cli(p, "--out", self.out, "--reuse"), 0)
        self.assertTrue((self.out / "a.txt.md").exists())     # 결과가 없으면 다시 만든다


CHART_XML = f"""<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" {A}>
 <c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>월별 매출</a:t></a:r></a:p></c:rich></c:tx></c:title>
  <c:plotArea><c:barChart>
   <c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>2026년</c:v></c:pt></c:strCache></c:strRef></c:tx>
    <c:cat><c:strRef><c:strCache>
      <c:pt idx="0"><c:v>1월</c:v></c:pt><c:pt idx="1"><c:v>2월</c:v></c:pt>
    </c:strCache></c:strRef></c:cat>
    <c:val><c:numRef><c:numCache>
      <c:pt idx="0"><c:v>1234.5</c:v></c:pt><c:pt idx="1"><c:v>0</c:v></c:pt>
    </c:numCache></c:numRef></c:val></c:ser>
  </c:barChart></c:plotArea></c:chart></c:chartSpace>"""

SLIDE2_CHART = f"""<?xml version="1.0"?>
<p:sld {P} {A} {R}><p:cSld><p:spTree>
 <p:sp><p:nvSpPr><p:cNvPr id="2" name="본문"/></p:nvSpPr>
  <p:txBody><a:p><a:r><a:t>두번째슬라이드지만먼저온다</a:t></a:r></a:p></p:txBody></p:sp>
 <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="차트"/></p:nvGraphicFramePr>
  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
   <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId7"/>
  </a:graphicData></a:graphic></p:graphicFrame>
</p:spTree></p:cSld></p:sld>"""

SLIDE2_RELS = """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="/ppt/charts/chart1.xml"/>
</Relationships>"""


class TestPptxRealWorld(Tmp):
    """실제 파워포인트 파일에서 드러난 결함들."""

    def _make(self):
        p = self.inp / "차트발표.pptx"
        make_pptx(p)
        z = zipfile.ZipFile(p)
        parts = {n: z.read(n) for n in z.namelist()}
        z.close()
        parts["ppt/slides/slide2.xml"] = SLIDE2_CHART.encode()
        parts["ppt/slides/_rels/slide2.xml.rels"] = SLIDE2_RELS.encode()   # 절대형 Target
        parts["ppt/charts/chart1.xml"] = CHART_XML.encode()
        parts["ppt/media/image2.png"] = PNG + b"orphan"                    # 아무도 참조 안 함
        with zipfile.ZipFile(p, "w") as o:
            for n, d in parts.items():
                o.writestr(n, d)
        return p

    def test_chart_cached_values_are_kept(self):
        p = self._make()
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)
        md = self.md_of("차트발표.pptx")
        self.assertIn("**차트: 월별 매출**", md)
        self.assertIn("| 1월 | 1234.5 |", md)
        self.assertIn("| 2월 | 0 |", md)              # 0을 빈칸으로 만들지 않는다
        self.assertTrue((self.out / "차트발표.pptx.assets" / "structure" / "chart1.xml").exists())

    def test_absolute_rel_target_resolved(self):
        p = self._make()
        self.run_cli(p, "--out", self.out)
        self.assertNotIn("차트 부분을 읽지 못했다", self.md_of("차트발표.pptx"))

    def test_orphan_media_preserved_without_demoting(self):
        p = self._make()
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)   # 잔여 이미지로 등급을 내리지 않는다
        md = self.md_of("차트발표.pptx")
        self.assertIn("문서 어느 부분도 참조하지 않는 이미지 1개", md)
        self.assertEqual((self.out / "차트발표.pptx.assets" / "media" / "image2.png").read_bytes(),
                         PNG + b"orphan")

    def test_referenced_but_unrendered_media_still_demotes(self):
        p = self._make()
        z = zipfile.ZipFile(p)
        parts = {n: z.read(n) for n in z.namelist()}
        z.close()
        parts["ppt/slides/_rels/slide2.xml.rels"] = SLIDE2_RELS.replace(
            "</Relationships>",
            '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/image" Target="../media/image2.png"/></Relationships>').encode()
        with zipfile.ZipFile(p, "w") as o:
            for n, d in parts.items():
                o.writestr(n, d)
        self.assertEqual(self.run_cli(p, "--out", self.out), 1)
        self.assertIn("본문 위치를 확인하지 못한 자산", self.md_of("차트발표.pptx", incomplete=True))


class TestOffline(Tmp):
    """네트워크를 막은 상태에서 모든 지원 경로가 도는지 본다(가이드 3장·11장).
    외부 링크가 들어 있는 문서도 조회 없이 처리해야 한다."""

    def setUp(self):
        super().setUp()
        import socket
        import urllib.request          # ssl이 socket을 상속하므로 패치 전에 먼저 읽어 둔다
        self.urllib = urllib.request
        self.calls = []

        def deny(*a, **k):
            self.calls.append(a[:1])
            raise OSError("네트워크 사용 금지(오프라인 검증)")

        self._saved = {}
        for mod, name in ((socket, "socket"), (socket, "create_connection"),
                          (socket, "getaddrinfo"), (socket, "gethostbyname")):
            self._saved[name] = getattr(mod, name)
            setattr(mod, name, deny)
        self._saved["urlopen"] = self.urllib.urlopen
        self.urllib.urlopen = deny

    def tearDown(self):
        import socket, urllib.request
        for name, fn in self._saved.items():
            setattr(urllib.request if name == "urlopen" else socket, name, fn)
        super().tearDown()

    def test_every_format_converts_without_network(self):
        make_docx(self.inp / "a.docx")          # 외부 하이퍼링크·이미지 포함
        make_xlsx(self.inp / "b.xlsx")
        make_hwp(self.inp / "c.hwp")
        make_hwpx(self.inp / "d.hwpx")
        make_pptx(self.inp / "e.pptx")
        make_pdf(self.inp / "f.pdf")            # 외부 URI 주석 포함
        (self.inp / "g.txt").write_text("가나다", encoding="utf-8")
        (self.inp / "h.csv").write_text("a,b\n1,2\n", encoding="utf-8")
        (self.inp / "i.png").write_bytes(PNG)
        self.run_cli(self.inp, "--out", self.out, "--recursive")
        made = sorted(p.name for p in list(self.out.rglob("*.md")))
        self.assertEqual(len(made), 9, made)
        self.assertEqual(self.calls, [])        # 네트워크를 한 번도 건드리지 않는다

    def test_external_link_is_kept_but_not_fetched(self):
        make_docx(self.inp / "a.docx")
        self.assertEqual(self.run_cli(self.inp / "a.docx", "--out", self.out), 0)
        self.assertIn("https://example.invalid/a", self.md_of("a.docx"))
        self.assertEqual(self.calls, [])

    def test_url_input_still_rejected_offline(self):
        self.assertEqual(self.run_cli("https://example.invalid/a.docx", "--out", self.out), 2)
        self.assertEqual(self.calls, [])


class TestRunInfo(Tmp):
    def test_processing_info_recorded(self):
        p = self.inp / "a.txt"
        p.write_text("가나다", encoding="utf-8")
        self.run_cli(p, "--out", self.out)
        meta = json.loads((self.out / "a.txt.assets" / "meta.json").read_text(encoding="utf-8"))
        info = meta["info"]
        for k in ("elapsed_sec", "source_bytes", "ocr", "platform", "python",
                  "converter", "check_version", "chars"):
            self.assertIn(k, info)
        self.assertEqual(info["source_bytes"], p.stat().st_size)


def _text_image(text: str, size=(420, 90)) -> bytes:
    """OCR이 읽을 수 있는 글자 이미지를 만든다. Pillow가 없으면 None."""
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return b""
    im = Image.new("L", size, 255)
    d = ImageDraw.Draw(im)
    try:
        from PIL import ImageFont
        font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial.ttf", 40)
    except (ImportError, OSError):
        font = None
    d.text((10, 20), text, fill=0, font=font)
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    return buf.getvalue()


class TestOcr(Tmp):
    def setUp(self):
        super().setUp()
        if not M.ocr_available():
            self.skipTest("tesseract 가 없다")
        self.img = _text_image("MDMAKER OCR 2026")
        if not self.img:
            self.skipTest("Pillow 가 없다")

    def test_ocr_reads_image_and_marks_unverified(self):
        p = self.inp / "글자.png"
        p.write_bytes(self.img)
        self.assertEqual(self.run_cli(p, "--out", self.out, "--ocr", "force", "--ocr-lang", "eng"), 1)
        md = self.md_of("글자.png", incomplete=True)
        self.assertIn("MDMAKER", md.upper())
        self.assertIn("검수 전", md)
        self.assertIn("unverified", md)          # OCR 결과는 성공으로 올리지 않는다

    def test_ocr_off_by_default(self):
        p = self.inp / "글자.png"
        p.write_bytes(self.img)
        self.run_cli(p, "--out", self.out)
        md = self.md_of("글자.png", incomplete=True)
        self.assertNotIn("OCR 결과", md)
        self.assertIn("OCR 미수행", md)

    def test_original_image_kept_alongside_ocr(self):
        p = self.inp / "글자.png"
        p.write_bytes(self.img)
        self.run_cli(p, "--out", self.out, "--ocr", "force", "--ocr-lang", "eng")
        kept = self.out / "_incomplete" / "글자.png.assets" / "media" / "글자.png"
        self.assertEqual(kept.read_bytes(), self.img)


class TestPdfImageRebuild(Tmp):
    def test_raw_bitmap_becomes_viewable_png(self):
        """PDF 이미지는 파일이 아니라 원시 표본이다. 원본 바이트는 두고 PNG를 따로 만든다."""
        import zlib
        raw = bytes([(x * 7 + y * 3) % 256 for y in range(8) for x in range(8) for _ in range(3)])
        objs = {
            1: "<< /Type /Catalog /Pages 2 0 R >>",
            2: "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            3: ("<< /Type /Page /Parent 2 0 R /Contents 4 0 R"
                " /Resources << /XObject << /Im1 5 0 R >> >> >>"),
            4: _pdf_stream("", b"q 100 0 0 100 0 0 cm /Im1 Do Q\n"),
            5: _pdf_stream("/Type /XObject /Subtype /Image /Width 8 /Height 8"
                           " /ColorSpace /DeviceRGB /BitsPerComponent 8", raw),
        }
        p = self.inp / "비트맵.pdf"
        p.write_bytes(_pdf_build(objs))
        self.assertEqual(self.run_cli(p, "--out", self.out), 1)
        adir = self.out / "_incomplete" / "비트맵.pdf.assets" / "media"
        png = list(adir.glob("*.png"))
        self.assertEqual(len(png), 1)
        self.assertEqual(png[0].read_bytes()[:8], PNG[:8])      # 진짜 PNG 서명
        self.assertTrue(list(adir.glob("*.raw")))               # 원본 표본도 남는다
        meta = json.loads((adir.parent / "meta.json").read_text(encoding="utf-8"))
        self.assertEqual(len(meta["structure"]["rebuilt_images"]), 1)


class TestConcurrency(Tmp):
    def test_two_runs_do_not_corrupt_output(self):
        """덮어쓰기 금지 상태에서 동시 실행이 겹쳐도 결과가 깨지지 않는다(가이드 6장)."""
        import threading
        p = self.inp / "a.txt"
        p.write_text("동시 실행 시험 " * 200, encoding="utf-8")
        codes = []
        barrier = threading.Barrier(4)

        def run():
            barrier.wait()
            codes.append(M.main([str(p), "--out", str(self.out)]))

        threads = [threading.Thread(target=run) for _ in range(4)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()
        md = self.md_of("a.txt")
        self.assertIn("동시 실행 시험", md)
        self.assertIn("상태: **success", md)            # 잘린 결과가 남지 않는다
        self.assertEqual(sorted(codes), [0, 1, 1, 1])   # 하나만 쓰고 나머지는 충돌 보고
        leftovers = [q for q in self.out.iterdir() if q.name.startswith(".mdmaker-")]
        self.assertEqual(leftovers, [])


class TestPdfUnmappedGlyphs(Tmp):
    def test_unmappable_codes_are_counted_and_recorded(self):
        """대응표가 없는 글자는 잃었다고 보고하고, 어떤 코드였는지 남긴다."""
        objs = {
            1: "<< /Type /Catalog /Pages 2 0 R >>",
            2: "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            3: ("<< /Type /Page /Parent 2 0 R /Contents 4 0 R"
                " /Resources << /Font << /F1 5 0 R >> >> >>"),
            4: _pdf_stream("", b"BT /F1 12 Tf 72 720 Td <00110022> Tj ET\n"),
            # ToUnicode도 내장 폰트도 없는 Identity-H: 코드를 글자로 바꿀 방법이 없다
            5: ("<< /Type /Font /Subtype /Type0 /BaseFont /NoMap /Encoding /Identity-H"
                " /DescendantFonts [6 0 R] >>"),
            6: "<< /Type /Font /Subtype /CIDFontType2 /BaseFont /NoMap >>",
        }
        p = self.inp / "대응표없음.pdf"
        p.write_bytes(_pdf_build(objs))
        self.assertEqual(self.run_cli(p, "--out", self.out), 1)
        md = self.md_of("대응표없음.pdf", incomplete=True)
        self.assertIn("유니코드로 바꾸지 못했다", md)
        self.assertIn("partial", md)
        meta = json.loads((self.out / "_incomplete" / "대응표없음.pdf.assets" / "meta.json")
                          .read_text(encoding="utf-8"))
        codes = meta["structure"]["pages"][0]["unmapped_codes"]
        self.assertEqual(sorted(codes), ["0x0011×1", "0x0022×1"])


class TestGui(Tmp):
    """창을 숨긴 채 실제 Tk 위젯으로 돌린다. 화면이 없으면 건너뛴다."""

    def setUp(self):
        super().setUp()
        try:
            import tkinter as tk
            import gui
        except ImportError as exc:
            self.skipTest("tkinter 없음: %s" % exc)
        self.tk, self.gui = tk, gui
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest("화면 없음: %s" % exc)
        self.root.withdraw()
        self.app = gui.App(self.root)

    def tearDown(self):
        try:
            self.root.destroy()
        except Exception:
            pass
        super().tearDown()

    def _drain_now(self):
        import queue as q
        items = []
        while True:
            try:
                items.append(self.app.q.get_nowait())
            except q.Empty:
                break
        for it in items:
            if not (it.get("done") or it.get("stopped") or it.get("error")):
                self.app._add(it)
        self.app._summarize(final=True)
        return items

    def test_converts_through_same_path_as_cli(self):
        (self.inp / "a.txt").write_text("가나다", encoding="utf-8")
        (self.inp / "b.xyz").write_bytes(b"??")           # 미지원 형식
        make_docx(self.inp / "c.docx")
        opts = argparse.Namespace(out=str(self.out), recursive=True, overwrite=False,
                                  reuse=False, encoding=None, xlsx_table="auto",
                                  max_cells=2_000_000, ocr="off", ocr_lang="kor+eng",
                                  split_chars=0)
        files = M.collect(self.inp, self.out, True)
        self.app._work(files, self.inp, self.out, opts, single=False)
        self._drain_now()
        rows = [self.app.tree.item(i)["values"] for i in self.app.tree.get_children()]
        self.assertEqual(len(rows), 3)
        got = {r[2]: r[0] for r in rows}
        self.assertEqual(got["a.txt"], M.SUCCESS)
        self.assertEqual(got["c.docx"], M.SUCCESS)
        self.assertEqual(got["b.xyz"], M.UNSUPPORTED)
        self.assertTrue((self.out / "a.txt.md").exists())

    def test_summary_says_incomplete_when_not_all_verified(self):
        (self.inp / "b.xyz").write_bytes(b"??")
        opts = argparse.Namespace(out=str(self.out), recursive=True, overwrite=False,
                                  reuse=False, encoding=None, xlsx_table="auto",
                                  max_cells=2_000_000, ocr="off", ocr_lang="kor+eng",
                                  split_chars=0)
        self.app._work(M.collect(self.inp, self.out, True), self.inp, self.out, opts, False)
        self._drain_now()
        self.assertIn("success가 아닌 결과는 완료가 아니다", self.app.summary.get())

    def test_status_labels_cover_every_status(self):
        for st in (M.SUCCESS, M.PARTIAL, M.UNVERIFIED, M.FAILED, M.UNSUPPORTED,
                   M.SKIPPED, "conflict"):
            self.assertIn(st, self.gui.STATUS_TEXT)

    def test_worker_error_is_surfaced_not_swallowed(self):
        def boom(*a, **k):
            raise RuntimeError("일부러 낸 오류")
        saved, M.run_batch = M.run_batch, boom
        try:
            self.app._work([], self.inp, self.out, argparse.Namespace(max_cells=1), False)
        finally:
            M.run_batch = saved
        items = self._drain_now()
        self.assertTrue(any("일부러 낸 오류" in (i.get("error") or "") for i in items))


class TestAppBundle(Tmp):
    """앱 묶음이 실제로 서는지 본다. macOS가 아니면 건너뛴다."""

    def setUp(self):
        super().setUp()
        if sys.platform != "darwin":
            self.skipTest("macOS 전용 묶음")
        try:
            import build_app
        except ImportError as exc:
            self.skipTest(str(exc))
        self.build_app = build_app

    def test_bundle_structure_and_selftest(self):
        import plistlib
        import subprocess
        app = self.build_app.build(Path(__file__).resolve().parent, self.d)
        info = plistlib.loads((app / "Contents" / "Info.plist").read_bytes())
        self.assertEqual(info["CFBundleName"], "온글")
        self.assertEqual(info["CFBundleExecutable"], "ongeul")
        self.assertTrue(info["NSHighResolutionCapable"])
        packed = {f.name for f in (app / "Contents" / "Resources" / "app").glob("*.py")}
        here = {f.name for f in Path(__file__).resolve().parent.glob("*.py")
                if not f.name.startswith("test_") and f.name != "build_app.py"}
        self.assertEqual(packed, here)      # 새 모듈을 빠뜨리면 여기서 걸린다
        launcher = app / "Contents" / "MacOS" / "ongeul"
        self.assertTrue(os.access(launcher, os.X_OK))
        # 최소 환경에서도 파이썬을 찾아 창을 세울 수 있어야 한다
        r = subprocess.run([str(launcher), "--selftest"], capture_output=True, shell=False,
                           env={"PATH": "/usr/bin:/bin", "HOME": os.environ.get("HOME", "")},
                           timeout=120)
        self.assertEqual(r.returncode, 0, r.stderr.decode("utf-8", "replace")[:300])
        self.assertIn("정상", r.stdout.decode("utf-8", "replace"))

    def test_icon_is_drawn_at_small_size(self):
        try:
            from PIL import Image  # noqa: F401
        except ImportError:
            self.skipTest("Pillow 없음")
        im = self.build_app.draw_icon(32)
        self.assertEqual(im.size, (32, 32))
        colors = {p[:3] for p in im.getdata() if p[3] > 200}
        self.assertGreater(len(colors), 4)       # 단색 사각형이 아니라 그림이 들어있다


class TestFailureDetection(Tmp):
    """일부러 망가뜨린 결과에서 success와 종료 코드 0이 나오지 않는지 본다(가이드 11장).
    검사기가 통과시키는 것만 보는 시험은 보존 증명이 아니다."""

    def _docx_with_repeats(self, times=3):
        p = self.inp / "반복.docx"
        make_docx(p)
        z = zipfile.ZipFile(p)
        parts = {n: z.read(n) for n in z.namelist()}
        z.close()
        rep = "<w:p><w:r><w:t>반복되는 문장</w:t></w:r></w:p>" * times
        parts["word/document.xml"] = parts["word/document.xml"].replace(
            "<w:p><w:r><w:t>본문 뒤 문단</w:t></w:r></w:p>".encode(), rep.encode())
        with zipfile.ZipFile(p, "w") as o:
            for n, d in parts.items():
                o.writestr(n, d)
        return p

    def _convert(self, p):
        opts = argparse.Namespace(encoding=None, xlsx_table="auto", max_cells=10 ** 7,
                                  ocr="off", ocr_lang="kor+eng")
        return M.convert(p, opts, M.Limits())

    def test_repeated_text_removal_is_caught(self):
        p = self._docx_with_repeats(3)
        good = self._convert(p)
        self.assertEqual(good.markdown.count("반복되는 문장"), 3)
        lossy = M.Res(p, "docx")
        lossy.meta = {"images": []}
        lossy.markdown = good.markdown.replace(
            "반복되는 문장\n\n반복되는 문장\n\n반복되는 문장", "반복되는 문장")
        with M.open_zip(p, M.Limits()) as zf:
            M.verify_ooxml_text(lossy, zf, M.Limits(), "docx")
        self.assertEqual(lossy.status, M.PARTIAL)
        self.assertIn("확인되지 않았다", " ".join(lossy.warnings))

    def test_original_copy_is_not_success(self):
        """원본을 그대로 복사해 놓은 것은 변환이 아니다."""
        p = self.inp / "그대로.docx"
        make_docx(p)
        fake = M.Res(p, "docx")
        fake.meta = {"images": []}
        fake.markdown = p.read_bytes().decode("latin-1")   # 원본 바이트만 붙인 결과
        with M.open_zip(p, M.Limits()) as zf:
            M.verify_ooxml_text(fake, zf, M.Limits(), "docx")
        self.assertEqual(fake.status, M.PARTIAL)

    def test_asset_bytes_changed_is_caught(self):
        p = self.inp / "자산.docx"
        make_docx(p)
        res = self._convert(p)
        self.assertEqual(res.status, M.SUCCESS)
        broken = M.Res(p, "docx")
        broken.markdown = res.markdown
        broken.meta = res.meta
        broken.assets = dict(res.assets)
        name = next(k for k in broken.assets if k.startswith("media/"))
        broken.assets[name] = b"\x00" * 10          # 바이트가 바뀌었다
        with M.open_zip(p, M.Limits()) as zf:
            M.verify_ooxml_text(broken, zf, M.Limits(), "docx")
        self.assertEqual(broken.status, M.PARTIAL)
        self.assertIn("자산 해시가 일치하지 않는다", " ".join(broken.warnings))

    def test_equation_loss_is_caught(self):
        p = self.inp / "수식.hwpx"
        make_hwpx(p)
        res = self._convert(p)
        self.assertIn("E=mc^2", res.markdown)
        lossy = M.Res(p, "hwpx")
        lossy.meta = {"images": []}
        lossy.markdown = res.markdown.replace("`[수식] E=mc^2`", "")
        with M.open_zip(p, M.Limits()) as zf:
            M.verify_hwpx(lossy, zf, M.Limits())
        self.assertEqual(lossy.status, M.PARTIAL)
        eq = [c for c in lossy.checks if c["item"] == "수식"][0]
        self.assertFalse(eq["ok"])

    def test_split_piece_lost_or_reordered_is_caught(self):
        p = self.inp / "긴글.txt"
        p.write_text("\n\n".join("문단 %d %s" % (i, "가나다" * 30) for i in range(30)),
                     encoding="utf-8")
        orig = M.split_markdown          # 원본을 잡아두고 흉내 낸다(재귀 방지)
        for name, bad in (("조각 누락", lambda parts: parts[:-1]),
                          ("순서 바뀜", lambda parts: list(reversed(parts)))):
            with self.subTest(name):
                saved = M.split_markdown
                M.split_markdown = lambda body, n, _b=bad: _b(orig(body, n))
                try:
                    code = self.run_cli(p, "--out", self.out / name, "--split-chars", "900")
                finally:
                    M.split_markdown = saved
                self.assertEqual(code, 1, name)
                md = (self.out / name / "_incomplete" / "긴글.txt.md").read_text(encoding="utf-8")
                self.assertIn("다시 합쳤을 때 원본과 달라", md)

    def test_cell_value_change_is_caught(self):
        p = self.inp / "값바뀜.xlsx"
        make_xlsx(p)
        res = self._convert(p)
        self.assertEqual(res.status, M.SUCCESS)
        tampered = M.Res(p, "xlsx")
        tampered.markdown = res.markdown.replace("항목", "바뀐값")
        tampered.meta = {"images": []}
        cellmap = {("매출", "A1"): {"raw": "바뀐값", "formula": None, "rendered": "바뀐값",
                                   "nf": None, "type": "s"}}
        M.verify_xlsx(tampered, p, cellmap, M.Limits())
        self.assertEqual(tampered.status, M.PARTIAL)
        self.assertIn("불일치", " ".join(tampered.warnings))

    def test_ocr_and_pdf_never_reach_success(self):
        make_pdf(self.inp / "a.pdf")
        (self.inp / "b.png").write_bytes(PNG)
        for name in ("a.pdf", "b.png"):
            res = self._convert(self.inp / name)
            self.assertNotEqual(res.status, M.SUCCESS, name)

    def test_whole_run_exit_code_is_not_zero(self):
        """한 파일이라도 미달이면 일괄 실행 전체가 0을 내지 않는다."""
        (self.inp / "ok.txt").write_text("정상", encoding="utf-8")
        make_pdf(self.inp / "bad.pdf")
        self.assertEqual(self.run_cli(self.inp, "--out", self.out, "--recursive"), 1)
        self.assertTrue((self.out / "ok.txt.md").exists())
        self.assertTrue((self.out / "_incomplete" / "bad.pdf.md").exists())


class TestXlsxVerification(Tmp):
    def test_number_format_mismatch_is_caught(self):
        """표시 형식이 원본과 다르면 잡아낸다(복구본을 무조건 믿지 않는다)."""
        p = self.inp / "서식.xlsx"
        make_xlsx(p)
        res = M.convert(p, argparse.Namespace(encoding=None, xlsx_table="auto",
                                              max_cells=10 ** 7, ocr="off", ocr_lang="k"),
                        M.Limits())
        self.assertEqual(res.status, M.SUCCESS)
        fmt = [c for c in res.checks if c["item"] == "표시 형식"][0]
        self.assertGreater(fmt["count"], 0)      # 실제로 대조한 셀이 있다
        self.assertTrue(fmt["ok"])
        # 형식을 몰래 바꾼 셈 치면 불일치로 잡혀야 한다
        tampered = M.Res(p, "xlsx")
        tampered.markdown = res.markdown
        tampered.meta = {"images": []}
        cellmap = {}
        import openpyxl
        wb = openpyxl.load_workbook(p)
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if c.value is not None:
                        cellmap[(ws.title, c.coordinate)] = {
                            "raw": c.value, "formula": None,
                            "rendered": M._xl_value(c.value),
                            "nf": "엉뚱한서식", "type": c.data_type}
        M.verify_xlsx(tampered, p, cellmap, M.Limits())
        self.assertEqual(tampered.status, M.PARTIAL)
        self.assertIn("number_format", " ".join(tampered.warnings))

    def test_image_anchor_cell_is_recorded(self):
        """그림이 어느 칸에 붙었는지 원본에서 읽어 본문에 적는다."""
        import openpyxl
        from openpyxl.drawing.image import Image as XLImage
        try:
            from PIL import Image  # noqa: F401
        except ImportError:
            self.skipTest("Pillow 없음")
        p = self.inp / "그림.xlsx"
        make_xlsx(p)
        wb = openpyxl.load_workbook(p)
        img_path = self.inp / "점.png"
        img_path.write_bytes(PNG)
        wb["매출"].add_image(XLImage(str(img_path)), "D4")
        wb.save(p)
        self.assertEqual(self.run_cli(p, "--out", self.out), 0)
        md = self.md_of("그림.xlsx")
        self.assertIn("시트 `매출` 의 `D4` 칸", md)
        meta = json.loads((self.out / "그림.xlsx.assets" / "meta.json").read_text(encoding="utf-8"))
        self.assertEqual(meta["structure"]["image_anchors"]["매출"][0]["cell"], "D4")


class TestLegacyXls(Tmp):
    """구형 .xls: 중간 변환기를 믿지 않고 원본 BIFF 를 직접 읽어 대조한다."""

    def setUp(self):
        super().setUp()
        if not M.soffice_path():
            self.skipTest("LibreOffice 없음")
        src = self.inp / "원본.xlsx"
        make_xlsx(src)
        import subprocess
        tmp = self.d / "conv"
        tmp.mkdir()
        r = subprocess.run([M.soffice_path(), "--headless", "--norestore",
                            "-env:UserInstallation=file://%s/profile" % tmp,
                            "--convert-to", "xls", "--outdir", str(tmp), str(src)],
                           capture_output=True, shell=False, timeout=180)
        made = list(tmp.glob("*.xls"))
        if r.returncode != 0 or not made:
            self.skipTest("LibreOffice 가 .xls 를 만들지 못했다")
        self.p = self.inp / "옛엑셀.xls"
        self.p.write_bytes(made[0].read_bytes())
        src.unlink()

    def test_biff_reader_sees_original_strings(self):
        raw = self.p.read_bytes()
        strings = xls.sst_strings(raw)
        self.assertIn("항목", strings)
        self.assertIn("줄바꿈|파이프", strings)      # 특수문자도 그대로
        self.assertIn("매출", xls.sheet_names(raw))

    def test_converts_and_is_checked_against_the_original(self):
        self.assertEqual(self.run_cli(self.p, "--out", self.out), 0)
        md = self.md_of("옛엑셀.xls")
        self.assertIn("항목", md)
        self.assertIn("LibreOffice로 xlsx를 거쳐 변환했다", md)
        self.assertIn("원본 문자열(.xls 직접 읽기)", md)
        self.assertIn("원본 .xls 를 직접 읽어 글자", md)

    def test_middle_conversion_loss_is_caught(self):
        opts = argparse.Namespace(encoding=None, xlsx_table="auto", max_cells=10 ** 7,
                                  ocr="off", ocr_lang="kor+eng")
        good = M.convert(self.p, opts, M.Limits())
        self.assertEqual(good.status, M.SUCCESS)
        lossy = M.Res(self.p, "xls")
        lossy.markdown = good.markdown.replace("항목", "")    # 중간 변환이 글자를 흘린 셈
        M.verify_legacy_xls(lossy, self.p)
        self.assertEqual(lossy.status, M.PARTIAL)
        self.assertIn("중간 변환에서 원본 글자", " ".join(lossy.warnings))


MD_SAMPLE = """# 온글 PDF 시험

본문 문단이다. 한글과 English와 漢字를 같이 쓴다.

- 첫째 항목
- 둘째 항목

| 열1 | 열2 |
| --- | --- |
| 표셀가 | 표셀나 |

> 인용문이다.

```text
코드 블록 안의 글자
```

![그림 설명](media/점.png)
"""


KOREAN_PROBE = "온글 한글 가나다 체육대회 漢字"


class TestPdfWrite(Tmp):
    def setUp(self):
        super().setUp()
        self.font, why = pdfwrite.find_font(sample=KOREAN_PROBE)
        if self.font is None:
            self.skipTest("한글을 그릴 글꼴이 없다: %s" % why)

    def test_subset_is_small_and_keeps_used_glyphs(self):
        text = "온글 가나다 ABC"
        gids = {self.font.gid(c) for c in text}
        sub = self.font.subset(gids)
        self.assertLess(len(sub), len(self.font.d) / 4)      # 통째로 넣지 않는다
        self.assertEqual(sub[:4], b"\x00\x01\x00\x00")       # 제대로 된 TrueType
        self.assertTrue(all(g > 0 for g in gids))

    def test_refuses_font_that_forbids_embedding(self):
        saved = self.font.fs_type
        try:
            self.font.fs_type = 0x0002
            self.assertFalse(self.font.embeddable)
        finally:
            self.font.fs_type = saved
        self.assertTrue(self.font.embeddable)

    def test_render_and_read_back(self):
        """만든 PDF를 우리 PDF 리더로 다시 읽어 글자가 들어갔는지 본다."""
        (self.inp / "media").mkdir()
        (self.inp / "media" / "점.png").write_bytes(PNG)
        data = pdfwrite.render_markdown(MD_SAMPLE, self.font, "시험", base_dir=self.inp)
        out = self.d / "a.pdf"
        out.write_bytes(data)
        doc = pdf.Pdf(out.read_bytes())
        got = "".join("".join(pdf.extract_page(doc, p).parts) for p in doc.pages())
        for probe in ("온글 PDF 시험", "漢字", "첫째 항목", "표셀가", "인용문이다",
                      "코드 블록 안의 글자", "그림 설명"):
            self.assertIn(probe, got, probe)

    def test_decomposed_hangul_is_composed_when_drawn(self):
        import unicodedata
        nfd = unicodedata.normalize("NFD", "체육대회")
        self.assertNotEqual(nfd, "체육대회")
        data = pdfwrite.render_markdown("# " + nfd, self.font, "x")
        doc = pdf.Pdf(data)
        got = "".join("".join(pdf.extract_page(doc, p).parts) for p in doc.pages())
        self.assertIn("체육대회", got)          # 자모로 흩어지지 않는다

    def test_long_document_breaks_into_pages(self):
        md = "\n\n".join("문단 %d %s" % (i, "가나다라마바사" * 8) for i in range(120))
        doc = pdf.Pdf(pdfwrite.render_markdown(md, self.font, "x"))
        self.assertGreater(len(doc.pages()), 2)


class TestPdfExport(Tmp):
    def setUp(self):
        super().setUp()
        if pdfwrite.find_font(sample=KOREAN_PROBE)[0] is None:
            self.skipTest("한글을 그릴 글꼴이 없다")

    def test_result_pdf_is_made_and_checked(self):
        make_docx(self.inp / "계약서.docx")
        self.assertEqual(self.run_cli(self.inp / "계약서.docx", "--out", self.out,
                                      "--pdf", "result"), 0)
        made = self.out / "계약서.docx.pdf"
        self.assertTrue(made.exists())
        meta = json.loads((self.out / "계약서.docx.assets" / "meta.json").read_text(encoding="utf-8"))
        info = meta["info"]["pdf"]["result"]
        self.assertEqual(info["engine"], "builtin")
        self.assertGreaterEqual(info["pages"], 1)
        self.assertEqual(info["missing_lines"], 0)       # 본문이 다 들어갔다
        doc = pdf.Pdf(made.read_bytes())
        got = "".join("".join(pdf.extract_page(doc, p).parts) for p in doc.pages())
        self.assertIn("계약서 제목", got)

    def test_pdf_off_by_default(self):
        make_docx(self.inp / "a.docx")
        self.run_cli(self.inp / "a.docx", "--out", self.out)
        self.assertEqual(list(self.out.glob("*.pdf")), [])

    def test_source_pdf_needs_libreoffice(self):
        make_docx(self.inp / "a.docx")
        if M.soffice_path():
            self.skipTest("LibreOffice 가 있다")
        self.assertEqual(self.run_cli(self.inp / "a.docx", "--out", self.out,
                                      "--pdf", "source"), 2)

    def test_source_pdf_with_libreoffice(self):
        """원본 PDF는 LibreOffice 몫이다. 제대로 나왔거나, 아니면 이유를 남기고 버려야 한다.
        (LibreOffice 가 한글 글꼴을 못 찾는 환경이 실제로 있다.)"""
        if not M.soffice_path():
            self.skipTest("LibreOffice 없음")
        make_xlsx(self.inp / "매출.xlsx")
        self.assertEqual(self.run_cli(self.inp / "매출.xlsx", "--out", self.out,
                                      "--pdf", "both"), 0)
        self.assertTrue((self.out / "매출.xlsx.pdf").exists())     # 결과 PDF는 우리 몫
        md = self.md_of("매출.xlsx")
        made = (self.out / "매출.xlsx.source.pdf").exists()
        self.assertTrue(made or "쓸 수 없는 PDF를 남기지 않는다" in md,
                        "원본 PDF가 없으면 이유가 적혀 있어야 한다")

    def test_garbage_pdf_is_discarded_not_kept(self):
        self.assertTrue(M._pdf_is_garbage(10, 9))
        self.assertTrue(M._pdf_is_garbage(10, -1))      # 아예 읽지 못한 경우
        self.assertFalse(M._pdf_is_garbage(10, 2))
        self.assertFalse(M._pdf_is_garbage(2, 2))       # 표본이 너무 적으면 판단하지 않는다

    def test_broken_render_is_reported(self):
        """PDF에 본문이 빠지면 경고한다. Markdown 자체는 영향받지 않는다."""
        make_docx(self.inp / "b.docx")
        saved = pdfwrite.render_markdown
        pdfwrite.render_markdown = (
            lambda md, font, title, base_dir=None, report=None, summary=None:
            saved("# 제목만 남긴다", font, title, base_dir, report, summary))
        try:
            code = self.run_cli(self.inp / "b.docx", "--out", self.out, "--pdf", "result")
        finally:
            pdfwrite.render_markdown = saved
        self.assertEqual(code, 0)                         # Markdown 은 그대로 success
        md = self.md_of("b.docx")
        self.assertIn("본문이", md)
        self.assertIn("Markdown 은 그대로다", md)
        self.assertFalse((self.out / "b.docx.pdf").exists())   # 쓰레기를 남기지 않는다


class TestPdfFontCoverage(Tmp):
    def test_font_without_korean_is_refused(self):
        """한글을 못 그리는 글꼴로 PDF를 만들면 빈 네모만 남는다. 미리 거른다."""
        font, why = pdfwrite.find_font(sample="ABC 123")
        if font is None:
            self.skipTest(why)
        self.assertGreater(pdfwrite.covers(font, "ABC"), 0.9)
        fake = font.gid
        try:
            font.gid = lambda c: 0 if ord(c) > 0x2000 else fake(c)   # 한글이 없는 셈
            self.assertLess(pdfwrite.covers(font, "가나다"), 0.1)
        finally:
            font.gid = fake

    def test_missing_glyphs_are_reported(self):
        font, why = pdfwrite.find_font(sample="ABC")
        if font is None:
            self.skipTest(why)
        report = {}
        pdfwrite.render_markdown("보통 글자\n\n\ue000\ue001 사용자 영역", font, "x", report=report)
        self.assertGreaterEqual(report["missing_total"], 2)   # 글꼴에 없는 글자를 센다


class TestPdfReadability(Tmp):
    """PDF는 읽기용 산출물이다. 기계용 표기가 아니라 사람이 읽는 형태여야 한다."""

    def setUp(self):
        super().setUp()
        if pdfwrite.find_font(sample=KOREAN_PROBE)[0] is None:
            self.skipTest("한글을 그릴 글꼴이 없다")
        self.p = self.inp / "보고서.hwpx"
        make_hwpx(self.p)
        self.assertEqual(self.run_cli(self.p, "--out", self.out, "--pdf", "result"), 0)
        doc = pdf.Pdf((self.out / "보고서.hwpx.pdf").read_bytes())
        self.pages = doc.pages()
        self.text = ["".join(pdf.extract_page(doc, p).parts) for p in self.pages]
        self.all = "".join(self.text)

    def test_status_is_human_readable_not_json(self):
        self.assertIn("변환 상태", self.all)
        self.assertIn("전체 보존 검증 통과", self.all)
        self.assertIn("검사 항목", self.all)
        self.assertNotIn('{"item"', self.all)          # 날것의 JSON 을 보여주지 않는다
        self.assertNotIn('"ok": true', self.all)

    def test_says_what_the_pdf_is_for(self):
        self.assertIn("읽기용 산출물", self.all)
        self.assertIn("meta.json", self.all)           # 기록이 어디 있는지 알려준다

    def test_machine_comments_are_not_drawn(self):
        self.assertNotIn("<!--", self.all)
        self.assertNotIn("Contents/section", self.all)  # 구역 파일 이름은 기계용이다

    def test_page_number_shows_total(self):
        self.assertIn("1 / %d" % len(self.pages), self.all)

    def test_running_header_after_first_page(self):
        if len(self.pages) < 2:
            self.skipTest("한 쪽짜리")
        self.assertIn("보고서.hwpx", self.text[1])      # 둘째 쪽부터 문서 이름을 달고 다닌다

    def test_control_characters_are_not_drawn_as_boxes(self):
        report = {}
        font, _ = pdfwrite.find_font(sample=KOREAN_PROBE)
        pdfwrite.render_markdown("앞\t뒤\x0b가운데", font, "x", report=report)
        self.assertEqual(report["missing_total"], 0)   # 탭·제어문자를 글리프로 찾지 않는다
