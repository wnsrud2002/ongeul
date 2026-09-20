#!/usr/bin/env python3
"""mdmaker: 로컬 문서를 Markdown으로 변환한다 (txt/md/csv/tsv/docx/xlsx/hwpx).

완전 보존이 최우선이다. 변환 결과는 항상 "원본을 변환기와 다른 경로로 다시 읽어"
만든 요소 목록과 대조하며, 대조를 통과하지 못하면 success로 올리지 않는다.
외부 API·LLM·네트워크를 사용하지 않는다.
"""
from __future__ import annotations

import argparse
import copy
import csv
import datetime as _dt
import hashlib
import io
import json
import os
import platform
import re
import sys
import tempfile
import time
import shutil
import subprocess
import zipfile
import xml.etree.ElementTree as ET

import hwp5
import pdf
from pathlib import Path

VERSION = "0.1.0"
CHECK_VERSION = "1"

SUCCESS, PARTIAL, UNVERIFIED, FAILED, UNSUPPORTED, SKIPPED = (
    "success", "partial", "unverified", "failed", "unsupported", "skipped")
# 완료로 인정하는 상태는 success 하나뿐이다.
RANK = {SUCCESS: 0, UNVERIFIED: 1, PARTIAL: 2, UNSUPPORTED: 3, SKIPPED: 3, FAILED: 4}

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def q(ns: str, tag: str) -> str:
    return "{%s}%s" % (NS[ns], tag)


# --- 한도: 조용히 자르지 말고 partial로 보고하기 위한 안전장치 ---
class Limits:
    def __init__(self, max_bytes=512 * 1024 * 1024, max_zip_entries=20000,
                 max_unzipped=2 * 1024 * 1024 * 1024, max_xml_bytes=256 * 1024 * 1024,
                 max_cells=2_000_000, max_pages=5000, ocr_timeout=120,
                 convert_timeout=180, ocr_max_images=3):
        self.max_bytes = max_bytes
        self.max_zip_entries = max_zip_entries
        self.max_unzipped = max_unzipped
        self.max_xml_bytes = max_xml_bytes
        self.max_cells = max_cells
        self.max_pages = max_pages
        self.ocr_timeout = ocr_timeout
        self.convert_timeout = convert_timeout
        self.ocr_max_images = ocr_max_images


class Res:
    """형식별 변환 함수의 공통 반환값 (가이드 4장의 표)."""

    def __init__(self, src: Path, fmt: str):
        self.src = src
        self.fmt = fmt
        self.markdown = ""
        self.status = SUCCESS   # 검사에서 내려갈 수만 있다
        self.warnings: list[str] = []
        self.checks: list[dict] = []
        self.assets: dict[str, bytes] = {}       # 보조폴더 상대경로 -> 원본 바이트
        self.meta: dict = {}                     # 구조·서식 정보
        self.info: dict = {"converter": "mdmaker/%s" % VERSION, "check_version": CHECK_VERSION}

    def warn(self, msg: str) -> None:
        self.warnings.append(msg)

    def demote(self, status: str) -> None:
        if RANK[status] > RANK[self.status]:
            self.status = status


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


# ---------------------------------------------------------------- 안전한 XML/ZIP
_DOCTYPE = re.compile(rb"<!(DOCTYPE|ENTITY)", re.I)


def parse_xml(data: bytes, limits: Limits) -> ET.Element:
    """외부 엔티티·DTD를 원천 차단한다. expat은 기본적으로 외부 엔티티를 읽지
    않지만 내부 엔티티 확장은 허용하므로 선언 자체를 거부한다."""
    if len(data) > limits.max_xml_bytes:
        raise ValueError("XML 크기 한도 초과: %d바이트" % len(data))
    if _DOCTYPE.search(data[:65536]):
        raise ValueError("DTD/엔티티 선언이 있는 XML은 처리하지 않는다")
    return ET.fromstring(data)


def open_zip(path: Path, limits: Limits) -> zipfile.ZipFile:
    zf = zipfile.ZipFile(path)
    infos = zf.infolist()
    if len(infos) > limits.max_zip_entries:
        zf.close()
        raise ValueError("ZIP 항목 수 한도 초과: %d" % len(infos))
    total = 0
    for i in infos:
        name = i.filename
        if name.startswith("/") or ".." in Path(name).parts or (len(name) > 1 and name[1] == ":"):
            zf.close()
            raise ValueError("ZIP 경로 탈출 시도: %r" % name)
        total += i.file_size
        if total > limits.max_unzipped:
            zf.close()
            raise ValueError("ZIP 해제 크기 한도 초과")
    return zf


SIGNATURES = {".docx": b"PK\x03\x04", ".xlsx": b"PK\x03\x04", ".hwpx": b"PK\x03\x04",
              ".pptx": b"PK\x03\x04"}


def check_signature(path: Path) -> str | None:
    want = SIGNATURES.get(path.suffix.lower())
    if not want:
        return None
    with path.open("rb") as fh:
        got = fh.read(len(want))
    return None if got == want else "파일 시그니처 불일치: %r 기대, %r" % (want, got)


# ---------------------------------------------------------------- Markdown 유틸
_MD_LINE_START = re.compile(r"^([#>\-+*=]|\d+[.)]|\|)")


def esc_inline(text: str) -> str:
    """내용은 바꾸지 않고 표현만 안전하게 만든다. normalize_md()로 되돌릴 수 있어야 한다."""
    return text.replace("\\", "\\\\").replace("|", "\\|")


def esc_block(text: str) -> str:
    """줄 첫 글자만 막는다. 역슬래시는 esc_inline에서 이미 처리했으므로 다시 건드리지 않는다."""
    out = []
    for line in text.split("\n"):
        e = line
        if _MD_LINE_START.match(e.lstrip()):
            stripped = e.lstrip()
            e = e[: len(e) - len(stripped)] + "\\" + stripped
        out.append(e)
    return "\n".join(out)


def esc_cell(text: str) -> str:
    """원본 값(셀 값·CSV 필드)을 표 칸에 넣을 때 쓴다."""
    return esc_inline(text).replace("\n", "<br>")


def cell_inline(text: str) -> str:
    r"""이미 이스케이프를 거친 문단 텍스트를 표 칸에 넣을 때 쓴다.
    여기서 또 이스케이프하면 `\|` 가 `\\\|` 가 되어 원문과 달라진다."""
    return text.replace("\n", "<br>")


_TAG = re.compile(r"</?span[^>]*>")
_ESCAPED = re.compile(r"\\(.)")


def present(text: str, md: str, norm: str) -> bool:
    """정규화본이든 원본 Markdown이든 한쪽에 그대로 있으면 보존된 것이다.
    원문에 `*`나 백틱이 들어있으면 정규화가 그 글자까지 지워 오판이 난다."""
    if text in md or text in norm:
        return True
    # 원문에 `**`나 백틱이 있으면 결과에는 이스케이프가 붙는다. 같은 규칙을 원문에도
    # 적용해 양쪽을 같은 기준으로 비교한다.
    same = normalize_md(text)
    return bool(same.strip()) and same in norm


def normalize_md(md: str) -> str:
    """보존 검사용: 우리가 넣은 마크업만 걷어낸 평문을 만든다."""
    s = _TAG.sub("", md)
    s = s.replace("<br>", "\n")
    s = s.replace("~~", "").replace("**", "").replace("`", "")
    s = _ESCAPED.sub(r"\1", s)
    return s


def fence_for(text: str) -> str:
    longest = max((len(m) for m in re.findall(r"`+", text)), default=0)
    return "`" * max(3, longest + 1)


def header_block(res: Res, extra: dict | None = None) -> str:
    lines = ["# %s" % res.src.name, "",
             "- 원본 형식: `%s`" % res.fmt,
             "- 변환기: `%s`" % res.info["converter"]]
    for k, v in (extra or {}).items():
        lines.append("- %s: %s" % (k, v))
    lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------- 텍스트 / CSV
ENCODINGS = ("utf-8-sig", "cp949", "utf-16")


def read_text(path: Path, encoding: str | None) -> tuple[str, str]:
    raw = path.read_bytes()
    if encoding:
        return raw.decode(encoding), encoding  # 실패하면 예외를 그대로 올린다
    for enc in ENCODINGS:
        try:
            return raw.decode(enc), enc
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ValueError("인코딩을 추정하지 못했다. --encoding 으로 지정할 것")


def convert_text(path: Path, opts) -> Res:
    fmt = path.suffix.lower().lstrip(".") or "txt"
    res = Res(path, fmt)
    text, enc = read_text(path, opts.encoding)
    res.info["encoding"] = enc
    res.info["chars"] = len(text)
    if fmt == "md":
        # 이미 Markdown이므로 그대로 둔다. 재작성하면 원문이 바뀐다.
        body = text
    else:
        f = fence_for(text)
        body = "%s text\n%s\n%s" % (f, text, f)
    res.markdown = header_block(res, {"인코딩": "`%s`" % enc}) + "\n" + body + "\n"
    # 검사: 원문 전체가 결과에 글자 그대로 들어있는지 본다.
    ok = text in res.markdown
    res.checks.append({"item": "본문 전체", "chars": len(text), "ok": ok})
    if not ok:
        res.demote(PARTIAL)
        res.warn("본문이 결과에 그대로 포함되지 않았다")
    return res


def convert_csv(path: Path, opts) -> Res:
    fmt = path.suffix.lower().lstrip(".")
    res = Res(path, fmt)
    text, enc = read_text(path, opts.encoding)
    res.info["encoding"] = enc
    delim = "\t" if fmt == "tsv" else ","
    rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=delim))
    if not rows:
        res.markdown = header_block(res, {"인코딩": "`%s`" % enc}) + "\n(빈 파일)\n"
        res.checks.append({"item": "행", "count": 0, "ok": True})
        return res
    width = max(len(r) for r in rows)
    ragged = sorted({len(r) for r in rows}) if len({len(r) for r in rows}) > 1 else None
    out = [header_block(res, {"인코딩": "`%s`" % enc, "구분자": "`%s`" % ("TAB" if delim == "\t" else delim),
                              "행/열": "%d행 × 최대 %d열" % (len(rows), width)}), ""]
    out.append("> 표 헤더 정보가 없는 형식이므로 첫 행을 헤더로 바꾸지 않는다. 아래 헤더는 열 번호다.")
    out.append("")
    out.append("| 행 | " + " | ".join("열%d" % (i + 1) for i in range(width)) + " |")
    out.append("| --- | " + " | ".join("---" for _ in range(width)) + " |")
    for n, row in enumerate(rows, 1):
        cells = [esc_cell(c) for c in row] + ["" for _ in range(width - len(row))]
        out.append("| %d | %s |" % (n, " | ".join(cells)))
    if ragged:
        out.append("")
        out.append("> 행마다 열 수가 다르다(%s). 원본 열 수는 보조 정보에 기록했다."
                   % ", ".join(str(x) for x in ragged))
    res.markdown = "\n".join(out) + "\n"
    res.meta = {"format": fmt, "delimiter": delim, "encoding": enc,
                "row_widths": [len(r) for r in rows]}
    # 검사: csv 모듈과 무관하게 원본 텍스트에서 셀을 다시 세어 대조한다.
    norm = normalize_md(res.markdown)
    missing = [(n, i) for n, row in enumerate(rows, 1) for i, c in enumerate(row)
               if c.strip() and not present(c, res.markdown, norm)]
    res.checks.append({"item": "셀", "count": sum(len(r) for r in rows),
                       "missing": len(missing), "ok": not missing})
    if missing:
        res.demote(PARTIAL)
        res.warn("셀 %d개가 결과에서 확인되지 않았다: %r" % (len(missing), missing[:5]))
    return res


# ---------------------------------------------------------------- DOCX
def _rels(zf, limits, part: str) -> dict:
    name = str(Path(part).parent / "_rels" / (Path(part).name + ".rels"))
    try:
        root = parse_xml(zf.read(name), limits)
    except (KeyError, ET.ParseError):
        return {}
    out = {}
    for rel in root.findall(q("pr", "Relationship")):
        out[rel.get("Id")] = (rel.get("Target", ""), rel.get("TargetMode", "Internal"),
                              rel.get("Type", ""))
    return out


def join_part(base: str, target: str) -> str:
    """OOXML 관계의 Target을 패키지 경로로 바꾼다.
    `/ppt/...` 처럼 절대형으로 쓰는 생성기가 있어 앞의 /를 떼야 한다."""
    if target.startswith("/"):
        return os.path.normpath(target[1:]).replace(os.sep, "/")
    return os.path.normpath(str(Path(base) / target)).replace(os.sep, "/")


REL_TARGET = re.compile(rb'Target="([^"]+)"')


def referenced_parts(zf, prefix: str) -> set:
    """패키지의 모든 .rels가 가리키는 대상 이름(파일명 기준)."""
    out = set()
    for n in zf.namelist():
        if not n.endswith(".rels"):
            continue
        try:
            data = zf.read(n)
        except KeyError:
            continue
        for m in REL_TARGET.finditer(data):
            t = m.group(1).decode("utf-8", "replace")
            if prefix in t:
                out.add(Path(t).name)
    return out


def check_orphan_media(res: Res, zf, media_prefix: str, rendered: set) -> None:
    """본문에 그리지 못한 자산을 구분해 보고한다.
    어디에서도 참조하지 않는 잔여 파일은 문서 내용이 아니므로 등급을 내리지 않는다."""
    media = [n for n in zf.namelist() if n.startswith(media_prefix) and not n.endswith("/")]
    refs = referenced_parts(zf, "media/")
    orphan = []
    for n in media:
        if n in rendered:
            continue
        res.assets.setdefault("media/" + Path(n).name, zf.read(n))
        if Path(n).name in refs:
            res.warn("문서가 참조하지만 본문 위치를 확인하지 못한 자산을 보존했다: %s" % n)
            res.demote(UNVERIFIED)
        else:
            orphan.append(n)
    if orphan:
        res.info["orphan_media"] = len(orphan)
        res.markdown += ("\n> 패키지에 남아 있지만 문서 어느 부분도 참조하지 않는 이미지 %d개를 "
                         "보조 폴더에 원본 그대로 보존했다(본문 내용은 아니다).\n" % len(orphan))
    res.checks.append({"item": "자산", "count": len(media), "saved": len(res.assets),
                       "orphan": len(orphan), "ok": len(res.assets) >= len(media)})


def _part(zf, limits, name: str):
    """없거나 망가진 XML 조각은 None. 빈 파일(0바이트)인 _rels 도 실제로 존재한다."""
    try:
        return parse_xml(zf.read(name), limits)
    except (KeyError, ET.ParseError):
        return None


class _DocxCtx:
    def __init__(self, res, zf, limits):
        self.res, self.zf, self.limits = res, zf, limits
        self.rels: dict = {}
        self.numfmt: dict = {}
        self.styles: dict = {}
        self.images: list = []


def _tag(el) -> str:
    return el.tag.split("}")[-1]


def _dx_run(r, ctx) -> str:
    pr = r.find(q("w", "rPr"))
    pre, post = "", ""
    style_note = ""
    if pr is not None:
        if pr.find(q("w", "b")) is not None and pr.find(q("w", "b")).get(q("w", "val"), "1") not in ("0", "false"):
            pre, post = pre + "**", "**" + post
        it = pr.find(q("w", "i"))
        if it is not None and it.get(q("w", "val"), "1") not in ("0", "false"):
            pre, post = pre + "*", "*" + post
        st = pr.find(q("w", "strike")) if pr.find(q("w", "strike")) is not None else pr.find(q("w", "dstrike"))
        if st is not None and st.get(q("w", "val"), "1") not in ("0", "false"):
            pre, post = pre + "~~", "~~" + post
        col = pr.find(q("w", "color"))
        hl = pr.find(q("w", "highlight"))
        css = []
        if col is not None and col.get(q("w", "val"), "auto") != "auto":
            css.append("color:#%s" % col.get(q("w", "val")))
        if hl is not None:
            css.append("background:%s" % hl.get(q("w", "val")))
        if css:
            style_note = ";".join(css)
    body = _dx_children(r, ctx)
    if not body:
        return ""
    if style_note:
        body = '<span style="%s">%s</span>' % (style_note, body)
    return pre + body + post


def _dx_children(node, ctx) -> str:
    out = []
    for el in node:
        t = _tag(el)
        if t in ("t", "delText", "instrText"):
            out.append(esc_inline(el.text or ""))
        elif t == "tab":
            out.append("\t")
        elif t in ("br", "cr"):
            out.append("\n")
        elif t == "noBreakHyphen":
            out.append("-")
        elif t == "softHyphen":
            out.append("­")
        elif t == "sym":
            ch = el.get(q("w", "char"))
            out.append(chr(int(ch, 16)) if ch else "")
        elif t == "r":
            out.append(_dx_run(el, ctx))
        elif t == "hyperlink":
            inner = _dx_children(el, ctx)
            rid = el.get(q("r", "id"))
            target = ctx.rels.get(rid, ("", "", ""))[0] if rid else ""
            anchor = el.get(q("w", "anchor"))
            href = target or ("#" + anchor if anchor else "")
            out.append("[%s](%s)" % (inner, href) if href else inner)
        elif t == "ins":
            out.append("{+삽입:%s+}" % _dx_children(el, ctx))
        elif t == "del":
            out.append("{-삭제:%s-}" % _dx_children(el, ctx))
        elif t in ("moveFrom", "moveTo"):
            out.append("{%s:%s}" % (t, _dx_children(el, ctx)))
        elif t == "footnoteReference":
            out.append("[^fn%s]" % el.get(q("w", "id")))
        elif t == "endnoteReference":
            out.append("[^en%s]" % el.get(q("w", "id")))
        elif t == "commentReference":
            out.append("[^cm%s]" % el.get(q("w", "id")))
        elif t in ("drawing", "pict", "object"):
            out.append(_dx_graphic(el, ctx))
        elif t in ("smartTag", "sdt", "sdtContent", "fldSimple", "customXml", "bdo", "dir"):
            out.append(_dx_children(el, ctx))
        elif t in ("rPr", "pPr", "bookmarkStart", "bookmarkEnd", "proofErr", "lastRenderedPageBreak",
                   "commentRangeStart", "commentRangeEnd", "fldChar", "rPrChange", "pPrChange",
                   "sectPr", "tblPr", "tblGrid", "tcPr", "trPr", "footnoteRef", "endnoteRef",
                   "annotationRef", "separator", "continuationSeparator"):
            continue
        else:
            # 모르는 요소라도 안에 글자가 있으면 잃지 않는다.
            inner = _dx_children(el, ctx)
            if inner:
                out.append(inner)
            elif (el.text or "").strip():
                out.append(esc_inline(el.text))
    return "".join(out)


def _dx_graphic(el, ctx) -> str:
    """그림·텍스트상자·임베디드 개체. 자산은 원본 바이트로 보존한다."""
    parts = []
    for tb in el.iter():
        if _tag(tb) == "txbxContent":
            inner = "<br>".join(_dx_block(c, ctx) for c in tb)
            parts.append("[텍스트상자: %s]" % inner)
    for blip in el.iter():
        t = _tag(blip)
        rid = blip.get(q("r", "embed")) or blip.get(q("r", "id"))
        if t in ("blip", "imagedata", "OLEObject") and rid:
            target = ctx.rels.get(rid, ("", "", ""))[0]
            if not target:
                continue
            name = _dx_save_asset(ctx, target)
            if name:
                alt = ""
                for anc in el.iter():
                    if _tag(anc) in ("docPr", "cNvPr"):
                        alt = anc.get("descr") or anc.get("title") or anc.get("name") or ""
                        break
                label = "임베디드개체" if t == "OLEObject" else "이미지"
                parts.append("![%s: %s](%s)" % (label, esc_inline(alt) or "대체 텍스트 없음", name))
                if not alt:
                    parts.append("[이미지: 내용 해석 안 됨]")
    if not parts:
        # 글자도 이미지 데이터도 없는 도형·빈 그림틀. 잃은 내용은 없지만 도형의 생김새는
        # Markdown으로 확인할 수 없으므로 unverified로 남긴다(누락과는 구분한다).
        label = next((x.get("title") or x.get("descr") or x.get("alt") or x.get("name")
                      for x in el.iter()
                      if x.get("title") or x.get("descr") or x.get("alt") or x.get("name")), "")
        kinds = sorted({_tag(x) for x in el.iter()}
                       & {"wsp", "rect", "line", "oval", "shape", "group", "chart",
                          "diagramData", "imagedata", "roundrect", "polyline"})
        parts.append("[도형/빈 그림틀: %s%s]"
                     % (",".join(kinds) or _tag(el),
                        (" · " + esc_inline(label)) if label else ""))
        ctx.res.warn("도형 개체에 글자·이미지 데이터가 없어 생김새를 검증하지 못했다: <%s>"
                     % _tag(el))
        ctx.res.demote(UNVERIFIED)
    return "".join(parts)


def _dx_save_asset(ctx, target: str) -> str | None:
    rel = str(Path("word") / target) if not target.startswith("/") else target.lstrip("/")
    rel = os.path.normpath(rel).replace(os.sep, "/")
    try:
        data = ctx.zf.read(rel)
    except KeyError:
        ctx.res.warn("자산을 찾지 못했다: %s" % target)
        ctx.res.demote(PARTIAL)
        return None
    name = "media/" + Path(rel).name
    ctx.res.assets[name] = data
    ctx.images.append({"part": rel, "asset": name, "sha256": sha256(data), "bytes": len(data)})
    return name


def _dx_para(p, ctx) -> str:
    text = _dx_children(p, ctx)
    pr = p.find(q("w", "pPr"))
    level, numid, ilvl = 0, None, 0
    if pr is not None:
        ps = pr.find(q("w", "pStyle"))
        if ps is not None:
            level = ctx.styles.get(ps.get(q("w", "val")), 0)
        ol = pr.find(q("w", "outlineLvl"))
        if not level and ol is not None:
            level = int(ol.get(q("w", "val"), "9")) + 1
            level = level if level <= 6 else 0
        np = pr.find(q("w", "numPr"))
        if np is not None:
            n = np.find(q("w", "numId"))
            il = np.find(q("w", "ilvl"))
            numid = n.get(q("w", "val")) if n is not None else None
            ilvl = int(il.get(q("w", "val"), "0")) if il is not None else 0
    if not text.strip() and numid is None:
        return esc_block(text)
    if level:
        return "#" * min(level, 6) + " " + text
    if numid is not None:
        fmt = ctx.numfmt.get((numid, ilvl), "bullet")
        marker = "-" if fmt == "bullet" else "1."
        return "  " * ilvl + marker + " " + text
    return esc_block(text)


def _dx_cell_text(tc, ctx) -> tuple[str, bool]:
    blocks, complex_ = [], False
    for c in tc:
        t = _tag(c)
        if t == "p":
            blocks.append(_dx_para(c, ctx))
        elif t == "tbl":
            complex_ = True
            blocks.append(_dx_table(c, ctx, nested=True))
        elif t == "tcPr":
            continue
        else:
            inner = _dx_children(c, ctx)
            if inner:
                blocks.append(inner)
    txt = "\n".join(blocks)
    return txt, complex_ or len(blocks) > 1


def _dx_table(tbl, ctx, nested=False) -> str:
    rows = []
    complex_ = nested
    for tr in tbl.findall(q("w", "tr")):
        row = []
        for tc in tr.findall(q("w", "tc")):
            txt, cx = _dx_cell_text(tc, ctx)
            complex_ = complex_ or cx
            pr = tc.find(q("w", "tcPr"))
            span, vmerge = 1, None
            if pr is not None:
                gs = pr.find(q("w", "gridSpan"))
                if gs is not None:
                    span = int(gs.get(q("w", "val"), "1"))
                vm = pr.find(q("w", "vMerge"))
                if vm is not None:
                    vmerge = vm.get(q("w", "val"), "continue")
            if span > 1 or vmerge:
                complex_ = True
            row.append({"text": txt, "span": span, "vmerge": vmerge})
        rows.append(row)
    if not rows:
        return ""
    header_marked = False
    tr0 = tbl.find(q("w", "tr"))
    if tr0 is not None:
        trpr = tr0.find(q("w", "trPr"))
        header_marked = trpr is not None and trpr.find(q("w", "tblHeader")) is not None
    if not complex_:
        width = max(len(r) for r in rows)
        out = []
        if header_marked:
            out.append("| " + " | ".join(cell_inline(c["text"]) for c in rows[0]) + " |")
            body = rows[1:]
        else:
            out.append("| " + " | ".join("열%d" % (i + 1) for i in range(width)) + " |")
            body = rows
        out.append("| " + " | ".join("---" for _ in range(width)) + " |")
        for r in body:
            cells = [cell_inline(c["text"]) for c in r] + [""] * (width - len(r))
            out.append("| " + " | ".join(cells) + " |")
        if not header_marked:
            out.insert(0, "> 원본에 헤더 행 표시가 없어 첫 행을 데이터로 둔다.")
            out.insert(1, "")
        return "\n".join(out)
    # 병합·중첩·다문단 표는 좌표를 남기는 행별 표현으로 바꾼다.
    out = ["> 병합·중첩 때문에 Markdown 표 대신 행별 표현으로 변환했다. 좌표는 원본 기준이다.", ""]
    for ri, r in enumerate(rows, 1):
        ci = 1
        for c in r:
            coord = "R%dC%d" % (ri, ci)
            extra = []
            if c["span"] > 1:
                extra.append("가로병합 %d칸" % c["span"])
            if c["vmerge"] == "restart":
                extra.append("세로병합 시작")
            elif c["vmerge"]:
                extra.append("세로병합 계속")
            tag = " (%s)" % ", ".join(extra) if extra else ""
            out.append("- **%s**%s: %s" % (coord, tag, cell_block(c["text"])))
            ci += c["span"]
    return "\n".join(out)


def _dx_block(el, ctx) -> str:
    t = _tag(el)
    if t == "p":
        return _dx_para(el, ctx)
    if t == "tbl":
        return _dx_table(el, ctx)
    if t in ("sdt", "sdtContent", "customXml"):
        return "\n\n".join(x for x in (_dx_block(c, ctx) for c in el) if x)
    if t == "sectPr":
        return ""
    inner = _dx_children(el, ctx)
    return inner


def _dx_body(root, ctx) -> str:
    body = root.find(q("w", "body")) if _tag(root) == "document" else root
    if body is None:
        return ""
    return "\n\n".join(x for x in (_dx_block(c, ctx) for c in body) if x != "")


def convert_docx(path: Path, opts, limits: Limits) -> Res:
    res = Res(path, "docx")
    zf = open_zip(path, limits)
    try:
        ctx = _DocxCtx(res, zf, limits)
        ctx.rels = _rels(zf, limits, "word/document.xml")
        st = _part(zf, limits, "word/styles.xml")
        if st is not None:
            for s in st.findall(q("w", "style")):
                sid = s.get(q("w", "styleId"), "")
                nm = s.find(q("w", "name"))
                label = (nm.get(q("w", "val")) if nm is not None else "") or sid
                m = re.match(r"^(?:heading|제목)\s*([1-9])$", label.strip(), re.I)
                if m:
                    ctx.styles[sid] = int(m.group(1))
        num = _part(zf, limits, "word/numbering.xml")
        if num is not None:
            abstract = {}
            for an in num.findall(q("w", "abstractNum")):
                aid = an.get(q("w", "abstractNumId"))
                for lvl in an.findall(q("w", "lvl")):
                    nf = lvl.find(q("w", "numFmt"))
                    abstract[(aid, int(lvl.get(q("w", "ilvl"), "0")))] = (
                        nf.get(q("w", "val"), "bullet") if nf is not None else "bullet")
            for n in num.findall(q("w", "num")):
                nid = n.get(q("w", "numId"))
                a = n.find(q("w", "abstractNumId"))
                aid = a.get(q("w", "val")) if a is not None else None
                for (ai, il), fmt in abstract.items():
                    if ai == aid:
                        ctx.numfmt[(nid, il)] = fmt

        doc = _part(zf, limits, "word/document.xml")
        if doc is None:
            res.demote(FAILED)
            res.warn("word/document.xml 이 없다")
            return res
        parts_md = [header_block(res)]
        parts_md.append(_dx_body(doc, ctx))

        # 각주·미주·주석·머리글·바닥글도 본문과 함께 보존한다.
        for part, title, prefix, wrap in (
                ("word/footnotes.xml", "각주", "fn", q("w", "footnote")),
                ("word/endnotes.xml", "미주", "en", q("w", "endnote")),
                ("word/comments.xml", "주석", "cm", q("w", "comment"))):
            root = _part(zf, limits, part)
            if root is None:
                continue
            saved, ctx.rels = ctx.rels, _rels(zf, limits, part) or ctx.rels
            items = []
            for node in root.findall(wrap):
                nid = node.get(q("w", "id"))
                typ = node.get(q("w", "type"), "")
                if typ in ("separator", "continuationSeparator"):
                    continue
                txt = "\n\n".join(x for x in (_dx_block(c, ctx) for c in node) if x)
                who = node.get(q("w", "author"))
                items.append("[^%s%s]: %s%s" % (prefix, nid, ("(%s) " % who) if who else "", txt))
            ctx.rels = saved
            if items:
                parts_md.append("## %s" % title)
                parts_md.append("\n\n".join(items))

        hf = [(rid, t) for rid, (t, mode, typ) in ctx.rels.items()
              if typ.endswith("/header") or typ.endswith("/footer")]
        for rid, target in sorted(hf, key=lambda x: x[1]):
            part = "word/" + target
            root = _part(zf, limits, part)
            if root is None:
                continue
            saved, ctx.rels = ctx.rels, _rels(zf, limits, part) or ctx.rels
            txt = _dx_body(root, ctx)
            ctx.rels = saved
            if txt.strip():
                parts_md.append("## %s: %s" % ("머리글" if "header" in target else "바닥글", target))
                parts_md.append(txt)

        res.markdown = "\n\n".join(x for x in parts_md if x) + "\n"
        res.meta = {"format": "docx", "images": ctx.images,
                    "note": "변경추적은 {+삽입:...+} {-삭제:...-} 로 보존한다. 자동 수락/거절하지 않는다."}
        res.info["chars"] = len(res.markdown)
        verify_ooxml_text(res, zf, limits, "docx")
        return res
    finally:
        zf.close()


# ---------------------------------------------------------------- 보존 검사(OOXML)
_XML_TEXT = re.compile(rb"<(?:w:)?(t|delText|instrText)(?:\s[^>]*)?>(.*?)</(?:w:)?\1>", re.S)
_ENT = re.compile(r"&(#x?[0-9A-Fa-f]+|amp|lt|gt|quot|apos);")


def _unent(s: str) -> str:
    def sub(m):
        g = m.group(1)
        if g.startswith("#x") or g.startswith("#X"):
            return chr(int(g[2:], 16))
        if g.startswith("#"):
            return chr(int(g[1:]))
        return {"amp": "&", "lt": "<", "gt": ">", "quot": '"', "apos": "'"}[g]
    return _ENT.sub(sub, s).replace("\r\n", "\n").replace("\r", "\n")


def verify_ooxml_text(res: Res, zf, limits: Limits, kind: str) -> None:
    """변환기와 다른 경로(원본 XML 직접 스캔)로 요소 목록을 만들고 결과와 대조한다."""
    norm = normalize_md(res.markdown)
    want_parts = [n for n in zf.namelist()
                  if n.startswith("word/") and n.endswith(".xml")
                  and not n.startswith("word/_rels")
                  and Path(n).name not in ("settings.xml", "styles.xml", "numbering.xml",
                                           "fontTable.xml", "webSettings.xml", "theme1.xml")]
    missing, total = [], 0
    for part in sorted(want_parts):
        data = zf.read(part)
        if len(data) > limits.max_xml_bytes:
            res.warn("%s 가 XML 한도를 넘어 검사하지 못했다" % part)
            res.demote(UNVERIFIED)
            continue
        for m in _XML_TEXT.finditer(data):
            txt = _unent(m.group(2).decode("utf-8"))
            if not txt.strip():
                continue
            total += 1
            if not present(txt, res.markdown, norm):
                missing.append((part, txt[:60]))
    res.checks.append({"item": "텍스트 노드", "count": total, "missing": len(missing),
                       "ok": not missing})
    if missing:
        res.demote(PARTIAL)
        res.warn("원본 텍스트 %d개가 결과에서 확인되지 않았다: %r" % (len(missing), missing[:5]))

    saved_parts = {i["part"] for i in res.meta.get("images", [])} if res.meta else set()
    ok_assets = all(sha256(res.assets[i["asset"]]) == i["sha256"]
                    for i in (res.meta.get("images", []) if res.meta else [])
                    if i["asset"] in res.assets)
    if not ok_assets:
        res.demote(PARTIAL)
        res.warn("자산 해시가 일치하지 않는다")
    check_orphan_media(res, zf, "word/media/", saved_parts)


# ---------------------------------------------------------------- XLSX
def _xl_num(v) -> str:
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        return str(int(v))
    return str(v)


def _xl_serial(v):
    from openpyxl.utils.datetime import to_excel, CALENDAR_WINDOWS_1900
    try:
        if isinstance(v, _dt.datetime):
            return to_excel(v, CALENDAR_WINDOWS_1900)
        if isinstance(v, _dt.date):
            return to_excel(_dt.datetime(v.year, v.month, v.day), CALENDAR_WINDOWS_1900)
        if isinstance(v, _dt.time):
            return to_excel(v, CALENDAR_WINDOWS_1900)
        if isinstance(v, _dt.timedelta):
            return v.total_seconds() / 86400.0
    except Exception:
        return None
    return None


def _xl_value(v) -> str:
    """값 자체를 바꾸지 않는다. 0 / False / 빈 문자열 / 빈 셀을 구별해 표기한다."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, str):
        return "[빈 문자열]" if v == "" else v
    if isinstance(v, (_dt.datetime, _dt.date, _dt.time, _dt.timedelta)):
        return v.isoformat() if not isinstance(v, _dt.timedelta) else str(v)
    return _xl_num(v)


def _xl_esc_tsv(s: str) -> str:
    return s.replace("\\", "\\\\").replace("\t", "\\t").replace("\n", "\\n").replace("\r", "\\r")


_XL_STYLE_REF = re.compile(rb'\ss="(\d+)"')
_XL_ITEM = {"cellXfs": "xf", "cellStyleXfs": "xf", "fonts": "font", "fills": "fill",
            "borders": "border", "numFmts": "numFmt", "cellStyles": "cellStyle", "dxfs": "dxf"}


def _xlsx_repair_styles(path: Path, limits: Limits) -> Path | None:
    """엑셀이 아닌 도구가 만든 파일은 styles.xml이 openpyxl의 검사를 통과하지 못한다
    (이름 없는 cellStyle, 셀이 참조하는 범위 밖 xf 인덱스). 원본은 그대로 두고
    복사본에서 그 두 가지만 고친다. 서식·수식·값은 손대지 않는다."""
    zf = open_zip(path, limits)
    try:
        names = zf.namelist()
        if "xl/styles.xml" not in names:
            return None
        root = parse_xml(zf.read("xl/styles.xml"), limits)
        need = 0
        for n in names:
            if n.startswith("xl/worksheets/") and n.endswith(".xml"):
                for m in _XL_STYLE_REF.finditer(zf.read(n)):
                    need = max(need, int(m.group(1)) + 1)
        changed = False
        for i, cs in enumerate(x for x in root.iter() if _tag(x) == "cellStyle"):
            if cs.get("name") is None:
                cs.set("name", "Style_%d" % i)
                changed = True
        # <AlternateContent>로 감싼 항목은 openpyxl이 건너뛰므로 뒤 항목의 색인이
        # 통째로 밀린다(셀이 참조하는 스타일 번호가 어긋나 IndexError). Fallback 쪽
        # 실제 항목을 꺼내 제자리에 돌려놓는다.
        for holder in [x for x in root.iter() if _tag(x) in _XL_ITEM]:
            item = _XL_ITEM[_tag(holder)]
            for i, child in enumerate(list(holder)):
                if _tag(child) != "AlternateContent":
                    continue
                pick = None
                for branch in ("Fallback", "Choice"):
                    b = next((x for x in child if _tag(x) == branch), None)
                    if b is not None:
                        pick = next((e for e in b.iter() if _tag(e) == item), None)
                        if pick is not None:
                            break
                if pick is None:
                    pick = next((e for e in child.iter() if _tag(e) == item), None)
                holder.remove(child)
                holder.insert(i, pick if pick is not None else ET.Element(q("x", item)))
                changed = True
            if changed:
                holder.set("count", str(len(holder)))
        xfs = next((x for x in root.iter() if _tag(x) == "cellXfs"), None)
        if xfs is not None and len(xfs) < need:
            filler = xfs[0] if len(xfs) else None
            for _ in range(need - len(xfs)):
                xfs.append(copy.deepcopy(filler) if filler is not None
                           else ET.Element(q("x", "xf")))
            xfs.set("count", str(len(xfs)))
            changed = True
        if not changed:
            return None
        ET.register_namespace("", NS["x"])
        fixed = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
        tmp = Path(tempfile.mkdtemp(prefix="mdmaker-xlsx-")) / path.name
        with zipfile.ZipFile(tmp, "w") as out:
            for n in names:
                out.writestr(n, fixed if n == "xl/styles.xml" else zf.read(n))
        return tmp
    finally:
        zf.close()


def _xl_sheet_roots(zf, limits: Limits):
    """워크북 순서대로 (시트 이름, 시트 XML 루트)를 돌려준다."""
    wbx = parse_xml(zf.read("xl/workbook.xml"), limits)
    rels = _rels(zf, limits, "xl/workbook.xml")
    for sh in wbx.iter(q("x", "sheet")):
        target = rels.get(sh.get(q("r", "id")), ("", "", ""))[0]
        target = target[1:] if target.startswith("/") else "xl/" + target.lstrip("./")
        try:
            yield sh.get("name"), parse_xml(zf.read(os.path.normpath(target).replace(os.sep, "/")),
                                            limits)
        except (KeyError, ET.ParseError):
            continue


def _xl_merged_leftovers(path: Path, limits: Limits) -> dict:
    """병합 범위의 좌상단이 아닌 칸에 값이 남아 있는 경우를 찾는다.

    엑셀 화면에는 보이지 않고 openpyxl도 MergedCell로 가려서 None을 주지만, 파일에는
    분명히 저장된 값이므로 버리지 않는다."""
    from openpyxl.utils import range_boundaries, get_column_letter
    out: dict = {}
    zf = open_zip(path, limits)
    try:
        shared = []
        try:
            ss = parse_xml(zf.read("xl/sharedStrings.xml"), limits)
            shared = [_si_text(si) for si in ss.findall(q("x", "si"))]
        except (KeyError, ET.ParseError):
            pass
        for name, sx in _xl_sheet_roots(zf, limits):
            anchors, inside = set(), set()
            for mc in sx.iter(q("x", "mergeCell")):
                ref = mc.get("ref") or ""
                try:
                    c1, r1, c2, r2 = range_boundaries(ref)
                except ValueError:
                    continue
                anchors.add("%s%d" % (get_column_letter(c1), r1))
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        inside.add("%s%d" % (get_column_letter(c), r))
            if not inside:
                continue
            found = {}
            for c in sx.iter(q("x", "c")):
                coord = c.get("r")
                if coord not in inside or coord in anchors:
                    continue
                ve = c.find(q("x", "v"))
                ise = c.find(q("x", "is"))
                if ve is None and ise is None:
                    continue
                if c.get("t") == "s" and ve is not None:
                    try:
                        val = shared[int(ve.text)]
                    except (ValueError, IndexError):
                        val = ve.text or ""
                elif ise is not None:
                    val = "".join(x.text or "" for x in ise.iter(q("x", "t")))
                else:
                    val = ve.text or ""
                if str(val).strip():
                    found[coord] = str(val)
            if found:
                out[name] = found
        return out
    finally:
        zf.close()


def convert_xlsx(path: Path, opts, limits: Limits) -> Res:
    import openpyxl
    from openpyxl.utils import get_column_letter

    res = Res(path, "xlsx")
    res.info["converter"] += " + openpyxl/%s" % openpyxl.__version__
    repaired = None
    try:
        wb = openpyxl.load_workbook(path, data_only=False, rich_text=False)
        wbv = openpyxl.load_workbook(path, data_only=True, rich_text=False)
    except (TypeError, IndexError, KeyError, ValueError) as exc:
        repaired = _xlsx_repair_styles(path, limits)
        if repaired is None:
            raise
        wb = openpyxl.load_workbook(repaired, data_only=False, rich_text=False)
        wbv = openpyxl.load_workbook(repaired, data_only=True, rich_text=False)
        res.warn("styles.xml을 그대로는 읽지 못해(%s: %s) 복사본에서 스타일 참조만 고쳐 읽었다. "
                 "값·수식은 원본 그대로지만 표시 형식은 검증하지 못했다."
                 % (type(exc).__name__, str(exc)[:60]))
        res.demote(UNVERIFIED)
    cellmap: dict = {}
    special: list = []
    sheets_meta: list = []
    out = [header_block(res, {"시트 수": len(wb.sheetnames)})]
    out.append("> 수식은 `=수식` 과 저장된 계산값을 함께 보존한다. 저장된 값이 최신이라는 보장은 없으며, "
               "mdmaker는 수식을 계산하지 않는다.")
    cells_done = 0
    truncated = False
    leftovers = _xl_merged_leftovers(path, limits)

    for name in wb.sheetnames:
        ws, wsv = wb[name], wbv[name]
        if ws.__class__.__name__ != "Worksheet":  # 차트시트 등
            out.append("\n## 시트: %s\n\n> 워크시트가 아닌 시트(%s)라 셀 내용이 없다."
                       % (esc_inline(name), ws.__class__.__name__))
            res.warn("워크시트가 아닌 시트는 구조만 기록했다: %s" % name)
            res.demote(UNVERIFIED)
            sheets_meta.append({"name": name, "type": ws.__class__.__name__})
            continue
        merged = [str(r) for r in ws.merged_cells.ranges]
        hidden_rows = [i for i, d in ws.row_dimensions.items() if d.hidden]
        hidden_cols = [k for k, d in ws.column_dimensions.items() if d.hidden]
        head = ["", "## 시트: %s%s" % (esc_inline(name),
                                      " (숨김: %s)" % ws.sheet_state if ws.sheet_state != "visible" else ""),
                "", "- 범위: `%s` (%d행 × %d열)" % (ws.dimensions or "", ws.max_row, ws.max_column)]
        if merged:
            head.append("- 병합 셀: %s" % ", ".join("`%s`" % m for m in merged))
        if hidden_rows:
            head.append("- 숨김 행: %s" % ", ".join(str(r) for r in hidden_rows))
        if hidden_cols:
            head.append("- 숨김 열: %s" % ", ".join(hidden_cols))
        out.extend(head)
        sheets_meta.append({"name": name, "state": ws.sheet_state, "dimensions": ws.dimensions,
                            "max_row": ws.max_row, "max_col": ws.max_column, "merged": merged,
                            "hidden_rows": hidden_rows, "hidden_cols": hidden_cols})

        rows_txt, empty_styled = [], []
        for row in ws.iter_rows():
            vrow = wsv[row[0].row] if row else ()
            line = []
            for idx, c in enumerate(row):
                cells_done += 1
                if cells_done > limits.max_cells:
                    truncated = True
                    break
                cached = vrow[idx].value if idx < len(vrow) else None
                f = None
                if isinstance(c.value, str) and c.value.startswith("="):
                    f = c.value
                elif c.data_type == "f":
                    f = str(getattr(c.value, "text", c.value))
                if f is not None:
                    shown = "%s → %s" % (f, _xl_value(cached) if cached is not None
                                         else "[원본에 계산값 없음]")
                    raw = cached
                else:
                    raw = c.value
                    if raw is None and c.data_type in ("s", "str", "inlineStr"):
                        raw = ""      # 저장된 빈 문자열: 빈 셀과 구별한다
                    shown = _xl_value(raw)
                nf = c.number_format
                if shown == "" and f is None:
                    if nf not in ("General", None) or c.has_style:
                        empty_styled.append(c.coordinate)
                    line.append("")
                    continue
                cellmap[(name, c.coordinate)] = {"raw": raw, "formula": f, "rendered": shown,
                                                 "nf": nf, "type": c.data_type}
                if f is not None or nf not in ("General", None) or isinstance(
                        raw, (_dt.datetime, _dt.date, _dt.time)) or (
                        isinstance(raw, str) and raw.startswith("#")):
                    special.append({"sheet": name, "coord": c.coordinate, "formula": f,
                                    "number_format": nf, "type": c.data_type,
                                    "raw": _xl_value(raw),
                                    "serial": _xl_serial(raw)})
                line.append(shown)
                if c.comment is not None:
                    note = c.comment.text
                    cellmap[(name, c.coordinate + "#comment")] = {
                        "raw": note, "formula": None, "rendered": note, "nf": None, "type": "cm"}
                    line[-1] = line[-1] + " [메모: %s]" % note
            if truncated:
                break
            rows_txt.append((row[0].row if row else 0, line))
        width = ws.max_column
        mode = opts.xlsx_table
        if mode == "auto":
            mode = "tsv" if (width > 12 or len(rows_txt) > 200) else "md"
        out.append("")
        if mode == "md":
            out.append("| 행 | " + " | ".join(get_column_letter(i + 1) for i in range(width)) + " |")
            out.append("| --- | " + " | ".join("---" for _ in range(width)) + " |")
            for rn, line in rows_txt:
                cells = [esc_cell(x) for x in line] + [""] * (width - len(line))
                out.append("| %d | %s |" % (rn, " | ".join(cells)))
        else:
            out.append("원본 행 번호를 첫 열에 둔다. 셀 안의 탭·줄바꿈은 `\\t` `\\n` 으로 표기했다.")
            out.append("")
            out.append("```tsv")
            out.append("행\t" + "\t".join(get_column_letter(i + 1) for i in range(width)))
            for rn, line in rows_txt:
                cells = [_xl_esc_tsv(x) for x in line] + [""] * (width - len(line))
                out.append("%d\t%s" % (rn, "\t".join(cells)))
            out.append("```")
        for coord, val in sorted(leftovers.get(name, {}).items()):
            cellmap[(name, coord)] = {"raw": val, "formula": None, "rendered": val,
                                      "nf": None, "type": "merged-hidden"}
        if leftovers.get(name):
            out.append("")
            out.append("> 병합 범위 안(좌상단이 아닌 칸)에 값이 남아 있다. 화면에는 보이지 않지만 "
                       "파일에 저장된 값이라 그대로 옮긴다:")
            for coord, val in sorted(leftovers[name].items()):
                out.append("> - `%s`: %s" % (coord, esc_cell(val)))
            sheets_meta[-1]["merged_hidden"] = sorted(leftovers[name])
        if empty_styled:
            out.append("")
            out.append("> 값 없이 서식만 저장된 셀: %s%s"
                       % (", ".join("`%s`" % c for c in empty_styled[:200]),
                          " 외 %d개" % (len(empty_styled) - 200) if len(empty_styled) > 200 else ""))
            sheets_meta[-1]["empty_styled"] = empty_styled

    if truncated:
        res.demote(PARTIAL)
        res.warn("셀 한도(%d)에 도달해 이후 셀을 변환하지 못했다" % limits.max_cells)
        out.append("\n> **셀 한도 초과로 이 결과는 불완전하다.**")

    # 워크북 밖의 자산·외부 링크도 원본 그대로 보존한다(실행·조회하지 않는다).
    zf = open_zip(path, limits)
    try:
        media = [n for n in zf.namelist() if n.startswith("xl/media/")]
        for n in media:
            res.assets["media/" + Path(n).name] = zf.read(n)
        drawings = [n for n in zf.namelist() if n.startswith("xl/drawings/") and n.endswith(".xml")]
        for n in drawings:
            res.assets["structure/" + Path(n).name] = zf.read(n)
        ext = [n for n in zf.namelist() if n.startswith("xl/externalLinks/")]
        for n in ext:
            res.assets["structure/external/" + Path(n).name] = zf.read(n)
        if media or drawings:
            out.append("\n> 이미지·도형 %d개와 배치 정보를 보조 폴더에 원본 그대로 보존했다. "
                       "Markdown 본문 안의 위치까지는 검증하지 못했다." % len(media))
            res.warn("이미지·도형의 셀 배치는 자동 검증하지 못했다")
            res.demote(UNVERIFIED)
        if ext:
            out.append("\n> 외부 링크 정의가 있다. 값을 가져오지 않고 정의만 보존했다.")
            res.warn("외부 링크 정의가 있다. 저장된 값만 사용했다")
            res.demote(UNVERIFIED)
    finally:
        zf.close()

    res.markdown = "\n".join(out) + "\n"
    res.meta = {"format": "xlsx", "sheets": sheets_meta, "cells": special,
                "note": "raw는 저장된 값, serial은 엑셀 내부 숫자 표현이다. 표시 형식(number_format)은 "
                        "문자열로 보존하며 mdmaker가 화면 표시를 재현하지는 않는다."}
    res.info["chars"] = len(res.markdown)
    verify_xlsx(res, path, cellmap, limits)      # 검사는 언제나 원본 파일로 한다
    if repaired is not None:
        shutil.rmtree(repaired.parent, ignore_errors=True)
    return res


_A1_REF = re.compile(r"\$?[A-Za-z]{1,3}\$?[0-9]{1,7}")


def formula_shape(f: str) -> str:
    """셀 주소만 지운 수식의 뼈대. 공유 수식은 파일에 본문이 한 번만 저장되므로
    칸마다 주소가 달라도 뼈대는 같아야 한다."""
    return _A1_REF.sub("@", (f or "").lstrip("=").upper()).replace(" ", "")


def _si_text(si) -> str:
    """sharedStrings 항목의 본문만 모은다. rPh(후리가나)는 본문이 아니다."""
    out = []
    for el in si:
        tag = _tag(el)
        if tag == "t":
            out.append(el.text or "")
        elif tag == "r":
            out.extend(t.text or "" for t in el.findall(q("x", "t")))
    return "".join(out)


def verify_xlsx(res: Res, path: Path, cellmap: dict, limits: Limits) -> None:
    """openpyxl과 무관하게 xlsx 내부 XML을 직접 읽어 셀 목록을 만들고 대조한다."""
    zf = open_zip(path, limits)
    try:
        wbx = parse_xml(zf.read("xl/workbook.xml"), limits)
        rels = _rels(zf, limits, "xl/workbook.xml")
        sheets = []
        for sh in wbx.iter(q("x", "sheet")):
            rid = sh.get(q("r", "id"))
            target = rels.get(rid, ("", "", ""))[0]
            # openpyxl은 절대형(/xl/...), Excel은 상대형(worksheets/...)으로 쓴다
            target = target[1:] if target.startswith("/") else "xl/" + target.lstrip("./")
            sheets.append((sh.get("name"), os.path.normpath(target).replace(os.sep, "/")))
        shared = []
        try:
            ss = parse_xml(zf.read("xl/sharedStrings.xml"), limits)
            for si in ss.findall(q("x", "si")):
                shared.append(_si_text(si))
        except (KeyError, ET.ParseError):
            pass

        norm = normalize_md(res.markdown)
        total, missing, mismatch, shared_ok = 0, [], [], 0
        for name, target in sheets:
            try:
                sx = parse_xml(zf.read(target), limits)
            except (KeyError, ET.ParseError):
                res.warn("시트 XML을 찾지 못했다: %s" % target)
                res.demote(UNVERIFIED)
                continue
            shared_f: dict = {}                 # 공유 수식 본문은 대표 셀에만 있다
            for c in sx.iter(q("x", "c")):
                fe = c.find(q("x", "f"))
                if fe is not None and fe.get("t") == "shared" and fe.get("si") and (fe.text or ""):
                    shared_f[fe.get("si")] = fe.text
            for c in sx.iter(q("x", "c")):
                coord = c.get("r")
                t = c.get("t")
                fe = c.find(q("x", "f"))
                ve = c.find(q("x", "v"))
                ise = c.find(q("x", "is"))
                if fe is None and ve is None and ise is None:
                    continue
                total += 1
                got = cellmap.get((name, coord))
                if got is None:
                    missing.append((name, coord))
                    continue
                if fe is not None:
                    ftext = fe.text or ""
                    if not ftext and fe.get("t") == "shared":
                        master = shared_f.get(fe.get("si"))
                        if master is None:
                            res.demote(UNVERIFIED)
                            res.warn("공유 수식의 대표 셀을 찾지 못했다: %s!%s" % (name, coord))
                        elif not (got["formula"] or "").strip():
                            mismatch.append((name, coord, "shared-formula-missing"))
                        elif formula_shape(got["formula"]) != formula_shape(master):
                            mismatch.append((name, coord, "shared-formula-shape"))
                        else:
                            shared_ok += 1
                    elif not ftext:
                        res.demote(UNVERIFIED)
                        res.warn("본문 없는 수식을 대조하지 못했다: %s!%s" % (name, coord))
                    elif (got["formula"] or "").lstrip("=") != ftext:
                        mismatch.append((name, coord, "formula"))
                    continue  # 캐시 값은 아래 표기 검사로 갈음한다
                raw = got["raw"]
                if ise is not None or t == "inlineStr":
                    want = "".join(x.text or "" for x in (ise.iter(q("x", "t")) if ise is not None else []))
                    if str(raw) != want:
                        mismatch.append((name, coord, "inline"))
                elif t == "s":
                    want = shared[int(ve.text)] if ve is not None else ""
                    if str(raw) != want:
                        mismatch.append((name, coord, "shared:%r!=%r" % (raw, want)))
                elif t == "b":
                    if bool(raw) != (ve is not None and ve.text == "1"):
                        mismatch.append((name, coord, "bool"))
                elif t == "e":
                    if str(raw) != (ve.text or ""):
                        mismatch.append((name, coord, "error"))
                elif t == "str":
                    if str(raw) != (ve.text or ""):
                        mismatch.append((name, coord, "str"))
                else:
                    want = float(ve.text)
                    serial = _xl_serial(raw)
                    try:
                        have = float(raw) if not isinstance(raw, bool) else None
                    except (TypeError, ValueError):
                        have = None
                    if serial is not None and abs(serial - want) < 1e-9:
                        pass
                    elif have is not None and abs(have - want) < 1e-9:
                        pass
                    else:
                        mismatch.append((name, coord, "number:%r!=%s" % (raw, ve.text)))
                shown = got["rendered"]
                # tsv 모드는 \t \n 로 표기하므로 원본 Markdown에서도 찾아본다
                if shown.strip() and not present(shown, res.markdown, norm) \
                        and _xl_esc_tsv(shown) not in res.markdown:
                    missing.append((name, coord + "(표기)"))
        res.checks.append({"item": "셀", "count": total, "missing": len(missing),
                           "mismatch": len(mismatch), "shared_formula": shared_ok,
                           "ok": not missing and not mismatch})
        if missing or mismatch:
            res.demote(PARTIAL)
            res.warn("셀 대조 실패 누락 %d개 %r / 불일치 %d개 %r"
                     % (len(missing), missing[:5], len(mismatch), mismatch[:5]))
    finally:
        zf.close()


# ---------------------------------------------------------------- HWPX
# 네임스페이스 URI는 한글 버전에 따라 달라질 수 있어 지역 이름(local name)으로만 판단한다.
HWPX_SKIP = {"linesegarray", "charPr", "paraPr", "ctrlheader", "parameters", "shapeobject",
             "imgRect", "imgDim", "imgClip", "inMargin", "outMargin", "sz", "pos", "offset",
             "orgSz", "curSz", "flip", "rotationInfo", "renderingInfo", "lineShape",
             "fillBrush", "drawText", "shadow", "effects", "cellSz", "cellMargin", "borderFillIDRef"}


class _HwpxCtx:
    def __init__(self, res, zf, limits):
        self.res, self.zf, self.limits = res, zf, limits
        self.manifest: dict = {}      # 항목 id -> zip 경로
        self.style_name: dict = {}    # 스타일 id -> 이름
        self.para_head: dict = {}     # paraPr id -> (종류, 수준)
        self.notes: list = []         # (표시, 본문)
        self.images: list = []
        self.equations = 0


def _hx_text(el, ctx) -> str:
    """hp:t 안의 글자와 탭·줄바꿈·전각공백 같은 명시적 기호를 구분해 읽는다."""
    parts = [esc_inline(el.text or "")]
    for ch in el:
        t = _tag(ch)
        if t == "tab":
            parts.append("\t")
        elif t in ("lineBreak", "lineBreakCell"):
            parts.append("\n")
        elif t == "nbSpace":
            parts.append(" ")
        elif t == "fwSpace":
            parts.append("　")
        elif t == "hyphen":
            parts.append("-")
        elif t in ("markpenBegin", "markpenEnd", "titleMark", "insertBegin", "insertEnd",
                   "deleteBegin", "deleteEnd"):
            pass
        else:
            parts.append(_hx_text(ch, ctx))
        parts.append(esc_inline(ch.tail or ""))
    return "".join(parts)


def _hx_binary(ctx, item_id: str, kind: str) -> str | None:
    href = ctx.manifest.get(item_id)
    if not href:
        ctx.res.warn("자산 항목을 찾지 못했다: %s" % item_id)
        ctx.res.demote(PARTIAL)
        return None
    try:
        data = ctx.zf.read(href)
    except KeyError:
        ctx.res.warn("자산 파일이 없다: %s" % href)
        ctx.res.demote(PARTIAL)
        return None
    name = "media/" + Path(href).name
    ctx.res.assets[name] = data
    ctx.images.append({"part": href, "asset": name, "sha256": sha256(data),
                       "bytes": len(data), "kind": kind})
    return name


def _hx_inline(node, ctx, blocks: list) -> str:
    out = []
    for el in node:
        t = _tag(el)
        if t == "t":
            out.append(_hx_text(el, ctx))
        elif t == "tbl":
            blocks.append(_hx_table(el, ctx))
        elif t in ("footNote", "endNote"):
            mark = "hfn%d" % (len(ctx.notes) + 1)
            body = "\n\n".join(_hx_paras(el, ctx))
            ctx.notes.append(("[^%s]" % mark, body))
            out.append("[^%s]" % mark)
        elif t == "equation":
            ctx.equations += 1
            script = ""
            for s in el.iter():
                if _tag(s) == "script":
                    script = "".join(s.itertext())
                    break
            out.append("`[수식] %s`" % esc_inline(script))
        elif t == "fieldBegin":
            cmd = ""
            for s in el.iter():
                if _tag(s) in ("stringParam", "parameter") and (s.text or "").strip():
                    cmd = s.text.strip()
                    break
            typ = el.get("type", "")
            if cmd:
                out.append("[%s: %s]" % (typ or "필드", esc_inline(cmd)))
            inner = [b for sub in el if _tag(sub) in ("subList", "p") for b in _hx_paras([sub], ctx)]
            if inner:
                out.append("<br>".join(inner))
        elif t in ("pic", "ole", "container", "rect", "ellipse", "line", "arc", "polygon",
                   "curve", "connectLine", "textart", "chart", "video"):
            out.append(_hx_shape(el, ctx, blocks))
        elif t in HWPX_SKIP:
            continue
        else:
            out.append(_hx_inline(el, ctx, blocks))
    return "".join(out)


def _hx_shape(el, ctx, blocks: list) -> str:
    """그림·도형·글상자. 안에 글자가 있으면 그대로, 이미지는 원본 바이트로 보존한다."""
    parts = []
    for sub in el.iter():
        if _tag(sub) == "img" and sub.get("binaryItemIDRef"):
            name = _hx_binary(ctx, sub.get("binaryItemIDRef"), _tag(el))
            if name:
                alt = el.get("alt") or el.get("desc") or ""
                parts.append("![%s: %s](%s)" % (_tag(el), esc_inline(alt) or "대체 텍스트 없음", name))
                if not alt:
                    parts.append("[이미지: 내용 해석 안 됨]")
    inner_blocks: list = []
    text = []
    for sub in el:
        if _tag(sub) in ("drawText", "subList", "textBox"):
            text.extend(_hx_paras(sub, ctx))
    if text:
        parts.append("[도형 글상자: %s]" % "<br>".join(text))
    blocks.extend(inner_blocks)
    if not parts:
        ctx.res.warn("내용을 확인하지 못한 개체: <%s>" % _tag(el))
        ctx.res.demote(UNVERIFIED)
        parts.append("[개체: %s]" % _tag(el))
    return "".join(parts)


def _hx_paras(node, ctx) -> list[str]:
    """문단 컨테이너(sec, subList, footNote ...)를 블록 목록으로 바꾼다."""
    out = []
    for el in node:
        t = _tag(el)
        if t == "p":
            out.extend(_hx_para(el, ctx))
        elif t in ("subList", "sec"):
            out.extend(_hx_paras(el, ctx))
        elif t in HWPX_SKIP:
            continue
        elif t == "tbl":
            out.append(_hx_table(el, ctx))
    return [b for b in out if b != ""]


def _hx_para(p, ctx) -> list[str]:
    blocks: list = []
    text = _hx_inline(p, ctx, blocks)
    kind, level = ctx.para_head.get(p.get("paraPrIDRef"), (None, 0))
    if kind is None:
        name = ctx.style_name.get(p.get("styleIDRef"), "")
        m = re.match(r"^(?:개요|제목)\s*([1-9])$", name.strip())
        if m:
            kind, level = "OUTLINE", int(m.group(1)) - 1
    if text.strip():
        if kind == "OUTLINE":
            head = "#" * min(level + 1, 6) + " " + text
        elif kind == "NUMBER":
            head = "  " * level + "1. " + text
        elif kind == "BULLET":
            head = "  " * level + "- " + text
        else:
            head = esc_block(text)
        return [head] + blocks
    return ([esc_block(text)] if text else []) + blocks


def _hx_table(tbl, ctx) -> str:
    rows: dict = {}
    complex_ = False
    for tr in tbl:
        if _tag(tr) != "tr":
            continue
        for tc in tr:
            if _tag(tc) != "tc":
                continue
            addr = next((c for c in tc if _tag(c) == "cellAddr"), None)
            span = next((c for c in tc if _tag(c) == "cellSpan"), None)
            r = int(addr.get("rowAddr", "0")) if addr is not None else len(rows)
            c = int(addr.get("colAddr", "0")) if addr is not None else 0
            cs = int(span.get("colSpan", "1")) if span is not None else 1
            rs = int(span.get("rowSpan", "1")) if span is not None else 1
            blocks = _hx_paras(tc, ctx)
            if cs > 1 or rs > 1 or len(blocks) > 1 or any("\n" in b for b in blocks):
                complex_ = True
            rows.setdefault(r, {})[c] = {"text": "\n".join(blocks), "cs": cs, "rs": rs}
    if not rows:
        return ""
    ncol = int(tbl.get("colCnt") or (max(max(r) for r in rows.values()) + 1))
    order = sorted(rows)
    if not complex_:
        out = ["> 원본에 헤더 행 표시가 없어 첫 행을 데이터로 둔다.", "",
               "| " + " | ".join("열%d" % (i + 1) for i in range(ncol)) + " |",
               "| " + " | ".join("---" for _ in range(ncol)) + " |"]
        for r in order:
            out.append("| " + " | ".join(cell_inline(rows[r].get(c, {"text": ""})["text"])
                                           for c in range(ncol)) + " |")
        return "\n".join(out)
    out = ["> 병합·중첩 때문에 Markdown 표 대신 행별 표현으로 변환했다. 좌표는 원본 기준이다.", ""]
    for r in order:
        for c in sorted(rows[r]):
            cell = rows[r][c]
            extra = []
            if cell["cs"] > 1:
                extra.append("가로병합 %d칸" % cell["cs"])
            if cell["rs"] > 1:
                extra.append("세로병합 %d칸" % cell["rs"])
            tag = " (%s)" % ", ".join(extra) if extra else ""
            out.append("- **R%dC%d**%s: %s" % (r + 1, c + 1, tag, cell_block(cell["text"])))
    return "\n".join(out)


def _hx_find_package(zf, limits) -> str:
    try:
        root = parse_xml(zf.read("META-INF/container.xml"), limits)
        for rf in root.iter():
            if _tag(rf) == "rootfile" and rf.get("full-path"):
                return rf.get("full-path")
    except (KeyError, ValueError, ET.ParseError):
        pass
    return "Contents/content.hpf"


def convert_hwpx(path: Path, opts, limits: Limits) -> Res:
    res = Res(path, "hwpx")
    zf = open_zip(path, limits)
    try:
        names = zf.namelist()
        if "mimetype" in names:
            mt = zf.read("mimetype").decode("utf-8", "replace").strip()
            res.info["mimetype"] = mt
            if "hwp" not in mt:
                res.warn("mimetype이 HWPX가 아니다: %r" % mt)
                res.demote(UNVERIFIED)
        ctx = _HwpxCtx(res, zf, limits)
        pkg_path = _hx_find_package(zf, limits)
        try:
            pkg = parse_xml(zf.read(pkg_path), limits)
        except (KeyError, ET.ParseError):
            res.status = FAILED
            res.warn("패키지 정의(%s)가 없다. HWPX가 아닐 수 있다." % pkg_path)
            return res
        base = str(Path(pkg_path).parent)
        for it in pkg.iter():
            if _tag(it) == "item" and it.get("id") and it.get("href"):
                href = it.get("href")
                cand = os.path.normpath(os.path.join(base, href)).replace(os.sep, "/")
                ctx.manifest[it.get("id")] = cand if cand in names else href
        spine = [x.get("idref") for x in pkg.iter() if _tag(x) == "itemref" and x.get("idref")]
        sections = [ctx.manifest[i] for i in spine
                    if i in ctx.manifest and re.search(r"section\d+\.xml$", ctx.manifest[i] or "")]
        if not sections:
            # spine이 없거나 비정상인 파일: 이름순 정렬만으로 순서를 정하면 틀릴 수 있어 경고한다.
            sections = sorted(n for n in names if re.search(r"Contents/section\d+\.xml$", n))
            if sections:
                res.warn("문서 순서 정보(spine)를 쓰지 못해 파일명 순서로 읽었다")
                res.demote(UNVERIFIED)
        if not sections:
            res.status = FAILED
            res.warn("본문 구역(section*.xml)을 찾지 못했다")
            return res
        res.info["sections"] = sections

        head_path = next((v for k, v in ctx.manifest.items()
                          if v and v.endswith("header.xml")), "Contents/header.xml")
        head = _part(zf, limits, head_path)
        if head is not None:
            for el in head.iter():
                t = _tag(el)
                if t == "style" and el.get("id") is not None:
                    ctx.style_name[el.get("id")] = el.get("name") or el.get("engName") or ""
                elif t == "paraPr" and el.get("id") is not None:
                    h = next((x for x in el if _tag(x) == "heading"), None)
                    if h is not None and h.get("type", "NONE") != "NONE":
                        ctx.para_head[el.get("id")] = (h.get("type"), int(h.get("level", "0")))

        out = [header_block(res, {"구역 수": len(sections)})]
        for sec_path in sections:
            sec = _part(zf, limits, sec_path)
            if sec is None:
                res.warn("구역을 읽지 못했다: %s" % sec_path)
                res.demote(PARTIAL)
                continue
            out.append("")
            out.append("<!-- 구역: %s -->" % sec_path)
            out.extend(_hx_paras(sec, ctx))
        if ctx.notes:
            out.append("")
            out.append("## 각주·미주")
            out.append("")
            out.extend("%s: %s" % (mark, body) for mark, body in ctx.notes)

        res.markdown = "\n\n".join(x for x in out if x != "") + "\n"
        res.meta = {"format": "hwpx", "sections": sections, "images": ctx.images,
                    "equations": ctx.equations,
                    "note": "Preview/PrvText.txt는 미리보기이므로 본문으로 쓰지 않는다."}
        res.info["chars"] = len(res.markdown)
        verify_hwpx(res, zf, limits)
        return res
    finally:
        zf.close()


_HX_T = re.compile(rb"<[a-zA-Z0-9]*:?t(?:\s[^>]*)?>(.*?)</[a-zA-Z0-9]*:?t>", re.S)
_HX_SCRIPT = re.compile(rb"<[a-zA-Z0-9]*:?script(?:\s[^>]*)?>(.*?)</[a-zA-Z0-9]*:?script>", re.S)
_HX_INNER_TAG = re.compile(rb"<[^>]*>")


def verify_hwpx(res: Res, zf, limits: Limits) -> None:
    """변환기와 다른 경로(원본 XML 직접 스캔)로 본문 요소를 세어 결과와 대조한다.
    Preview/는 미리보기라 본문 대조 대상이 아니다."""
    norm = normalize_md(res.markdown)
    parts = [n for n in zf.namelist()
             if n.endswith(".xml") and not n.startswith("Preview/")
             and re.search(r"section\d+\.xml$", n)]
    missing, total, eq = [], 0, 0
    for part in sorted(parts):
        data = zf.read(part)
        if len(data) > limits.max_xml_bytes:
            res.warn("%s 가 XML 한도를 넘어 검사하지 못했다" % part)
            res.demote(UNVERIFIED)
            continue
        for m in _HX_T.finditer(data):
            for seg in _HX_INNER_TAG.split(m.group(1)):
                txt = _unent(seg.decode("utf-8"))
                if not txt.strip():
                    continue
                total += 1
                if not present(txt, res.markdown, norm):
                    missing.append((part, txt[:60]))
        for m in _HX_SCRIPT.finditer(data):
            txt = _unent(_HX_INNER_TAG.sub(b"", m.group(1)).decode("utf-8"))
            if txt.strip():
                eq += 1
                if not present(txt, res.markdown, norm):
                    missing.append((part, "수식:" + txt[:40]))
    res.checks.append({"item": "텍스트 노드", "count": total, "missing": len(missing),
                       "ok": not missing})
    res.checks.append({"item": "수식", "count": eq, "ok": True})
    if missing:
        res.demote(PARTIAL)
        res.warn("원본 텍스트 %d개가 결과에서 확인되지 않았다: %r" % (len(missing), missing[:5]))

    bins = [n for n in zf.namelist() if n.startswith("BinData/") and not n.endswith("/")]
    saved = {i["part"] for i in res.meta.get("images", [])}
    for n in bins:
        if n not in saved:
            res.assets.setdefault("media/" + Path(n).name, zf.read(n))
            res.warn("본문에서 참조를 확인하지 못한 자산을 그대로 보존했다: %s" % n)
            res.demote(UNVERIFIED)
    res.checks.append({"item": "자산", "count": len(bins), "saved": len(res.assets),
                       "ok": len(res.assets) >= len(bins)})


# ---------------------------------------------------------------- HWP 5.x
# 본문 글자를 담지 않는 배치·번호 컨트롤. 내용이 없으므로 표시하지 않는다.
CTRL_IGNORE = {"secd", "cold", "bokm", "idxm", "tcps", "tdut", "atno", "nwno",
               "pgct", "pghd", "pgnp", "pgnm"}


class _HwpCtx:
    def __init__(self, res, hdr, bins, style_names, streams):
        self.res, self.hdr, self.bins = res, hdr, bins
        self.styles, self.streams = style_names, streams
        self.notes: list = []
        self.tails: list = []          # 머리말·꼬리말처럼 본문 뒤에 붙이는 덩어리
        self.images: list = []


def _hp_heading(ctx, payload: bytes) -> int:
    if len(payload) < 11:
        return 0
    name = ctx.styles[payload[10]] if payload[10] < len(ctx.styles) else ""
    m = re.match(r"^(?:개요|제목)\s*([1-9])$", name.strip())
    return int(m.group(1)) if m else 0


def _hp_asset(ctx, bin_index: int) -> str | None:
    """그림이 참조하는 BinItem 색인(1-based)을 BinData 스트림으로 잇는다."""
    if not (1 <= bin_index <= len(ctx.bins)):
        ctx.res.warn("그림이 참조하는 BinItem 색인이 범위를 벗어난다: %r" % bin_index)
        ctx.res.demote(PARTIAL)
        return None
    item = ctx.bins[bin_index - 1]
    if item.get("type") == 0:
        ctx.res.warn("문서 밖 그림 링크는 가져오지 않는다: %s" % item.get("rel_path", ""))
        ctx.res.demote(UNVERIFIED)
        return None
    stream = "BinData/BIN%04X.%s" % (item.get("id", 0), item.get("ext", "bin"))
    raw = ctx.streams.get(stream)
    if raw is None:
        cand = [k for k in ctx.streams if k.startswith("BinData/BIN%04X." % item.get("id", 0))]
        raw = ctx.streams[cand[0]] if cand else None
        stream = cand[0] if cand else stream
    if raw is None:
        ctx.res.warn("BinData 스트림을 찾지 못했다: %s" % stream)
        ctx.res.demote(PARTIAL)
        return None
    try:
        data = hwp5.decompress(raw, ctx.hdr.compressed)
    except hwp5.HwpError:
        data = raw                     # 항목별로 압축 여부가 다를 수 있다
    name = "media/" + Path(stream).name
    ctx.res.assets[name] = data
    ctx.images.append({"part": stream, "asset": name, "sha256": sha256(data), "bytes": len(data)})
    return name


def _hp_lists(node, deep: bool = False) -> list:
    """CTRL_HEADER 아래의 LIST_HEADER 단위로 문단을 묶는다.

    HWP는 문단을 LIST_HEADER의 자식으로 넣지 않고 "같은 수준의 다음 형제"로 이어
    붙인다(실제 문서에서 확인). 자식만 보면 표 셀이 통째로 빈다."""
    out: list = []
    cur = None

    def walk(n):
        nonlocal cur
        for c in n["children"]:
            if c["tag"] == hwp5.TAG_LIST_HEADER:
                cur = {"header": c, "paras": []}
                out.append(cur)
            elif c["tag"] == hwp5.TAG_PARA_HEADER and cur is not None:
                cur["paras"].append(c)
            elif deep:
                walk(c)      # 글상자 문단은 SHAPE_COMPONENT 아래로 더 들어가 있다

    walk(node)
    return out


def _hp_list_blocks(node, ctx) -> list:
    out = []
    for grp in _hp_lists(node, deep=True):
        out.extend(_hp_paras(grp["paras"], ctx))
    return out


def _hp_find_all(node, tags) -> list:
    out = []
    for c in node["children"]:
        if c["tag"] in tags:
            out.append(c)
        out.extend(_hp_find_all(c, tags))
    return out


def _hp_find(node, tag):
    for c in node["children"]:
        if c["tag"] == tag:
            return c
        got = _hp_find(c, tag)
        if got is not None:
            return got
    return None


def _hp_table(node, ctx) -> str:
    tbl = _hp_find(node, hwp5.TAG_TABLE)
    nrow, ncol = hwp5.table_size(tbl["data"]) if tbl else (0, 0)
    cells, seq = {}, 0
    complex_ = False
    captions: list = []
    for grp in _hp_lists(node):
        addr = hwp5.cell_addr(grp["header"]["data"])
        if addr is None:                   # 표 설명(캡션) 등은 셀이 아니다
            captions.extend(_hp_paras(grp["paras"], ctx))
            continue
        seq += 1
        blocks = _hp_paras(grp["paras"], ctx)
        if addr["colspan"] > 1 or addr["rowspan"] > 1 or len(blocks) > 1 or any(
                "\n" in b for b in blocks):
            complex_ = True
        cells.setdefault(addr["row"], {})[addr["col"]] = {"text": "\n".join(blocks), **addr}
    if not cells:
        return "\n\n".join(captions)
    ncol = ncol or max(max(r) for r in cells.values()) + 1
    order = sorted(cells)
    if not complex_:
        out = ["> 원본에 헤더 행 표시가 없어 첫 행을 데이터로 둔다.", "",
               "| " + " | ".join("열%d" % (i + 1) for i in range(ncol)) + " |",
               "| " + " | ".join("---" for _ in range(ncol)) + " |"]
        for r in order:
            out.append("| " + " | ".join(cell_inline(cells[r].get(c, {"text": ""})["text"])
                                           for c in range(ncol)) + " |")
        if captions:
            out += ["", "표 설명: " + " ".join(captions)]
        return "\n".join(out)
    out = ["> 병합·중첩 때문에 Markdown 표 대신 행별 표현으로 변환했다. 좌표는 원본 기준이다.", ""]
    for r in order:
        for c in sorted(cells[r]):
            cell = cells[r][c]
            extra = []
            if cell["colspan"] > 1:
                extra.append("가로병합 %d칸" % cell["colspan"])
            if cell["rowspan"] > 1:
                extra.append("세로병합 %d칸" % cell["rowspan"])
            tag = " (%s)" % ", ".join(extra) if extra else ""
            out.append("- **R%dC%d**%s: %s" % (r + 1, c + 1, tag, cell_block(cell["text"])))
    return "\n".join(out + (["", "표 설명: " + " ".join(captions)] if captions else []))


def _hp_ctrl(node, ctx, blocks: list) -> str:
    cid = hwp5.ctrl_id(node["data"])
    if cid == "tbl ":
        blocks.append(_hp_table(node, ctx))
        return ""
    if cid in ("fn  ", "en  "):
        mark = "hfn%d" % (len(ctx.notes) + 1)
        body = "\n\n".join(_hp_list_blocks(node, ctx))
        ctx.notes.append(("[^%s]" % mark, body or "[각주 본문 없음]"))
        return "[^%s]" % mark
    if cid in ("head", "foot"):
        body = "\n\n".join(_hp_list_blocks(node, ctx))
        ctx.tails.append(("머리말" if cid == "head" else "꼬리말", body))
        return "[%s 위치]" % ("머리말" if cid == "head" else "꼬리말")
    if cid == "tcmt":                      # 메모: 본문 글자를 담고 있어 반드시 살린다
        body = "<br>".join(_hp_list_blocks(node, ctx))
        return "[메모: %s]" % body if body else ""
    if cid.startswith("%"):                # 필드(하이퍼링크·누름틀 등): 주소를 보존한다
        cmd = hwp5.field_command(node["data"])
        inner = "<br>".join(_hp_list_blocks(node, ctx))
        label = {"%hlk": "링크", "%fmu": "수식필드", "%unk": "필드"}.get(cid, "필드 " + cid)
        out = "[%s: %s]" % (label, esc_inline(cmd)) if cmd else ""
        return out + inner
    if cid in ("gso ", "eqed"):
        parts = []
        eq = _hp_find(node, hwp5.TAG_EQEDIT)
        if eq is not None:
            parts.append("`[수식] %s`" % esc_inline(hwp5.equation_script(eq["data"])))
        for pic in _hp_find_all(node, (hwp5.TAG_SHAPE_PICTURE, hwp5.TAG_SHAPE_OLE)):
            name = _hp_asset(ctx, hwp5.picture_bin_id(pic["data"], len(ctx.bins)))
            if name:
                kind = "그림" if pic["tag"] == hwp5.TAG_SHAPE_PICTURE else "삽입 개체"
                parts.append("![%s: 대체 텍스트 없음](%s)[이미지: 내용 해석 안 됨]" % (kind, name))
        inner = _hp_list_blocks(node, ctx)
        if inner:
            parts.append("[도형 글상자: %s]" % "<br>".join(inner))
        if not parts:
            ctx.res.warn("내용을 확인하지 못한 그리기 개체가 있다")
            ctx.res.demote(UNVERIFIED)
            parts.append("[개체: gso]")
        return "".join(parts)
    if cid in CTRL_IGNORE:
        return ""
    # 모르는 개체를 조용히 버리지 않는다. 안에 문단이 있으면 그것만이라도 살린다.
    inner = _hp_list_blocks(node, ctx)
    ctx.res.warn("해석하지 못한 컨트롤: %r" % cid)
    ctx.res.demote(UNVERIFIED)
    return "[개체: %s]%s" % (esc_inline(cid.strip() or "?"),
                             ("<br>".join(inner) if inner else ""))


def _hp_para(node, ctx) -> list[str]:
    text_rec = next((c for c in node["children"] if c["tag"] == hwp5.TAG_PARA_TEXT), None)
    ctrls = [c for c in node["children"] if c["tag"] == hwp5.TAG_CTRL_HEADER]
    blocks: list = []
    parts: list[str] = []
    ci = 0
    for kind, val in (hwp5.para_text(text_rec["data"]) if text_rec else []):
        if kind == "text":
            parts.append(esc_inline(val) if val not in ("\t", "\n") else val)
        elif kind == "ctrl":
            if ci < len(ctrls):
                parts.append(_hp_ctrl(ctrls[ci], ctx, blocks))
                ci += 1
            else:
                ctx.res.warn("문단의 컨트롤 수와 레코드 수가 맞지 않는다")
                ctx.res.demote(UNVERIFIED)
    if ci < len(ctrls):                 # 본문 글자에 대응이 없는 컨트롤도 버리지 않는다
        for extra in ctrls[ci:]:
            parts.append(_hp_ctrl(extra, ctx, blocks))
    # LIST_HEADER/MEMO_LIST가 컨트롤과 같은 수준에 붙는 경우, 그 문단들은 이 문단의
    # 자식으로 들어온다. 렌더링에서 빠뜨리면 메모 본문이 통째로 사라진다.
    attached = [c for c in node["children"] if c["tag"] == hwp5.TAG_PARA_HEADER]
    if attached:
        body = _hp_paras(attached, ctx)
        if any(c["tag"] == hwp5.TAG_MEMO_LIST for c in node["children"]):
            blocks.append("[메모: %s]" % "<br>".join(body))
        else:
            blocks.extend(body)
    text = "".join(parts)
    level = _hp_heading(ctx, node["data"])
    if text.strip():
        head = ("#" * min(level, 6) + " " + text) if level else esc_block(text)
        return [head] + [b for b in blocks if b]
    return ([text] if text else []) + [b for b in blocks if b]


def _hp_paras(nodes, ctx) -> list[str]:
    out = []
    for n in nodes:
        if n["tag"] == hwp5.TAG_PARA_HEADER:
            out.extend(_hp_para(n, ctx))
    return [b for b in out if b != ""]


def convert_hwp(path: Path, opts, limits: Limits) -> Res:
    res = Res(path, "hwp")
    try:
        cfb = hwp5.Cfb(path.read_bytes(), max_stream=limits.max_bytes)
        streams = cfb.streams()
    except hwp5.HwpError as exc:
        res.status = FAILED
        res.warn("HWP 5.x로 읽지 못했다: %s" % exc)
        return res
    if "FileHeader" not in streams:
        res.status = FAILED
        res.warn("FileHeader 스트림이 없다. HWP 5.x가 아니다.")
        return res
    hdr = hwp5.Header(streams["FileHeader"])
    res.info["hwp_version"] = ".".join(str(x) for x in hdr.version)
    res.info["converter"] += " + hwp5reader"
    if hdr.version[0] != 5:
        res.status = UNSUPPORTED
        res.warn("HWP %s는 지원 대상이 아니다 (5.x만 처리한다)"
                 % ".".join(str(x) for x in hdr.version))
        return res
    guard = hdr.protected
    if guard:
        res.status = UNSUPPORTED
        res.warn("보호된 문서다(%s). 보호 기능을 우회하지 않는다." % guard)
        return res

    try:
        docinfo = hwp5.decompress(streams.get("DocInfo", b""), hdr.compressed)
        ctx = _HwpCtx(res, hdr, hwp5.bin_items(docinfo), hwp5.styles(docinfo), streams)
        sec_names = sorted((k for k in streams if re.match(r"BodyText/Section\d+$", k)),
                           key=lambda k: int(k.rsplit("Section", 1)[1]))
        if not sec_names:
            res.status = FAILED
            res.warn("BodyText/Section 스트림이 없다")
            return res
        out = [header_block(res, {"HWP 버전": res.info["hwp_version"],
                                  "구역 수": len(sec_names)})]
        sections = {}
        for name in sec_names:
            body = hwp5.decompress(streams[name], hdr.compressed)
            sections[name] = body
            out.append("")
            out.append("<!-- 구역: %s -->" % name)
            out.extend(_hp_paras(hwp5.tree(hwp5.records(body)), ctx))
        for title, body in ctx.tails:
            out.append("")
            out.append("## %s" % title)
            out.append(body)
        if ctx.notes:
            out.append("")
            out.append("## 각주·미주")
            out.extend("%s: %s" % (m, b) for m, b in ctx.notes)
    except hwp5.HwpError as exc:
        res.status = FAILED
        res.warn("본문을 읽지 못했다: %s" % exc)
        return res

    res.markdown = "\n\n".join(x for x in out if x != "") + "\n"
    res.meta = {"format": "hwp", "version": res.info["hwp_version"],
                "compressed": hdr.compressed, "sections": sec_names,
                "images": ctx.images, "bin_items": ctx.bins,
                "note": "PrvText는 미리보기라 본문으로 쓰지 않는다. 문서 이력·스크립트는 변환 대상이 아니다."}
    if hdr.history or hdr.signed:
        res.warn("문서 이력 또는 서명 정보가 있다. 본문 변환 대상이 아니라 그대로 두었다.")
        res.demote(UNVERIFIED)
    res.info["chars"] = len(res.markdown)
    verify_hwp(res, sections, streams, hdr)
    return res


def verify_hwp(res: Res, sections: dict, streams: dict, hdr) -> None:
    """구조 순회와 별개로 PARA_TEXT 레코드만 평평하게 훑어 글자를 대조한다."""
    norm = normalize_md(res.markdown)
    missing, total = [], 0
    for name, body in sections.items():
        try:
            for tag, _lvl, payload in hwp5.records(body):
                if tag != hwp5.TAG_PARA_TEXT:
                    continue
                for kind, val in hwp5.para_text(payload):
                    if kind != "text" or not str(val).strip():
                        continue
                    total += 1
                    if not present(str(val), res.markdown, norm):
                        missing.append((name, str(val)[:60]))
        except hwp5.HwpError as exc:
            res.warn("%s 검사 중단: %s" % (name, exc))
            res.demote(UNVERIFIED)
    res.checks.append({"item": "텍스트 조각", "count": total, "missing": len(missing),
                       "ok": not missing})
    if missing:
        res.demote(PARTIAL)
        res.warn("원본 텍스트 %d개가 결과에서 확인되지 않았다: %r" % (len(missing), missing[:5]))

    bins = [k for k in streams if k.startswith("BinData/")]
    saved = {i["part"] for i in res.meta.get("images", [])}
    for k in bins:
        if k not in saved:
            try:
                data = hwp5.decompress(streams[k], hdr.compressed)
            except hwp5.HwpError:
                data = streams[k]
            res.assets.setdefault("media/" + Path(k).name, data)
            res.warn("본문에서 참조를 확인하지 못한 자산을 그대로 보존했다: %s" % k)
            res.demote(UNVERIFIED)
    res.checks.append({"item": "자산", "count": len(bins), "saved": len(res.assets),
                       "ok": len(res.assets) >= len(bins)})


# ---------------------------------------------------------------- PPTX
def _px_text(node, ctx, blocks: list) -> str:
    """도형 하나의 글자. 문단(a:p)은 줄로, 표(a:tbl)는 블록으로 뺀다."""
    out = []
    for el in node:
        t = _tag(el)
        if t == "t":
            out.append(esc_inline(el.text or ""))
        elif t == "br":
            out.append("\n")
        elif t == "tab":
            out.append("\t")
        elif t == "fld":                      # 슬라이드 번호·날짜 같은 필드
            inner = _px_text(el, ctx, blocks)
            if inner:
                out.append(inner)
        elif t == "p":
            inner = _px_text(el, ctx, blocks)
            out.append(inner + "\n")
        elif t in ("tbl",):
            blocks.append(_px_table(el, ctx))
        elif t in ("rPr", "pPr", "endParaRPr", "defRPr", "lstStyle"):
            continue
        else:
            out.append(_px_text(el, ctx, blocks))
    return "".join(out)


def _px_table(tbl, ctx) -> str:
    rows = []
    complex_ = False
    for tr in tbl:
        if _tag(tr) != "tr":
            continue
        row = []
        for tc in tr:
            if _tag(tc) != "tc":
                continue
            blocks: list = []
            txt = _px_text(tc, ctx, blocks).rstrip("\n")
            span = int(tc.get("gridSpan", "1")), int(tc.get("rowSpan", "1"))
            if tc.get("hMerge") or tc.get("vMerge") or span != (1, 1) or "\n" in txt or blocks:
                complex_ = True
            row.append({"text": txt, "cs": span[0], "rs": span[1]})
        rows.append(row)
    if not rows:
        return ""
    if not complex_:
        width = max(len(r) for r in rows)
        out = ["> 원본에 헤더 행 표시가 없어 첫 행을 데이터로 둔다.", "",
               "| " + " | ".join("열%d" % (i + 1) for i in range(width)) + " |",
               "| " + " | ".join("---" for _ in range(width)) + " |"]
        for r in rows:
            out.append("| " + " | ".join([cell_inline(c["text"]) for c in r]
                                         + [""] * (width - len(r))) + " |")
        return "\n".join(out)
    out = ["> 병합·여러 문단 때문에 Markdown 표 대신 행별 표현으로 변환했다. 좌표는 원본 기준이다.", ""]
    for ri, r in enumerate(rows, 1):
        for ci, c in enumerate(r, 1):
            extra = []
            if c["cs"] > 1:
                extra.append("가로병합 %d칸" % c["cs"])
            if c["rs"] > 1:
                extra.append("세로병합 %d칸" % c["rs"])
            tag = " (%s)" % ", ".join(extra) if extra else ""
            out.append("- **R%dC%d**%s: %s" % (ri, ci, tag, cell_block(c["text"])))
    return "\n".join(out)


def _px_shape(sp, ctx, order: list) -> None:
    """spTree를 문서 순서대로 훑는다. 그룹은 안으로 들어간다."""
    t = _tag(sp)
    if t == "grpSp":
        for child in sp:
            _px_shape(child, ctx, order)
        return
    name = ""
    descr = ""
    for nv in sp.iter():
        if _tag(nv) in ("cNvPr",):
            name = nv.get("name") or ""
            descr = nv.get("descr") or ""
            break
    pos = None
    for off in sp.iter():
        if _tag(off) == "off":
            pos = (off.get("x"), off.get("y"))
            break
    blocks: list = []
    txt = ""
    body = next((x for x in sp.iter() if _tag(x) == "txBody"), None)
    if body is not None:
        txt = _px_text(body, ctx, blocks).rstrip("\n")
    if t == "graphicFrame":
        for tb in sp.iter():
            if _tag(tb) == "tbl":
                blocks.append(_px_table(tb, ctx))
                break
        else:
            chart = _px_chart(sp, ctx)
            if chart:
                blocks.append(chart)
            elif not txt:
                ctx.res.warn("표·글자가 없는 개체(차트·다이어그램 등)는 원본 데이터만 보존했다: %s"
                             % (name or "이름 없음"))
                ctx.res.demote(UNVERIFIED)
    img = None
    for blip in sp.iter():
        if _tag(blip) == "blip" and (blip.get(q("r", "embed")) or blip.get(q("r", "link"))):
            img = blip.get(q("r", "embed")) or blip.get(q("r", "link"))
            break
    if img:
        asset = _px_asset(ctx, img)
        if asset:
            order.append("![그림: %s](%s)%s"
                         % (esc_inline(descr or name) or "대체 텍스트 없음", asset,
                            "" if descr else "[이미지: 내용 해석 안 됨]"))
    if txt.strip():
        order.append(esc_block(txt))
    order.extend(b for b in blocks if b)
    if pos:
        ctx.shapes.append({"name": name, "descr": descr, "x": pos[0], "y": pos[1],
                           "kind": t, "chars": len(txt)})


def _px_cache(node) -> list[str]:
    """c:strCache / c:numCache 의 <c:pt> 값을 idx 순서대로."""
    pts = {}
    for pt in node.iter():
        if _tag(pt) == "pt" and pt.get("idx") is not None:
            v = next((x.text for x in pt if _tag(x) == "v"), "")
            pts[int(pt.get("idx"))] = v or ""
    return [pts.get(i, "") for i in range(max(pts) + 1)] if pts else []


def _px_chart(sp, ctx) -> str:
    """차트에 저장된 수치를 표로 옮긴다. 원본 xml도 보조 파일로 남긴다."""
    rid = None
    for el in sp.iter():
        if _tag(el) == "chart" and el.get(q("r", "id")):
            rid = el.get(q("r", "id"))
            break
    if not rid:
        return ""
    target = ctx.rels.get(rid, ("", "", ""))[0]
    if not target:
        return ""
    part = join_part(ctx.part_dir, target)
    root = _part(ctx.zf, ctx.limits, part)
    if root is None:
        ctx.res.warn("차트 부분을 읽지 못했다: %s" % part)
        ctx.res.demote(UNVERIFIED)
        return ""
    ctx.res.assets["structure/" + Path(part).name] = ctx.zf.read(part)
    title = ""
    tnode = next((x for x in root.iter() if _tag(x) == "title"), None)
    if tnode is not None:
        title = "".join(t.text or "" for t in tnode.iter() if _tag(t) == "t")
    series = []
    for ser in root.iter():
        if _tag(ser) != "ser":
            continue
        nm = next((x for x in ser if _tag(x) == "tx"), None)
        cats = next((x for x in ser if _tag(x) == "cat"), None)
        vals = next((x for x in ser if _tag(x) == "val"), None)
        series.append({"name": "".join(_px_cache(nm)) if nm is not None else "",
                       "cats": _px_cache(cats) if cats is not None else [],
                       "vals": _px_cache(vals) if vals is not None else []})
    if not series:
        return ""
    cats = next((s0["cats"] for s0 in series if s0["cats"]), [])
    out = ["> 차트에 저장된 값이다(원본 xml은 보조 폴더에 그대로 보존). 그림 자체는 재현하지 않는다."]
    if title:
        out.append("**차트: %s**" % esc_inline(title))
    out.append("")
    out.append("| 항목 | " + " | ".join(esc_cell(s0["name"] or "계열%d" % (i + 1))
                                        for i, s0 in enumerate(series)) + " |")
    out.append("| --- | " + " | ".join("---" for _ in series) + " |")
    rows = max([len(s0["vals"]) for s0 in series] + [len(cats)])
    for i in range(rows):
        out.append("| %s | %s |" % (esc_cell(cats[i] if i < len(cats) else str(i + 1)),
                                    " | ".join(esc_cell(s0["vals"][i] if i < len(s0["vals"]) else "")
                                               for s0 in series)))
    ctx.charts.append({"part": part, "series": len(series), "rows": rows})
    return "\n".join(out)


def _px_asset(ctx, rid: str) -> str | None:
    target = ctx.rels.get(rid, ("", "", ""))[0]
    if not target:
        return None
    if ctx.rels.get(rid, ("", "", ""))[1] == "External":
        ctx.res.warn("문서 밖 그림 링크는 가져오지 않는다: %s" % target)
        ctx.res.demote(UNVERIFIED)
        return None
    part = join_part(ctx.part_dir, target)
    try:
        data = ctx.zf.read(part)
    except KeyError:
        ctx.res.warn("자산을 찾지 못했다: %s" % target)
        ctx.res.demote(PARTIAL)
        return None
    name = "media/" + Path(part).name
    ctx.res.assets[name] = data
    ctx.images.append({"part": part, "asset": name, "sha256": sha256(data), "bytes": len(data)})
    return name


class _PptxCtx:
    def __init__(self, res, zf, limits):
        self.res, self.zf, self.limits = res, zf, limits
        self.rels: dict = {}
        self.part_dir = "ppt/slides"
        self.images: list = []
        self.shapes: list = []
        self.charts: list = []


def convert_pptx(path: Path, opts, limits: Limits) -> Res:
    res = Res(path, "pptx")
    zf = open_zip(path, limits)
    try:
        ctx = _PptxCtx(res, zf, limits)
        pres = _part(zf, limits, "ppt/presentation.xml")
        if pres is None:
            res.status = FAILED
            res.warn("ppt/presentation.xml 이 없다")
            return res
        prels = _rels(zf, limits, "ppt/presentation.xml")
        slides = []
        for sld in pres.iter():
            if _tag(sld) == "sldId" and sld.get(q("r", "id")):
                target = prels.get(sld.get(q("r", "id")), ("", "", ""))[0]
                if target:
                    slides.append(join_part("ppt", target))
        if not slides:
            slides = sorted(n for n in zf.namelist()
                            if re.match(r"ppt/slides/slide\d+\.xml$", n))
            if slides:
                res.warn("슬라이드 순서 정보를 쓰지 못해 파일명 순서로 읽었다")
                res.demote(UNVERIFIED)
        out = [header_block(res, {"슬라이드 수": len(slides)})]
        for i, part in enumerate(slides, 1):
            root = _part(zf, limits, part)
            if root is None:
                res.warn("슬라이드를 읽지 못했다: %s" % part)
                res.demote(PARTIAL)
                continue
            ctx.rels = _rels(zf, limits, part)
            ctx.part_dir = str(Path(part).parent)
            out.append("")
            out.append("## 슬라이드 %d (%s)" % (i, part))
            tree = next((x for x in root.iter() if _tag(x) == "spTree"), None)
            order: list = []
            if tree is not None:
                for sp in tree:
                    if _tag(sp) in ("nvGrpSpPr", "grpSpPr"):
                        continue
                    _px_shape(sp, ctx, order)
            out.extend(order)
            # 발표자 노트
            for rid, (target, mode, typ) in _rels(zf, limits, part).items():
                if not typ.endswith("/notesSlide"):
                    continue
                npart = join_part(str(Path(part).parent), target)
                nroot = _part(zf, limits, npart)
                if nroot is None:
                    continue
                saved_rels, saved_dir = ctx.rels, ctx.part_dir
                ctx.rels = _rels(zf, limits, npart)
                ctx.part_dir = str(Path(npart).parent)
                ntree = next((x for x in nroot.iter() if _tag(x) == "spTree"), None)
                norder: list = []
                if ntree is not None:
                    for sp in ntree:
                        if _tag(sp) in ("nvGrpSpPr", "grpSpPr"):
                            continue
                        _px_shape(sp, ctx, norder)
                ctx.rels, ctx.part_dir = saved_rels, saved_dir
                if norder:
                    out.append("")
                    out.append("### 슬라이드 %d 발표자 노트" % i)
                    out.extend(norder)
        res.markdown = "\n\n".join(x for x in out if x != "") + "\n"
        res.meta = {"format": "pptx", "slides": slides, "images": ctx.images,
                    "shapes": ctx.shapes, "charts": ctx.charts,
                    "note": "도형 위치(x,y)는 EMU 단위 원본 값이다. 읽기 순서는 원본 도형 순서를 따른다."}
        res.info["chars"] = len(res.markdown)
        verify_pptx(res, zf, limits)
        return res
    finally:
        zf.close()


_PX_T = re.compile(rb"<a:t(?:\s[^>]*)?>(.*?)</a:t>", re.S)


def verify_pptx(res: Res, zf, limits: Limits) -> None:
    norm = normalize_md(res.markdown)
    parts = [n for n in zf.namelist()
             if re.match(r"ppt/(slides|notesSlides)/[^/]+\.xml$", n)]
    missing, total = [], 0
    for part in sorted(parts):
        data = zf.read(part)
        if len(data) > limits.max_xml_bytes:
            res.warn("%s 가 XML 한도를 넘어 검사하지 못했다" % part)
            res.demote(UNVERIFIED)
            continue
        for m in _PX_T.finditer(data):
            txt = _unent(m.group(1).decode("utf-8"))
            if not txt.strip():
                continue
            total += 1
            if not present(txt, res.markdown, norm):
                missing.append((part, txt[:60]))
    res.checks.append({"item": "텍스트 노드", "count": total, "missing": len(missing),
                       "ok": not missing})
    if missing:
        res.demote(PARTIAL)
        res.warn("원본 텍스트 %d개가 결과에서 확인되지 않았다: %r" % (len(missing), missing[:5]))
    check_orphan_media(res, zf, "ppt/media/", {i["part"] for i in res.meta.get("images", [])})


# ---------------------------------------------------------------- OCR (외부 엔진)
def ocr_available() -> str | None:
    return shutil.which("tesseract")


def run_ocr(data: bytes, ext: str, lang: str, res: Res, limits: Limits) -> str | None:
    """로컬 Tesseract로만 돌린다. 없으면 없다고 말한다(조용히 건너뛰지 않는다)."""
    exe = ocr_available()
    if not exe:
        res.warn("OCR 엔진(tesseract)이 없어 이미지의 글자를 읽지 못했다. "
                 "설치 후 --ocr 로 다시 실행할 것.")
        res.demote(UNVERIFIED)
        return None
    tmp = Path(tempfile.mkdtemp(prefix="mdmaker-ocr-"))
    try:
        img = tmp / ("page." + (ext if ext.isalnum() else "bin"))
        img.write_bytes(data)
        try:
            out = subprocess.run([exe, str(img), "stdout", "-l", lang],
                                 capture_output=True, timeout=limits.ocr_timeout,
                                 shell=False, cwd=str(tmp))
        except subprocess.TimeoutExpired:
            res.warn("OCR 시간 초과(%d초)" % limits.ocr_timeout)
            res.demote(PARTIAL)
            return None
        except OSError as exc:
            res.warn("OCR 실행 실패: %s" % exc)
            res.demote(PARTIAL)
            return None
        if out.returncode != 0:
            res.warn("OCR 실패(코드 %d): %s"
                     % (out.returncode, out.stderr.decode("utf-8", "replace")[:120]))
            res.demote(PARTIAL)
            return None
        return out.stdout.decode("utf-8", "replace")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ---------------------------------------------------------------- PDF
def _pdf_annots(doc, page, res: Res) -> list[str]:
    out = []
    annots = doc.resolve(page.get("Annots"))
    if not isinstance(annots, list):
        return out
    for a in annots:
        a = doc.resolve(a)
        if not isinstance(a, dict):
            continue
        sub = str(doc.resolve(a.get("Subtype")) or "")
        uri = ""
        act = doc.resolve(a.get("A"))
        if isinstance(act, dict):
            u = doc.resolve(act.get("URI"))
            if isinstance(u, bytes):
                uri = u.decode("utf-8", "replace")
        contents = doc.resolve(a.get("Contents"))
        text = contents.decode("utf-8", "replace") if isinstance(contents, bytes) else ""
        if text.startswith("\xfe\xff"):
            text = contents.decode("utf-16", "replace")
        if uri or text.strip():
            out.append("- [%s] %s%s" % (sub or "주석", esc_inline(text.strip()),
                                        (" → %s" % esc_inline(uri)) if uri else ""))
    return out


def convert_pdf(path: Path, opts, limits: Limits) -> Res:
    res = Res(path, "pdf")
    res.info["converter"] += " + pdfreader"
    try:
        doc = pdf.Pdf(path.read_bytes())
    except pdf.PdfError as exc:
        res.status = FAILED
        res.warn("PDF로 읽지 못했다: %s" % exc)
        return res
    res.info["pdf_version"] = doc.version
    if doc.encrypted:
        res.status = UNSUPPORTED
        res.warn("암호화된 PDF다. 보호 기능을 우회하지 않는다.")
        return res
    pages = doc.pages()
    if not pages:
        res.status = FAILED
        res.warn("페이지를 찾지 못했다")
        return res
    if len(pages) > limits.max_pages:
        res.demote(PARTIAL)
        res.warn("페이지 한도(%d)를 넘어 뒷부분을 변환하지 못했다" % limits.max_pages)
        pages = pages[:limits.max_pages]

    out = [header_block(res, {"페이지 수": len(pages), "PDF 버전": doc.version})]
    out.append("> PDF는 문단·표·읽기 순서를 파일에 담지 않는다. 아래 본문은 그리기 순서와 "
               "글자 좌표로 재구성한 것이라 자동으로 보존을 확정할 수 없다.")
    stats = {"codes": 0, "unmapped": 0, "chars": 0, "images": 0, "ocr_pages": 0,
             "empty_pages": 0}
    page_meta: list = []
    rebuilt: list = []
    for i, page in enumerate(pages, 1):
        out.append("")
        out.append("## 페이지 %d" % i)
        try:
            pt = pdf.extract_page(doc, page)
        except pdf.PdfError as exc:
            res.demote(PARTIAL)
            res.warn("%d쪽 본문을 읽지 못했다: %s" % (i, exc))
            out.append("> **이 페이지는 변환하지 못했다: %s**" % esc_inline(str(exc)))
            page_meta.append({"page": i, "status": "failed"})
            continue
        text = "".join(pt.parts)
        stats["codes"] += pt.codes
        stats["unmapped"] += pt.unmapped
        stats["chars"] += len(text)
        if pt.unmapped:
            res.warn("%d쪽에서 글자 %d개를 유니코드로 바꾸지 못했다(폰트에 대응표 없음)"
                     % (i, pt.unmapped))
            res.demote(PARTIAL)
        if pt.no_font:
            res.warn("%d쪽에 글꼴 지정 없이 그린 글자가 있다" % i)
            res.demote(UNVERIFIED)
        assets: list = []
        sized: list = []
        for name, st in pt.images:
            data, ext = pdf.image_bytes(doc, st)
            base = "media/p%d_%s" % (i, re.sub(r"\W+", "_", name))
            if ext in ("raw", "bin", "tif", "jbig2"):
                # PDF 안의 이미지는 파일이 아니라 원시 표본이라 그대로는 열리지 않는다.
                # 원본 바이트는 그대로 두고, 볼 수 있는 PNG를 따로 만든다.
                png = pdf.image_to_png(doc, st)
                res.assets[base + "." + ext] = data
                if png:
                    res.assets[base + ".png"] = png
                    assets.append(base + ".png")
                    rebuilt.append({"page": i, "raw": base + "." + ext,
                                    "png": base + ".png", "raw_sha256": sha256(data)})
                else:
                    assets.append(base + "." + ext)
                    res.warn("%d쪽 이미지를 볼 수 있는 형식으로 되돌리지 못했다(%s). "
                             "원본 바이트만 보존했다." % (i, ext))
                    res.demote(UNVERIFIED)
            else:
                res.assets[base + "." + ext] = data
                assets.append(base + "." + ext)
            w = doc.resolve(st.dict.get("Width")) or doc.resolve(st.dict.get("W")) or 0
            h = doc.resolve(st.dict.get("Height")) or doc.resolve(st.dict.get("H")) or 0
            sized.append((int(w) * int(h) if isinstance(w, int) and isinstance(h, int) else 0,
                          assets[-1]))
            stats["images"] += 1
        if text.strip():
            out.append(esc_block(text))
        for a in assets:
            out.append("![그림: 내용 해석 안 됨](%s)" % a)
        if not text.strip():
            stats["empty_pages"] += 1
            if assets:
                res.warn("%d쪽에 글자가 없고 이미지만 있다. 스캔 페이지일 수 있다." % i)
            else:
                res.warn("%d쪽에서 글자도 이미지도 찾지 못했다." % i)
            res.demote(UNVERIFIED)
        ocr_mode = getattr(opts, "ocr", "off")
        want_ocr = ocr_mode == "force" or (ocr_mode == "auto" and not text.strip() and assets)
        if want_ocr and assets:
            # 스캔 문서는 한 쪽이 이미지 수백 개로 쪼개져 있기도 하다. 전부 돌리면
            # 몇 시간이 걸리고 결과도 조각난다. 큰 것부터 정해진 개수만 읽는다.
            picked = [a for _, a in sorted(sized, key=lambda x: -x[0])[:limits.ocr_max_images]]
            if len(assets) > len(picked):
                res.warn("%d쪽 이미지 %d개 중 큰 %d개만 OCR했다(나머지는 읽지 않음)"
                         % (i, len(assets), len(picked)))
                res.demote(UNVERIFIED)
            got = []
            for aname in picked:
                txt = run_ocr(res.assets[aname], aname.rsplit(".", 1)[-1],
                              getattr(opts, "ocr_lang", "kor+eng"), res, limits)
                if txt and txt.strip():
                    got.append(txt)
            if got:
                stats["ocr_pages"] += 1
                out.append("")
                out.append("### %d쪽 OCR 결과 (검수 전, unverified)" % i)
                out.append("> 원본 이미지에서 기계가 읽은 글자다. 숫자·고유명사·표에 오류가 "
                           "있을 수 있으며 사람이 검수하기 전에는 확정된 내용이 아니다.")
                out.append(esc_block("\n".join(got)))
                res.demote(UNVERIFIED)
        ann = _pdf_annots(doc, page, res)
        if ann:
            out.append("")
            out.append("### %d쪽 주석·링크" % i)
            out.extend(ann)
        page_meta.append({"page": i, "chars": len(text), "codes": pt.codes,
                          "unmapped": pt.unmapped, "images": assets,
                          "fonts": sorted(pt.fonts),
                          "unmapped_codes": ["0x%04X×%d" % (c, n)
                                             for c, n in pt.missing_codes.most_common(50)]})

    res.markdown = "\n\n".join(x for x in out if x != "") + "\n"
    res.meta = {"format": "pdf", "version": doc.version, "pages": page_meta, "stats": stats,
                "rebuilt_images": rebuilt,
                "note": "글자 위치로 줄을 나눈 결과다. 다단·표·수식의 원래 구조는 복원하지 않는다."}
    res.info["chars"] = len(res.markdown)
    verify_pdf(res, doc, pages, stats)
    return res


def verify_pdf(res: Res, doc, pages, stats: dict) -> None:
    """본문 추출과 별개로 내용 스트림을 다시 훑어 글자 코드 수를 센다."""
    codes = 0
    checked = 0
    for page in pages:
        try:
            content = pdf.content_of(doc, page)
        except pdf.PdfError:
            res.demote(UNVERIFIED)
            continue
        checked += 1
        for m in re.finditer(rb"\((?:\\.|[^()\\])*\)|<[0-9A-Fa-f\s]*>", content):
            tok = m.group(0)
            if tok.startswith(b"("):
                codes += len(re.sub(rb"\\.", b".", tok[1:-1]))
            else:
                hx = re.sub(rb"[^0-9A-Fa-f]", b"", tok[1:-1])
                codes += len(hx) // 2
    res.checks.append({"item": "글자 코드", "stream_bytes": codes, "decoded": stats["codes"],
                       "unmapped": stats["unmapped"], "pages": checked,
                       "ok": stats["unmapped"] == 0 and stats["codes"] > 0})
    res.checks.append({"item": "페이지", "count": len(pages),
                       "empty": stats["empty_pages"], "ok": stats["empty_pages"] == 0})
    # PDF는 구조 정보를 담지 않으므로 검사에 통과해도 완전 보존을 확정하지 않는다.
    res.demote(UNVERIFIED)
    res.warn("PDF는 문단·표·읽기 순서를 파일에 담지 않아 완전 보존을 자동 확정할 수 없다. "
             "사람이 원본과 대조해야 한다.")


# ---------------------------------------------------------------- 이미지
def convert_image(path: Path, opts, limits: Limits) -> Res:
    res = Res(path, path.suffix.lower().lstrip("."))
    data = path.read_bytes()
    res.assets["media/" + path.name] = data
    res.info["bytes"] = len(data)
    out = [header_block(res, {"원본 보존": "`media/%s` (바이트 그대로)" % path.name})]
    ocr_mode = getattr(opts, "ocr", "off")
    text = None
    if ocr_mode in ("auto", "force"):
        text = run_ocr(data, res.fmt, getattr(opts, "ocr_lang", "kor+eng"), res, limits)
    if text and text.strip():
        out.append("## OCR 결과 (검수 전, unverified)")
        out.append("> 기계가 읽은 글자다. 사람이 검수하기 전에는 확정된 내용이 아니다.")
        out.append(esc_block(text))
        res.checks.append({"item": "OCR 문자", "count": len(text), "ok": False,
                           "note": "검수 전"})
    else:
        out.append("> 이미지 안의 글자는 해석하지 않았다. 원본 이미지만 보존했다. "
                   "`--ocr force` 와 로컬 OCR 엔진이 있어야 글자를 읽는다.")
        res.checks.append({"item": "OCR 문자", "count": 0, "ok": False, "note": "OCR 미수행"})
    res.markdown = "\n\n".join(out) + "\n"
    res.meta = {"format": res.fmt, "sha256": sha256(data),
                "note": "이미지의 내용은 자동으로 해석하지 않는다. OCR 결과는 검수 전 상태다."}
    res.warn("이미지 내용은 사람이 검수해야 확정된다(OCR 여부와 무관).")
    res.demote(UNVERIFIED)
    return res


# ---------------------------------------------------------------- 구형 오피스
LEGACY_TARGET = {".doc": "docx", ".dot": "docx", ".rtf": "docx", ".odt": "docx",
                 ".xls": "xlsx", ".ods": "xlsx", ".ppt": "pptx", ".odp": "pptx",
                 ".wps": "docx"}


def soffice_path() -> str | None:
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
              "/usr/lib/libreoffice/program/soffice"):
        if Path(p).exists():
            return p
    return None


def convert_legacy(path: Path, opts, limits: Limits) -> Res:
    """중간 변환기로 현재 형식으로 바꾼 뒤 기존 경로를 재사용한다.
    중간 변환의 손실은 대조할 수 없으므로 결과는 unverified를 넘지 않는다."""
    ext = path.suffix.lower()
    target = LEGACY_TARGET[ext]
    res = Res(path, ext.lstrip("."))
    exe = soffice_path()
    if not exe:
        res.status = UNSUPPORTED
        res.warn("구형 %s 변환에는 LibreOffice가 필요하다. 설치돼 있지 않아 변환하지 않았다. "
                 "(원본을 %s로 저장한 뒤 다시 실행해도 된다)" % (ext, target))
        return res
    tmp = Path(tempfile.mkdtemp(prefix="mdmaker-lo-"))
    try:
        try:
            out = subprocess.run(
                [exe, "--headless", "--norestore", "--nolockcheck", "--nodefault",
                 "-env:UserInstallation=file://%s/profile" % tmp,
                 "--convert-to", target, "--outdir", str(tmp), str(path)],
                capture_output=True, timeout=limits.convert_timeout, shell=False, cwd=str(tmp))
        except subprocess.TimeoutExpired:
            res.status = FAILED
            res.warn("중간 변환 시간 초과(%d초)" % limits.convert_timeout)
            return res
        except OSError as exc:
            res.status = FAILED
            res.warn("중간 변환 실행 실패: %s" % exc)
            return res
        made = sorted(tmp.glob("*." + target))
        if out.returncode != 0 or not made or made[0].stat().st_size == 0:
            res.status = FAILED
            res.warn("중간 변환 실패(코드 %d): %s"
                     % (out.returncode, out.stderr.decode("utf-8", "replace")[:150]))
            return res
        inner = _convert_as(target, made[0], opts, limits)
        inner.src = path
        inner.fmt = res.fmt
        inner.info["intermediate"] = "%s -> %s (LibreOffice)" % (ext, target)
        inner.markdown = inner.markdown.replace(made[0].name, path.name, 1)
        inner.warn("LibreOffice로 %s를 거쳐 변환했다. 중간 변환에서 생긴 차이는 대조할 수 없다."
                   % target)
        inner.demote(UNVERIFIED)
        return inner
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ---------------------------------------------------------------- 파이프라인
TEXT_EXT = {".txt", ".md", ".markdown", ".log"}
CSV_EXT = {".csv", ".tsv"}
IMAGE_EXT = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif", ".webp"}
SUPPORTED = (TEXT_EXT | CSV_EXT | IMAGE_EXT | set(LEGACY_TARGET)
             | {".docx", ".xlsx", ".hwpx", ".hwp", ".pptx", ".pdf"})
STATUS_LABEL = {SUCCESS: "success (전체 보존 검증 통과)", PARTIAL: "partial (누락·변형 확인)",
                UNVERIFIED: "unverified (보존 확인 불가)", FAILED: "failed",
                UNSUPPORTED: "unsupported", SKIPPED: "skipped"}


def detect_format(path: Path) -> str | None:
    """확장자를 믿지 않고 내용으로 형식을 본다(가이드 9장)."""
    head = path.read_bytes()[:8]
    if head.startswith(hwp5.CFB_SIGNATURE):
        return "hwp"
    if not head.startswith(b"PK\x03\x04"):
        return None
    try:
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
    except zipfile.BadZipFile:
        return None
    if "word/document.xml" in names:
        return "docx"
    if "xl/workbook.xml" in names:
        return "xlsx"
    if "ppt/presentation.xml" in names:
        return "pptx"
    if any(n.endswith("content.hpf") for n in names) or "mimetype" in names:
        return "hwpx"
    return None


def convert(path: Path, opts, limits: Limits) -> Res:
    ext = path.suffix.lower()
    sig = check_signature(path)
    if sig:
        real = detect_format(path)
        if real is None:
            res = Res(path, ext.lstrip("."))
            res.status = FAILED
            res.warn(sig)
            return res
        # 확장자가 틀린 파일을 확장자만 보고 실패시키지 않는다. 다만 사실을 기록한다.
        res = _convert_as(real, path, opts, limits)
        res.warn("확장자는 %s지만 내용은 %s다. 내용 기준으로 변환했다." % (ext, real))
        res.demote(UNVERIFIED)
        return res
    return _convert_as(ext.lstrip("."), path, opts, limits)


def _convert_as(fmt: str, path: Path, opts, limits: Limits) -> Res:
    ext = "." + fmt
    if ext in TEXT_EXT:
        return convert_text(path, opts)
    if ext in CSV_EXT:
        return convert_csv(path, opts)
    if ext == ".docx":
        return convert_docx(path, opts, limits)
    if ext == ".xlsx":
        return convert_xlsx(path, opts, limits)
    if ext == ".pptx":
        return convert_pptx(path, opts, limits)
    if ext == ".pdf":
        return convert_pdf(path, opts, limits)
    if ext in IMAGE_EXT:
        return convert_image(path, opts, limits)
    if ext in LEGACY_TARGET:
        return convert_legacy(path, opts, limits)
    if ext == ".hwpx":
        return convert_hwpx(path, opts, limits)
    if ext == ".hwp":
        return convert_hwp(path, opts, limits)
    res = Res(path, ext.lstrip(".") or "?")
    res.status = UNSUPPORTED
    res.warn("아직 지원하지 않는 형식이다 (현재: txt/md/csv/tsv/docx/xlsx/hwpx)")
    return res


def cell_block(text: str) -> str:
    """행별 표 표현의 셀 본문. 들여쓰기를 끼워 넣으면 원문 문자열이 끊기므로
    중첩 표 같은 블록 마크업이 있을 때만 들여쓴다."""
    lines = text.split("\n")
    if any(l.lstrip().startswith(("|", ">")) for l in lines):
        return text.replace("\n", "\n    ")
    return "<br>".join(lines)


def status_section(res: Res) -> str:
    lines = ["", "---", "", "## 변환 상태", "",
             "- 상태: **%s**" % STATUS_LABEL[res.status]]
    for c in res.checks:
        lines.append("- 보존 검사 · %s: %s" % (c["item"], json.dumps(c, ensure_ascii=False)))
    for w in res.warnings:
        lines.append("- 경고: %s" % w)
    if res.assets:
        lines.append("- 보조 파일: %d개 (같은 이름의 `.assets` 폴더). "
                     "이미지는 링크만으로 AI에 내용이 전달되지 않는다." % len(res.assets))
    if res.status != SUCCESS:
        lines.append("")
        lines.append("> 이 결과는 완전 보존 검증을 통과하지 못했다. 원본을 대체하지 않는다.")
    return "\n".join(lines) + "\n"


def split_markdown(body: str, limit: int) -> list[str]:
    """제목·문단 경계에서만 자른다. 한 덩어리가 한도보다 크면 줄 단위로 더 자른다."""
    blocks = body.split("\n\n")
    parts: list[list[str]] = [[]]
    size = 0
    for b in blocks:
        if len(b) > limit:
            lines = b.split("\n")
            chunk: list[str] = []
            for ln in lines:
                if chunk and size + len(ln) + 1 > limit:
                    parts[-1].append("\n".join(chunk))
                    parts.append([])
                    chunk, size = [], 0
                chunk.append(ln)
                size += len(ln) + 1
            if chunk:
                parts[-1].append("\n".join(chunk))
            continue
        if parts[-1] and size + len(b) + 2 > limit:
            parts.append([])
            size = 0
        parts[-1].append(b)
        size += len(b) + 2
    return ["\n\n".join(p) for p in parts if p]


def rejoin(parts: list[str]) -> str:
    return "\n\n".join(parts)


def options_key(opts) -> str:
    return json.dumps({k: getattr(opts, k, None)
                       for k in ("encoding", "xlsx_table", "max_cells", "ocr", "ocr_lang",
                                 "split_chars")}, ensure_ascii=False, sort_keys=True)


def cache_path(out_root: Path) -> Path:
    return out_root / ".mdmaker-cache.json"


def load_cache(out_root: Path) -> dict:
    try:
        return json.loads(cache_path(out_root).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def cache_valid(entry: dict, src: Path, out_root: Path, opts) -> bool:
    """내용 해시·옵션·버전이 모두 같고 결과 묶음이 온전할 때만 재사용한다.
    파일명이나 수정 시각은 근거로 쓰지 않는다."""
    if not entry or entry.get("status") != SUCCESS:
        return False
    if entry.get("options") != options_key(opts) or entry.get("version") != VERSION:
        return False
    if entry.get("check_version") != CHECK_VERSION:
        return False
    try:
        if entry.get("source_sha256") != sha256(src.read_bytes()):
            return False
        md = out_root / entry["md"]
        if not md.exists() or sha256(md.read_bytes()) != entry.get("md_sha256"):
            return False
        for rel, want in (entry.get("assets") or {}).items():
            f = out_root / rel
            if not f.exists() or sha256(f.read_bytes()) != want:
                return False
    except (OSError, KeyError):
        return False
    return True


def write_result(res: Res, src: Path, root: Path, out_root: Path, opts,
                 ours: bool = False) -> tuple[Path, str | None]:
    rel = src.relative_to(root) if root != src else Path(src.name)
    base = out_root if res.status == SUCCESS else out_root / "_incomplete"
    target = base / rel.parent / (rel.name + ".md")
    assets_dir = base / rel.parent / (rel.name + ".assets")
    # ours=True는 재사용 기록에 남아 있는, 우리가 만든 결과라는 뜻이다.
    if not (opts.overwrite or ours) and (target.exists() or assets_dir.exists()):
        return target, "출력이 이미 있다 (--overwrite 없이는 덮어쓰지 않는다): %s" % target
    target.parent.mkdir(parents=True, exist_ok=True)
    # 중간 실패로 기존 결과를 훼손하지 않도록 임시 위치에 묶음을 만든 뒤 옮긴다.
    tmp = Path(tempfile.mkdtemp(prefix=".mdmaker-", dir=str(target.parent)))
    made_parts: list[Path] = []
    try:
        limit = getattr(opts, "split_chars", 0) or 0
        if limit > 0:
            parts = split_markdown(res.markdown, limit)
            if rejoin(parts) != res.markdown:          # 재결합 검사: 내용·순서가 같아야 한다
                res.demote(PARTIAL)
                res.warn("분할 조각을 다시 합쳤을 때 원본과 달라 분할 결과를 신뢰할 수 없다")
            res.checks.append({"item": "분할", "parts": len(parts), "limit": limit,
                               "ok": rejoin(parts) == res.markdown})
            for i, part in enumerate(parts, 1):
                head = ("<!-- %s 조각 %d/%d · 문자 기준 분할(토큰 수와 다르다) -->\n"
                        % (src.name, i, len(parts)))
                fn = "%s.part%02d.md" % (rel.name, i)
                (tmp / fn).write_text(head + part, encoding="utf-8", newline="\n")
                made_parts.append(tmp / fn)
        (tmp / "doc.md").write_text(res.markdown + status_section(res), encoding="utf-8", newline="\n")
        adir = tmp / "assets"        # meta.json이 보존 검사 기록이라 항상 만든다
        adir.mkdir()
        listing = []
        for name, data in sorted(res.assets.items()):
            f2 = adir / name
            f2.parent.mkdir(parents=True, exist_ok=True)
            f2.write_bytes(data)
            listing.append({"path": name, "sha256": sha256(data), "bytes": len(data)})
        (adir / "meta.json").write_text(json.dumps({
            "source": src.name, "source_sha256": sha256(src.read_bytes()),
            "format": res.fmt, "status": res.status, "warnings": res.warnings,
            "checks": res.checks, "info": res.info, "assets": listing,
            "structure": res.meta}, ensure_ascii=False, indent=2), encoding="utf-8")
        if opts.overwrite or ours:
            os.replace(tmp / "doc.md", target)
        else:
            # 검사와 쓰기 사이에 다른 실행이 끼어들 수 있다. link는 원자적으로
            # "없을 때만 만들기"라서 동시 실행 충돌을 여기서 잡는다.
            try:
                os.link(tmp / "doc.md", target)
            except FileExistsError:
                return target, "다른 실행이 방금 같은 결과를 만들었다(덮어쓰지 않음): %s" % target
        # 본문 자리를 잡은 뒤에야 보조 폴더를 옮긴다(동시 실행에서 순서가 중요하다).
        for part in made_parts:
            os.replace(part, target.parent / part.name)
        if assets_dir.exists():
            shutil.rmtree(assets_dir, ignore_errors=True)
        try:
            os.replace(adir, assets_dir)
        except OSError as exc:
            return target, "보조 폴더를 놓지 못했다(동시 실행 충돌일 수 있다): %s" % exc
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return target, None


def collect(root: Path, out_root: Path, recursive: bool) -> list[Path]:
    files = []
    out_res = out_root.resolve()
    for p in sorted(root.rglob("*") if recursive else root.glob("*")):
        if p.is_symlink() or not p.is_file():
            continue
        try:
            rp = p.resolve()
        except OSError:
            continue
        if rp == out_res or out_res in rp.parents:   # 출력 폴더를 다시 읽지 않는다
            continue
        files.append(p)
    return files


def check_options(files: list[Path], opts) -> str | None:
    exts = {f.suffix.lower() for f in files}
    if opts.encoding and not (exts & (TEXT_EXT | CSV_EXT)):
        return "--encoding 은 텍스트/CSV 입력에만 쓴다"
    if opts.xlsx_table != "auto" and ".xlsx" not in exts:
        return "--xlsx-table 은 xlsx 입력에만 쓴다"
    if opts.ocr != "off" and not (exts & (IMAGE_EXT | {".pdf"})):
        return "--ocr 은 PDF·이미지 입력에만 쓴다"
    if opts.split_chars < 0:
        return "--split-chars 는 0보다 커야 한다"
    return None


def run_batch(files, root: Path, out_root: Path, opts, limits: Limits, single: bool = False):
    """파일 하나를 끝낼 때마다 결과를 내보낸다. CLI와 GUI가 같은 경로를 쓰게 하려는 것이다.

    yield: {"rel", "status", "target", "warnings", "elapsed"} — conflict면 status가
    "conflict"이고 warnings에 이유가 들어간다."""
    cache = load_cache(out_root) if opts.reuse else {}
    new_cache = dict(cache)
    for f in files:
        rel = Path(f.name) if single else f.relative_to(root)
        if f.suffix.lower() not in SUPPORTED:
            yield {"rel": rel, "status": UNSUPPORTED, "target": None, "elapsed": 0.0,
                   "warnings": ["아직 지원하지 않는 형식이다"]}
            continue
        key = str(rel)
        if opts.reuse and cache_valid(cache.get(key, {}), f, out_root, opts):
            yield {"rel": rel, "status": SKIPPED, "target": out_root / cache[key]["md"],
                   "elapsed": 0.0,
                   "warnings": ["내용·옵션·결과가 이전 성공과 같다(재검증 통과)"]}
            continue
        t0 = time.perf_counter()
        try:
            if f.stat().st_size > limits.max_bytes:
                raise ValueError("파일 크기 한도 초과: %d바이트" % f.stat().st_size)
            res = convert(f, opts, limits)
        except Exception as exc:                      # 한 파일이 실패해도 나머지는 계속한다
            res = Res(f, f.suffix.lower().lstrip("."))
            res.status = FAILED
            res.warn("%s: %s" % (type(exc).__name__, exc))
        # 성능 비교에 필요한 조건을 결과와 함께 남긴다(가이드 11장).
        res.info.update({"elapsed_sec": round(time.perf_counter() - t0, 3),
                         "source_bytes": f.stat().st_size, "ocr": opts.ocr,
                         "platform": platform.platform(),
                         "python": sys.version.split()[0]})
        target, conflict = write_result(res, f, root, out_root, opts,
                                        ours=bool(opts.reuse and key in cache))
        if conflict:
            yield {"rel": rel, "status": "conflict", "target": target,
                   "elapsed": res.info["elapsed_sec"], "warnings": [conflict]}
            continue
        if opts.reuse:
            if res.status == SUCCESS:
                adir = target.parent / (rel.name + ".assets")
                assets = {}
                for a in sorted(res.assets):
                    f2 = adir / a
                    if f2.exists():
                        assets[str(f2.relative_to(out_root))] = sha256(f2.read_bytes())
                new_cache[key] = {"source_sha256": sha256(f.read_bytes()),
                                  "options": options_key(opts), "version": VERSION,
                                  "check_version": CHECK_VERSION, "status": res.status,
                                  "md": str(target.relative_to(out_root)),
                                  "md_sha256": sha256(target.read_bytes()),
                                  "assets": assets}
            else:
                new_cache.pop(key, None)
        yield {"rel": rel, "status": res.status, "target": target,
               "elapsed": res.info["elapsed_sec"], "warnings": list(res.warnings)}
    if opts.reuse and new_cache != cache:
        out_root.mkdir(parents=True, exist_ok=True)
        cache_path(out_root).write_text(json.dumps(new_cache, ensure_ascii=False, indent=1),
                                        encoding="utf-8")


def exit_code(counts: dict) -> int:
    """재검증을 통과한 재사용 결과(skipped)만 성공과 함께 센다."""
    return 1 if sum(v for k, v in counts.items() if k not in (SUCCESS, SKIPPED)) else 0


def peak_memory_mb() -> float:
    """이 실행의 최대 메모리. ru_maxrss 단위가 macOS는 바이트, 리눅스는 KB다."""
    try:
        import resource
        v = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    except (ImportError, OSError):
        return 0.0
    return v / (1024 * 1024) if sys.platform == "darwin" else v / 1024


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="로컬 문서를 완전 보존 기준으로 Markdown으로 변환한다 (외부 API 없음)")
    ap.add_argument("input", help="입력 파일 또는 폴더 (URL 불가)")
    ap.add_argument("--out", required=True, help="출력 폴더")
    ap.add_argument("--recursive", action="store_true", help="하위 폴더까지 처리")
    ap.add_argument("--overwrite", action="store_true", help="기존 결과를 덮어쓴다")
    ap.add_argument("--encoding", help="텍스트/CSV 입력 인코딩을 직접 지정")
    ap.add_argument("--xlsx-table", choices=("auto", "md", "tsv"), default="auto",
                    help="엑셀 표 출력 방식")
    ap.add_argument("--max-cells", type=int, default=2_000_000, help="셀 처리 한도")
    ap.add_argument("--ocr", choices=("off", "auto", "force"), default="off",
                    help="이미지 글자 읽기 (로컬 엔진 필요, 결과는 항상 검수 전 상태)")
    ap.add_argument("--ocr-lang", default="kor+eng", help="OCR 언어 데이터 이름")
    ap.add_argument("--split-chars", type=int, default=0,
                    help="결과를 문자 수 기준으로 나눠 조각 파일도 함께 만든다(토큰 수 아님)")
    ap.add_argument("--reuse", action="store_true",
                    help="이전에 전체 보존 검증을 통과한 결과를 재검증 후 건너뛴다")
    opts = ap.parse_args(argv)

    if re.match(r"^[a-zA-Z][a-zA-Z0-9+.\-]*://", opts.input):
        print("오류: URL 입력은 지원하지 않는다. 로컬 파일만 처리한다.", file=sys.stderr)
        return 2
    src = Path(opts.input)
    if not src.exists():
        print("오류: 입력을 찾을 수 없다: %s" % src, file=sys.stderr)
        return 2
    out_root = Path(opts.out)
    limits = Limits(max_cells=opts.max_cells)

    if src.is_dir():
        root, files = src, collect(src, out_root, opts.recursive)
    else:
        root, files = src.parent, [src]
    err = check_options(files, opts)
    if err:
        print("오류: %s" % err, file=sys.stderr)
        return 2
    if not files:
        print("처리할 파일이 없다: %s" % src, file=sys.stderr)
        return 2

    counts: dict = {}
    elapsed_total = 0.0
    for item in run_batch(files, root, out_root, opts, limits, single=not src.is_dir()):
        counts[item["status"]] = counts.get(item["status"], 0) + 1
        elapsed_total += item["elapsed"]
        note = (" — " + "; ".join(item["warnings"][:2])) if item["warnings"] else ""
        if item["target"] is None:
            print("[%s] %s%s" % (item["status"], item["rel"], note))
        else:
            print("[%s] %s -> %s%s" % (item["status"], item["rel"], item["target"], note))

    print("\n처리 시간 %.1f초 · 최대 메모리 %.0fMB" % (elapsed_total, peak_memory_mb()))
    print("요약: " + ", ".join("%s %d" % (k, v) for k, v in sorted(counts.items())))
    return exit_code(counts)


if __name__ == "__main__":
    sys.exit(main())
