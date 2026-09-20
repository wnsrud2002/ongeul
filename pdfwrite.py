#!/usr/bin/env python3
"""PDF 쓰기: 글꼴 부분집합 추출과 PDF 뼈대. 표준 라이브러리만 쓴다.

한글을 PDF에 넣으려면 글꼴을 파일 안에 넣어야 한다. 통째로 넣으면 20MB가 되므로
쓰인 글자의 외곽선만 남긴다. 글꼴 임베딩 허가(OS/2 fsType)를 확인하고, 금지된
글꼴은 넣지 않는다.
"""
from __future__ import annotations

import struct
import zlib

HEAD_TABLES = ("head", "hhea", "maxp", "hmtx", "loca", "glyf", "cvt ", "fpgm", "prep")


class FontError(Exception):
    pass


def _u16(d, o):
    return struct.unpack_from(">H", d, o)[0]


def _u32(d, o):
    return struct.unpack_from(">I", d, o)[0]


class TtfFont:
    """PDF에 넣을 수 있는 TrueType(외곽선 glyf) 글꼴."""

    def __init__(self, data: bytes, index: int = 0):
        self.d = data
        base = _u32(data, 12) if data[:4] == b"ttcf" else 0
        if data[base:base + 4] not in (b"\x00\x01\x00\x00", b"true", b"ttcf"):
            if data[base:base + 4] == b"OTTO":
                raise FontError("CFF(OTTO) 글꼴은 아직 넣지 못한다. TrueType 외곽선이 필요하다.")
        n = _u16(data, base + 4)
        self.tables: dict[str, tuple[int, int]] = {}
        for i in range(n):
            rec = base + 12 + i * 16
            tag = data[rec:rec + 4].decode("latin-1")
            self.tables[tag] = (_u32(data, rec + 8), _u32(data, rec + 12))
        for need in ("head", "hhea", "maxp", "hmtx", "loca", "glyf", "cmap"):
            if need not in self.tables:
                raise FontError("%s 테이블이 없어 쓸 수 없다" % need)
        head = self.tables["head"][0]
        self.units = _u16(data, head + 18) or 1000
        self.index_to_loc = _u16(data, head + 50)
        hhea = self.tables["hhea"][0]
        self.ascent = struct.unpack_from(">h", data, hhea + 4)[0]
        self.descent = struct.unpack_from(">h", data, hhea + 6)[0]
        self.num_hmetrics = _u16(data, hhea + 34)
        self.num_glyphs = _u16(data, self.tables["maxp"][0] + 4)
        self.fs_type = None
        if "OS/2" in self.tables:
            self.fs_type = _u16(data, self.tables["OS/2"][0] + 8)
        self._cmap: dict[int, int] | None = None
        self._loca: list[int] | None = None

    # -- 허가 --
    @property
    def embeddable(self) -> bool:
        """fsType 비트 1이 켜져 있으면 임베딩 금지다. 값이 없으면 제한 표시가 없는 것이다."""
        return not (self.fs_type is not None and (self.fs_type & 0x0002))

    # -- 글자 -> 글리프 --
    def _build_cmap(self) -> dict[int, int]:
        d = self.d
        cm = self.tables["cmap"][0]
        best = None
        for i in range(_u16(d, cm + 2)):
            pid, eid, sub = struct.unpack_from(">HHI", d, cm + 4 + i * 8)
            fmt = _u16(d, cm + sub)
            score = {(3, 10): 4, (3, 1): 3, (0, 4): 3, (0, 3): 2}.get((pid, eid), 0)
            if fmt in (4, 12) and score and (best is None or score > best[0]):
                best = (score, cm + sub, fmt)
        if best is None:
            raise FontError("쓸 수 있는 cmap 하위표가 없다")
        _, base, fmt = best
        out: dict[int, int] = {}
        if fmt == 4:
            segx2 = _u16(d, base + 6)
            seg = segx2 // 2
            ends = struct.unpack_from(">%dH" % seg, d, base + 14)
            starts = struct.unpack_from(">%dH" % seg, d, base + 16 + segx2)
            deltas = struct.unpack_from(">%dh" % seg, d, base + 16 + segx2 * 2)
            ro_base = base + 16 + segx2 * 3
            ros = struct.unpack_from(">%dH" % seg, d, ro_base)
            for i in range(seg):
                for c in range(starts[i], min(ends[i], 0xFFFE) + 1):
                    if ros[i] == 0:
                        g = (c + deltas[i]) & 0xFFFF
                    else:
                        gi = ro_base + i * 2 + ros[i] + (c - starts[i]) * 2
                        if gi + 2 > len(d):
                            continue
                        g = _u16(d, gi)
                        if g:
                            g = (g + deltas[i]) & 0xFFFF
                    if g:
                        out.setdefault(c, g)
        else:
            for i in range(_u32(d, base + 12)):
                s, e, g = struct.unpack_from(">III", d, base + 16 + i * 12)
                if e - s > 0x10FFFF:
                    continue
                for k in range(s, e + 1):
                    out.setdefault(k, g + (k - s))
        return out

    def gid(self, ch: str) -> int:
        if self._cmap is None:
            self._cmap = self._build_cmap()
        return self._cmap.get(ord(ch), 0)

    def loca(self) -> list[int]:
        if self._loca is None:
            off, ln = self.tables["loca"]
            n = self.num_glyphs + 1
            if self.index_to_loc:
                self._loca = list(struct.unpack_from(">%dI" % n, self.d, off))
            else:
                self._loca = [x * 2 for x in struct.unpack_from(">%dH" % n, self.d, off)]
        return self._loca

    def advance(self, gid: int) -> int:
        off = self.tables["hmtx"][0]
        i = min(gid, self.num_hmetrics - 1)
        if off + i * 4 + 2 > len(self.d):
            return self.units // 2
        return _u16(self.d, off + i * 4)

    def width1000(self, gid: int) -> int:
        return int(round(self.advance(gid) * 1000.0 / self.units))

    # -- 부분집합 --
    def _components(self, gid: int, seen: set) -> None:
        """복합 글리프는 다른 글리프를 참조한다. 같이 넣지 않으면 글자가 깨진다."""
        loca = self.loca()
        if gid + 1 >= len(loca):
            return
        start, end = self.tables["glyf"][0] + loca[gid], self.tables["glyf"][0] + loca[gid + 1]
        if end - start < 10:
            return
        if struct.unpack_from(">h", self.d, start)[0] >= 0:
            return
        pos = start + 10
        while pos + 4 <= end:
            flags, index = struct.unpack_from(">HH", self.d, pos)
            pos += 4
            if index not in seen:
                seen.add(index)
                self._components(index, seen)
            pos += 4 if flags & 0x0001 else 2
            if flags & 0x0008:
                pos += 2
            elif flags & 0x0040:
                pos += 4
            elif flags & 0x0080:
                pos += 8
            if not flags & 0x0020:
                break

    def subset(self, gids: set[int]) -> bytes:
        """쓰인 글리프만 남긴 글꼴. 글리프 번호는 그대로 두어 PDF 쪽을 단순하게 한다."""
        used = {0} | {g for g in gids if 0 <= g < self.num_glyphs}
        for g in list(used):
            self._components(g, used)
        loca = self.loca()
        gbase = self.tables["glyf"][0]
        glyf = bytearray()
        new_loca = [0]
        for gid in range(self.num_glyphs):
            if gid in used and gid + 1 < len(loca) and loca[gid + 1] > loca[gid]:
                blob = self.d[gbase + loca[gid]:gbase + loca[gid + 1]]
                glyf += blob
                if len(glyf) % 4:
                    glyf += b"\x00" * (4 - len(glyf) % 4)
            new_loca.append(len(glyf))
        tables = {}
        for tag in HEAD_TABLES:
            if tag in ("loca", "glyf") or tag not in self.tables:
                continue
            off, ln = self.tables[tag]
            tables[tag] = bytearray(self.d[off:off + ln])
        head = tables["head"]
        struct.pack_into(">I", head, 8, 0)             # checksumAdjustment
        struct.pack_into(">h", head, 50, 1)            # indexToLocFormat = long
        tables["loca"] = bytearray(b"".join(struct.pack(">I", x) for x in new_loca))
        tables["glyf"] = glyf
        return _build_sfnt(tables)


def _checksum(data: bytes) -> int:
    pad = data + b"\x00" * (-len(data) % 4)
    return sum(struct.unpack(">%dI" % (len(pad) // 4), pad)) & 0xFFFFFFFF


def _build_sfnt(tables: dict) -> bytes:
    tags = sorted(tables)
    n = len(tags)
    search = 1
    while search * 2 <= n:
        search *= 2
    out = bytearray(struct.pack(">IHHHH", 0x00010000, n, search * 16,
                                (search).bit_length() - 1, (n - search) * 16))
    offset = 12 + n * 16
    records = bytearray()
    body = bytearray()
    for tag in tags:
        data = bytes(tables[tag])
        records += tag.encode("latin-1") + struct.pack(">III", _checksum(data), offset, len(data))
        body += data + b"\x00" * (-len(data) % 4)
        offset += len(data) + (-len(data) % 4)
    return bytes(out + records + body)


# ---------------------------------------------------------------- PDF 뼈대
class PdfDoc:
    """번호 매긴 객체와 xref 를 갖춘 최소 PDF. 글꼴은 Identity-H 로 넣는다."""

    def __init__(self):
        self.objects: list[bytes | None] = [None]      # 1번부터 쓴다

    def add(self, body: bytes) -> int:
        self.objects.append(body)
        return len(self.objects) - 1

    def reserve(self) -> int:
        self.objects.append(b"null")
        return len(self.objects) - 1

    def set(self, num: int, body: bytes) -> None:
        self.objects[num] = body

    def stream(self, d: str, data: bytes, compress: bool = True) -> int:
        if compress:
            data = zlib.compress(data)
            d = d + " /Filter /FlateDecode"
        return self.add(("<< %s /Length %d >>\nstream\n" % (d, len(data))).encode("latin-1")
                        + data + b"\nendstream")

    def build(self, root: int, info: int | None = None) -> bytes:
        out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
        offsets = [0] * len(self.objects)
        for num in range(1, len(self.objects)):
            offsets[num] = len(out)
            out += ("%d 0 obj\n" % num).encode("latin-1")
            out += self.objects[num] or b"null"
            out += b"\nendobj\n"
        xref = len(out)
        out += ("xref\n0 %d\n" % len(self.objects)).encode("latin-1")
        out += b"0000000000 65535 f \n"
        for num in range(1, len(self.objects)):
            out += ("%010d 00000 n \n" % offsets[num]).encode("latin-1")
        trailer = "trailer\n<< /Size %d /Root %d 0 R%s >>\nstartxref\n%d\n%%%%EOF\n" % (
            len(self.objects), root, (" /Info %d 0 R" % info) if info else "", xref)
        out += trailer.encode("latin-1")
        return bytes(out)


def pdf_text(s: str) -> bytes:
    """PDF 문자열. 한글처럼 Latin-1 밖의 글자는 UTF-16BE 16진 표기로 쓴다."""
    try:
        body = s.encode("latin-1")
    except UnicodeEncodeError:
        return b"<FEFF" + s.encode("utf-16-be").hex().upper().encode("ascii") + b">"
    esc = body.replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)")
    return b"(" + esc + b")"


def embed_font(doc: PdfDoc, font: TtfFont, used: dict[int, str], name: str) -> int:
    """부분집합 글꼴을 Type0(Identity-H)로 넣는다. ToUnicode 도 같이 넣어야
    만든 PDF에서 글자를 다시 뽑을 수 있다(우리 검사기가 그걸 확인한다)."""
    sub = font.subset(set(used))
    file_ref = doc.stream("/Length1 %d" % len(sub), sub)
    scale = 1000.0 / font.units
    widths = []
    for gid in sorted(used):
        widths.append("%d [%d]" % (gid, font.width1000(gid)))
    desc = doc.add(("<< /Type /FontDescriptor /FontName /%s /Flags 4 /FontBBox [0 %d 1000 %d]"
                    " /ItalicAngle 0 /Ascent %d /Descent %d /CapHeight %d /StemV 80"
                    " /FontFile2 %d 0 R >>"
                    % (name, int(font.descent * scale), int(font.ascent * scale),
                       int(font.ascent * scale), int(font.descent * scale),
                       int(font.ascent * scale * 0.7), file_ref)).encode("latin-1"))
    tou = ["/CIDInit /ProcSet findresource begin 12 dict begin begincmap",
           "1 begincodespacerange <0000> <FFFF> endcodespacerange"]
    items = [(g, t) for g, t in sorted(used.items()) if t]
    for i in range(0, len(items), 100):
        chunk = items[i:i + 100]
        tou.append("%d beginbfchar" % len(chunk))
        for gid, ch in chunk:
            tou.append("<%04X> <%s>" % (gid, "".join("%04X" % ord(c) for c in ch[:4])))
        tou.append("endbfchar")
    tou.append("endcmap CMapName currentdict /CMap defineresource pop end end")
    tou_ref = doc.stream("", "\n".join(tou).encode("utf-8"))
    cid = doc.add(("<< /Type /Font /Subtype /CIDFontType2 /BaseFont /%s"
                   " /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >>"
                   " /FontDescriptor %d 0 R /DW 1000 /W [%s] /CIDToGIDMap /Identity >>"
                   % (name, desc, " ".join(widths))).encode("latin-1"))
    return doc.add(("<< /Type /Font /Subtype /Type0 /BaseFont /%s /Encoding /Identity-H"
                    " /DescendantFonts [%d 0 R] /ToUnicode %d 0 R >>"
                    % (name, cid, tou_ref)).encode("latin-1"))


# ---------------------------------------------------------------- 마크다운 배치
A4 = (595.28, 841.89)
SIZES = {1: 19, 2: 16, 3: 14, 4: 12.5, 5: 11.5, 6: 11}
BODY = 10
LEADING = 1.45
MONO_BG = (0.96, 0.96, 0.97)

_INLINE = [
    (r"<br\s*/?>", "\n"), (r"</?span[^>]*>", ""), (r"\*\*", ""), (r"~~", ""), (r"`", ""),
]
_LINK = __import__("re").compile(r"\[([^\]]*)\]\(([^)]*)\)")
_IMG = __import__("re").compile(r"!\[([^\]]*)\]\(([^)]*)\)")
_ESC = __import__("re").compile(r"\\(.)")


def plain(text: str) -> str:
    import re
    import unicodedata
    # macOS 파일 이름은 자모가 분해된 형태(NFD)로 온다. 그대로 그리면 "체육"이
    # "ㅊㅔㅇㅠㄱ"으로 보인다. 그릴 때만 합친다(원본 Markdown 은 건드리지 않는다).
    s = unicodedata.normalize("NFC", text)
    for pat, rep in _INLINE:
        s = re.sub(pat, rep, s)
    s = _LINK.sub(lambda m: "%s (%s)" % (m.group(1), m.group(2)) if m.group(2) else m.group(1), s)
    return _ESC.sub(r"\1", s)


class Writer:
    def __init__(self, font: TtfFont, title: str = "", base_dir=None):
        self.font = font
        self.title = title
        self.base = base_dir
        self.doc = PdfDoc()
        self.used: dict[int, str] = {}
        self.missing: dict[str, int] = {}
        self.pages: list[int] = []
        self.images: dict[str, tuple] = {}
        self.ops: list[str] = []
        self.x0, self.y0 = 50.0, 50.0
        self.w = A4[0] - self.x0 * 2
        self.y = A4[1] - 60.0

    # -- 글자 --
    def width(self, s: str, size: float) -> float:
        f = self.font
        s = s.replace("\t", "    ")
        return sum(f.width1000(f.gid(c)) for c in s if ord(c) >= 32) * size / 1000.0

    def _hex(self, s: str) -> str:
        out = []
        for c in s:
            g = self.font.gid(c)
            if g == 0 and not c.isspace() and ord(c) >= 32:
                self.missing[c] = self.missing.get(c, 0) + 1   # 빈 네모로 그려질 글자
            else:
                self.used.setdefault(g, c)
            out.append("%04X" % g)
        return "".join(out)

    def draw(self, s: str, x: float, y: float, size: float, bold=False, gray=0.0) -> None:
        # 탭·제어문자는 글꼴에 글리프가 없어 빈 네모로 그려진다. 빈칸으로 바꾼다.
        s = s.replace("\t", "    ")
        s = "".join(" " if ord(c) < 32 else c for c in s)
        if not s.strip():
            return
        mode = "2 Tr 0.3 w" if bold else "0 Tr"
        self.ops.append("q %s %s BT /F1 %.2f Tf %.2f %.2f Td <%s> Tj ET Q"
                        % ("%.2f g %.2f G" % (gray, gray), mode, size, x, y, self._hex(s)))

    def rule(self, y: float, x1: float, x2: float, gray=0.75, w=0.6) -> None:
        self.ops.append("q %.2f G %.2f w %.2f %.2f m %.2f %.2f l S Q" % (gray, w, x1, y, x2, y))

    def box(self, x, y, w, h, fill=None, stroke=0.8) -> None:
        if fill:
            self.ops.append("q %.3f %.3f %.3f rg %.2f %.2f %.2f %.2f re f Q" % (*fill, x, y, w, h))
        self.ops.append("q %.2f G 0.5 w %.2f %.2f %.2f %.2f re S Q" % (stroke, x, y, w, h))

    # -- 쪽 --
    def need(self, h: float) -> None:
        if self.y - h < self.y0:
            self.end_page()

    def end_page(self) -> None:
        if not self.ops:
            return
        cs = self.doc.stream("", "\n".join(self.ops).encode("latin-1"))
        self.pages.append(cs)
        self.ops = []
        self.y = A4[1] - 60.0

    def _furniture(self, index: int, total: int) -> int:
        """머리글과 쪽번호. 전체 쪽수를 알아야 해서 마지막에 따로 그린다."""
        ops = []
        label = "%d / %d" % (index + 1, total)
        ops.append("q 0.45 g BT /F1 8.5 Tf %.2f %.2f Td <%s> Tj ET Q"
                   % (A4[0] / 2 - self.width(label, 8.5) / 2, 28, self._hex(label)))
        if index > 0 and self.title:
            head = self.title if self.width(self.title, 8) < self.w else self.title[:40] + "…"
            ops.append("q 0.55 g BT /F1 8 Tf %.2f %.2f Td <%s> Tj ET Q"
                       % (self.x0, A4[1] - 38, self._hex(head)))
            ops.append("q 0.88 G 0.5 w %.2f %.2f m %.2f %.2f l S Q"
                       % (self.x0, A4[1] - 44, self.x0 + self.w, A4[1] - 44))
        return self.doc.stream("", "\n".join(ops).encode("latin-1"))

    # -- 블록 --
    def wrap(self, text: str, size: float, width: float) -> list[str]:
        """한글은 아무데서나, 영문은 빈칸에서 끊는다."""
        lines: list[str] = []
        for para in text.split("\n"):
            cur = ""
            for ch in para:
                if self.width(cur + ch, size) > width and cur:
                    cut = cur.rfind(" ")
                    if cut > len(cur) * 0.4:
                        lines.append(cur[:cut])
                        cur = cur[cut + 1:] + ch
                    else:
                        lines.append(cur)
                        cur = ch
                else:
                    cur += ch
            lines.append(cur)
        return lines

    def paragraph(self, text: str, size=BODY, indent=0.0, bold=False, gray=0.0, bar=False) -> None:
        for line in self.wrap(text, size, self.w - indent):
            self.need(size * LEADING)
            if bar:
                self.ops.append("q 0.8 0.8 0.85 rg %.2f %.2f 2.5 %.2f re f Q"
                                % (self.x0 + indent - 10, self.y - 2, size * LEADING))
            self.draw(line, self.x0 + indent, self.y, size, bold=bold, gray=gray)
            self.y -= size * LEADING

    def heading(self, level: int, text: str) -> None:
        size = SIZES.get(level, 11)
        # 제목만 쪽 끝에 덩그러니 남지 않게 뒤따를 두 줄 자리까지 본다
        self.need(size * 2.4 + BODY * LEADING * 2)
        self.y -= size * 0.7
        self.paragraph(text, size=size, bold=level <= 2)
        if level <= 2:
            self.rule(self.y + size * 0.35, self.x0, self.x0 + self.w, gray=0.85)
        self.y -= size * 0.35

    def code(self, lines: list[str]) -> None:
        size = 8.6
        h = len(lines) * size * 1.35 + 8
        self.need(min(h, 200))
        top = self.y + size
        drawn = 0
        for line in lines:
            for piece in self.wrap(line, size, self.w - 16) or [""]:
                if self.y - size * 1.35 < self.y0:
                    self.ops.insert(len(self.ops) - drawn,
                                    "q %.3f %.3f %.3f rg %.2f %.2f %.2f %.2f re f Q"
                                    % (*MONO_BG, self.x0, self.y, self.w, top - self.y))
                    self.end_page()
                    top = self.y + size
                    drawn = 0
                self.draw(piece, self.x0 + 8, self.y, size, gray=0.15)
                self.y -= size * 1.35
                drawn += 1
        self.ops.insert(len(self.ops) - drawn,
                        "q %.3f %.3f %.3f rg %.2f %.2f %.2f %.2f re f Q"
                        % (*MONO_BG, self.x0, self.y, self.w, top - self.y))
        self.y -= 6

    def table(self, rows: list[list[str]]) -> None:
        size = 8.8
        cols = max(len(r) for r in rows)
        colw = self.w / cols
        header = rows[0] if rows else []
        page_at_start = len(self.pages)
        for ri, row in enumerate(rows):
            if ri and len(self.pages) != page_at_start:
                page_at_start = len(self.pages)
                self._table_row(header, cols, colw, size, is_header=True)
            self._table_row(row, cols, colw, size, is_header=(ri == 0))
        self.y -= 6

    def _table_row(self, row, cols, colw, size, is_header=False) -> None:
        cells = [self.wrap(c, size, colw - 8) for c in row] + [[""]] * (cols - len(row))
        h = max(len(c) for c in cells) * size * 1.3 + 6
        self.need(h)
        top = self.y + size
        for ci, lines in enumerate(cells):
            x = self.x0 + ci * colw
            self.box(x, top - h, colw, h, fill=(0.94, 0.94, 0.96) if is_header else None)
            yy = self.y
            for line in lines:
                self.draw(line, x + 4, yy, size, bold=is_header)
                yy -= size * 1.3
        self.y = top - h - size * 0.2

    def image(self, path, alt: str) -> None:
        info = load_image(path)
        if not info:
            self.paragraph("[그림: %s — 넣지 못했다]" % alt, size=9, gray=0.45)
            return
        data, w, h, kind, extra = info
        key = str(path)
        if key not in self.images:
            ref = self.doc.stream("/Type /XObject /Subtype /Image /Width %d /Height %d %s"
                                  % (w, h, extra), data, compress=(kind != "jpg"))
            self.images[key] = (ref, w, h)
        ref, w, h = self.images[key]
        draw_w = min(self.w, w * 0.75)
        draw_h = draw_w * h / w
        if draw_h > A4[1] - 140:
            draw_h = A4[1] - 140
            draw_w = draw_h * w / h
        # 남은 자리가 넉넉하면 줄여서 넣는다. 무조건 쪽을 넘기면 앞쪽이 텅 빈다.
        avail = self.y - self.y0 - 10
        if draw_h > avail:
            if avail > (A4[1] - 140) * 0.45:
                draw_w *= avail / draw_h
                draw_h = avail
            else:
                self.end_page()
        self.need(draw_h + 10)
        self.ops.append("q %.2f 0 0 %.2f %.2f %.2f cm /I%d Do Q"
                        % (draw_w, draw_h, self.x0, self.y - draw_h, ref))
        self.y -= draw_h + 8
        if alt:
            self.paragraph(alt, size=8.5, gray=0.45)

    # -- 변환 상태 (사람이 읽는 형태) --
    def summary_block(self, summary: dict) -> None:
        if not summary:
            return
        self.y -= 10
        self.need(90)
        self.rule(self.y + 6, self.x0, self.x0 + self.w, gray=0.8)
        self.y -= 6
        self.paragraph("변환 상태", size=13, bold=True)
        ok = summary.get("status") == "success"
        badge = summary.get("label") or summary.get("status", "")
        pad = 5
        tw = self.width(badge, 9.5)
        self.need(20)
        fill = (0.90, 0.95, 0.90) if ok else (0.99, 0.94, 0.88)
        self.ops.append("q %.3f %.3f %.3f rg %.2f %.2f %.2f %.2f re f Q"
                        % (*fill, self.x0, self.y - 4, tw + pad * 2, 16))
        self.draw(badge, self.x0 + pad, self.y, 9.5, gray=0.1)
        self.y -= 24
        rows = summary.get("rows") or []
        if rows:
            self.table([["검사 항목", "결과"]] + rows)
        for w in summary.get("warnings", []):
            self.paragraph("• " + w, size=9, indent=6, gray=0.25)
        note = summary.get("note")
        if note:
            self.y -= 4
            self.paragraph(note, size=8.5, gray=0.45)

    # -- 마무리 --
    def save(self) -> bytes:
        self.end_page()
        if not self.pages:
            self.ops.append("")
            self.end_page()
        furniture = [self._furniture(i, len(self.pages)) for i in range(len(self.pages))]
        font_ref = embed_font(self.doc, self.font, self.used or {0: ""}, "OngeulSubset")
        xobj = " ".join("/I%d %d 0 R" % (r, r) for r, _w, _h in self.images.values())
        parent = self.doc.reserve()
        kids = []
        for i, cs in enumerate(self.pages):
            kids.append(self.doc.add(
                ("<< /Type /Page /Parent %d 0 R /MediaBox [0 0 %.2f %.2f]"
                 " /Contents [%d 0 R %d 0 R]"
                 " /Resources << /Font << /F1 %d 0 R >> /XObject << %s >> >> >>"
                 % (parent, A4[0], A4[1], cs, furniture[i], font_ref, xobj)).encode("latin-1")))
        self.doc.set(parent, ("<< /Type /Pages /Kids [%s] /Count %d >>"
                              % (" ".join("%d 0 R" % k for k in kids), len(kids))).encode())
        root = self.doc.add(("<< /Type /Catalog /Pages %d 0 R >>" % parent).encode())
        info = self.doc.add(b"<< /Producer " + pdf_text("ongeul") + b" /Title "
                            + pdf_text(self.title[:120]) + b" >>")
        return self.doc.build(root, info)


def load_image(path):
    """(데이터, 너비, 높이, 종류, PDF 추가 항목). JPEG는 그대로, 나머지는 Pillow로."""
    try:
        raw = path.read_bytes()
    except OSError:
        return None
    if raw[:3] == b"\xff\xd8\xff":
        i, n = 2, len(raw)
        while i + 9 < n:
            if raw[i] != 0xFF:
                i += 1
                continue
            marker = raw[i + 1]
            if marker in (0xC0, 0xC1, 0xC2, 0xC3):
                h, w = struct.unpack_from(">HH", raw, i + 5)
                comp = raw[i + 9]
                cs = {1: "/DeviceGray", 3: "/DeviceRGB", 4: "/DeviceCMYK"}.get(comp, "/DeviceRGB")
                return raw, w, h, "jpg", "/ColorSpace %s /BitsPerComponent 8 /Filter /DCTDecode" % cs
            i += 2 + struct.unpack_from(">H", raw, i + 2)[0]
        return None
    try:
        from PIL import Image
        import io
        im = Image.open(io.BytesIO(raw))
        im = im.convert("RGB")
        return im.tobytes(), im.width, im.height, "raw", "/ColorSpace /DeviceRGB /BitsPerComponent 8"
    except Exception:
        return None


FONT_CANDIDATES = (
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
    "/Library/Fonts/NanumGothic.ttf",
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "C:/Windows/Fonts/malgun.ttf",
    "C:/Windows/Fonts/arialuni.ttf",
)
_FONT_CACHE: dict = {}


def covers(font: "TtfFont", sample: str) -> float:
    """글꼴이 이 글자들을 실제로 갖고 있는 비율. 없는 글자는 빈 네모로 그려져
    멀쩡해 보이는 엉터리 PDF가 나오므로 미리 본다."""
    chars = {c for c in sample if not c.isspace()}
    if not chars:
        return 1.0
    return sum(1 for c in chars if font.gid(c)) / float(len(chars))


def find_font(explicit: str | None = None, sample: str = ""):
    """PDF에 넣을 글꼴을 찾는다. (글꼴, 사유).
    임베딩이 금지됐거나 본문 글자를 못 그리는 글꼴은 쓰지 않는다."""
    from pathlib import Path as _Path
    tried = []
    probe = sample[:4000]
    for cand in ([explicit] if explicit else []) + list(FONT_CANDIDATES):
        if not cand:
            continue
        p = _Path(cand)
        if not p.exists():
            tried.append("%s: 없음" % p.name)
            continue
        key = str(p)
        font = _FONT_CACHE.get(key)
        if font is None:
            try:
                font = TtfFont(p.read_bytes())
            except (FontError, struct.error, ValueError) as exc:
                tried.append("%s: %s" % (p.name, exc))
                continue
            if not font.embeddable:
                tried.append("%s: 글꼴 제작자가 임베딩을 금지했다" % p.name)
                continue
            font.source_name = p.name
            _FONT_CACHE[key] = font
        if probe:
            rate = covers(font, probe)
            if rate < 0.98:
                tried.append("%s: 본문 글자의 %.0f%%만 갖고 있다" % (p.name, rate * 100))
                continue
        return font, ""
    return None, "본문을 그릴 수 있는 글꼴을 찾지 못했다 (%s)" % "; ".join(tried[:4])


def render_markdown(md: str, font: TtfFont, title: str, base_dir=None,
                    report: dict | None = None, summary: dict | None = None) -> bytes:
    """온글이 만든 Markdown 을 PDF 로 그린다. 우리가 쓰는 표기만 다룬다.
    report 를 주면 글꼴에 없어 못 그린 글자를 담아 돌려준다."""
    import re
    from pathlib import Path as _P
    w = Writer(font, title=title, base_dir=base_dir)
    lines = md.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if stripped.startswith("<!--"):        # 기계용 메모는 읽는 사람에게 필요 없다
            w.y -= 2
            w.rule(w.y, w.x0, w.x0 + w.w, gray=0.9)
            w.y -= 8
            i += 1
            continue
        if stripped.startswith("```"):
            j = i + 1
            block = []
            while j < len(lines) and not lines[j].strip().startswith("```"):
                block.append(lines[j])
                j += 1
            w.code(block)
            i = j + 1
            continue
        m = re.match(r"^(#{1,6})\s+(.*)$", stripped)
        if m:
            w.heading(len(m.group(1)), plain(m.group(2)))
            i += 1
            continue
        if stripped.startswith("|") and stripped.endswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [plain(c.strip()) for c in lines[i].strip().strip("|").split("|")]
                if not all(set(c) <= set("-: ") for c in cells):
                    rows.append(cells)
                i += 1
            if rows:
                w.table(rows)
            continue
        # 그림은 문단 한가운데에 끼어 있기도 하다. 글은 글대로 쓰고 그림은 따로 그린다.
        shots = _IMG.findall(line) if base_dir is not None else []
        if shots:
            rest = plain(_IMG.sub("", line)).strip()
            if rest:
                w.paragraph(rest)
            for alt, target in shots:
                if target.startswith(("http://", "https://")):
                    w.paragraph("[바깥 그림 링크: %s]" % target, size=9, gray=0.45)
                else:
                    w.image(_P(base_dir) / target, plain(alt))
            i += 1
            continue
        if re.match(r"^(-{3,}|\*{3,}|={3,})$", stripped):
            w.need(12)
            w.rule(w.y, w.x0, w.x0 + w.w)
            w.y -= 12
            i += 1
            continue
        if stripped.startswith(">"):
            w.paragraph(plain(stripped.lstrip("> ")), size=9.5, indent=14, gray=0.3, bar=True)
            i += 1
            continue
        m = re.match(r"^(\s*)([-*+]|\d+[.)])\s+(.*)$", line)
        if m:
            depth = len(m.group(1)) // 2
            bullet = "•" if m.group(2) in "-*+" else m.group(2)
            w.paragraph("%s %s" % (bullet, plain(m.group(3))), indent=14 + depth * 14)
            i += 1
            continue
        if not stripped:
            w.y -= BODY * 0.5
            i += 1
            continue
        w.paragraph(plain(line))
        i += 1
    w.summary_block(summary or {})
    out = w.save()
    if report is not None:
        report["missing_chars"] = dict(sorted(w.missing.items(), key=lambda kv: -kv[1])[:20])
        report["missing_total"] = sum(w.missing.values())
        report["pages"] = len(w.pages)
        report["font"] = getattr(font, "source_name", "?")
    return out
