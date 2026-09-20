#!/usr/bin/env python3
"""PDF 저수준 리더: 객체·스트림·폰트 인코딩만 다룬다(표준 라이브러리만 사용).

mdmaker.py를 import하지 않는다. 암호화된 문서는 해제하지 않고 식별만 한다.
xref 대신 파일 전체에서 "N G obj"를 훑는다. 깨진 xref나 증분 저장에도 견디고,
xref 스트림 구현을 생략할 수 있어서다(같은 번호는 뒤에 나온 것이 이긴다).
"""
from __future__ import annotations

import re
import struct
import zlib

WS = b"\x00\t\n\x0c\r "
DELIM = b"()<>[]{}/%"
MAX_DEPTH = 64


class PdfError(Exception):
    pass


class Name(str):
    __slots__ = ()


class PdfString(bytes):
    """PDF 문자열. 연산자(bytes)와 구분하지 않으면 `(글자) Tj` 가 연산자로 오인된다."""
    __slots__ = ()


class Ref:
    __slots__ = ("num", "gen")

    def __init__(self, num: int, gen: int = 0):
        self.num, self.gen = num, gen

    def __repr__(self):
        return "Ref(%d,%d)" % (self.num, self.gen)

    def __eq__(self, o):
        return isinstance(o, Ref) and (self.num, self.gen) == (o.num, o.gen)

    def __hash__(self):
        return hash((self.num, self.gen))


class Stream:
    __slots__ = ("dict", "raw")

    def __init__(self, d: dict, raw: bytes):
        self.dict, self.raw = d, raw


# ---------------------------------------------------------------- 토크나이저
class Lexer:
    def __init__(self, data: bytes, pos: int = 0):
        self.d, self.i = data, pos

    def skip(self):
        while self.i < len(self.d):
            c = self.d[self.i]
            if c in WS:
                self.i += 1
            elif c == 0x25:                       # % 주석
                while self.i < len(self.d) and self.d[self.i] not in b"\r\n":
                    self.i += 1
            else:
                return

    def token(self):
        """다음 토큰: bytes(연산자/키워드) 또는 파이썬 값."""
        self.skip()
        if self.i >= len(self.d):
            return None
        c = self.d[self.i]
        if c == 0x2F:                             # /Name
            self.i += 1
            start = self.i
            while self.i < len(self.d) and self.d[self.i] not in WS and self.d[self.i] not in DELIM:
                self.i += 1
            raw = self.d[start:self.i]
            return Name(re.sub(rb"#([0-9A-Fa-f]{2})",
                               lambda m: bytes([int(m.group(1), 16)]), raw).decode("latin-1"))
        if c == 0x28:                             # (문자열)
            return self._literal()
        if c == 0x3C:                             # <hex> 또는 <<dict>>
            if self.d[self.i:self.i + 2] == b"<<":
                self.i += 2
                return b"<<"
            return self._hex()
        if self.d[self.i:self.i + 2] == b">>":
            self.i += 2
            return b">>"
        if c in b"[]{}":
            self.i += 1
            return bytes([c])
        start = self.i
        while self.i < len(self.d) and self.d[self.i] not in WS and self.d[self.i] not in DELIM:
            self.i += 1
        if self.i == start:
            self.i += 1
            return bytes([c])
        tok = self.d[start:self.i]
        if re.fullmatch(rb"[+-]?\d+", tok):
            return int(tok)
        if re.fullmatch(rb"[+-]?(\d*\.\d*|\d+)", tok) and tok not in (b".", b"-", b"+"):
            try:
                return float(tok)
            except ValueError:
                return tok
        return tok                                 # 연산자·키워드

    def _literal(self) -> "PdfString":
        self.i += 1
        out, depth = bytearray(), 1
        while self.i < len(self.d):
            c = self.d[self.i]
            if c == 0x5C:                         # 역슬래시 이스케이프
                self.i += 1
                if self.i >= len(self.d):
                    break
                e = self.d[self.i]
                mapping = {0x6E: 10, 0x72: 13, 0x74: 9, 0x62: 8, 0x66: 12}
                if e in mapping:
                    out.append(mapping[e])
                    self.i += 1
                elif 0x30 <= e <= 0x37:
                    oct_ = 0
                    for _ in range(3):
                        if self.i < len(self.d) and 0x30 <= self.d[self.i] <= 0x37:
                            oct_ = oct_ * 8 + (self.d[self.i] - 0x30)
                            self.i += 1
                        else:
                            break
                    out.append(oct_ & 0xFF)
                elif e in b"\r\n":
                    self.i += 1
                    if self.i < len(self.d) and self.d[self.i - 1] == 13 and self.d[self.i] == 10:
                        self.i += 1
                else:
                    out.append(e)
                    self.i += 1
                continue
            if c == 0x28:
                depth += 1
            elif c == 0x29:
                depth -= 1
                if depth == 0:
                    self.i += 1
                    break
            out.append(c)
            self.i += 1
        return PdfString(out)

    def _hex(self) -> bytes:
        self.i += 1
        start = self.i
        while self.i < len(self.d) and self.d[self.i] != 0x3E:
            self.i += 1
        raw = re.sub(rb"[^0-9A-Fa-f]", b"", self.d[start:self.i])
        self.i += 1
        if len(raw) % 2:
            raw += b"0"
        return PdfString(bytes.fromhex(raw.decode("ascii")))


def parse_object(lex: Lexer, depth: int = 0):
    tok = lex.token()
    return _parse_from(lex, tok, depth)


def _parse_from(lex: Lexer, tok, depth: int = 0):
    if depth > MAX_DEPTH:
        raise PdfError("객체 중첩이 너무 깊다")
    if tok is None:
        return None
    if tok == b"<<":
        d = {}
        while True:
            k = lex.token()
            if k is None or k == b">>":
                break
            if not isinstance(k, Name):
                continue
            d[str(k)] = _parse_from(lex, lex.token(), depth + 1)
        save = lex.i
        nxt = lex.token()
        if nxt == b"stream":
            if lex.d[lex.i:lex.i + 2] == b"\r\n":
                lex.i += 2
            elif lex.i < len(lex.d) and lex.d[lex.i] in b"\n\r":
                lex.i += 1
            ln = d.get("Length")
            start = lex.i
            if isinstance(ln, int) and 0 <= ln <= len(lex.d) - start:
                raw = lex.d[start:start + ln]
                tail = lex.d[start + ln:start + ln + 20]
                if b"endstream" not in tail:      # Length가 틀린 파일이 많다
                    raw = None
            else:
                raw = None
            if raw is None:
                end = lex.d.find(b"endstream", start)
                raw = lex.d[start:end if end >= 0 else len(lex.d)].rstrip(b"\r\n")
                lex.i = (end if end >= 0 else len(lex.d)) + 9
            else:
                lex.i = start + ln
                e = lex.d.find(b"endstream", lex.i)
                lex.i = (e + 9) if e >= 0 else lex.i
            return Stream(d, raw)
        lex.i = save
        return d
    if tok == b"[":
        arr = []
        while True:
            t = lex.token()
            if t is None or t == b"]":
                break
            arr.append(_parse_from(lex, t, depth + 1))
        _collapse_refs(arr)
        return arr
    if type(tok) is bytes:
        if tok == b"true":
            return True
        if tok == b"false":
            return False
        if tok == b"null":
            return None
        return tok
    if isinstance(tok, int):
        save = lex.i
        t2 = lex.token()
        if isinstance(t2, int):
            save2 = lex.i
            t3 = lex.token()
            if t3 == b"R":
                return Ref(tok, t2)
            lex.i = save2
            lex.i = save
            return tok
        lex.i = save
        return tok
    return tok


def _collapse_refs(arr: list) -> None:
    """배열 안에서 'n g R' 가 세 토큰으로 남은 경우를 Ref로 합친다."""
    i = 0
    while i + 2 < len(arr):
        if isinstance(arr[i], int) and isinstance(arr[i + 1], int) and arr[i + 2] == b"R":
            arr[i:i + 3] = [Ref(arr[i], arr[i + 1])]
        i += 1


# ---------------------------------------------------------------- 필터
def apply_predictor(data: bytes, parms: dict) -> bytes:
    pred = parms.get("Predictor", 1)
    if not isinstance(pred, int) or pred < 2:
        return data
    colors = parms.get("Colors", 1) or 1
    bpc = parms.get("BitsPerComponent", 8) or 8
    columns = parms.get("Columns", 1) or 1
    bpp = max(1, (colors * bpc + 7) // 8)
    rowlen = (columns * colors * bpc + 7) // 8
    if pred == 2:
        return data
    out = bytearray()
    prev = bytearray(rowlen)
    i = 0
    while i + 1 + rowlen <= len(data) + rowlen and i < len(data):
        ft = data[i]
        row = bytearray(data[i + 1:i + 1 + rowlen])
        i += 1 + rowlen
        if len(row) < rowlen:
            row.extend(b"\x00" * (rowlen - len(row)))
        for j in range(rowlen):
            a = row[j - bpp] if j >= bpp else 0
            b = prev[j]
            c = prev[j - bpp] if j >= bpp else 0
            x = row[j]
            if ft == 0:
                v = x
            elif ft == 1:
                v = x + a
            elif ft == 2:
                v = x + b
            elif ft == 3:
                v = x + (a + b) // 2
            elif ft == 4:
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                v = x + (a if pa <= pb and pa <= pc else (b if pb <= pc else c))
            else:
                v = x
            row[j] = v & 0xFF
        out.extend(row)
        prev = row
    return bytes(out)


def ascii85(data: bytes) -> bytes:
    data = re.sub(rb"\s", b"", data)
    if data.startswith(b"<~"):
        data = data[2:]
    end = data.find(b"~>")
    if end >= 0:
        data = data[:end]
    out = bytearray()
    i = 0
    while i < len(data):
        if data[i:i + 1] == b"z":
            out += b"\x00\x00\x00\x00"
            i += 1
            continue
        chunk = data[i:i + 5]
        i += 5
        pad = 5 - len(chunk)
        chunk = chunk + b"u" * pad
        v = 0
        for ch in chunk:
            v = v * 85 + (ch - 33)
        four = struct.pack(">I", v & 0xFFFFFFFF)
        out += four[:4 - pad]
    return bytes(out)


def runlength(data: bytes) -> bytes:
    out = bytearray()
    i = 0
    while i < len(data):
        n = data[i]
        i += 1
        if n == 128:
            break
        if n < 128:
            out += data[i:i + n + 1]
            i += n + 1
        else:
            if i < len(data):
                out += bytes([data[i]]) * (257 - n)
                i += 1
    return bytes(out)


IMAGE_FILTERS = {"DCTDecode", "JPXDecode", "JBIG2Decode", "CCITTFaxDecode"}


def decode_stream(st: Stream, resolve) -> tuple[bytes, str | None]:
    """(내용, 남은 이미지 필터 이름). 이미지 필터는 풀지 않고 이름만 돌려준다."""
    data = st.raw
    filters = resolve(st.dict.get("Filter"))
    if filters is None:
        return data, None
    if isinstance(filters, (Name, str)):
        filters = [filters]
    parms = resolve(st.dict.get("DecodeParms")) or resolve(st.dict.get("DP"))
    if not isinstance(parms, list):
        parms = [parms]
    for idx, f in enumerate(filters):
        f = str(resolve(f))
        pm = resolve(parms[idx]) if idx < len(parms) else None
        pm = {k: resolve(v) for k, v in pm.items()} if isinstance(pm, dict) else {}
        if f in ("FlateDecode", "Fl"):
            try:
                data = zlib.decompress(data)
            except zlib.error:
                try:
                    data = zlib.decompressobj().decompress(data)
                except zlib.error as exc:
                    raise PdfError("FlateDecode 실패: %s" % exc)
            data = apply_predictor(data, pm)
        elif f in ("ASCIIHexDecode", "AHx"):
            hx = re.sub(rb"[^0-9A-Fa-f]", b"", data.split(b">")[0])
            data = bytes.fromhex((hx + b"0" if len(hx) % 2 else hx).decode("ascii"))
        elif f in ("ASCII85Decode", "A85"):
            data = ascii85(data)
        elif f in ("RunLengthDecode", "RL"):
            data = runlength(data)
        elif f in IMAGE_FILTERS:
            return data, f
        else:
            raise PdfError("지원하지 않는 필터: %s" % f)
    return data, None


# ---------------------------------------------------------------- 문서
OBJ_RE = re.compile(rb"(?<![0-9])(\d{1,10})\s+(\d{1,5})\s+obj\b")
ENCRYPT_RE = re.compile(rb"/Encrypt\s+\d+\s+\d+\s+R")


class Pdf:
    def __init__(self, data: bytes, max_objects: int = 500000):
        if not data.lstrip()[:5].startswith(b"%PDF-"):
            raise PdfError("PDF 시그니처가 아니다")
        self.data = data
        self.version = data[5:8].decode("latin-1", "replace")
        self.encrypted = bool(ENCRYPT_RE.search(data))
        self.offsets: dict[int, int] = {}
        for m in OBJ_RE.finditer(data):
            if len(self.offsets) > max_objects:
                raise PdfError("객체 수 한도 초과")
            self.offsets[int(m.group(1))] = m.start()
        self._cache: dict[int, object] = {}
        self._objstm_loaded: set[int] = set()

    # -- 객체 접근 --
    def obj(self, num: int):
        if num in self._cache:
            return self._cache[num]
        off = self.offsets.get(num)
        if off is None:
            self._load_objstms()
            return self._cache.get(num)
        lex = Lexer(self.data, off)
        lex.token(), lex.token(), lex.token()      # num gen obj
        try:
            val = parse_object(lex)
        except (PdfError, ValueError):
            val = None
        self._cache[num] = val
        return val

    def resolve(self, o, depth: int = 0):
        while isinstance(o, Ref) and depth < 32:
            o = self.obj(o.num)
            depth += 1
        return o

    def _load_objstms(self) -> None:
        """압축 객체 스트림(ObjStm) 안의 객체를 펼친다."""
        for num in list(self.offsets):
            if num in self._objstm_loaded:
                continue
            self._objstm_loaded.add(num)
            o = self.obj(num)
            if not isinstance(o, Stream) or str(self.resolve(o.dict.get("Type"))) != "ObjStm":
                continue
            try:
                data, img = decode_stream(o, self.resolve)
            except PdfError:
                continue
            if img:
                continue
            n = self.resolve(o.dict.get("N")) or 0
            first = self.resolve(o.dict.get("First")) or 0
            head = Lexer(data[:first])
            pairs = []
            for _ in range(int(n)):
                a, b = head.token(), head.token()
                if not isinstance(a, int) or not isinstance(b, int):
                    break
                pairs.append((a, b))
            for onum, rel in pairs:
                if onum in self._cache:
                    continue
                try:
                    self._cache[onum] = parse_object(Lexer(data, first + rel))
                except (PdfError, ValueError):
                    pass

    def catalog(self) -> dict | None:
        self._load_objstms()
        best = None
        for num in sorted(set(list(self.offsets) + list(self._cache))):
            o = self.resolve(self.obj(num))
            if isinstance(o, dict) and str(self.resolve(o.get("Type"))) == "Catalog":
                best = o
        return best

    def pages(self) -> list[dict]:
        """페이지 사전을 문서 순서대로. Pages 트리를 못 쓰면 Type/Page를 번호순으로."""
        cat = self.catalog()
        out: list[dict] = []
        seen: set[int] = set()

        def walk(node, depth=0):
            node = self.resolve(node)
            if not isinstance(node, dict) or depth > MAX_DEPTH:
                return
            t = str(self.resolve(node.get("Type")) or "")
            if t == "Page":
                out.append(node)
                return
            kids = self.resolve(node.get("Kids"))
            if isinstance(kids, list):
                for k in kids:
                    key = k.num if isinstance(k, Ref) else id(k)
                    if key in seen:
                        continue
                    seen.add(key)
                    walk(k, depth + 1)

        if cat is not None:
            walk(cat.get("Pages"))
        if not out:
            self._load_objstms()
            for num in sorted(set(list(self.offsets) + list(self._cache))):
                o = self.resolve(self.obj(num))
                if isinstance(o, dict) and str(self.resolve(o.get("Type"))) == "Page":
                    out.append(o)
        return out

    def inherited(self, page: dict, key: str, depth: int = 0):
        node = page
        while isinstance(node, dict) and depth < MAX_DEPTH:
            if key in node:
                return self.resolve(node[key])
            node = self.resolve(node.get("Parent"))
            depth += 1
        return None


# ---------------------------------------------------------------- 폰트/인코딩
STD_ENC = {0o101: "A"}          # 자리표시: 실제 표는 아래 LATIN으로 대체
LATIN = {i: chr(i) for i in range(32, 127)}
GLYPH_RE = re.compile(r"^uni([0-9A-Fa-f]{4,6})$")
NAMED = {"space": " ", "period": ".", "comma": ",", "hyphen": "-", "colon": ":",
         "semicolon": ";", "slash": "/", "parenleft": "(", "parenright": ")",
         "quotesingle": "'", "quotedbl": '"', "percent": "%", "plus": "+", "equal": "=",
         "asterisk": "*", "ampersand": "&", "question": "?", "exclam": "!",
         "bracketleft": "[", "bracketright": "]", "braceleft": "{", "braceright": "}",
         "less": "<", "greater": ">", "at": "@", "numbersign": "#", "dollar": "$",
         "underscore": "_", "bar": "|", "asciitilde": "~", "grave": "`", "backslash": "\\",
         "zero": "0", "one": "1", "two": "2", "three": "3", "four": "4", "five": "5",
         "six": "6", "seven": "7", "eight": "8", "nine": "9", "endash": "–",
         "emdash": "—", "quoteleft": "‘", "quoteright": "’",
         "quotedblleft": "“", "quotedblright": "”", "bullet": "•"}


def glyph_to_char(name: str) -> str | None:
    m = GLYPH_RE.match(name)
    if m:
        try:
            return chr(int(m.group(1), 16))
        except ValueError:
            return None
    if name in NAMED:
        return NAMED[name]
    if len(name) == 1:
        return name
    return None


CMAP_CHAR = re.compile(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]*)>")
CMAP_RANGE = re.compile(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]*)>")


def parse_cmap(data: bytes) -> tuple[dict, set]:
    """ToUnicode CMap → ({코드: 문자열}, {코드 바이트 길이})."""
    table: dict[int, str] = {}
    widths: set[int] = set()
    for m in re.finditer(rb"begincodespacerange(.*?)endcodespacerange", data, re.S):
        for lo, _hi in CMAP_CHAR.findall(m.group(1)):
            widths.add(max(1, len(lo) // 2))
    for m in re.finditer(rb"beginbfchar(.*?)endbfchar", data, re.S):
        for src, dst in CMAP_CHAR.findall(m.group(1)):
            widths.add(max(1, len(src) // 2))
            table[int(src, 16)] = _utf16(dst)
    for m in re.finditer(rb"beginbfrange(.*?)endbfrange", data, re.S):
        body = m.group(1)
        for lo, hi, dst in CMAP_RANGE.findall(body):
            widths.add(max(1, len(lo) // 2))
            a, b = int(lo, 16), int(hi, 16)
            if b - a > 65535 or not dst:
                continue
            base = _utf16(dst)
            for k in range(a, b + 1):
                if base and len(base) == 1:
                    table[k] = chr(ord(base[0]) + (k - a))
                else:
                    table[k] = base
        for m2 in re.finditer(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*\[(.*?)\]", body, re.S):
            a = int(m2.group(1), 16)
            widths.add(max(1, len(m2.group(1)) // 2))
            for off, dm in enumerate(re.findall(rb"<([0-9A-Fa-f]*)>", m2.group(3))):
                table[a + off] = _utf16(dm)
    return table, widths


def _utf16(h: bytes) -> str:
    try:
        raw = bytes.fromhex((h + b"0" if len(h) % 2 else h).decode("ascii"))
    except ValueError:
        return ""
    if len(raw) >= 2:
        try:
            return raw.decode("utf-16-be")
        except UnicodeDecodeError:
            return raw.decode("latin-1")
    return raw.decode("latin-1")


class Font:
    def __init__(self, doc: Pdf, fd: dict):
        self.two_byte = False
        self.table: dict[int, str] = {}
        self.diff: dict[int, str] = {}
        self.has_tounicode = False
        self._ttf: dict | None = None
        self.ttf_used = 0
        self.subtype = str(doc.resolve(fd.get("Subtype")) or "")
        enc = doc.resolve(fd.get("Encoding"))
        if self.subtype == "Type0":
            self.two_byte = True
            if isinstance(enc, (Name, str)) and "Identity" not in str(enc):
                pass
        tu = doc.resolve(fd.get("ToUnicode"))
        if isinstance(tu, Stream):
            try:
                data, img = decode_stream(tu, doc.resolve)
                if not img:
                    self.table, widths = parse_cmap(data)
                    self.has_tounicode = bool(self.table)
                    if widths == {2}:
                        self.two_byte = True
            except PdfError:
                pass
        self._doc, self._fd = doc, fd     # 내장 폰트 cmap은 필요할 때만 읽는다(비싸다)
        if isinstance(enc, dict):
            diffs = doc.resolve(enc.get("Differences"))
            if isinstance(diffs, list):
                code = 0
                for item in diffs:
                    item = doc.resolve(item)
                    if isinstance(item, (int, float)):
                        code = int(item)
                    elif isinstance(item, Name):
                        ch = glyph_to_char(str(item))
                        if ch:
                            self.diff[code] = ch
                        code += 1

    def _ttf_table(self) -> dict:
        """ToUnicode에 빠진 코드를 내장 TrueType cmap으로 보충한다
        (Identity-H에서는 코드가 곧 글리프 번호다)."""
        if self._ttf is None:
            prog = embedded_font_file(self._doc, self._fd)
            self._ttf = ttf_gid_to_unicode(prog) if prog else {}
        return self._ttf

    def decode(self, raw: bytes) -> tuple[str, int, int]:
        """(문자열, 코드 수, 매핑 실패 수)"""
        out, n, bad = [], 0, 0
        step = 2 if self.two_byte else 1
        for i in range(0, len(raw) - (step - 1), step):
            code = raw[i] if step == 1 else (raw[i] << 8) | raw[i + 1]
            n += 1
            if code in self.table:
                out.append(self.table[code])
            elif code in self.diff:
                out.append(self.diff[code])
            elif code in self._ttf_table():
                out.append(self._ttf_table()[code])
                self.ttf_used += 1
            elif step == 1 and 32 <= code < 127:
                out.append(chr(code))
            elif step == 1 and code >= 160:
                out.append(bytes([code]).decode("cp1252", "replace"))
            else:
                out.append("�")
                bad += 1
        return "".join(out), n, bad


# ---------------------------------------------------------------- 본문 추출
TEXT_OPS = {b"Tj", b"TJ", b"'", b'"'}


def content_of(doc: Pdf, page: dict) -> bytes:
    c = doc.resolve(page.get("Contents"))
    parts = []
    for st in (c if isinstance(c, list) else [c]):
        st = doc.resolve(st)
        if isinstance(st, Stream):
            try:
                data, img = decode_stream(st, doc.resolve)
            except PdfError as exc:
                raise PdfError("페이지 내용 스트림을 풀지 못했다: %s" % exc)
            if not img:
                parts.append(data)
    return b"\n".join(parts)


class PageText:
    def __init__(self):
        self.parts: list[str] = []
        self.codes = 0
        self.unmapped = 0
        self.images: list = []
        self.no_font = 0
        self.fonts: set = set()


def extract_page(doc: Pdf, page: dict, limits_chars: int = 5_000_000) -> PageText:
    out = PageText()
    res = doc.inherited(page, "Resources") or {}
    _run(doc, content_of(doc, page), res, out, 0, limits_chars)
    return out


def _fonts_of(doc: Pdf, res: dict) -> dict:
    fonts = {}
    fd = doc.resolve(res.get("Font")) if isinstance(res, dict) else None
    if isinstance(fd, dict):
        for k, v in fd.items():
            v = doc.resolve(v)
            if isinstance(v, dict):
                fonts[k] = Font(doc, v)
    return fonts


def _run(doc: Pdf, content: bytes, res: dict, out: PageText, depth: int, cap: int) -> None:
    if depth > 8:
        return
    fonts = _fonts_of(doc, res)
    xobjs = doc.resolve(res.get("XObject")) if isinstance(res, dict) else None
    xobjs = xobjs if isinstance(xobjs, dict) else {}
    lex = Lexer(content)
    stack: list = []
    font = None
    leading = 0.0
    x = y = 0.0
    last_y = None
    last_x = None
    size = 0.0

    def emit(s: str):
        if sum(len(p) for p in out.parts) < cap:
            out.parts.append(s)

    def show(raw: bytes):
        nonlocal last_x
        if font is None:
            out.no_font += 1
            txt = raw.decode("latin-1", "replace")
            n, bad = len(raw), 0
        else:
            txt, n, bad = font.decode(raw)
        out.codes += n
        out.unmapped += bad
        emit(txt)

    while True:
        try:
            tok = lex.token()
        except (PdfError, ValueError):
            break
        if tok is None:
            break
        if type(tok) is bytes and tok in (b"<<", b"["):
            stack.append(_parse_from(lex, tok))
            continue
        if type(tok) is not bytes or tok in (b"]", b">>"):
            stack.append(tok)
            if len(stack) > 64:
                del stack[:-32]
            continue
        op = tok
        if op == b"Tf":
            if len(stack) >= 2:
                size = stack[-1] if isinstance(stack[-1], (int, float)) else size
                key = stack[-2]
                font = fonts.get(str(key)) if isinstance(key, Name) else font
                if isinstance(key, Name):
                    out.fonts.add(str(key))
        elif op == b"TL":
            leading = stack[-1] if stack and isinstance(stack[-1], (int, float)) else leading
        elif op in (b"Td", b"TD"):
            if len(stack) >= 2 and all(isinstance(v, (int, float)) for v in stack[-2:]):
                x += stack[-2]
                y += stack[-1]
                if op == b"TD":
                    leading = -stack[-1]
        elif op == b"Tm":
            if len(stack) >= 6 and all(isinstance(v, (int, float)) for v in stack[-6:]):
                x, y = stack[-2], stack[-1]
        elif op == b"T*":
            y -= leading
        elif op == b"BT":
            x = y = 0.0
            last_y = None
        if op in TEXT_OPS:
            # 줄이 바뀌면 줄바꿈, 같은 줄에서 많이 건너뛰면 빈칸을 넣는다(위치 기반 추정).
            if last_y is not None and abs(y - last_y) > max(0.5, size * 0.3):
                emit("\n")
            elif last_x is not None and x - last_x > max(2.0, size * 0.9):
                emit(" ")
            last_y, last_x = y, x
            if op == b"Tj" and stack and isinstance(stack[-1], PdfString):
                show(stack[-1])
            elif op in (b"'", b'"') and stack and isinstance(stack[-1], PdfString):
                emit("\n")
                show(stack[-1])
            elif op == b"TJ" and stack and isinstance(stack[-1], list):
                for item in stack[-1]:
                    if isinstance(item, PdfString):
                        show(item)
                    elif isinstance(item, (int, float)) and item < -170:
                        emit(" ")
        elif op == b"Do":
            name = stack[-1] if stack else None
            xo = doc.resolve(xobjs.get(str(name))) if isinstance(name, Name) else None
            if isinstance(xo, Stream):
                sub = str(doc.resolve(xo.dict.get("Subtype")) or "")
                if sub == "Form":
                    try:
                        data, img = decode_stream(xo, doc.resolve)
                        if not img:
                            _run(doc, data, doc.resolve(xo.dict.get("Resources")) or res,
                                 out, depth + 1, cap)
                    except PdfError:
                        pass
                elif sub == "Image":
                    out.images.append((str(name), xo))
        stack.clear()        # 연산자가 피연산자를 모두 소비한다


def image_bytes(doc: Pdf, st: Stream) -> tuple[bytes, str]:
    """이미지 XObject의 원본 바이트와 확장자. 푸는 대신 있는 그대로 보존한다."""
    try:
        data, img = decode_stream(st, doc.resolve)
    except PdfError:
        return st.raw, "bin"
    ext = {"DCTDecode": "jpg", "JPXDecode": "jp2", "JBIG2Decode": "jbig2",
           "CCITTFaxDecode": "tif"}.get(img or "", None)
    if ext:
        return data, ext
    return data, "raw"


# ---------------------------------------------------------------- 내장 TrueType cmap
def ttf_gid_to_unicode(data: bytes) -> dict[int, str]:
    """내장 TrueType의 cmap을 뒤집어 글리프 번호 → 유니코드 표를 만든다.

    ToUnicode가 없는 Identity-H 한글 PDF에서 코드는 곧 글리프 번호라서,
    이 표만 있으면 글자를 되살릴 수 있다."""
    out: dict[int, str] = {}
    try:
        if data[:4] == b"ttcf":
            off = struct.unpack_from(">I", data, 12)[0]
        else:
            off = 0
        num = struct.unpack_from(">H", data, off + 4)[0]
        tables = {}
        for i in range(num):
            rec = off + 12 + i * 16
            tag = data[rec:rec + 4].decode("latin-1")
            toff, tlen = struct.unpack_from(">II", data, rec + 8)
            tables[tag] = (toff, tlen)
        if "cmap" not in tables:
            return out
        cm = tables["cmap"][0]
        n = struct.unpack_from(">H", data, cm + 2)[0]
        best = None
        for i in range(n):
            pid, eid, sub = struct.unpack_from(">HHI", data, cm + 4 + i * 8)
            fmt = struct.unpack_from(">H", data, cm + sub)[0]
            score = {(3, 10): 4, (3, 1): 3, (0, 4): 3, (0, 3): 2, (3, 0): 1}.get((pid, eid), 0)
            if fmt in (4, 12) and score and (best is None or score > best[0]):
                best = (score, cm + sub, fmt, pid, eid)
        if best is None:
            return out
        _, base, fmt, pid, eid = best
        if fmt == 4:
            segx2 = struct.unpack_from(">H", data, base + 6)[0]
            seg = segx2 // 2
            ends = struct.unpack_from(">%dH" % seg, data, base + 14)
            starts = struct.unpack_from(">%dH" % seg, data, base + 16 + segx2)
            deltas = struct.unpack_from(">%dh" % seg, data, base + 16 + segx2 * 2)
            ro_base = base + 16 + segx2 * 3
            ros = struct.unpack_from(">%dH" % seg, data, ro_base)
            for i in range(seg):
                if starts[i] > ends[i] or ends[i] == 0xFFFF and starts[i] == 0xFFFF:
                    continue
                for c in range(starts[i], min(ends[i], 0xFFFE) + 1):
                    if ros[i] == 0:
                        gid = (c + deltas[i]) & 0xFFFF
                    else:
                        gi = ro_base + i * 2 + ros[i] + (c - starts[i]) * 2
                        if gi + 2 > len(data):
                            continue
                        gid = struct.unpack_from(">H", data, gi)[0]
                        if gid:
                            gid = (gid + deltas[i]) & 0xFFFF
                    ch = c - 0xF000 if pid == 3 and eid == 0 and 0xF000 <= c <= 0xF0FF else c
                    if gid and gid not in out:
                        out[gid] = chr(ch)
        else:
            ngroups = struct.unpack_from(">I", data, base + 12)[0]
            for i in range(min(ngroups, 200000)):
                s, e, g = struct.unpack_from(">III", data, base + 16 + i * 12)
                if e < s or e - s > 65535:
                    continue
                for k in range(s, e + 1):
                    gid = g + (k - s)
                    if gid not in out and k <= 0x10FFFF:
                        out[gid] = chr(k)
    except (struct.error, IndexError, ValueError, UnicodeDecodeError):
        return out
    return out


def embedded_font_file(doc: Pdf, fd: dict):
    """폰트 사전에서 내장 TrueType 프로그램(FontFile2)을 찾는다."""
    cand = [fd]
    desc = doc.resolve(fd.get("DescendantFonts"))
    if isinstance(desc, list):
        cand += [doc.resolve(d) for d in desc]
    for f in cand:
        if not isinstance(f, dict):
            continue
        d = doc.resolve(f.get("FontDescriptor"))
        if not isinstance(d, dict):
            continue
        ff = doc.resolve(d.get("FontFile2"))
        if isinstance(ff, Stream):
            try:
                data, img = decode_stream(ff, doc.resolve)
                if not img:
                    return data
            except PdfError:
                continue
    return None
