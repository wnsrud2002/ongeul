#!/usr/bin/env python3
"""HWP 5.x 저수준 리더: CFB(OLE) 복합 파일과 HWP 레코드 스트림만 다룬다.

mdmaker.py를 import하지 않는다(순환 참조 방지). Markdown 변환은 mdmaker 쪽 책임이다.
참고 구조: CFB는 MS-CFB, 레코드/문자 규칙은 한컴 HWP 5.0 포맷 공개 문서.
보호(암호·배포용) 문서는 해제하지 않고 식별만 한다.
"""
from __future__ import annotations

import struct
import zlib

CFB_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
HWP_SIGNATURE = b"HWP Document File"
ENDOFCHAIN, FREESECT = 0xFFFFFFFE, 0xFFFFFFFF
MAX_CHAIN = 1 << 22            # 손상 파일의 순환 체인 방어


class HwpError(Exception):
    pass


# ---------------------------------------------------------------- CFB
class Cfb:
    def __init__(self, data: bytes, max_stream: int = 512 * 1024 * 1024):
        if not data.startswith(CFB_SIGNATURE):
            raise HwpError("CFB 시그니처가 아니다")
        self.d = data
        self.max_stream = max_stream
        (self.ssz, self.msz) = (1 << struct.unpack_from("<H", data, 30)[0],
                                1 << struct.unpack_from("<H", data, 32)[0])
        if self.ssz not in (512, 4096) or self.msz != 64:
            raise HwpError("지원하지 않는 섹터 크기: %d/%d" % (self.ssz, self.msz))
        (self.n_dir, self.n_fat, self.dir_start, _, self.cutoff,
         self.mini_start, self.n_mini, self.difat_start, self.n_difat) = struct.unpack_from(
            "<IIIIIIIII", data, 40)
        self.fat = self._read_fat()
        self.dirs = self._read_dirs()
        self.mini = self._read_chain(self.dirs[0]["start"], self.dirs[0]["size"]) if self.dirs else b""
        self.minifat = self._read_minifat()

    def _sector(self, n: int) -> bytes:
        off = (n + 1) * self.ssz
        if off + self.ssz > len(self.d):
            raise HwpError("섹터 %d가 파일 범위를 벗어난다" % n)
        return self.d[off:off + self.ssz]

    def _read_fat(self) -> list[int]:
        nums = list(struct.unpack_from("<109I", self.d, 76))
        nxt, guard = self.difat_start, 0
        while nxt not in (ENDOFCHAIN, FREESECT) and guard < MAX_CHAIN:
            sec = self._sector(nxt)
            per = self.ssz // 4 - 1
            nums.extend(struct.unpack_from("<%dI" % per, sec, 0))
            nxt = struct.unpack_from("<I", sec, per * 4)[0]
            guard += 1
        fat: list[int] = []
        for n in nums:
            if n in (ENDOFCHAIN, FREESECT):
                continue
            fat.extend(struct.unpack_from("<%dI" % (self.ssz // 4), self._sector(n), 0))
        return fat

    def _read_chain(self, start: int, size: int | None = None) -> bytes:
        out, n, guard = bytearray(), start, 0
        while n not in (ENDOFCHAIN, FREESECT) and guard < MAX_CHAIN:
            out += self._sector(n)
            if size is not None and len(out) >= size:
                break
            if n >= len(self.fat):
                raise HwpError("FAT 범위를 벗어난 섹터 %d" % n)
            n = self.fat[n]
            guard += 1
            if len(out) > self.max_stream:
                raise HwpError("스트림 크기 한도 초과")
        return bytes(out[:size]) if size is not None else bytes(out)

    def _read_minifat(self) -> list[int]:
        if self.mini_start in (ENDOFCHAIN, FREESECT):
            return []
        raw = self._read_chain(self.mini_start)
        return list(struct.unpack_from("<%dI" % (len(raw) // 4), raw, 0))

    def _read_mini(self, start: int, size: int) -> bytes:
        out, n, guard = bytearray(), start, 0
        while n not in (ENDOFCHAIN, FREESECT) and len(out) < size and guard < MAX_CHAIN:
            off = n * self.msz
            out += self.mini[off:off + self.msz]
            if n >= len(self.minifat):
                break
            n = self.minifat[n]
            guard += 1
        return bytes(out[:size])

    def _read_dirs(self) -> list[dict]:
        raw = self._read_chain(self.dir_start,
                               self.n_dir * self.ssz if self.n_dir else None)
        entries = []
        for i in range(len(raw) // 128):
            e = raw[i * 128:(i + 1) * 128]
            nlen = struct.unpack_from("<H", e, 64)[0]
            name = e[:max(0, nlen - 2)].decode("utf-16-le", "replace")
            entries.append({"name": name, "type": e[66], "left": struct.unpack_from("<I", e, 68)[0],
                            "right": struct.unpack_from("<I", e, 72)[0],
                            "child": struct.unpack_from("<I", e, 76)[0],
                            "start": struct.unpack_from("<I", e, 116)[0],
                            "size": struct.unpack_from("<Q", e, 120)[0]})
        return entries

    def streams(self) -> dict[str, bytes]:
        """경로 -> 원본 바이트. 이름은 'BodyText/Section0' 형태다."""
        out: dict[str, bytes] = {}
        seen: set[int] = set()

        def walk(idx: int, prefix: str) -> None:
            if idx in (ENDOFCHAIN, FREESECT) or idx >= len(self.dirs) or idx in seen:
                return
            seen.add(idx)
            e = self.dirs[idx]
            walk(e["left"], prefix)
            path = prefix + e["name"]
            if e["type"] == 2:                      # 스트림
                if e["size"] > self.max_stream:
                    raise HwpError("스트림 크기 한도 초과: %s" % path)
                out[path] = (self._read_mini(e["start"], e["size"])
                             if e["size"] < self.cutoff else
                             self._read_chain(e["start"], e["size"]))
            elif e["type"] == 1:                    # 스토리지
                walk(e["child"], path + "/")
            walk(e["right"], prefix)

        if self.dirs:
            walk(self.dirs[0]["child"], "")
        return out


# ---------------------------------------------------------------- HWP 레코드
HWPTAG_BEGIN = 0x10
TAG_BIN_DATA = HWPTAG_BEGIN + 2
TAG_STYLE = HWPTAG_BEGIN + 10
TAG_PARA_HEADER = HWPTAG_BEGIN + 50
TAG_PARA_TEXT = HWPTAG_BEGIN + 51
TAG_CTRL_HEADER = HWPTAG_BEGIN + 55
TAG_LIST_HEADER = HWPTAG_BEGIN + 56
TAG_TABLE = HWPTAG_BEGIN + 61
TAG_SHAPE_OLE = HWPTAG_BEGIN + 68        # 84
TAG_SHAPE_PICTURE = HWPTAG_BEGIN + 69    # 85 (+67=83은 CURVE다)
TAG_SHAPE_CONTAINER = HWPTAG_BEGIN + 70
TAG_EQEDIT = HWPTAG_BEGIN + 72
TAG_MEMO_SHAPE = HWPTAG_BEGIN + 76      # 92
TAG_MEMO_LIST = HWPTAG_BEGIN + 77       # 93

ONE_WCHAR = {0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31}
EXTENDED = {1, 2, 3, 11, 12, 14, 15, 16, 17, 18, 21, 22, 23}
CHAR_TEXT = {9: "\t", 10: "\n", 24: "-", 30: " ", 31: "　"}


class Header:
    def __init__(self, raw: bytes):
        if not raw.startswith(HWP_SIGNATURE):
            raise HwpError("HWP 문서 시그니처가 아니다")
        v = struct.unpack_from("<I", raw, 32)[0]
        self.version = (v >> 24 & 0xFF, v >> 16 & 0xFF, v >> 8 & 0xFF, v & 0xFF)
        p = struct.unpack_from("<I", raw, 36)[0]
        self.compressed = bool(p & 1)
        self.password = bool(p & 2)
        self.distribution = bool(p & 4)
        self.drm = bool(p & 0x10)
        self.history = bool(p & 0x40)
        self.signed = bool(p & 0x80)
        self.cert_encrypted = bool(p & 0x100)
        self.cert_drm = bool(p & 0x400)

    @property
    def protected(self) -> str | None:
        for flag, label in ((self.password, "암호 설정"), (self.distribution, "배포용"),
                            (self.drm, "DRM"), (self.cert_encrypted, "인증서 암호화"),
                            (self.cert_drm, "인증서 DRM")):
            if flag:
                return label
        return None


def decompress(data: bytes, compressed: bool) -> bytes:
    if not compressed:
        return data
    try:
        return zlib.decompressobj(-15).decompress(data)
    except zlib.error as exc:
        raise HwpError("스트림 압축 해제 실패: %s" % exc) from exc


def records(data: bytes):
    """(태그, 수준, payload) 순서대로 돌려준다."""
    i, n = 0, len(data)
    while i + 4 <= n:
        h = struct.unpack_from("<I", data, i)[0]
        i += 4
        tag, level, size = h & 0x3FF, (h >> 10) & 0x3FF, (h >> 20) & 0xFFF
        if size == 0xFFF:
            if i + 4 > n:
                break
            size = struct.unpack_from("<I", data, i)[0]
            i += 4
        if i + size > n:
            raise HwpError("레코드 크기가 스트림을 넘는다 (태그 %d)" % tag)
        yield tag, level, data[i:i + size]
        i += size


def tree(recs) -> list[dict]:
    """레코드의 level 값으로 부모-자식 구조를 만든다."""
    root: list = []
    stack: list = [(-1, root)]
    for tag, level, payload in recs:
        node = {"tag": tag, "level": level, "data": payload, "children": []}
        while len(stack) > 1 and stack[-1][0] >= level:
            stack.pop()
        stack[-1][1].append(node)
        stack.append((level, node["children"]))
    return root


def para_text(payload: bytes) -> list[tuple[str, object]]:
    """PARA_TEXT를 ('text', 글자) / ('ctrl', 코드) / ('inline', 코드) 토큰으로 나눈다.

    제어 문자는 1 WCHAR짜리와 16바이트(8 WCHAR)짜리가 섞여 있어 길이를 구분해야
    이후 글자가 통째로 밀린다."""
    out: list[tuple[str, object]] = []
    buf = bytearray()
    i, n = 0, len(payload)

    def flush():
        if buf:
            out.append(("text", buf.decode("utf-16-le", "replace")))
            buf.clear()

    while i + 1 < n:
        c = payload[i] | (payload[i + 1] << 8)
        if c >= 32:
            buf += payload[i:i + 2]
            i += 2
            continue
        flush()
        if c in ONE_WCHAR:
            if c in CHAR_TEXT:
                out.append(("text", CHAR_TEXT[c]))
            elif c == 13:
                out.append(("para", None))
            i += 2
        else:
            if c in CHAR_TEXT:
                out.append(("text", CHAR_TEXT[c]))
            out.append(("ctrl" if c in EXTENDED else "inline", c))
            i += 16
    flush()
    return out


def wstr(data: bytes, off: int) -> tuple[str, int]:
    """WORD 길이(문자 수) + UTF-16LE 문자열. (문자열, 다음 오프셋)"""
    if off + 2 > len(data):
        return "", off
    ln = struct.unpack_from("<H", data, off)[0]
    end = off + 2 + ln * 2
    return data[off + 2:end].decode("utf-16-le", "replace"), end


def ctrl_id(payload: bytes) -> str:
    """CTRL_HEADER의 앞 4바이트는 뒤집힌 ASCII 식별자다 ('tbl ', 'gso ', 'fn  ' 등)."""
    if len(payload) < 4:
        return ""
    return payload[:4][::-1].decode("ascii", "replace")


def bin_items(docinfo: bytes) -> list[dict]:
    """DocInfo의 BIN_DATA 레코드 순서가 곧 그림이 참조하는 1-based 색인이다."""
    items = []
    for tag, _lvl, p in records(docinfo):
        if tag != TAG_BIN_DATA:
            continue
        prop = struct.unpack_from("<H", p, 0)[0] if len(p) >= 2 else 0
        kind = prop & 0x0F
        item: dict = {"type": kind, "property": prop}
        off = 2
        if kind == 0:                                  # LINK
            item["abs_path"], off = wstr(p, off)
            item["rel_path"], off = wstr(p, off)
        else:                                          # EMBEDDING / STORAGE
            if off + 2 <= len(p):
                item["id"] = struct.unpack_from("<H", p, off)[0]
                off += 2
            if kind == 1:
                item["ext"], off = wstr(p, off)
        items.append(item)
    return items


def styles(docinfo: bytes) -> list[str]:
    out = []
    for tag, _lvl, p in records(docinfo):
        if tag == TAG_STYLE:
            name, _ = wstr(p, 0)
            out.append(name)
    return out


def picture_bin_id(payload: bytes, max_id: int = 0) -> int | None:
    """SHAPE_COMPONENT_PICTURE에서 BinItem 색인을 읽는다.

    테두리(12)+사각형(32)+자르기(16)+여백(8)=68 뒤에 밝기·명암·효과 1바이트씩이
    먼저 오므로 실제 위치는 71이다. 실제 한글 문서 382건에서 모두 71이 맞았고 68은
    전부 0이었다. 문서마다 다를 가능성에 대비해 68도 확인한다."""
    for off in (71, 68):
        if len(payload) >= off + 2:
            v = struct.unpack_from("<H", payload, off)[0]
            if v and (not max_id or v <= max_id):
                return v
    return None


def table_size(payload: bytes) -> tuple[int, int]:
    if len(payload) < 8:
        return 0, 0
    return struct.unpack_from("<H", payload, 4)[0], struct.unpack_from("<H", payload, 6)[0]


def cell_addr(payload: bytes) -> dict | None:
    """표 셀의 LIST_HEADER: paraCount(4) + property(4) 뒤에 셀 좌표가 온다."""
    if len(payload) < 16:
        return None
    col, row, cs, rs = struct.unpack_from("<HHHH", payload, 8)
    if max(col, row) > 4096 or cs > 4096 or rs > 4096:
        return None
    # 실제 문서에 span 0이 들어있는 경우가 있다(38건). 좌표 자체는 멀쩡하므로 1로 본다.
    return {"col": col, "row": row, "colspan": cs or 1, "rowspan": rs or 1}


def field_command(payload: bytes) -> str:
    """필드 컨트롤(%hlk 등)의 Command 문자열. ctrl id(4)+속성(4)+미지(1) 다음이다."""
    for off in (9, 8):
        if off + 2 <= len(payload):
            ln = struct.unpack_from("<H", payload, off)[0]
            if 0 < ln < 4096 and off + 2 + ln * 2 <= len(payload):
                return payload[off + 2:off + 2 + ln * 2].decode("utf-16-le", "replace")
    return ""


def equation_script(payload: bytes) -> str:
    """EQEDIT: property(4) 다음이 수식 문자열이다."""
    s, _ = wstr(payload, 4)
    return s
