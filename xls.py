#!/usr/bin/env python3
"""구형 엑셀(.xls, BIFF8)에서 문자열만 직접 읽는다.

중간 변환기(LibreOffice)가 내용을 흘리지 않았는지 대조하려면 원본을 우리 손으로
한 번 읽어야 한다. 완전한 .xls 파서가 아니라 검사용 문자열 추출기다.

.xls 도 CFB(OLE) 복합 파일이라 hwp5 의 리더를 그대로 쓴다(같은 컨테이너 형식).
"""
from __future__ import annotations

import struct

import hwp5

BIFF_SST = 0x00FC
BIFF_CONTINUE = 0x003C
BIFF_BOF = 0x0809
BIFF_LABEL = 0x0204          # BIFF5 이하의 직접 문자열
BIFF_RSTRING = 0x00D6


class XlsError(Exception):
    pass


def workbook_stream(data: bytes) -> bytes:
    """.xls 안의 Workbook(또는 구버전 Book) 스트림."""
    try:
        streams = hwp5.Cfb(data).streams()
    except hwp5.HwpError as exc:
        raise XlsError("CFB로 읽지 못했다: %s" % exc) from exc
    for name in ("Workbook", "Book"):
        if name in streams:
            return streams[name]
    raise XlsError("Workbook 스트림이 없다. .xls가 아닐 수 있다.")


def _records(data: bytes):
    i, n = 0, len(data)
    while i + 4 <= n:
        rid, ln = struct.unpack_from("<HH", data, i)
        i += 4
        if i + ln > n:
            break
        yield rid, data[i:i + ln]
        i += ln


def _read_unicode(chunks: list, pos: list, cch: int, flags: int) -> str:
    """BIFF8 문자열 본문. CONTINUE 경계에서 인코딩 플래그가 다시 나온다."""
    out = []
    left = cch
    wide = bool(flags & 0x01)
    while left > 0:
        buf = chunks[pos[0]]
        off = pos[1]
        avail = len(buf) - off
        if avail <= 0:
            pos[0] += 1
            pos[1] = 0
            if pos[0] >= len(chunks):
                break
            buf = chunks[pos[0]]
            wide = bool(buf[0] & 0x01)        # 이어지는 조각의 새 플래그
            pos[1] = 1
            continue
        step = 2 if wide else 1
        take = min(left, avail // step)
        if take <= 0:
            pos[0] += 1
            pos[1] = 0
            continue
        raw = buf[off:off + take * step]
        out.append(raw.decode("utf-16-le" if wide else "latin-1", "replace"))
        pos[1] = off + take * step
        left -= take
    return "".join(out)


def sst_strings(data: bytes, limit: int = 200000) -> list[str]:
    """공유 문자열 표(SST)의 문자열. 시트 글자 대부분이 여기에 모여 있다."""
    wb = workbook_stream(data)
    recs = list(_records(wb))
    out: list[str] = []
    for idx, (rid, payload) in enumerate(recs):
        if rid != BIFF_SST:
            continue
        chunks = [payload]
        for rid2, more in recs[idx + 1:]:
            if rid2 != BIFF_CONTINUE:
                break
            chunks.append(more)
        if len(chunks[0]) < 8:
            break
        unique = struct.unpack_from("<I", chunks[0], 4)[0]
        pos = [0, 8]
        for _ in range(min(unique, limit)):
            while pos[0] < len(chunks) and pos[1] + 3 > len(chunks[pos[0]]):
                pos[0] += 1
                pos[1] = 0
            if pos[0] >= len(chunks):
                break
            buf = chunks[pos[0]]
            cch = struct.unpack_from("<H", buf, pos[1])[0]
            flags = buf[pos[1] + 2]
            pos[1] += 3
            nruns, extlen = 0, 0
            if flags & 0x08:
                nruns = struct.unpack_from("<H", buf, pos[1])[0]
                pos[1] += 2
            if flags & 0x04:
                extlen = struct.unpack_from("<i", buf, pos[1])[0]
                pos[1] += 4
            out.append(_read_unicode(chunks, pos, cch, flags))
            for skip in (nruns * 4, max(0, extlen)):
                left = skip
                while left > 0 and pos[0] < len(chunks):
                    avail = len(chunks[pos[0]]) - pos[1]
                    if avail <= 0:
                        pos[0] += 1
                        pos[1] = 0
                        continue
                    take = min(left, avail)
                    pos[1] += take
                    left -= take
        break
    # BIFF5 이하는 SST 없이 LABEL 레코드에 글자를 담는다
    for rid, payload in recs:
        if rid in (BIFF_LABEL, BIFF_RSTRING) and len(payload) > 8:
            cch = struct.unpack_from("<H", payload, 6)[0]
            body = payload[8:8 + cch]
            if body:
                out.append(body.decode("cp949", "replace"))
    return [s for s in out if s.strip()]


BIFF_BOUNDSHEET = 0x0085


def sheet_names(data: bytes) -> list[str]:
    """BOUNDSHEET 레코드의 시트 이름. 중간 변환이 시트를 통째로 흘렸는지 본다."""
    out = []
    for rid, payload in _records(workbook_stream(data)):
        if rid != BIFF_BOUNDSHEET or len(payload) < 8:
            continue
        cch = payload[6]
        flags = payload[7]
        body = payload[8:]
        if flags & 0x01:
            out.append(body[:cch * 2].decode("utf-16-le", "replace"))
        else:
            out.append(body[:cch].decode("cp949", "replace"))
    return [s for s in out if s.strip()]
