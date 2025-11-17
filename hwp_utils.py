from __future__ import annotations

from io import BytesIO
import re
import struct
import zipfile
import zlib
from typing import Iterable, Sequence


def _parse_hwp_records(data: bytes) -> str:
    texts: list[str] = []
    offset = 0
    length = len(data)
    while offset + 4 <= length:
        try:
            header = struct.unpack('<I', data[offset:offset + 4])[0]
        except struct.error:
            break
        tag_id = header & 0x3FF
        size = (header >> 20) & 0xFFF
        offset += 4
        if size == 0xFFF:
            if offset + 4 > length:
                break
            size = struct.unpack('<I', data[offset:offset + 4])[0]
            offset += 4
        if size <= 0 or offset + size > length:
            break
        payload = data[offset:offset + size]
        offset += size
        if tag_id in {66, 67, 68, 80, 0x50}:
            for encoding in ('utf-16le', 'utf-8'):
                try:
                    text = payload.decode(encoding, errors='ignore')
                    break
                except Exception:
                    continue
            else:
                continue
            filtered = ''.join(
                '\n' if c in {'\r', '\v'} else c
                for c in text
                if c.isprintable() or c.isspace()
            )
            if filtered and len(filtered.strip()) > 0:
                texts.append(filtered)
    return '\n'.join(texts).strip()


def _iter_candidate_payloads(raw: bytes) -> Iterable[bytes]:
    yield raw
    try:
        yield zlib.decompress(raw, -zlib.MAX_WBITS)
    except Exception:
        pass
    try:
        yield zlib.decompress(raw)
    except Exception:
        pass


def extract_text_from_hwp_ole(file_bytes: bytes) -> tuple[str | None, str]:
    try:
        import olefile  # type: ignore
    except Exception as exc:  # pragma: no cover - exercised when olefile missing
        return None, f"OLE 모듈 없음: {exc}"

    try:
        ole = olefile.OleFileIO(BytesIO(file_bytes))
    except Exception as exc:
        return None, f"OLE 열기 실패: {exc}"

    texts: list[str] = []
    try:
        for path in ole.listdir():
            if not path or path[0] != 'BodyText':
                continue
            try:
                raw = ole.openstream(path).read()
            except Exception:
                continue
            for payload in _iter_candidate_payloads(raw):
                parsed = _parse_hwp_records(payload)
                if parsed:
                    texts.append(parsed)
                    break
    finally:
        try:
            ole.close()
        except Exception:
            pass

    merged = '\n'.join(t for t in texts if t).strip()
    return (merged, 'OK[OLE]') if merged else (None, 'OLE 추출 실패')


def convert_hwp_local_to_text(file_bytes: bytes) -> tuple[str | None, str]:
    text, status = extract_text_from_hwp_ole(file_bytes)
    if text:
        return text, status
    return None, status


def extract_text_from_hwpx_bytes(file_bytes: bytes) -> tuple[str | None, str]:
    """Parse the XML payload inside a HWPX archive without external tools."""

    try:
        texts: list[str] = []
        with zipfile.ZipFile(BytesIO(file_bytes)) as zf:
            for name in zf.namelist():
                if not name.lower().endswith('.xml'):
                    continue
                try:
                    raw = zf.read(name)
                except KeyError:
                    continue
                for encoding in ('utf-8', 'utf-16le', 'cp949', 'euc-kr'):
                    try:
                        xml = raw.decode(encoding)
                        break
                    except Exception:
                        continue
                else:
                    xml = raw.decode('utf-8', errors='ignore')
                cleaned = re.sub(r'<[^>]+>', ' ', xml)
                cleaned = re.sub(r'\s+', ' ', cleaned).strip()
                if cleaned:
                    texts.append(cleaned)
        if not texts:
            return None, 'HWPX 추출 결과 비어있음'
        return '\n'.join(texts), 'OK[HWPX]'
    except zipfile.BadZipFile:
        return None, 'HWPX ZIP 인식 실패'
    except Exception as exc:
        return None, f'HWPX 추출 실패: {exc}'


def contains_korean(text: str) -> bool:
    return any("가" <= ch <= "힣" for ch in text)


def collect_text_statistics(
    text: str | None,
    keywords: Sequence[str] | None = None,
) -> dict[str, object]:
    """Return quick stats for debugging/CLI usage."""

    keywords = tuple(keywords or ())
    normalized = (text or "").strip()
    found = [kw for kw in keywords if kw and kw in normalized]
    missing = [kw for kw in keywords if kw and kw not in normalized]
    korean_chars = sum(1 for ch in normalized if "가" <= ch <= "힣")
    stats = {
        "length": len(normalized),
        "korean_characters": korean_chars,
        "contains_korean": bool(korean_chars),
        "found_keywords": found,
        "missing_keywords": missing,
    }
    return stats
