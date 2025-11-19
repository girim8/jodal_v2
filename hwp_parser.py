"""HWP/HWPX 파일에서 텍스트를 뽑아내는 도우미 함수 모음."""

from __future__ import annotations

import io
import os
import struct
import subprocess
import sys
import tempfile
import zipfile
import zlib
from pathlib import Path

import olefile
from xml.etree import ElementTree

from bs4 import BeautifulSoup


def _maybe_decompress_raw(data: bytes) -> bytes:
    """필요 시 RAW zlib 으로 한 번만 압축 해제한다."""
    try:
        return zlib.decompress(data, -zlib.MAX_WBITS)
    except zlib.error:
        return data


def _strip_control_chars(text: str) -> str:
    """줄바꿈·탭을 제외한 제어 문자는 공백으로 치환하고 NULL 은 제거한다."""
    cleaned_chars = []
    for ch in text:
        code = ord(ch)
        if ch in ("\n", "\t"):
            cleaned_chars.append(ch)
        elif ch == "\x00":
            # UTF-16LE 디코딩 과정에서 남은 NULL 제거
            continue
        elif code < 32:
            cleaned_chars.append(" ")
        else:
            cleaned_chars.append(ch)
    return "".join(cleaned_chars)


def _clean_text(text: str) -> str:
    """제어문자를 정리하면서 연속 공백을 축소하고 빈 줄은 제거한다."""
    filtered = _strip_control_chars(text)
    normalized_lines = []
    for line in filtered.splitlines():
        compacted = " ".join(line.split())
        if compacted:
            normalized_lines.append(compacted)
    return "\n".join(normalized_lines).strip()


def _parse_body_records(data: bytes) -> str:
    """BodyText 레코드를 순회하면서 사람이 읽을 수 있는 텍스트만 모은다."""
    HWPTAG_PARA_TEXT = 67

    text_chunks: list[str] = []
    offset = 0
    length = len(data)

    while offset + 4 <= length:
        header = struct.unpack_from("<I", data, offset)[0]
        tag_id = header & 0x3FF
        size = (header >> 20) & 0xFFF
        offset += 4

        if size == 0xFFF:
            if offset + 4 > length:
                break
            size = struct.unpack_from("<I", data, offset)[0]
            offset += 4

        if offset + size > length or size < 0:
            break

        payload = data[offset : offset + size]
        offset += size

        if tag_id != HWPTAG_PARA_TEXT:
            continue

        decoded = payload.decode("utf-16le", errors="ignore")
        cleaned = _clean_text(decoded)
        if cleaned:
            text_chunks.append(cleaned)

    return "\n".join(text_chunks)


def _detect_body_compressed(file_header: bytes) -> bool:
    """FileHeader[36] 비트 0 이 1이면 BodyText 가 zlib RAW 로 압축됨."""
    return len(file_header) >= 37 and (file_header[36] & 0x01) != 0


def _list_body_sections(entries: list[list[str]]) -> list[str]:
    """BodyText/SectionN 스트림 이름을 순서대로 반환한다."""
    section_ids: list[int] = []
    for entry in entries:
        if len(entry) == 2 and entry[0] == "BodyText" and entry[1].startswith("Section"):
            try:
                section_ids.append(int(entry[1].replace("Section", "")))
            except ValueError:
                continue

    section_ids.sort()
    return [f"BodyText/Section{idx}" for idx in section_ids]


def extract_text_from_hwp(data: bytes) -> str:
    """OLE 스트림을 직접 읽어 바이너리 HWP 문서에서 텍스트를 뽑는다."""
    text_parts: list[str] = []

    with olefile.OleFileIO(io.BytesIO(data)) as ole:
        try:
            header_bytes = ole.openstream("FileHeader").read()
        except OSError:
            header_bytes = b""

        compressed = _detect_body_compressed(header_bytes)
        section_streams = _list_body_sections(ole.listdir())

        for stream_name in section_streams:
            try:
                raw_stream = ole.openstream(stream_name).read()
            except OSError:
                continue

            stream_data = _maybe_decompress_raw(raw_stream) if compressed else raw_stream
            parsed = _parse_body_records(stream_data)
            if parsed:
                text_parts.append(f"[{stream_name}]\n{parsed}")

    return _clean_text("\n\n".join(text_parts))


def _find_hwp5_cli(base_dir: Path) -> Path | None:
    """현재 파일 기준으로 hwp5-table-extractor/cli.py 경로를 추정한다."""
    candidates = [
        base_dir / "hwp5-table-extractor" / "cli.py",
        base_dir.parent / "hwp5-table-extractor" / "cli.py",
    ]
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


def _html_tables_to_markdown(html: str) -> str:
    """hwp5-table-extractor 출력 HTML을 Markdown 테이블로 단순 변환한다."""
    soup = BeautifulSoup(html, "html.parser")
    tables = soup.find_all("table")
    md_blocks: list[str] = []

    for idx, table in enumerate(tables, start=1):
        rows = []
        for tr in table.find_all("tr"):
            cells = tr.find_all(["th", "td"])
            row = []
            for cell in cells:
                text = " ".join(cell.get_text(separator=" ", strip=True).split())
                row.append(text)
            if row:
                rows.append(row)

        if not rows:
            continue

        width = max(len(r) for r in rows)
        padded_rows = [r + [""] * (width - len(r)) for r in rows]
        header, *body = padded_rows

        md_lines = []
        md_lines.append(f"<!-- Table {idx} -->")
        md_lines.append("| " + " | ".join(header) + " |")
        md_lines.append("| " + " | ".join(["---"] * width) + " |")
        for row in body:
            md_lines.append("| " + " | ".join(row) + " |")

        md_blocks.append("\n".join(md_lines))

    return "\n\n".join(md_blocks)


def _extract_tables_markdown_from_hwp_bytes(data: bytes) -> str:
    """바이너리 HWP 파일에서 표를 Markdown 텍스트로 추출한다."""
    base_dir = Path(__file__).resolve().parent
    cli_path = _find_hwp5_cli(base_dir)
    if cli_path is None:
        return ""

    with tempfile.NamedTemporaryFile(delete=False, suffix=".hwp") as tmp_in:
        tmp_in.write(data)
        in_path = tmp_in.name

    tmp_html = tempfile.NamedTemporaryFile(delete=False, suffix=".html")
    out_path = tmp_html.name
    tmp_html.close()

    try:
        subprocess.run(
            [sys.executable, str(cli_path), in_path, out_path],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

        html = Path(out_path).read_text(encoding="utf-8")
        return _html_tables_to_markdown(html).strip()
    except (subprocess.CalledProcessError, OSError):
        return ""
    finally:
        for path in (in_path, out_path):
            try:
                os.unlink(path)
            except OSError:
                pass


def extract_text_from_hwpx(data: bytes) -> str:
    """ZIP 안의 XML 파일들을 합쳐 HWPX 문서의 텍스트를 만든다."""
    text_chunks: list[str] = []

    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        for name in archive.namelist():
            if not name.endswith(".xml"):
                continue

            try:
                xml_data = archive.read(name)
            except KeyError:
                continue

            try:
                root = ElementTree.fromstring(xml_data)
            except ElementTree.ParseError:
                continue

            # itertext()로 태그 사이의 문자만 빠르게 추출한다.
            text_chunks.append("".join(root.itertext()))

    return _clean_text("\n".join(text_chunks))


def convert_to_text(
    data: bytes, filename: str | None = None, *, include_tables: bool = True
) -> tuple[str, str]:
    """파일 포맷을 판별하고 텍스트와 판별 결과를 함께 돌려준다."""
    name = (filename or "").lower()
    is_hwpx = name.endswith("hwpx") or data[:2] == b"PK"

    if is_hwpx:
        text = extract_text_from_hwpx(data)
        fmt = "HWPX"
    else:
        text = extract_text_from_hwp(data)
        fmt = "HWP"

    if include_tables and fmt == "HWP":
        tables_md = _extract_tables_markdown_from_hwp_bytes(data)
        if tables_md:
            separator = "\n\n---\n\n" if text else ""
            text = f"{text}{separator}{tables_md}".strip()

    if not text:
        raise ValueError("Unable to extract text from the provided file.")

    return text, fmt


__all__ = [
    "convert_to_text",
    "extract_text_from_hwp",
    "extract_text_from_hwpx",
]
