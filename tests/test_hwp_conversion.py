import io
from pathlib import Path
import struct
import sys
import types

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import hwp_utils  # noqa: E402


@pytest.fixture
def fake_hwp_bytes():
    text = "계약 번호 확인"
    payload = text.encode("utf-16le")
    size = len(payload)
    tag_id = 67  # recognized as text-containing record
    header = struct.pack("<I", tag_id | (size << 20))
    return header + payload, text


def test_convert_hwp_local_to_text(monkeypatch, fake_hwp_bytes):
    body_stream, expected = fake_hwp_bytes

    class FakeOleFile:
        def __init__(self, _file_bytes):
            self._closed = False

        def listdir(self):
            return [["BodyText", "Section0"]]

        def openstream(self, path):
            assert path == ["BodyText", "Section0"]
            return io.BytesIO(body_stream)

        def close(self):
            self._closed = True

    fake_module = types.SimpleNamespace(OleFileIO=lambda _: FakeOleFile(_))
    monkeypatch.setitem(sys.modules, "olefile", fake_module)

    text, status = hwp_utils.convert_hwp_local_to_text(b"dummy")

    assert expected in text
    for keyword in ("계약", "번호"):
        assert keyword in text
    assert status == "OK[OLE]"


def test_collect_text_statistics():
    sample = "계약 번호와 스쿨넷 사업"
    stats = hwp_utils.collect_text_statistics(sample, ["계약", "번호", "미존재"])
    assert stats["contains_korean"] is True
    assert stats["korean_characters"] >= 3
    assert stats["found_keywords"] == ["계약", "번호"]
    assert stats["missing_keywords"] == ["미존재"]


@pytest.mark.skipif(
    not (Path(__file__).resolve().parents[1] / "공고서_국가_일반_국내_유선통신서비스_1458+0036_공동이행_하도급사전승인_장기계속계약_차등점수제_수요기관평가.hwp").exists(),
    reason="샘플 HWP 파일이 존재하지 않습니다.",
)
def test_sample_hwp_contains_keywords():
    sample = (
        Path(__file__).resolve().parents[1]
        / "공고서_국가_일반_국내_유선통신서비스_1458+0036_공동이행_하도급사전승인_장기계속계약_차등점수제_수요기관평가.hwp"
    )
    data = sample.read_bytes()
    text, status = hwp_utils.convert_hwp_local_to_text(data)
    assert text, f"텍스트 추출 실패: {status}"
    for keyword in ("계약", "번호"):
        assert keyword in text
