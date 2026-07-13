from __future__ import annotations

import binascii
import struct
import zlib
from pathlib import Path

from heuthesis_latex2word import eps_fallback
from heuthesis_latex2word.eps_fallback import convert_imagemagick_eps_to_png


REPO_ROOT = Path(__file__).resolve().parents[1]


def _synthetic_eps(
    width: int,
    height: int,
    image_class: int,
    compression: int,
    body: str,
) -> bytes:
    return (
        "%!PS-Adobe-3.0 EPSF-3.0\n"
        "%%Creator: (ImageMagick)\n"
        "%%DocumentData: Clean7Bit\n"
        "%%EndProlog\n"
        "DisplayImage\n"
        "0 0\n"
        f"{width} {height}\n"
        "12\n"
        f"{width} {height}\n"
        f"{image_class}\n"
        f"{compression}\n"
        f"{body}\n"
        "end\n"
        "%%PageTrailer\n"
        "%%Trailer\n"
        "%%EOF\n"
    ).encode("ascii")


def _read_png(path: Path) -> tuple[int, int, int, bytes]:
    data = path.read_bytes()
    assert data.startswith(b"\x89PNG\r\n\x1a\n")
    position = 8
    chunks: list[tuple[bytes, bytes]] = []
    while position < len(data):
        length = struct.unpack(">I", data[position : position + 4])[0]
        chunk_type = data[position + 4 : position + 8]
        payload = data[position + 8 : position + 8 + length]
        checksum = struct.unpack(">I", data[position + 8 + length : position + 12 + length])[0]
        assert checksum == binascii.crc32(chunk_type + payload) & 0xFFFFFFFF
        chunks.append((chunk_type, payload))
        position += 12 + length
    assert position == len(data)
    assert [chunk_type for chunk_type, _ in chunks] == [b"IHDR", b"IDAT", b"IEND"]
    width, height, depth, color_type, compression, filtering, interlace = struct.unpack(
        ">IIBBBBB", chunks[0][1]
    )
    assert (depth, compression, filtering, interlace) == (8, 0, 0, 0)
    return width, height, color_type, zlib.decompress(chunks[1][1])


def test_converts_repository_imagemagick_grayscale_eps(tmp_path: Path) -> None:
    source = (
        REPO_ROOT
        / "HeuThesis_Overleaf"
        / "examples"
        / "book"
        / "bachelor"
        / "figures"
        / "word.eps"
    )
    target = tmp_path / "word.png"

    assert convert_imagemagick_eps_to_png(source, target)

    width, height, color_type, scanlines = _read_png(target)
    assert (width, height, color_type) == (443, 55, 0)
    rows = [
        scanlines[offset : offset + width + 1]
        for offset in range(0, len(scanlines), width + 1)
    ]
    assert len(rows) == height
    assert all(row[0] == 0 for row in rows)
    pixels = b"".join(row[1:] for row in rows)
    assert len(pixels) == width * height
    assert 0 in pixels
    assert 255 in pixels


def test_converts_directclass_rle_packets(tmp_path: Path) -> None:
    source = tmp_path / "direct.eps"
    target = tmp_path / "direct.png"
    # Two red pixels followed by one green pixel.
    source.write_bytes(_synthetic_eps(3, 1, 0, 1, "FF000001 00FF0000"))

    assert convert_imagemagick_eps_to_png(source, target)
    assert _read_png(target) == (
        3,
        1,
        2,
        b"\x00\xff\x00\x00\xff\x00\x00\x00\xff\x00",
    )


def test_converts_pseudoclass_rle_packets(tmp_path: Path) -> None:
    source = tmp_path / "palette.eps"
    target = tmp_path / "palette.png"
    # PseudoClass subtype 0, two colors, then red x2 and blue x1 packets.
    source.write_bytes(
        _synthetic_eps(3, 1, 1, 1, "0\n2\nFF0000 0000FF\n0001 0100")
    )

    assert convert_imagemagick_eps_to_png(source, target)
    assert _read_png(target) == (
        3,
        1,
        2,
        b"\x00\xff\x00\x00\xff\x00\x00\x00\x00\xff",
    )


def test_rejects_truncated_or_unsupported_eps_without_output(tmp_path: Path) -> None:
    truncated = tmp_path / "truncated.eps"
    truncated.write_bytes(_synthetic_eps(3, 1, 1, 1, "1\n8\n00FF"))
    unsupported = tmp_path / "unsupported.eps"
    unsupported.write_bytes(_synthetic_eps(1, 1, 1, 0, "1\n4\n00"))

    truncated_target = tmp_path / "truncated.png"
    unsupported_target = tmp_path / "unsupported.png"
    assert not convert_imagemagick_eps_to_png(truncated, truncated_target)
    assert not truncated_target.exists()
    assert not convert_imagemagick_eps_to_png(unsupported, unsupported_target)
    assert not unsupported_target.exists()


def test_failed_conversion_preserves_existing_target(tmp_path: Path) -> None:
    source = tmp_path / "not-eps.bin"
    source.write_bytes(b"not an EPS")
    target = tmp_path / "existing.png"
    target.write_bytes(b"existing output")

    assert not convert_imagemagick_eps_to_png(source, target)
    assert target.read_bytes() == b"existing output"


def test_rejects_input_above_byte_limit(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr(eps_fallback, "_MAX_EPS_BYTES", 32)
    source = tmp_path / "oversized.eps"
    source.write_bytes(b"%!PS-Adobe-3.0 EPSF-3.0\n" + b"x" * 32)
    target = tmp_path / "oversized.png"

    assert not convert_imagemagick_eps_to_png(source, target)
    assert not target.exists()
