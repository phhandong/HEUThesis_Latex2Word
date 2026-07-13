"""Dependency-free fallback for the raster EPS files emitted by ImageMagick.

This module intentionally does not try to interpret arbitrary PostScript.  It
only decodes the well-defined pixel stream written after ImageMagick's
``DisplayImage`` invocation.  Unsupported or malformed input is rejected so a
caller can safely continue to another conversion backend.
"""

from __future__ import annotations

import binascii
import math
import os
import re
import struct
import tempfile
import zlib
from dataclasses import dataclass
from pathlib import Path

__all__ = ["convert_imagemagick_eps_to_png"]


_MAX_PIXELS = 25_000_000
_MAX_EPS_BYTES = 256 * 1024 * 1024
_PS_WHITESPACE = frozenset((0, 9, 10, 12, 13, 32))
_INTEGER_RE = re.compile(rb"[+-]?\d+\Z")


class _EPSDecodeError(ValueError):
    """The EPS is not one of the supported ImageMagick raster encodings."""


@dataclass(frozen=True)
class _Raster:
    width: int
    height: int
    color_type: int
    stride: int
    pixels: bytes


class _LineCursor:
    def __init__(self, data: bytes, position: int) -> None:
        self.data = data
        self.position = position

    def read(self) -> bytes:
        if self.position >= len(self.data):
            raise _EPSDecodeError("unexpected end of EPS header")
        end = self.data.find(b"\n", self.position)
        if end < 0:
            end = len(self.data)
            next_position = end
        else:
            next_position = end + 1
        line = self.data[self.position : end]
        self.position = next_position
        if line.endswith(b"\r"):
            line = line[:-1]
        return line.strip()


class _HexCursor:
    def __init__(self, data: bytes, position: int) -> None:
        self.data = data
        self.position = position

    def _nibble(self) -> int:
        while self.position < len(self.data):
            value = self.data[self.position]
            self.position += 1
            if value in _PS_WHITESPACE:
                continue
            if 48 <= value <= 57:
                return value - 48
            if 65 <= value <= 70:
                return value - 65 + 10
            if 97 <= value <= 102:
                return value - 97 + 10
            raise _EPSDecodeError("non-hexadecimal byte in raster stream")
        raise _EPSDecodeError("truncated raster stream")

    def read(self, size: int) -> bytes:
        result = bytearray(size)
        for index in range(size):
            result[index] = (self._nibble() << 4) | self._nibble()
        return bytes(result)

    def read_byte(self) -> int:
        return (self._nibble() << 4) | self._nibble()


def _tokens(line: bytes, count: int) -> list[bytes]:
    values = line.split()
    if len(values) != count:
        raise _EPSDecodeError("unexpected ImageMagick EPS header layout")
    return values


def _floats(line: bytes, count: int) -> list[float]:
    try:
        values = [float(value) for value in _tokens(line, count)]
    except ValueError as exc:
        raise _EPSDecodeError("invalid numeric EPS header value") from exc
    if not all(math.isfinite(value) for value in values):
        raise _EPSDecodeError("non-finite EPS header value")
    return values


def _integers(line: bytes, count: int) -> list[int]:
    values = _tokens(line, count)
    if not all(_INTEGER_RE.fullmatch(value) for value in values):
        raise _EPSDecodeError("invalid integer EPS header value")
    try:
        return [int(value) for value in values]
    except ValueError as exc:
        raise _EPSDecodeError("invalid integer EPS header value") from exc


def _decode_rle_direct(cursor: _HexCursor, pixel_count: int) -> bytes:
    pixels = bytearray()
    decoded = 0
    while decoded < pixel_count:
        color = cursor.read(3)
        run = cursor.read_byte() + 1
        if decoded + run > pixel_count:
            raise _EPSDecodeError("DirectClass RLE packet exceeds image bounds")
        pixels.extend(color * run)
        decoded += run
    return bytes(pixels)


def _decode_rle_palette(
    cursor: _HexCursor, pixel_count: int, colormap: bytes, colors: int
) -> bytes:
    pixels = bytearray()
    decoded = 0
    while decoded < pixel_count:
        color_index = cursor.read_byte()
        run = cursor.read_byte() + 1
        if color_index >= colors:
            raise _EPSDecodeError("palette index is outside the colormap")
        if decoded + run > pixel_count:
            raise _EPSDecodeError("PseudoClass RLE packet exceeds image bounds")
        color = colormap[color_index * 3 : color_index * 3 + 3]
        pixels.extend(color * run)
        decoded += run
    return bytes(pixels)


def _decode_palette(
    cursor: _HexCursor,
    pixel_count: int,
    compression: int,
    line_cursor: _LineCursor,
) -> bytes:
    colors = _integers(line_cursor.read(), 1)[0]
    if not 1 <= colors <= 256:
        raise _EPSDecodeError("unsupported ImageMagick colormap size")
    cursor.position = line_cursor.position
    colormap = cursor.read(colors * 3)
    if compression == 1:
        return _decode_rle_palette(cursor, pixel_count, colormap, colors)

    indexes = cursor.read(pixel_count)
    pixels = bytearray(pixel_count * 3)
    for pixel, color_index in enumerate(indexes):
        if color_index >= colors:
            raise _EPSDecodeError("palette index is outside the colormap")
        source = color_index * 3
        destination = pixel * 3
        pixels[destination : destination + 3] = colormap[source : source + 3]
    return bytes(pixels)


def _valid_trailer(data: bytes, position: int) -> bool:
    try:
        trailer = data[position:].decode("ascii")
    except UnicodeDecodeError:
        return False

    lines = [line.strip() for line in trailer.splitlines() if line.strip()]
    if not lines or lines[-1] != "%%EOF":
        return False

    executable_lines = []
    for line in lines:
        if line in {"%%PageTrailer", "%%Trailer", "%%EOF"}:
            continue
        if line in {"end", "grestore", "showpage"}:
            executable_lines.append(line)
            continue
        return False
    return "end" in executable_lines


def _decode_imagemagick_eps(data: bytes) -> _Raster:
    first_line = data.splitlines()[0] if data else b""
    if not first_line.startswith(b"%!PS-Adobe-") or b"EPSF-" not in first_line:
        raise _EPSDecodeError("not an EPS file")

    end_prolog = re.search(rb"(?m)^%%EndProlog[ \t]*\r?\n", data)
    if end_prolog is None:
        raise _EPSDecodeError("missing ImageMagick EPS prolog terminator")
    if re.search(
        rb"(?mi)^%%Creator:[^\r\n]*ImageMagick[^\r\n]*\r?$",
        data[: end_prolog.start()],
    ) is None:
        raise _EPSDecodeError("EPS was not emitted by ImageMagick")

    display = re.search(
        rb"(?m)^DisplayImage[ \t]*\r?\n", data[end_prolog.end() :]
    )
    if display is None:
        raise _EPSDecodeError("missing ImageMagick DisplayImage invocation")
    header_position = end_prolog.end() + display.end()
    line_cursor = _LineCursor(data, header_position)

    _floats(line_cursor.read(), 2)  # translation
    scale = _floats(line_cursor.read(), 2)
    point_size = _floats(line_cursor.read(), 1)[0]
    if scale[0] <= 0 or scale[1] <= 0 or point_size < 0:
        raise _EPSDecodeError("invalid ImageMagick image geometry")

    width, height = _integers(line_cursor.read(), 2)
    if width <= 0 or height <= 0:
        raise _EPSDecodeError("invalid raster dimensions")
    pixel_count = width * height
    if pixel_count > _MAX_PIXELS:
        raise _EPSDecodeError("raster exceeds fallback decoder limits")

    image_class = _integers(line_cursor.read(), 1)[0]
    compression = _integers(line_cursor.read(), 1)[0]
    if image_class not in (0, 1) or compression not in (0, 1):
        raise _EPSDecodeError("unsupported ImageMagick image class or compression")

    cursor = _HexCursor(data, line_cursor.position)
    if image_class == 0:
        if compression == 0:
            pixels = cursor.read(pixel_count * 3)
        else:
            pixels = _decode_rle_direct(cursor, pixel_count)
        raster = _Raster(width, height, 2, width * 3, pixels)
    else:
        pseudo_class = _integers(line_cursor.read(), 1)[0]
        if pseudo_class == 1:
            depth = _integers(line_cursor.read(), 1)[0]
            if depth != 8:
                raise _EPSDecodeError("only 8-bit ImageMagick grayscale EPS is supported")
            cursor.position = line_cursor.position
            # ImageMagick's PostScript prolog ignores the outer compression
            # flag in this grayscale branch and reads raw scanlines directly.
            pixels = cursor.read(pixel_count)
            raster = _Raster(width, height, 0, width, pixels)
        elif pseudo_class == 0:
            pixels = _decode_palette(cursor, pixel_count, compression, line_cursor)
            raster = _Raster(width, height, 2, width * 3, pixels)
        else:
            raise _EPSDecodeError("unsupported ImageMagick PseudoClass subtype")

    if not _valid_trailer(data, cursor.position):
        raise _EPSDecodeError("unexpected data after ImageMagick raster stream")
    return raster


def _png_chunk(chunk_type: bytes, payload: bytes) -> bytes:
    checksum = binascii.crc32(chunk_type + payload) & 0xFFFFFFFF
    return (
        struct.pack(">I", len(payload))
        + chunk_type
        + payload
        + struct.pack(">I", checksum)
    )


def _encode_png(raster: _Raster) -> bytes:
    expected_size = raster.stride * raster.height
    if len(raster.pixels) != expected_size:
        raise _EPSDecodeError("decoded raster has an invalid byte count")

    scanlines = bytearray(expected_size + raster.height)
    source = 0
    destination = 0
    for _ in range(raster.height):
        scanlines[destination] = 0  # PNG filter method: None
        destination += 1
        scanlines[destination : destination + raster.stride] = raster.pixels[
            source : source + raster.stride
        ]
        source += raster.stride
        destination += raster.stride

    ihdr = struct.pack(">IIBBBBB", raster.width, raster.height, 8, raster.color_type, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + _png_chunk(b"IHDR", ihdr)
        + _png_chunk(b"IDAT", zlib.compress(bytes(scanlines), level=9))
        + _png_chunk(b"IEND", b"")
    )


def _write_atomic(target: Path, payload: bytes) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    descriptor = -1
    temporary_name: str | None = None
    try:
        descriptor, temporary_name = tempfile.mkstemp(
            prefix=f".{target.name}.", suffix=".tmp", dir=str(target.parent)
        )
        with os.fdopen(descriptor, "wb") as stream:
            descriptor = -1
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary_name, target)
        temporary_name = None
    finally:
        if descriptor >= 0:
            os.close(descriptor)
        if temporary_name is not None:
            try:
                Path(temporary_name).unlink()
            except FileNotFoundError:
                pass


def convert_imagemagick_eps_to_png(source: Path, target: Path) -> bool:
    """Convert a supported ImageMagick raster EPS to a PNG atomically.

    ``False`` means that the input is unsupported, malformed, unreadable, or
    the output could not be written.  A failed call never installs a partial
    output file; an already existing target is left unchanged.
    """

    source_path = Path(source)
    target_path = Path(target)
    if os.path.normcase(os.path.abspath(source_path)) == os.path.normcase(
        os.path.abspath(target_path)
    ):
        return False
    try:
        with source_path.open("rb") as stream:
            if os.fstat(stream.fileno()).st_size > _MAX_EPS_BYTES:
                return False
            data = stream.read(_MAX_EPS_BYTES + 1)
        if len(data) > _MAX_EPS_BYTES:
            return False
        raster = _decode_imagemagick_eps(data)
        png = _encode_png(raster)
        _write_atomic(target_path, png)
    except (OSError, _EPSDecodeError, OverflowError, zlib.error):
        return False
    return True
