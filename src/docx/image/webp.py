"""Objects related to parsing headers of WebP image streams."""

from __future__ import annotations

from struct import Struct

from docx.image.constants import MIME_TYPE
from docx.image.image import BaseImageHeader


class Webp(BaseImageHeader):
    """Image header parser for WebP images."""

    @classmethod
    def from_stream(cls, stream):
        """Return |Webp| instance having header properties parsed from WebP image in
        `stream`.

        Resolution metadata is not surfaced; both DPI values default to 72.
        """
        px_width, px_height = _WebpParser.parse(stream).dimensions
        return cls(px_width, px_height, 72, 72)

    @property
    def content_type(self):
        """MIME content type for this image, unconditionally `image/webp`."""
        return MIME_TYPE.WEBP

    @property
    def default_ext(self):
        """Default filename extension, always 'webp' for WebP images."""
        return "webp"


class _WebpParser:
    """Parses a WebP image stream to extract canvas dimensions."""

    _UINT32_LE = Struct("<I")
    _UINT16_LE = Struct("<H")

    def __init__(self, stream):
        super(_WebpParser, self).__init__()
        self._stream = stream

    @classmethod
    def parse(cls, stream):
        """Return a |_WebpParser| instance ready to extract properties from `stream`."""
        return cls(stream)

    @property
    def dimensions(self):
        """(px_width, px_height) pair parsed from the WebP stream."""
        stream = self._stream
        stream.seek(0)

        if stream.read(4) != b"RIFF":
            raise ValueError("Not a valid WebP file")

        stream.read(4)  # file size

        if stream.read(4) != b"WEBP":
            raise ValueError("Not a valid WebP file")

        chunk_type = stream.read(4)
        if chunk_type == b"VP8 ":
            return self._parse_vp8_chunk(stream)
        if chunk_type == b"VP8L":
            return self._parse_vp8l_chunk(stream)
        if chunk_type == b"VP8X":
            return self._parse_vp8x_chunk(stream)
        raise ValueError("Unsupported WebP format")

    def _parse_vp8_chunk(self, stream):
        """Return dimensions from a lossy VP8 chunk."""
        stream.read(4)  # chunk size
        stream.read(3)  # frame tag
        if stream.read(3) != b"\x9d\x01\x2a":
            raise ValueError("Invalid VP8 WebP stream")

        width = self._UINT16_LE.unpack(stream.read(2))[0] & 0x3FFF
        height = self._UINT16_LE.unpack(stream.read(2))[0] & 0x3FFF
        return width, height

    def _parse_vp8l_chunk(self, stream):
        """Return dimensions from a lossless VP8L chunk."""
        stream.read(4)  # chunk size
        if stream.read(1) != b"\x2f":
            raise ValueError("Invalid VP8L WebP stream")

        bits = self._UINT32_LE.unpack(stream.read(4))[0]
        width = (bits & 0x3FFF) + 1
        height = ((bits >> 14) & 0x3FFF) + 1
        return width, height

    def _parse_vp8x_chunk(self, stream):
        """Return dimensions from an extended VP8X chunk."""
        stream.read(4)  # chunk size
        stream.read(4)  # feature flags + reserved bytes
        width = int.from_bytes(stream.read(3), "little") + 1
        height = int.from_bytes(stream.read(3), "little") + 1
        return width, height
