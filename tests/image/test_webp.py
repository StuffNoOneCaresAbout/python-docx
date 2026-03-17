"""Unit test suite for docx.image.webp module."""

import io

import pytest

from docx.image.constants import MIME_TYPE
from docx.image.webp import Webp, _WebpParser

from ..unitutil.file import test_file
from ..unitutil.mock import ANY, initializer_mock


class DescribeWebp:
    def it_can_construct_from_a_lossy_webp_stream(self, Webp__init_):
        with open(test_file("python.webp"), "rb") as stream:
            webp = Webp.from_stream(stream)

        Webp__init_.assert_called_once_with(ANY, 2, 3, 72, 72)
        assert isinstance(webp, Webp)

    def it_knows_its_content_type(self):
        webp = Webp(None, None, None, None)
        assert webp.content_type == MIME_TYPE.WEBP

    def it_knows_its_default_ext(self):
        webp = Webp(None, None, None, None)
        assert webp.default_ext == "webp"

    @pytest.fixture
    def Webp__init_(self, request):
        return initializer_mock(request, Webp)


class Describe_WebpParser:
    @pytest.mark.parametrize(
        ("filename", "expected_dimensions"),
        [
            ("python.webp", (2, 3)),
            ("python-vp8l.webp", (7, 5)),
            ("python-vp8x.webp", (11, 13)),
        ],
    )
    def it_parses_dimensions_from_supported_webp_variants(self, filename, expected_dimensions):
        with open(test_file(filename), "rb") as stream:
            parser = _WebpParser.parse(stream)

            assert parser.dimensions == expected_dimensions

    @pytest.mark.parametrize(
        "blob",
        [
            b"NOPE" + b"\x00" * 28,
            b"RIFF\x00\x00\x00\x00NOPE" + b"\x00" * 16,
            b"RIFF\x0e\x00\x00\x00WEBPZZZZ" + b"\x00" * 6,
        ],
    )
    def it_raises_on_invalid_or_unsupported_webp_streams(self, blob):
        with pytest.raises(ValueError):
            _WebpParser.parse(io.BytesIO(blob)).dimensions
