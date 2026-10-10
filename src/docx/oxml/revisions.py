"""Custom element classes for revision tracking markup."""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, List

from docx.oxml.ns import qn
from docx.oxml.simpletypes import ST_String, XsdInt
from docx.oxml.text.run import CT_Text
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.text.run import CT_R


class CT_TrackChange(BaseOxmlElement):
    """Base class for tracked revision elements like `w:ins` and `w:del`."""

    id: int = RequiredAttribute("w:id", XsdInt)  # pyright: ignore[reportAssignmentType]
    author: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:author", ST_String
    )
    date: str | None = OptionalAttribute("w:date", ST_String)  # pyright: ignore[reportAssignmentType]

    @property
    def date_value(self) -> dt.datetime | None:
        date_str = self.date
        if date_str is None:
            return None
        try:
            return dt.datetime.fromisoformat(date_str.replace("Z", "+00:00"))
        except ValueError:
            return None

    @date_value.setter
    def date_value(self, value: dt.datetime | None):
        if value is None:
            date_qn = qn("w:date")
            if date_qn in self.attrib:  # pyright: ignore[reportUnknownMemberType]
                del self.attrib[date_qn]  # pyright: ignore[reportUnknownMemberType]
            return
        self.date = value.astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


class CT_RunTrackChange(CT_TrackChange):
    """`w:ins` or `w:del` container for tracked content."""

    p_lst: List[CT_P]
    tbl_lst: List[CT_Tbl]
    r_lst: List[CT_R]

    p = ZeroOrMore("w:p")
    tbl = ZeroOrMore("w:tbl")
    r = ZeroOrMore("w:r")

    @property
    def inner_content_elements(self) -> List[CT_P | CT_Tbl]:
        return self.xpath("./w:p | ./w:tbl")

    @property
    def run_content_elements(self) -> List[CT_R]:
        return self.xpath("./w:r")

    @property
    def text(self) -> str:  # pyright: ignore[reportIncompatibleMethodOverride]
        """Plain-text equivalent of the tracked content in this revision wrapper."""
        return "".join(
            str(e)
            for e in self.xpath(
                ".//w:br | .//w:cr | .//w:delText | .//w:noBreakHyphen"
                " | .//w:ptab | .//w:t | .//w:tab"
            )
        )


class CT_RPrChange(CT_TrackChange):
    rPr: BaseOxmlElement | None = ZeroOrOne("w:rPr")  # pyright: ignore[reportAssignmentType]


class CT_PPrChange(CT_TrackChange):
    pPr: BaseOxmlElement | None = ZeroOrOne("w:pPr")  # pyright: ignore[reportAssignmentType]


class CT_SectPrChange(CT_TrackChange):
    sectPr: BaseOxmlElement | None = ZeroOrOne("w:sectPr")  # pyright: ignore[reportAssignmentType]


class CT_TblPrChange(CT_TrackChange):
    tblPr: BaseOxmlElement | None = ZeroOrOne("w:tblPr")  # pyright: ignore[reportAssignmentType]


class CT_TcPrChange(CT_TrackChange):
    tcPr: BaseOxmlElement | None = ZeroOrOne("w:tcPr")  # pyright: ignore[reportAssignmentType]


class CT_TrPrChange(CT_TrackChange):
    trPr: BaseOxmlElement | None = ZeroOrOne("w:trPr")  # pyright: ignore[reportAssignmentType]


__all__ = [
    "CT_PPrChange",
    "CT_RPrChange",
    "CT_RunTrackChange",
    "CT_SectPrChange",
    "CT_TblPrChange",
    "CT_TcPrChange",
    "CT_TrackChange",
    "CT_TrPrChange",
    "CT_Text",
]
