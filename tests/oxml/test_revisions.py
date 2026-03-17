from __future__ import annotations

import datetime as dt
from typing import cast

from docx.oxml.document import CT_Body
from docx.oxml.revisions import CT_RunTrackChange
from docx.oxml.text.paragraph import CT_P

from ..unitutil.cxml import element


class DescribeCT_RunTrackChange:
    def it_can_round_trip_its_datetime_value(self):
        change = cast(
            CT_RunTrackChange,
            element('w:ins{w:id=1,w:author=TestAuthor,w:date=2026-03-17T12:34:56Z}/w:r/w:t"x"'),
        )

        assert change.date_value == dt.datetime(2026, 3, 17, 12, 34, 56, tzinfo=dt.timezone.utc)


class DescribeCT_P:
    def it_distinguishes_original_accepted_and_deleted_text(self):
        paragraph = cast(
            CT_P,
            element(
                'w:p/(w:r/w:t"keep",w:del{w:id=1,w:author=TestAuthor}/w:r/w:delText"gone",w:ins{w:id=2,w:author=TestAuthor}/w:r/w:t"new")'
            ),
        )

        assert paragraph.text == "keepgone"
        assert paragraph.accepted_text == "keepnew"
        assert paragraph.deleted_text == "gone"

    def it_ignores_comment_markers_when_computing_text_views(self):
        paragraph = cast(
            CT_P,
            element(
                "w:p/("
                "w:commentRangeStart{w:id=7},"
                'w:r/w:t"keep",'
                "w:commentRangeEnd{w:id=7},"
                "w:r/(w:rPr/w:rStyle{w:val=CommentReference},w:commentReference{w:id=7}),"
                'w:del{w:id=1,w:author=TestAuthor}/w:r/w:delText"gone",'
                'w:ins{w:id=2,w:author=TestAuthor}/w:r/w:t"new"'
                ")"
            ),
        )

        assert paragraph.text == "keepgone"
        assert paragraph.accepted_text == "keepnew"
        assert paragraph.deleted_text == "gone"


class DescribeCT_Body:
    def it_includes_deleted_block_items_in_inner_content(self):
        body = cast(
            CT_Body,
            element(
                "w:body/(w:p,w:del{w:id=1,w:author=TestAuthor}/w:p,w:tbl/(w:tblPr,w:tblGrid),w:del{w:id=2,w:author=TestAuthor}/w:tbl/(w:tblPr,w:tblGrid))"
            ),
        )

        assert [elm.tag.rsplit("}", 1)[-1] for elm in body.inner_content_elements] == [
            "p",
            "p",
            "tbl",
            "tbl",
        ]
