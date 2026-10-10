# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.comments` module."""

from __future__ import annotations

import datetime as dt
from typing import cast

import pytest

from docx.oxml.comments import (
    CT_Comments,
    CT_CommentsEx,
    CT_CommentsExtensible,
    CT_CommentsIds,
    CT_People,
)
from docx.oxml.parser import parse_xml

from ..unitutil.cxml import element


class DescribeCT_Comments:
    """Unit-test suite for `docx.oxml.comments.CT_Comments`."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:comments", 0),
            ("w:comments/(w:comment{w:id=1})", 2),
            ("w:comments/(w:comment{w:id=4},w:comment{w:id=2147483646})", 2147483647),
            ("w:comments/(w:comment{w:id=1},w:comment{w:id=2147483647})", 0),
            ("w:comments/(w:comment{w:id=1},w:comment{w:id=2},w:comment{w:id=3})", 4),
        ],
    )
    def it_finds_the_next_available_comment_id_to_help(self, cxml: str, expected_value: int):
        comments_elm = cast(CT_Comments, element(cxml))
        assert comments_elm._next_available_comment_id() == expected_value

    def it_can_add_a_comment_with_an_explicit_id(self):
        comments_elm = cast(CT_Comments, element("w:comments"))

        comment = comments_elm.add_comment(42)

        assert comment.id == 42
        assert len(comments_elm.comment_lst) == 1
        assert comment.author == ""
        assert comment.p_lst[0].style == "CommentText"
        assert comment.p_lst[0].r_lst[0].style == "CommentReference"

    def it_can_add_a_comment_with_a_para_id(self):
        comments_elm = cast(CT_Comments, element("w:comments"))

        comment = comments_elm.add_comment(42, para_id="0000002A")

        assert comment.p_lst[0].paraId == "0000002A"
        assert comment.p_lst[0].textId == "77777777"

    def it_can_write_a_word_style_local_comment_date(self):
        comments_elm = cast(CT_Comments, element("w:comments"))
        comment = comments_elm.add_comment(42)

        comment.set_local_date(
            dt.datetime(2025, 6, 11, 20, 42, 30, tzinfo=dt.timezone(dt.timedelta(hours=8)))
        )

        date_attr = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date"
        assert comment.get(date_attr) == "2025-06-11T20:42:30"


class DescribeCT_CommentsEx:
    def it_can_add_a_comment_extension(self):
        comments_ex = cast(
            CT_CommentsEx,
            parse_xml(
                '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>'
            ),
        )

        comment_ex = comments_ex.add_comment_ex("0000002A", parent_para_id="00000001", done=True)

        assert comment_ex.paraId == "0000002A"
        assert comment_ex.paraIdParent == "00000001"
        assert comment_ex.done is True
        assert len(comments_ex.commentEx_lst) == 1


class DescribeCT_CommentsIds:
    def it_can_add_a_durable_comment_id(self):
        comments_ids = cast(
            CT_CommentsIds,
            parse_xml(
                '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"/>'
            ),
        )

        comment_id = comments_ids.add_comment_id("0000002A", "0000000A")

        assert comment_id.paraId == "0000002A"
        assert comment_id.durableId == "0000000A"
        assert comments_ids.get_comment_id_by_para_id("0000002A") is comment_id


class DescribeCT_CommentsExtensible:
    def it_can_add_an_extensible_comment_entry(self):
        comments_extensible = cast(
            CT_CommentsExtensible,
            parse_xml(
                '<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"/>'
            ),
        )

        comment_extensible = comments_extensible.add_comment_extensible(durable_id="0000000A")

        assert comment_extensible.durableId == "0000000A"
        assert comment_extensible.dateUtc is not None
        assert (
            comments_extensible.get_comment_extensible_by_durable_id("0000000A")
            is comment_extensible
        )

    def it_normalizes_extensible_comment_timestamps_to_whole_seconds(self):
        comments_extensible = cast(
            CT_CommentsExtensible,
            parse_xml(
                '<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"/>'
            ),
        )
        timestamp = dt.datetime(
            2025,
            6,
            11,
            20,
            42,
            30,
            987654,
            tzinfo=dt.timezone(dt.timedelta(hours=8)),
        )

        comment_extensible = comments_extensible.add_comment_extensible("0000000A", timestamp)

        assert comment_extensible.dateUtc == dt.datetime(
            2025, 6, 11, 12, 42, 30, tzinfo=dt.timezone.utc
        )


class DescribeCT_People:
    def it_can_add_a_person(self):
        people = cast(
            CT_People,
            parse_xml(
                '<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>'
            ),
        )

        person = people.add_person("TestAuthor")

        assert person.author == "TestAuthor"
        assert person.presenceInfo is not None
        assert people.get_person_by_author("TestAuthor") is person
