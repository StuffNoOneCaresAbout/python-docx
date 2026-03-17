"""Unit test suite for the docx.parts.hdrftr module."""

from __future__ import annotations

import datetime as dt
from typing import cast

import pytest

from docx.comments import Comments
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import PartFactory
from docx.oxml.comments import (
    CT_Comments,
    CT_CommentsEx,
    CT_CommentsExtensible,
    CT_CommentsIds,
    CT_People,
)
from docx.oxml.parser import parse_xml
from docx.package import Package
from docx.parts.comments import (
    CommentsExtendedPart,
    CommentsExtensiblePart,
    CommentsIdsPart,
    CommentsPart,
    PeoplePart,
)

from ..unitutil.cxml import element
from ..unitutil.mock import FixtureRequest, Mock, class_mock, instance_mock, method_mock


class DescribeCommentsPart:
    """Unit test suite for `docx.parts.comments.CommentsPart` objects."""

    def it_is_used_by_the_part_loader_to_construct_a_comments_part(
        self, package_: Mock, CommentsPart_load_: Mock, comments_part_: Mock
    ):
        partname = PackURI("/word/comments.xml")
        content_type = CT.WML_COMMENTS
        reltype = RT.COMMENTS
        blob = b"<w:comments/>"
        CommentsPart_load_.return_value = comments_part_

        part = PartFactory(partname, content_type, reltype, blob, package_)

        CommentsPart_load_.assert_called_once_with(partname, content_type, blob, package_)
        assert part is comments_part_

    def it_provides_access_to_its_comments_collection(
        self, Comments_: Mock, comments_: Mock, package_: Mock
    ):
        Comments_.return_value = comments_
        comments_elm = cast(CT_Comments, element("w:comments"))
        comments_part = CommentsPart(
            PackURI("/word/comments.xml"), CT.WML_COMMENTS, comments_elm, package_
        )

        comments = comments_part.comments

        Comments_.assert_called_once_with(comments_part.element, comments_part)
        assert comments is comments_

    def it_constructs_a_default_comments_part_to_help(self):
        package = Package()

        comments_part = CommentsPart.default(package)

        assert isinstance(comments_part, CommentsPart)
        assert comments_part.partname == "/word/comments.xml"
        assert comments_part.content_type == CT.WML_COMMENTS
        assert comments_part.package is package
        assert comments_part.element.tag == (
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comments"
        )
        assert len(comments_part.element) == 0

    def it_can_match_comment_extension_metadata_by_position(self, package_: Mock):
        comments_elm = cast(
            CT_Comments,
            element("w:comments/(w:comment{w:id=1},w:comment{w:id=2})"),
        )
        comments_part = CommentsPart(
            PackURI("/word/comments.xml"), CT.WML_COMMENTS, comments_elm, package_
        )
        comments_ex_part = CommentsExtendedPart(
            PackURI("/word/commentsExtended.xml"),
            CT.WML_COMMENTS_EXTENDED,
            _comments_ex_elm(
                '<w15:commentEx w15:paraId="00000001"/><w15:commentEx w15:paraId="00000002"/>'
            ),
            package_,
        )
        package_.main_document_part._comments_extended_part = comments_ex_part

        comment_ex = comments_part.comment_ex_for(comments_elm.comment_lst[1])

        assert comment_ex is comments_ex_part.element.commentEx_lst[1]

    def it_can_ensure_comment_extension_metadata_exists(self, package_: Mock):
        comments_elm = cast(
            CT_Comments,
            element("w:comments/(w:comment{w:id=1},w:comment{w:id=2})"),
        )
        comments_part = CommentsPart(
            PackURI("/word/comments.xml"), CT.WML_COMMENTS, comments_elm, package_
        )
        comments_ex_part = CommentsExtendedPart(
            PackURI("/word/commentsExtended.xml"),
            CT.WML_COMMENTS_EXTENDED,
            _comments_ex_elm('<w15:commentEx w15:paraId="00000001"/>'),
            package_,
        )
        package_.main_document_part._comments_extended_part = comments_ex_part

        comment_ex = comments_part.ensure_comment_ex(comments_elm.comment_lst[1])

        assert len(comments_ex_part.element.commentEx_lst) == 2
        assert comment_ex.paraId == "00000002"

    def it_stamps_a_matching_para_id_when_adding_a_comment(self, package_: Mock):
        comments_elm = cast(CT_Comments, element("w:comments"))
        comments_part = CommentsPart(
            PackURI("/word/comments.xml"), CT.WML_COMMENTS, comments_elm, package_
        )
        comments_ex_part = CommentsExtendedPart(
            PackURI("/word/commentsExtended.xml"),
            CT.WML_COMMENTS_EXTENDED,
            _comments_ex_elm(),
            package_,
        )
        package_.main_document_part._comments_extended_part = comments_ex_part

        comment_elm, comment_ex = comments_part.add_comment_elm()

        assert comment_elm.p_lst[0].paraId == comment_ex.paraId == "00000001"
        assert comment_elm.p_lst[0].textId == "77777777"

    def it_can_ensure_a_durable_comment_id(self, package_: Mock):
        comments_elm = cast(CT_Comments, element("w:comments"))
        comments_part = CommentsPart(
            PackURI("/word/comments.xml"), CT.WML_COMMENTS, comments_elm, package_
        )
        comments_ex_part = CommentsExtendedPart(
            PackURI("/word/commentsExtended.xml"),
            CT.WML_COMMENTS_EXTENDED,
            _comments_ex_elm(),
            package_,
        )
        comments_ids_part = CommentsIdsPart(
            PackURI("/word/commentsIds.xml"),
            CT.WML_COMMENTS_IDS,
            _comments_ids_elm(),
            package_,
        )
        package_.main_document_part._comments_extended_part = comments_ex_part
        package_.main_document_part._comments_ids_part = comments_ids_part
        comment_elm, _ = comments_part.add_comment_elm()

        comment_id = comments_part.ensure_comment_id(comment_elm)

        assert comment_id.paraId == "00000001"
        assert comment_id.durableId == "00000001"

    def it_can_ensure_extensible_thread_metadata(self, package_: Mock):
        comments_elm = cast(CT_Comments, element("w:comments"))
        comments_part = CommentsPart(
            PackURI("/word/comments.xml"), CT.WML_COMMENTS, comments_elm, package_
        )
        comments_ex_part = CommentsExtendedPart(
            PackURI("/word/commentsExtended.xml"),
            CT.WML_COMMENTS_EXTENDED,
            _comments_ex_elm(),
            package_,
        )
        comments_ids_part = CommentsIdsPart(
            PackURI("/word/commentsIds.xml"),
            CT.WML_COMMENTS_IDS,
            _comments_ids_elm(),
            package_,
        )
        comments_extensible_part = CommentsExtensiblePart(
            PackURI("/word/commentsExtensible.xml"),
            CT.WML_COMMENTS_EXTENSIBLE,
            _comments_extensible_elm(),
            package_,
        )
        package_.main_document_part._comments_extended_part = comments_ex_part
        package_.main_document_part._comments_ids_part = comments_ids_part
        package_.main_document_part._comments_extensible_part = comments_extensible_part
        comment_elm, _ = comments_part.add_comment_elm()

        comment_extensible = comments_part.ensure_comment_extensible(comment_elm)

        assert comment_extensible.durableId == "00000001"
        assert comment_extensible.dateUtc is not None

    def it_updates_existing_extensible_metadata_when_a_new_timestamp_is_provided(
        self, package_: Mock
    ):
        comments_extensible_part = CommentsExtensiblePart(
            PackURI("/word/commentsExtensible.xml"),
            CT.WML_COMMENTS_EXTENSIBLE,
            _comments_extensible_elm(
                '<w16cex:commentExtensible w16cex:durableId="0000000A" '
                'w16cex:dateUtc="2025-06-11T12:42:30Z"/>'
            ),
            package_,
        )
        timestamp = dt.datetime(2025, 6, 12, 9, 15, 0, tzinfo=dt.timezone.utc)

        comment_extensible = comments_extensible_part.ensure_comment_extensible(
            "0000000A", timestamp
        )

        assert comment_extensible.dateUtc == timestamp

    def it_can_ensure_a_person_for_threaded_comment_authors(self, package_: Mock):
        comments_elm = cast(CT_Comments, element("w:comments"))
        comments_part = CommentsPart(
            PackURI("/word/comments.xml"), CT.WML_COMMENTS, comments_elm, package_
        )
        people_part = PeoplePart(
            PackURI("/word/people.xml"),
            CT.WML_PEOPLE,
            _people_elm(),
            package_,
        )
        package_.main_document_part._people_part = people_part

        person = comments_part.ensure_person("TestAuthor")

        assert person is not None
        assert person.author == "TestAuthor"


class DescribeCommentsExtendedPart:
    def it_constructs_a_default_comments_extended_part(self):
        package = Package()

        comments_extended_part = CommentsExtendedPart.default(package)

        assert isinstance(comments_extended_part, CommentsExtendedPart)
        assert comments_extended_part.partname == "/word/commentsExtended.xml"
        assert comments_extended_part.content_type == CT.WML_COMMENTS_EXTENDED
        assert comments_extended_part.element.tag == (
            "{http://schemas.microsoft.com/office/word/2012/wordml}commentsEx"
        )
        assert len(comments_extended_part.element) == 0

    def it_knows_the_next_available_para_id(self, package_: Mock):
        comments_extended_part = CommentsExtendedPart(
            PackURI("/word/commentsExtended.xml"),
            CT.WML_COMMENTS_EXTENDED,
            _comments_ex_elm(
                '<w15:commentEx w15:paraId="00000001"/><w15:commentEx w15:paraId="0000000F"/>'
            ),
            package_,
        )

        assert comments_extended_part.next_available_para_id == "00000010"


class DescribeCommentsIdsPart:
    def it_constructs_a_default_comments_ids_part(self):
        package = Package()

        comments_ids_part = CommentsIdsPart.default(package)

        assert isinstance(comments_ids_part, CommentsIdsPart)
        assert comments_ids_part.partname == "/word/commentsIds.xml"
        assert comments_ids_part.content_type == CT.WML_COMMENTS_IDS
        assert comments_ids_part.element.tag == (
            "{http://schemas.microsoft.com/office/word/2016/wordml/cid}commentsIds"
        )
        assert len(comments_ids_part.element) == 0


class DescribeCommentsExtensiblePart:
    def it_constructs_a_default_comments_extensible_part(self):
        package = Package()

        comments_extensible_part = CommentsExtensiblePart.default(package)

        assert isinstance(comments_extensible_part, CommentsExtensiblePart)
        assert comments_extensible_part.partname == "/word/commentsExtensible.xml"
        assert comments_extensible_part.content_type == CT.WML_COMMENTS_EXTENSIBLE
        assert comments_extensible_part.element.tag == (
            "{http://schemas.microsoft.com/office/word/2018/wordml/cex}commentsExtensible"
        )
        assert len(comments_extensible_part.element) == 0


class DescribePeoplePart:
    def it_constructs_a_default_people_part(self):
        package = Package()

        people_part = PeoplePart.default(package)

        assert isinstance(people_part, PeoplePart)
        assert people_part.partname == "/word/people.xml"
        assert people_part.content_type == CT.WML_PEOPLE
        assert people_part.element.tag == (
            "{http://schemas.microsoft.com/office/word/2012/wordml}people"
        )
        assert len(people_part.element) == 0


# -- fixtures --------------------------------------------------------------------------------


@pytest.fixture
def Comments_(request: FixtureRequest) -> Mock:
    return class_mock(request, "docx.parts.comments.Comments")


@pytest.fixture
def comments_(request: FixtureRequest) -> Mock:
    return instance_mock(request, Comments)


@pytest.fixture
def comments_part_(request: FixtureRequest) -> Mock:
    return instance_mock(request, CommentsPart)


@pytest.fixture
def CommentsPart_load_(request: FixtureRequest) -> Mock:
    return method_mock(request, CommentsPart, "load", autospec=False)


@pytest.fixture
def package_(request: FixtureRequest) -> Mock:
    return instance_mock(request, Package)


def _comments_ex_elm(inner_xml: str = "") -> CT_CommentsEx:
    return cast(
        CT_CommentsEx,
        parse_xml(
            '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
            f"{inner_xml}"
            "</w15:commentsEx>"
        ),
    )


def _comments_ids_elm(inner_xml: str = "") -> CT_CommentsIds:
    return cast(
        CT_CommentsIds,
        parse_xml(
            '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">'
            f"{inner_xml}"
            "</w16cid:commentsIds>"
        ),
    )


def _comments_extensible_elm(inner_xml: str = "") -> CT_CommentsExtensible:
    return cast(
        CT_CommentsExtensible,
        parse_xml(
            '<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex">'
            f"{inner_xml}"
            "</w16cex:commentsExtensible>"
        ),
    )


def _people_elm(inner_xml: str = "") -> CT_People:
    return cast(
        CT_People,
        parse_xml(
            '<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
            f"{inner_xml}"
            "</w15:people>"
        ),
    )
