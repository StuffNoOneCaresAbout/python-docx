# pyright: reportAssignmentType=false

"""Custom element classes related to document comments."""

from __future__ import annotations

import datetime as dt
import hashlib
from typing import TYPE_CHECKING, Callable, cast

from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.oxml.simpletypes import (
    ST_DateTime,
    ST_DecimalNumber,
    ST_LongHexNumber,
    ST_OnOff,
    ST_String,
)
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


class CT_Comments(BaseOxmlElement):
    """`w:comments` element, the root element for the comments part.

    Simply contains a collection of `w:comment` elements, each representing a single comment. Each
    contained comment is identified by a unique `w:id` attribute, used to reference the comment
    from the document text. The offset of the comment in this collection is arbitrary; it is
    essentially a _set_ implemented as a list.
    """

    # -- type-declarations to fill in the gaps for metaclass-added methods --
    comment_lst: list[CT_Comment]

    comment = ZeroOrMore("w:comment")

    def add_comment(self, comment_id: int | None = None, para_id: str | None = None) -> CT_Comment:
        """Return newly added `w:comment` child of this `w:comments`.

        The returned `w:comment` element is the minimum valid value, having a `w:id` value unique
        within the existing comments and the required `w:author` attribute present but set to the
        empty string. It's content is limited to a single run containing the necessary annotation
        reference but no text. Content is added by adding runs to this first paragraph and by
        adding additional paragraphs as needed.
        """
        next_id = self._next_available_comment_id() if comment_id is None else comment_id
        para_id_attr = f' w14:paraId="{para_id}" w14:textId="77777777"' if para_id else ""
        w14_ns = (
            ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"' if para_id else ""
        )
        comment = cast(
            CT_Comment,
            parse_xml(
                f'<w:comment {nsdecls("w")}{w14_ns} w:id="{next_id}" w:author="">'
                f"  <w:p{para_id_attr}>"
                f"    <w:pPr>"
                f'      <w:pStyle w:val="CommentText"/>'
                f"    </w:pPr>"
                f"    <w:r>"
                f"      <w:rPr>"
                f'        <w:rStyle w:val="CommentReference"/>'
                f"      </w:rPr>"
                f"      <w:annotationRef/>"
                f"    </w:r>"
                f"  </w:p>"
                f"</w:comment>"
            ),
        )
        self.append(comment)
        return comment

    def get_comment_by_id(self, comment_id: int) -> CT_Comment | None:
        """Return the `w:comment` element identified by `comment_id`, or |None| if not found."""
        comment_elms = self.xpath(f"(./w:comment[@w:id='{comment_id}'])[1]")
        return comment_elms[0] if comment_elms else None

    def _next_available_comment_id(self) -> int:
        """The next available comment id.

        According to the schema, this can be any positive integer, as big as you like, and the
        default mechanism is to use `max() + 1`. However, if that yields a value larger than will
        fit in a 32-bit signed integer, we take a more deliberate approach to use the first
        ununsed integer starting from 0.
        """
        used_ids = [int(x) for x in self.xpath("./w:comment/@w:id")]

        next_id = max(used_ids, default=-1) + 1

        if next_id <= 2**31 - 1:
            return next_id

        # -- fall-back to enumerating all used ids to find the first unused one --
        for expected, actual in enumerate(sorted(used_ids)):
            if expected != actual:
                return expected

        return len(used_ids)


class CT_Comment(BaseOxmlElement):
    """`w:comment` element, representing a single comment.

    A comment is a so-called "story" and can contain paragraphs and tables much like a table-cell.
    While probably most often used for a single sentence or phrase, a comment can contain rich
    content, including multiple rich-text paragraphs, hyperlinks, images, and tables.
    """

    # -- attributes on `w:comment` --
    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    author: str = RequiredAttribute("w:author", ST_String)  # pyright: ignore[reportAssignmentType]
    initials: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:initials", ST_String
    )
    date: dt.datetime | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:date", ST_DateTime
    )

    # -- children --

    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    # -- type-declarations for methods added by metaclass --

    add_p: Callable[[], CT_P]
    p_lst: list[CT_P]
    tbl_lst: list[CT_Tbl]
    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    @property
    def inner_content_elements(self) -> list[CT_P | CT_Tbl]:
        """Generate all `w:p` and `w:tbl` elements in this comment."""
        return self.xpath("./w:p | ./w:tbl")

    def set_local_date(self, value: dt.datetime) -> None:
        """Set `w:date` using Word-style local authored time serialization."""
        self.set(qn("w:date"), value.replace(tzinfo=None).isoformat(timespec="seconds"))


class CT_CommentsEx(BaseOxmlElement):
    """`w15:commentsEx` element, root element for comment extension metadata."""

    commentEx_lst: list[CT_CommentEx]

    commentEx = ZeroOrMore("w15:commentEx")

    def add_comment_ex(
        self, para_id: str, parent_para_id: str | None = None, done: bool = False
    ) -> CT_CommentEx:
        comment_ex = cast(
            CT_CommentEx,
            parse_xml(
                f'<w15:commentEx {nsdecls("w15")} w15:paraId="{para_id}"'
                f"{f' w15:paraIdParent="{parent_para_id}"' if parent_para_id else ''}"
                f' w15:done="{1 if done else 0}"/>'
            ),
        )
        self.append(comment_ex)
        return comment_ex


class CT_CommentEx(BaseOxmlElement):
    """`w15:commentEx` element, extra metadata for a single comment."""

    paraId: str = RequiredAttribute("w15:paraId", ST_LongHexNumber)  # pyright: ignore[reportAssignmentType]
    paraIdParent: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w15:paraIdParent", ST_LongHexNumber
    )
    done: bool = OptionalAttribute("w15:done", ST_OnOff, default=False)  # ty: ignore[invalid-assignment]


class CT_CommentsIds(BaseOxmlElement):
    """`w16cid:commentsIds` element, mapping paraId to durableId."""

    commentId_lst: list[CT_CommentId]

    commentId = ZeroOrMore("w16cid:commentId")

    def add_comment_id(self, para_id: str, durable_id: str) -> CT_CommentId:
        comment_id = cast(
            CT_CommentId,
            parse_xml(
                f"<w16cid:commentId {nsdecls('w16cid')} "
                f'w16cid:paraId="{para_id}" w16cid:durableId="{durable_id}"/>'
            ),
        )
        self.append(comment_id)
        return comment_id

    def get_comment_id_by_para_id(self, para_id: str) -> CT_CommentId | None:
        comment_ids = self.xpath(f"(./w16cid:commentId[@w16cid:paraId='{para_id}'])[1]")
        return comment_ids[0] if comment_ids else None


class CT_CommentId(BaseOxmlElement):
    """`w16cid:commentId` element."""

    paraId: str = RequiredAttribute("w16cid:paraId", ST_LongHexNumber)  # pyright: ignore[reportAssignmentType]
    durableId: str = RequiredAttribute("w16cid:durableId", ST_LongHexNumber)  # pyright: ignore[reportAssignmentType]


class CT_CommentsExtensible(BaseOxmlElement):
    """`w16cex:commentsExtensible` element, extra metadata for threaded replies."""

    commentExtensible_lst: list[CT_CommentExtensible]

    commentExtensible = ZeroOrMore("w16cex:commentExtensible")

    def add_comment_extensible(
        self, durable_id: str, date_utc: dt.datetime | None = None
    ) -> CT_CommentExtensible:
        if date_utc is None:
            date_utc = dt.datetime.now(dt.timezone.utc)
        date_utc = date_utc.astimezone(dt.timezone.utc).replace(microsecond=0)
        date_utc_str = date_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
        comment_extensible = cast(
            CT_CommentExtensible,
            parse_xml(
                f"<w16cex:commentExtensible {nsdecls('w16cex')} "
                f'w16cex:durableId="{durable_id}" w16cex:dateUtc="{date_utc_str}"/>'
            ),
        )
        self.append(comment_extensible)
        return comment_extensible

    def get_comment_extensible_by_durable_id(self, durable_id: str) -> CT_CommentExtensible | None:
        comment_extensibles = self.xpath(
            f"(./w16cex:commentExtensible[@w16cex:durableId='{durable_id}'])[1]"
        )
        return comment_extensibles[0] if comment_extensibles else None


class CT_CommentExtensible(BaseOxmlElement):
    """`w16cex:commentExtensible` element."""

    durableId: str = RequiredAttribute("w16cex:durableId", ST_LongHexNumber)  # pyright: ignore[reportAssignmentType]
    dateUtc: dt.datetime | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w16cex:dateUtc", ST_DateTime
    )


class CT_People(BaseOxmlElement):
    """`w15:people` element, root element for threaded comment authors."""

    person_lst: list[CT_Person]

    person = ZeroOrMore("w15:person")

    def add_person(self, author: str) -> CT_Person:
        person = cast(
            CT_Person,
            parse_xml(
                f'<w15:person {nsdecls("w15")} w15:author="{author}">'
                f'<w15:presenceInfo w15:providerId="Windows Live" '
                f'w15:userId="{hashlib.sha1(author.encode("utf-8")).hexdigest()[:16]}"/>'
                f"</w15:person>"
            ),
        )
        self.append(person)
        return person

    def get_person_by_author(self, author: str) -> CT_Person | None:
        people = self.xpath(f"(./w15:person[@w15:author='{author}'])[1]")
        return people[0] if people else None


class CT_Person(BaseOxmlElement):
    """`w15:person` element."""

    author: str = RequiredAttribute("w15:author", ST_String)  # pyright: ignore[reportAssignmentType]
    presenceInfo = ZeroOrOne("w15:presenceInfo")


class CT_PresenceInfo(BaseOxmlElement):
    """`w15:presenceInfo` element."""

    providerId: str = RequiredAttribute("w15:providerId", ST_String)  # pyright: ignore[reportAssignmentType]
    userId: str = RequiredAttribute("w15:userId", ST_String)  # pyright: ignore[reportAssignmentType]
