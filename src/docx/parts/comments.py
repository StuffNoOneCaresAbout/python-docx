"""Contains comments added to the document."""

from __future__ import annotations

import datetime as dt
import os
from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from docx.comments import Comments
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml.comments import (
    CT_Comment,
    CT_CommentEx,
    CT_CommentExtensible,
    CT_CommentId,
    CT_Comments,
    CT_CommentsEx,
    CT_CommentsExtensible,
    CT_CommentsIds,
    CT_People,
    CT_Person,
)
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement, parse_xml
from docx.package import Package
from docx.parts.story import StoryPart

if TYPE_CHECKING:
    from docx.oxml.text.paragraph import CT_P
    from docx.package import Package


class CommentsPart(StoryPart):
    """Container part for comments added to the document."""

    def __init__(
        self, partname: PackURI, content_type: str, element: CT_Comments, package: Package
    ):
        super().__init__(partname, content_type, element, package)
        self._comments = element

    @property
    def comments(self) -> Comments:
        """A |Comments| proxy object for the `w:comments` root element of this part."""
        return Comments(self._comments, self)

    @property
    def comments_extended_part(self) -> CommentsExtendedPart:
        """Office 2013+ comment extension metadata part."""
        return self._document_part._comments_extended_part

    def add_comment_elm(self, parent_para_id: str | None = None) -> tuple[CT_Comment, CT_CommentEx]:
        """Add comment XML and matching extension metadata, returning both elements."""
        comment_id = self.next_available_comment_id
        para_id = self.next_available_para_id
        comment_elm = self._comments.add_comment(comment_id, para_id=para_id)
        comment_ex_elm = self.comments_extended_part.add_comment_ex(para_id, parent_para_id)
        return comment_elm, comment_ex_elm

    @property
    def next_available_comment_id(self) -> int:
        return self._comments._next_available_comment_id()  # pyright: ignore[reportPrivateUsage]

    @property
    def next_available_para_id(self) -> str:
        return self.comments_extended_part.next_available_para_id

    @property
    def comments_extensible_part(self) -> CommentsExtensiblePart:
        """Office 2018+ extensible threaded-comment metadata part."""
        return self._document_part._comments_extensible_part

    @property
    def comments_ids_part(self) -> CommentsIdsPart:
        """Office 2016+ durable comment-id metadata part."""
        return self._document_part._comments_ids_part

    @property
    def people_part(self) -> PeoplePart:
        """Office 2013+ people part for threaded comment authors."""
        return self._document_part._people_part

    def comment_ex_for(self, comment_elm: CT_Comment) -> CT_CommentEx | None:
        comment_lst = self._comments.comment_lst
        try:
            idx = comment_lst.index(comment_elm)
        except ValueError:
            return None
        comment_ex_lst = self.comments_extended_part._comments_ex.commentEx_lst
        return comment_ex_lst[idx] if idx < len(comment_ex_lst) else None

    def ensure_comment_ex(self, comment_elm: CT_Comment) -> CT_CommentEx:
        """Ensure there is a `w15:commentEx` entry corresponding to `comment_elm`."""
        comment_lst = self._comments.comment_lst
        comment_ex_lst = self.comments_extended_part._comments_ex.commentEx_lst
        try:
            idx = comment_lst.index(comment_elm)
        except ValueError as e:
            raise KeyError("comment element not found in comments part") from e

        while len(comment_ex_lst) <= idx:
            missing_comment_elm = comment_lst[len(comment_ex_lst)]
            para_id = self._para_id_for_comment(missing_comment_elm)
            self.comments_extended_part.add_comment_ex(para_id)
            self._stamp_para_id(missing_comment_elm, para_id)
            comment_ex_lst = self.comments_extended_part._comments_ex.commentEx_lst

        return comment_ex_lst[idx]

    def ensure_comment_extensible(
        self, comment_elm: CT_Comment, date_utc: dt.datetime | None = None
    ) -> CT_CommentExtensible:
        """Ensure threaded extensible metadata exists for `comment_elm`."""
        comment_id = self.ensure_comment_id(comment_elm)
        return self.comments_extensible_part.ensure_comment_extensible(
            comment_id.durableId, date_utc
        )

    def comment_extensible_for(self, comment_elm: CT_Comment) -> CT_CommentExtensible | None:
        """Extensible metadata entry for `comment_elm`, or |None| if not present."""
        comment_ex = self.comment_ex_for(comment_elm)
        if comment_ex is None:
            return None
        comment_id = self.comments_ids_part.comment_id_for_para_id(comment_ex.paraId)
        if comment_id is None:
            return None
        comments_extensible = self.comments_extensible_part._comments_extensible
        return comments_extensible.get_comment_extensible_by_durable_id(
            comment_id.durableId,
        )

    def ensure_comment_id(self, comment_elm: CT_Comment) -> CT_CommentId:
        """Ensure durable comment-id metadata exists for `comment_elm`."""
        comment_ex = self.ensure_comment_ex(comment_elm)
        comment_id = self.comments_ids_part.comment_id_for_para_id(comment_ex.paraId)
        return (
            comment_id
            if comment_id is not None
            else self.comments_ids_part.add_comment_id(comment_ex.paraId)
        )

    def ensure_person(self, author: str) -> CT_Person | None:
        """Ensure `people.xml` contains a person entry for `author`."""
        return self.people_part.ensure_person(author)

    def anchor_comment_like(self, comment_id: int, parent_comment_id: int) -> None:
        """Anchor `comment_id` over the same document range as `parent_comment_id`.

        Reply comments appear in Word only when they are also referenced from the main story.
        This method mirrors the parent's `commentRangeStart`, `commentRangeEnd`, and
        `commentReference` placement, inserting the new markers immediately after the parent's.
        """
        document_elm = self._document_part.element
        start_markers = document_elm.xpath(f"//w:commentRangeStart[@w:id='{parent_comment_id}']")
        reference_elms = document_elm.xpath(f"//w:commentReference[@w:id='{parent_comment_id}']")

        if not start_markers or not reference_elms:
            return

        start_marker = start_markers[0]
        reference_run = reference_elms[0].getparent()
        if reference_run is None:
            return

        start_marker.addnext(
            OxmlElement("w:commentRangeStart", attrs={qn("w:id"): str(comment_id)})
        )
        reference_run.addnext(self._new_comment_reference_run(comment_id))
        reference_run.addnext(OxmlElement("w:commentRangeEnd", attrs={qn("w:id"): str(comment_id)}))

    def _comment_para(self, comment_elm: CT_Comment) -> CT_P | None:
        return comment_elm.p_lst[0] if comment_elm.p_lst else None

    def _para_id_for_comment(self, comment_elm: CT_Comment) -> str:
        comment_para = self._comment_para(comment_elm)
        if comment_para is not None and comment_para.paraId is not None:
            return comment_para.paraId
        return self.next_available_para_id

    def _stamp_para_id(self, comment_elm: CT_Comment, para_id: str) -> None:
        comment_para = self._comment_para(comment_elm)
        if comment_para is not None:
            comment_para.paraId = para_id
            if comment_para.textId is None:
                comment_para.textId = "77777777"

    def _new_comment_reference_run(self, comment_id: int):
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        rStyle = OxmlElement("w:rStyle", attrs={qn("w:val"): "CommentReference"})
        rPr.append(rStyle)
        r.append(rPr)
        r.append(OxmlElement("w:commentReference", attrs={qn("w:id"): str(comment_id)}))
        return r

    @classmethod
    def default(cls, package: Package) -> Self:
        """A newly created comments part, containing a default empty `w:comments` element."""
        partname = PackURI("/word/comments.xml")
        content_type = CT.WML_COMMENTS
        element = cast("CT_Comments", parse_xml(cls._default_comments_xml()))
        return cls(partname, content_type, element, package)

    @classmethod
    def _default_comments_xml(cls) -> bytes:
        """A byte-string containing XML for a default comments part."""
        path = os.path.join(os.path.split(__file__)[0], "..", "templates", "default-comments.xml")
        with open(path, "rb") as f:
            xml_bytes = f.read()
        return xml_bytes


class CommentsExtendedPart(StoryPart):
    """Container part for Office 2013+ comment extension metadata."""

    def __init__(
        self, partname: PackURI, content_type: str, element: CT_CommentsEx, package: Package
    ):
        super().__init__(partname, content_type, element, package)
        self._comments_ex = element

    def add_comment_ex(
        self, para_id: str | None = None, parent_para_id: str | None = None, done: bool = False
    ) -> CT_CommentEx:
        para_id = self.next_available_para_id if para_id is None else para_id
        return self._comments_ex.add_comment_ex(para_id, parent_para_id=parent_para_id, done=done)

    @classmethod
    def default(cls, package: Package) -> Self:
        partname = PackURI("/word/commentsExtended.xml")
        content_type = CT.WML_COMMENTS_EXTENDED
        element = cast("CT_CommentsEx", parse_xml(cls._default_comments_extended_xml()))
        return cls(partname, content_type, element, package)

    @property
    def next_available_para_id(self) -> str:
        used_para_ids = [
            int(comment_ex.paraId, 16)
            for comment_ex in self._comments_ex.commentEx_lst
            if comment_ex.paraId is not None
        ]
        next_id = max(used_para_ids, default=0) + 1
        return f"{next_id:08X}"

    @classmethod
    def _default_comments_extended_xml(cls) -> bytes:
        return (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            b"<w15:commentsEx "
            b'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            b'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
            b'mc:Ignorable="w15"/>\n'
        )


class CommentsIdsPart(XmlPart):
    """Container part for Office 2016+ durable comment IDs."""

    def __init__(
        self, partname: PackURI, content_type: str, element: CT_CommentsIds, package: Package
    ):
        super().__init__(partname, content_type, element, package)
        self._comments_ids = element

    def add_comment_id(self, para_id: str) -> CT_CommentId:
        return self._comments_ids.add_comment_id(para_id, self.next_available_durable_id)

    def comment_id_for_para_id(self, para_id: str) -> CT_CommentId | None:
        return self._comments_ids.get_comment_id_by_para_id(para_id)

    @classmethod
    def default(cls, package: Package) -> Self:
        partname = PackURI("/word/commentsIds.xml")
        content_type = CT.WML_COMMENTS_IDS
        element = cast("CT_CommentsIds", parse_xml(cls._default_comments_ids_xml()))
        return cls(partname, content_type, element, package)

    @property
    def next_available_durable_id(self) -> str:
        used_durable_ids = [
            int(comment_id.durableId, 16)
            for comment_id in self._comments_ids.commentId_lst
            if comment_id.durableId is not None
        ]
        next_id = max(used_durable_ids, default=0) + 1
        return f"{next_id:08X}"

    @classmethod
    def _default_comments_ids_xml(cls) -> bytes:
        return (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            b"<w16cid:commentsIds "
            b'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            b'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" '
            b'mc:Ignorable="w16cid"/>\n'
        )


class CommentsExtensiblePart(XmlPart):
    """Container part for Office 2018+ threaded-comment metadata."""

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        element: CT_CommentsExtensible,
        package: Package,
    ):
        super().__init__(partname, content_type, element, package)
        self._comments_extensible = element

    def add_comment_extensible(
        self, durable_id: str, date_utc: dt.datetime | None = None
    ) -> CT_CommentExtensible:
        return self._comments_extensible.add_comment_extensible(durable_id, date_utc)

    @classmethod
    def default(cls, package: Package) -> Self:
        partname = PackURI("/word/commentsExtensible.xml")
        content_type = CT.WML_COMMENTS_EXTENSIBLE
        element = cast("CT_CommentsExtensible", parse_xml(cls._default_comments_extensible_xml()))
        return cls(partname, content_type, element, package)

    def ensure_comment_extensible(
        self, durable_id: str, date_utc: dt.datetime | None = None
    ) -> CT_CommentExtensible:
        comment_extensible = self._comments_extensible.get_comment_extensible_by_durable_id(
            durable_id
        )
        if comment_extensible is None:
            return self.add_comment_extensible(durable_id, date_utc)

        if date_utc is not None:
            comment_extensible.dateUtc = date_utc.astimezone(dt.timezone.utc)

        return comment_extensible

    @classmethod
    def _default_comments_extensible_xml(cls) -> bytes:
        return (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            b"<w16cex:commentsExtensible "
            b'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            b'xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" '
            b'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" '
            b'mc:Ignorable="w16 w16cex"/>\n'
        )


class PeoplePart(XmlPart):
    """Container part for Office 2013+ people metadata."""

    def __init__(self, partname: PackURI, content_type: str, element: CT_People, package: Package):
        super().__init__(partname, content_type, element, package)
        self._people = element

    @classmethod
    def default(cls, package: Package) -> Self:
        partname = PackURI("/word/people.xml")
        content_type = CT.WML_PEOPLE
        element = cast("CT_People", parse_xml(cls._default_people_xml()))
        return cls(partname, content_type, element, package)

    def ensure_person(self, author: str) -> CT_Person | None:
        if author == "":
            return None
        person = self._people.get_person_by_author(author)
        return person if person is not None else self._people.add_person(author)

    @classmethod
    def _default_people_xml(cls) -> bytes:
        return (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            b'<w15:people xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            b'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
            b'mc:Ignorable="w15"/>\n'
        )
