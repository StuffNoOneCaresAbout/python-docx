"""Paragraph-related proxy types."""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, Iterator, List, cast

from docx.enum.style import WD_STYLE_TYPE
from docx.revisions import (
    TrackedChange,
    TrackedDeletion,
    TrackedInsertion,
    TrackedReplacement,
    paragraph_accepted_text,
    paragraph_add_tracked_deletion,
    paragraph_add_tracked_insertion,
    paragraph_add_tracked_insertion_after,
    paragraph_add_tracked_insertion_at,
    paragraph_add_tracked_insertion_before,
    paragraph_comment_range_runs,
    paragraph_deleted_text,
    paragraph_deletions,
    paragraph_has_track_changes,
    paragraph_insertions,
    paragraph_iter_inner_content,
    paragraph_replace_tracked,
    paragraph_replace_tracked_at,
    paragraph_track_changes,
)
from docx.shared import StoryChild
from docx.styles.style import ParagraphStyle
from docx.text.hyperlink import Hyperlink
from docx.text.pagebreak import RenderedPageBreak
from docx.text.parfmt import ParagraphFormat
from docx.text.run import Run

if TYPE_CHECKING:
    import docx.types as t
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml.text.paragraph import CT_P
    from docx.styles.style import CharacterStyle


class Paragraph(StoryChild):
    """Proxy object wrapping a `<w:p>` element."""

    def __init__(self, p: CT_P, parent: t.ProvidesStoryPart):
        super(Paragraph, self).__init__(parent)
        self._p = self._element = p

    def add_run(self, text: str | None = None, style: str | CharacterStyle | None = None) -> Run:
        """Append run containing `text` and having character-style `style`.

        `text` can contain tab (``\\t``) characters, which are converted to the
        appropriate XML form for a tab. `text` can also include newline (``\\n``) or
        carriage return (``\\r``) characters, each of which is converted to a line
        break. When `text` is `None`, the new run is empty.
        """
        r = self._p.add_r()
        run = Run(r, self)
        if text:
            run.text = text
        if style:
            run.style = style
        return run

    @property
    def alignment(self) -> WD_PARAGRAPH_ALIGNMENT | None:
        """A member of the :ref:`WdParagraphAlignment` enumeration specifying the
        justification setting for this paragraph.

        A value of |None| indicates the paragraph has no directly-applied alignment
        value and will inherit its alignment value from its style hierarchy. Assigning
        |None| to this property removes any directly-applied alignment value.
        """
        return self._p.alignment

    @alignment.setter
    def alignment(self, value: WD_PARAGRAPH_ALIGNMENT):
        self._p.alignment = value

    def clear(self):
        """Return this same paragraph after removing all its content.

        Paragraph-level formatting, such as style, is preserved.
        """
        self._p.clear_content()
        return self

    @property
    def contains_page_break(self) -> bool:
        """`True` when one or more rendered page-breaks occur in this paragraph."""
        return bool(self._p.lastRenderedPageBreaks)

    @property
    def hyperlinks(self) -> List[Hyperlink]:
        """A |Hyperlink| instance for each hyperlink in this paragraph."""
        return [Hyperlink(hyperlink, self) for hyperlink in self._p.hyperlink_lst]

    def insert_paragraph_before(
        self, text: str | None = None, style: str | ParagraphStyle | None = None
    ) -> Paragraph:
        """Return a newly created paragraph, inserted directly before this paragraph.

        If `text` is supplied, the new paragraph contains that text in a single run. If
        `style` is provided, that style is assigned to the new paragraph.
        """
        paragraph = self._insert_paragraph_before()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    def iter_inner_content(
        self, include_revisions: bool = False
    ) -> Iterator[Run | Hyperlink | TrackedInsertion | TrackedDeletion]:
        """Generate the runs and hyperlinks in this paragraph, in the order they appear.

        The content in a paragraph consists of both runs and hyperlinks. This method
        allows accessing each of those separately, in document order, for when the
        precise position of the hyperlink within the paragraph text is important. Note
        that a hyperlink itself contains runs.
        """
        yield from paragraph_iter_inner_content(self, include_revisions=include_revisions)

    @property
    def paragraph_format(self):
        """The |ParagraphFormat| object providing access to the formatting properties
        for this paragraph, such as line spacing and indentation."""
        return ParagraphFormat(self._element)

    @property
    def rendered_page_breaks(self) -> List[RenderedPageBreak]:
        """All rendered page-breaks in this paragraph.

        Most often an empty list, sometimes contains one page-break, but can contain
        more than one is rare or contrived cases.
        """
        return [RenderedPageBreak(lrpb, self) for lrpb in self._p.lastRenderedPageBreaks]

    @property
    def runs(self) -> List[Run]:
        """Sequence of |Run| instances corresponding to the <w:r> elements in this
        paragraph."""
        return [Run(r, self) for r in self._p.r_lst]

    @property
    def has_track_changes(self) -> bool:
        """True when this paragraph contains tracked insertions or deletions."""
        return paragraph_has_track_changes(self)

    @property
    def insertions(self) -> List[TrackedInsertion]:
        """Tracked insertions in this paragraph, in document order."""
        return paragraph_insertions(self)

    @property
    def deletions(self) -> List[TrackedDeletion]:
        """Tracked deletions in this paragraph, in document order."""
        return paragraph_deletions(self)

    @property
    def track_changes(self) -> List[TrackedChange]:
        """Tracked insertions and deletions in this paragraph, in document order."""
        return paragraph_track_changes(self)

    @property
    def style(self) -> ParagraphStyle | None:
        """Read/Write.

        |_ParagraphStyle| object representing the style assigned to this paragraph. If
        no explicit style is assigned to this paragraph, its value is the default
        paragraph style for the document. A paragraph style name can be assigned in lieu
        of a paragraph style object. Assigning |None| removes any applied style, making
        its effective value the default paragraph style for the document.
        """
        style_id = self._p.style
        style = self.part.get_style(style_id, WD_STYLE_TYPE.PARAGRAPH)
        return cast(ParagraphStyle, style)

    @style.setter
    def style(self, style_or_name: str | ParagraphStyle | None):
        style_id = self.part.get_style_id(style_or_name, WD_STYLE_TYPE.PARAGRAPH)
        self._p.style = style_id

    @property
    def text(self) -> str:
        """The textual content of this paragraph.

        The text includes the visible-text portion of any hyperlinks in the paragraph.
        Tabs and line breaks in the XML are mapped to ``\\t`` and ``\\n`` characters
        respectively. Deleted content in tracked revisions is included; inserted
        content is excluded.

        Assigning text to this property causes all existing paragraph content to be
        replaced with a single run containing the assigned text. A ``\\t`` character in
        the text is mapped to a ``<w:tab/>`` element and each ``\\n`` or ``\\r``
        character is mapped to a line break. Paragraph-level formatting, such as style,
        is preserved. All run-level formatting, such as bold or italic, is removed.
        """
        return self._p.text

    @property
    def accepted_text(self) -> str:
        """Visible paragraph text with tracked insertions included and deletions omitted."""
        return paragraph_accepted_text(self)

    @property
    def deleted_text(self) -> str:
        """Deleted-only text present in this paragraph's tracked revisions."""
        return paragraph_deleted_text(self)

    @text.setter
    def text(self, text: str | None):
        self.clear()
        self.add_run(text)

    def _insert_paragraph_before(self):
        """Return a newly created paragraph, inserted directly before this paragraph."""
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)

    def add_tracked_insertion(
        self,
        text: str | None = None,
        style: str | CharacterStyle | None = None,
        author: str = "",
        revision_id: int | None = None,
    ) -> TrackedInsertion:
        """Append a tracked insertion containing a run with the specified text."""
        return paragraph_add_tracked_insertion(
            self, text=text, style=style, author=author, revision_id=revision_id
        )

    def add_tracked_insertion_at(
        self,
        offset: int,
        text: str | None = None,
        style: str | CharacterStyle | None = None,
        author: str = "",
        revision_id: int | None = None,
    ) -> TrackedInsertion:
        """Insert a tracked insertion at `offset` in `accepted_text`."""
        return paragraph_add_tracked_insertion_at(
            self,
            offset,
            text=text,
            style=style,
            author=author,
            revision_id=revision_id,
        )

    def add_tracked_insertion_before(
        self,
        search_text: str,
        text: str | None = None,
        style: str | CharacterStyle | None = None,
        author: str = "",
        revision_id: int | None = None,
    ) -> TrackedInsertion:
        """Insert tracked text before a unique `search_text` match in `accepted_text`."""
        return paragraph_add_tracked_insertion_before(
            self,
            search_text,
            text=text,
            style=style,
            author=author,
            revision_id=revision_id,
        )

    def add_tracked_insertion_after(
        self,
        search_text: str,
        text: str | None = None,
        style: str | CharacterStyle | None = None,
        author: str = "",
        revision_id: int | None = None,
    ) -> TrackedInsertion:
        """Insert tracked text after a unique `search_text` match in `accepted_text`."""
        return paragraph_add_tracked_insertion_after(
            self,
            search_text,
            text=text,
            style=style,
            author=author,
            revision_id=revision_id,
        )

    def add_tracked_deletion(
        self, start: int, end: int, author: str = "", revision_id: int | None = None
    ) -> TrackedDeletion | None:
        """Wrap accepted-view text in a tracked deletion range."""
        return paragraph_add_tracked_deletion(
            self, start, end, author=author, revision_id=revision_id
        )

    def replace_tracked(
        self, search_text: str, replace_text: str, author: str = ""
    ) -> List[TrackedReplacement]:
        """Replace all accepted-view occurrences using tracked deletion + insertion.

        Returns the created replacements in document order.
        """
        return paragraph_replace_tracked(self, search_text, replace_text, author=author)

    def replace_tracked_at(
        self, start: int, end: int, replace_text: str, author: str = ""
    ) -> TrackedReplacement:
        """Replace accepted-view text at offsets using tracked deletion + insertion.

        Returns a :class:`.TrackedReplacement` containing the tracked deletion and insertion
        created for this single replacement, allowing comments or other follow-up
        operations to be applied immediately.
        """
        return paragraph_replace_tracked_at(self, start, end, replace_text, author=author)

    def add_comment(
        self,
        text: str | None = "",
        author: str = "",
        initials: str | None = "",
        timestamp: dt.datetime | None = None,
    ):
        """Add a comment spanning all runs in this paragraph."""
        document = self.part._document_part.document  # pyright: ignore[reportPrivateUsage]
        if not self.runs:
            run = self.add_run()
            return (
                document.add_comment(run, text=text, author=author, initials=initials)
                if timestamp is None
                else document.add_comment(
                    run,
                    text=text,
                    author=author,
                    initials=initials,
                    timestamp=timestamp,
                )
            )
        return (
            document.add_comment(self.runs, text=text, author=author, initials=initials)
            if timestamp is None
            else document.add_comment(
                self.runs,
                text=text,
                author=author,
                initials=initials,
                timestamp=timestamp,
            )
        )

    def add_comment_range(
        self,
        start: int,
        end: int,
        text: str | None = "",
        author: str = "",
        initials: str | None = "",
        timestamp: dt.datetime | None = None,
    ):
        """Add a comment spanning accepted-text offsets within this paragraph.

        Offsets are based on `accepted_text`, so tracked insertions are included and
        tracked deletions are excluded.
        """
        document = self.part._document_part.document  # pyright: ignore[reportPrivateUsage]
        first_run, last_run = paragraph_comment_range_runs(self, start, end)
        return (
            document.add_comment([first_run, last_run], text=text, author=author, initials=initials)
            if timestamp is None
            else document.add_comment(
                [first_run, last_run],
                text=text,
                author=author,
                initials=initials,
                timestamp=timestamp,
            )
        )
