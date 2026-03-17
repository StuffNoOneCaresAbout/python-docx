"""Public proxy objects and helpers for tracked changes."""

from __future__ import annotations

import contextlib
import datetime as dt
from copy import deepcopy
from dataclasses import dataclass
from typing import TYPE_CHECKING, Iterator, List

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.shared import Parented

if TYPE_CHECKING:
    from lxml.etree import _Element as etree_Element  # pyright: ignore[reportPrivateUsage]

    import docx.types as t
    from docx.comments import Comment
    from docx.oxml.revisions import CT_RunTrackChange
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run


def revision_attrs(rev_id: int, author: str, now: str) -> dict[str, str]:
    return {qn("w:id"): str(rev_id), qn("w:author"): author, qn("w:date"): now}


def make_text_run(text: str):
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    if text.startswith(" ") or text.endswith(" "):
        t.set(qn("xml:space"), "preserve")
    r.append(t)
    return r


def make_comment_reference_run(comment_id: int):
    r = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    r_style = OxmlElement("w:rStyle", attrs={qn("w:val"): "CommentReference"})
    r_pr.append(r_style)
    r.append(r_pr)
    r.append(OxmlElement("w:commentReference", attrs={qn("w:id"): str(comment_id)}))
    return r


def make_del_element(deleted_text: str, author: str, rev_id: int, now: str):
    del_elem = OxmlElement("w:del", attrs=revision_attrs(rev_id, author, now))
    del_r = OxmlElement("w:r")
    del_text_elem = OxmlElement("w:delText")
    del_text_elem.text = deleted_text
    if deleted_text.startswith(" ") or deleted_text.endswith(" "):
        del_text_elem.set(qn("xml:space"), "preserve")
    del_r.append(del_text_elem)
    del_elem.append(del_r)
    return del_elem


def make_ins_element(insert_text: str, author: str, rev_id: int, now: str):
    ins_elem = OxmlElement("w:ins", attrs=revision_attrs(rev_id, author, now))
    ins_elem.append(make_text_run(insert_text))
    return ins_elem


def next_revision_id(element: etree_Element) -> int:
    max_id = 0
    for ins_or_del in element.xpath("//w:ins | //w:del"):
        id_val = ins_or_del.get(qn("w:id"))
        if id_val is not None:
            with contextlib.suppress(ValueError):
                max_id = max(max_id, int(id_val))
    return max_id + 1


@dataclass(frozen=True)
class _AcceptedSpan:
    kind: str
    element: etree_Element
    text: str
    start: int
    end: int


class TrackedChange(Parented):
    """Base proxy type for a tracked insertion or deletion."""

    def __init__(self, element: CT_RunTrackChange, parent: t.ProvidesStoryPart):
        super().__init__(parent)
        self._element = element

    @property
    def author(self) -> str:
        return self._element.author

    @author.setter
    def author(self, value: str):
        self._element.author = value

    @property
    def date(self) -> dt.datetime | None:
        return self._element.date_value

    @date.setter
    def date(self, value: dt.datetime | None):
        self._element.date_value = value

    @property
    def revision_id(self) -> int:
        return self._element.id

    @revision_id.setter
    def revision_id(self, value: int):
        self._element.id = value

    @property
    def is_block_level(self) -> bool:
        return bool(self._element.inner_content_elements)

    @property
    def is_run_level(self) -> bool:
        return bool(self._element.run_content_elements)

    def add_comment(
        self,
        text: str | None = "",
        author: str = "",
        initials: str | None = "",
        timestamp: dt.datetime | None = None,
    ) -> Comment:
        """Add a comment anchored to the content of this tracked change."""
        document_part = getattr(self.part, "_document_part", None)
        if self.part is not document_part:
            raise ValueError(
                "comments can only be added to tracked changes in the main document story"
            )

        if self.is_run_level:
            comment = self.part.comments.add_comment(
                text=text or "", author=author, initials=initials, timestamp=timestamp
            )
            self._insert_comment_range_around_revision(comment.comment_id)
            return comment

        first_run, last_run = self._comment_anchor_runs()
        comment = self.part.comments.add_comment(
            text=text or "", author=author, initials=initials, timestamp=timestamp
        )
        first_run.mark_comment_range(last_run, comment.comment_id)
        return comment

    def iter_inner_content(self) -> Iterator[Paragraph | Table]:
        from docx.table import Table
        from docx.text.paragraph import Paragraph

        for element in self._element.inner_content_elements:
            if element.tag == qn("w:p"):
                yield Paragraph(element, self._parent)  # pyright: ignore[reportArgumentType]
            elif element.tag == qn("w:tbl"):
                yield Table(element, self._parent)  # pyright: ignore[reportArgumentType]

    def iter_runs(self) -> Iterator[Run]:
        from docx.text.run import Run

        for r in self._element.run_content_elements:
            yield Run(r, self._parent)  # pyright: ignore[reportArgumentType]

    @property
    def paragraphs(self) -> List[Paragraph]:
        from docx.text.paragraph import Paragraph

        return [Paragraph(p, self._parent) for p in self._element.p_lst]  # pyright: ignore[reportArgumentType]

    @property
    def runs(self) -> List[Run]:
        from docx.text.run import Run

        return [Run(r, self._parent) for r in self._element.r_lst]  # pyright: ignore[reportArgumentType]

    def accept(self) -> None:
        raise NotImplementedError()

    def reject(self) -> None:
        raise NotImplementedError()

    def _comment_anchor_runs(self) -> tuple[Run, Run]:
        """Return first and last runs suitable for anchoring a comment."""
        if self.is_block_level:
            block_items = list(self.iter_inner_content())
            if not block_items:
                raise ValueError("tracked change has no block content to anchor a comment")
            return (
                _first_block_item_comment_anchor_run(block_items[0]),
                _last_block_item_comment_anchor_run(block_items[-1]),
            )

        raise ValueError("tracked change has no anchorable content for a comment")

    def _insert_comment_range_around_revision(self, comment_id: int) -> None:
        """Insert comment markers around this run-level revision wrapper."""
        parent = self._element.getparent()
        if parent is None:
            raise ValueError("tracked change has no parent element")

        start_idx = list(parent).index(self._element)
        parent.insert(
            start_idx,
            OxmlElement("w:commentRangeStart", attrs={qn("w:id"): str(comment_id)}),
        )

        end_idx = list(parent).index(self._element) + 1
        while end_idx < len(parent) and _is_comment_range_trailer(parent[end_idx]):
            end_idx += 1

        parent.insert(
            end_idx,
            OxmlElement("w:commentRangeEnd", attrs={qn("w:id"): str(comment_id)}),
        )
        parent.insert(end_idx + 1, make_comment_reference_run(comment_id))


def _is_comment_range_trailer(element: etree_Element) -> bool:
    """True when `element` is an existing comment end marker or reference run."""
    if element.tag == qn("w:commentRangeEnd"):
        return True
    if element.tag != qn("w:r"):
        return False
    return bool(element.xpath("./w:commentReference"))


def _paragraph_comment_anchor_run(paragraph, *, last: bool):
    """Return an anchor run in `paragraph`, creating one if needed."""
    runs = paragraph.runs
    if runs:
        return runs[-1] if last else runs[0]
    return paragraph.add_run()


def _first_block_item_comment_anchor_run(block_item):
    """Return the first run that can anchor a comment in `block_item`."""
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run

    if isinstance(block_item, Paragraph):
        return _paragraph_comment_anchor_run(block_item, last=False)

    if isinstance(block_item, Table):
        return Run(block_item._tbl._first_comment_anchor_run, block_item)  # pyright: ignore[reportPrivateUsage]

    raise ValueError(f"unsupported block item for comment anchoring: {type(block_item).__name__}")


def _last_block_item_comment_anchor_run(block_item):
    """Return the last run that can anchor a comment in `block_item`."""
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run

    if isinstance(block_item, Paragraph):
        return _paragraph_comment_anchor_run(block_item, last=True)

    if isinstance(block_item, Table):
        return Run(block_item._tbl._last_comment_anchor_run, block_item)  # pyright: ignore[reportPrivateUsage]

    raise ValueError(f"unsupported block item for comment anchoring: {type(block_item).__name__}")


class TrackedInsertion(TrackedChange):
    @property
    def text(self) -> str:
        if self.is_block_level:
            return "\n".join(p.text for p in self.paragraphs)
        return "".join(r.text for r in self.runs)

    def accept(self) -> None:
        parent = self._element.getparent()
        if parent is None:
            return
        index = list(parent).index(self._element)
        for child in reversed(list(self._element)):
            parent.insert(index, child)
        parent.remove(self._element)

    def reject(self) -> None:
        parent = self._element.getparent()
        if parent is not None:
            parent.remove(self._element)


class TrackedDeletion(TrackedChange):
    @property
    def text(self) -> str:
        if self.is_block_level:
            return "\n".join(p.text for p in self.paragraphs)
        del_texts = self._element.xpath(".//w:delText")
        if del_texts:
            return "".join(t.text or "" for t in del_texts)
        return "".join(r.text for r in self.runs)

    def accept(self) -> None:
        parent = self._element.getparent()
        if parent is not None:
            parent.remove(self._element)

    def reject(self) -> None:
        parent = self._element.getparent()
        if parent is None:
            return
        for del_text in self._element.xpath(".//w:delText"):
            t_elem = OxmlElement("w:t")
            t_elem.text = del_text.text
            space_val = del_text.get(qn("xml:space"))
            if space_val:
                t_elem.set(qn("xml:space"), space_val)
            del_text_parent = del_text.getparent()
            if del_text_parent is not None:
                del_text_parent.replace(del_text, t_elem)
        index = list(parent).index(self._element)
        for child in reversed(list(self._element)):
            parent.insert(index, child)
        parent.remove(self._element)


def paragraph_has_track_changes(paragraph) -> bool:
    return bool(paragraph._p.xpath("./w:ins | ./w:del"))


def paragraph_insertions(paragraph) -> List[TrackedInsertion]:
    return [TrackedInsertion(e, paragraph) for e in paragraph._p.xpath("./w:ins")]  # pyright: ignore[reportArgumentType]


def paragraph_deletions(paragraph) -> List[TrackedDeletion]:
    return [TrackedDeletion(e, paragraph) for e in paragraph._p.xpath("./w:del")]  # pyright: ignore[reportArgumentType]


def paragraph_track_changes(paragraph) -> List[TrackedChange]:
    changes: List[TrackedChange] = []
    for e in paragraph._p.xpath("./w:ins | ./w:del"):
        if e.tag == qn("w:ins"):
            changes.append(TrackedInsertion(e, paragraph))  # pyright: ignore[reportArgumentType]
        elif e.tag == qn("w:del"):
            changes.append(TrackedDeletion(e, paragraph))  # pyright: ignore[reportArgumentType]
    return changes


def paragraph_iter_inner_content(paragraph, include_revisions: bool = False):
    from docx.oxml.text.run import CT_R
    from docx.text.hyperlink import Hyperlink
    from docx.text.run import Run

    if include_revisions:
        elements = paragraph._p.xpath("./w:r | ./w:hyperlink | ./w:ins | ./w:del")
    else:
        elements = paragraph._p.xpath("./w:r | ./w:hyperlink")

    for element in elements:
        if isinstance(element, CT_R):
            yield Run(element, paragraph)
        elif element.tag == qn("w:hyperlink"):
            yield Hyperlink(element, paragraph)  # pyright: ignore[reportArgumentType]
        elif element.tag == qn("w:ins"):
            yield TrackedInsertion(element, paragraph)  # pyright: ignore[reportArgumentType]
        elif element.tag == qn("w:del"):
            yield TrackedDeletion(element, paragraph)  # pyright: ignore[reportArgumentType]


def paragraph_accepted_text(paragraph) -> str:
    return paragraph._p.accepted_text


def paragraph_deleted_text(paragraph) -> str:
    return paragraph._p.deleted_text


def paragraph_add_tracked_insertion(
    paragraph,
    text=None,
    style=None,
    author: str = "",
    revision_id: int | None = None,
):
    if revision_id is None:
        revision_id = next_revision_id(paragraph._p)
    now = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    ins = OxmlElement("w:ins", attrs=revision_attrs(revision_id, author, now))
    r = OxmlElement("w:r")
    ins.append(r)
    paragraph._p.append(ins)
    tracked_insertion = TrackedInsertion(ins, paragraph)  # pyright: ignore[reportArgumentType]
    if text:
        for run in tracked_insertion.runs:
            run.text = text
    if style:
        for run in tracked_insertion.runs:
            run.style = style
    return tracked_insertion


def _new_tracked_insertion(
    paragraph,
    text=None,
    style=None,
    author: str = "",
    revision_id: int | None = None,
):
    """Return a tracked insertion element/proxy pair not yet placed in the paragraph."""
    if revision_id is None:
        revision_id = next_revision_id(paragraph._p)
    now = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    ins = OxmlElement("w:ins", attrs=revision_attrs(revision_id, author, now))
    r = OxmlElement("w:r")
    ins.append(r)
    tracked_insertion = TrackedInsertion(ins, paragraph)  # pyright: ignore[reportArgumentType]
    if text:
        for run in tracked_insertion.runs:
            run.text = text
    if style:
        for run in tracked_insertion.runs:
            run.style = style
    return ins, tracked_insertion


def _insert_element_at_accepted_offset(paragraph, offset: int, element) -> None:
    """Insert `element` at an accepted-text boundary in `paragraph`."""
    accepted_text = paragraph.accepted_text
    if offset < 0 or offset > len(accepted_text):
        raise ValueError(f"Invalid offset: offset={offset} for text of length {len(accepted_text)}")

    spans = _paragraph_accepted_spans(paragraph)
    if not spans:
        paragraph._p.append(element)
        return

    if offset == 0:
        spans[0].element.addprevious(element)
        return

    for idx, span in enumerate(spans):
        if offset == span.start:
            span.element.addprevious(element)
            return

        if span.start < offset < span.end:
            parent = span.element.getparent()
            if parent is None:
                raise ValueError("Paragraph element has no parent")

            split_at = offset - span.start
            before_text = span.text[:split_at]
            after_text = span.text[split_at:]
            index = list(parent).index(span.element)
            parent.remove(span.element)

            insert_idx = index
            if before_text:
                parent.insert(insert_idx, _make_visible_fragment(span, before_text))
                insert_idx += 1

            parent.insert(insert_idx, element)
            insert_idx += 1

            if after_text:
                parent.insert(insert_idx, _make_visible_fragment(span, after_text))
            return

        if offset == span.end:
            next_span = spans[idx + 1] if idx + 1 < len(spans) else None
            if next_span is not None and next_span.start == offset:
                next_span.element.addprevious(element)
            else:
                span.element.addnext(element)
            return

    raise ValueError(f"Invalid offset: offset={offset} for text of length {len(accepted_text)}")


def paragraph_add_tracked_insertion_at(
    paragraph,
    offset: int,
    text=None,
    style=None,
    author: str = "",
    revision_id: int | None = None,
):
    """Insert a tracked insertion at `offset` in accepted/visible paragraph text."""
    ins, tracked_insertion = _new_tracked_insertion(
        paragraph,
        text=text,
        style=style,
        author=author,
        revision_id=revision_id,
    )
    _insert_element_at_accepted_offset(paragraph, offset, ins)
    return tracked_insertion


def _tracked_insertion_search_offset(paragraph, search_text: str, *, after: bool) -> int:
    """Return insertion offset for a unique `search_text` match in accepted text."""
    if not search_text:
        raise ValueError("search text must not be empty")

    accepted_text = paragraph.accepted_text
    positions: list[int] = []
    start = 0
    while True:
        idx = accepted_text.find(search_text, start)
        if idx == -1:
            break
        positions.append(idx)
        start = idx + len(search_text)

    if not positions:
        raise ValueError(f"search text not found: {search_text!r}")
    if len(positions) > 1:
        raise ValueError(f"search text matched multiple occurrences: {search_text!r}")

    return positions[0] + (len(search_text) if after else 0)


def paragraph_add_tracked_insertion_before(
    paragraph,
    search_text: str,
    text=None,
    style=None,
    author: str = "",
    revision_id: int | None = None,
):
    """Insert tracked text before a unique `search_text` match in accepted text."""
    offset = _tracked_insertion_search_offset(paragraph, search_text, after=False)
    return paragraph_add_tracked_insertion_at(
        paragraph,
        offset,
        text=text,
        style=style,
        author=author,
        revision_id=revision_id,
    )


def paragraph_add_tracked_insertion_after(
    paragraph,
    search_text: str,
    text=None,
    style=None,
    author: str = "",
    revision_id: int | None = None,
):
    """Insert tracked text after a unique `search_text` match in accepted text."""
    offset = _tracked_insertion_search_offset(paragraph, search_text, after=True)
    return paragraph_add_tracked_insertion_at(
        paragraph,
        offset,
        text=text,
        style=style,
        author=author,
        revision_id=revision_id,
    )


def _paragraph_accepted_spans(paragraph) -> List[_AcceptedSpan]:
    from docx.text.run import Run

    spans: List[_AcceptedSpan] = []
    offset = 0
    for element in paragraph._p.xpath("./w:r | ./w:hyperlink | ./w:ins | ./w:del"):
        if element.tag == qn("w:r"):
            text = Run(element, paragraph).text
            kind = "run"
        elif element.tag == qn("w:hyperlink"):
            text = element.text
            kind = "hyperlink"
        elif element.tag == qn("w:ins"):
            text = TrackedInsertion(element, paragraph).text  # pyright: ignore[reportArgumentType]
            kind = "ins"
        else:
            continue
        if not text:
            continue
        spans.append(
            _AcceptedSpan(
                kind=kind,
                element=element,
                text=text,
                start=offset,
                end=offset + len(text),
            )
        )
        offset += len(text)
    return spans


def _make_visible_fragment(span: _AcceptedSpan, text: str):
    if span.kind == "run":
        run = deepcopy(span.element)
        run.text = text
        return run
    if span.kind == "hyperlink":
        hyperlink = deepcopy(span.element)
        for child in list(hyperlink):
            hyperlink.remove(child)
        hyperlink.append(make_text_run(text))
        return hyperlink
    ins = deepcopy(span.element)
    for child in list(ins):
        ins.remove(child)
    ins.append(make_text_run(text))
    return ins


def paragraph_comment_range_runs(paragraph, start: int, end: int):
    """Return the first and last runs spanning accepted-text offsets in `paragraph`.

    The paragraph is rewritten as needed so the range boundaries align with run
    boundaries. First-pass support is intentionally limited to plain paragraph runs.
    """
    from docx.text.run import Run

    accepted_text = paragraph.accepted_text
    if start < 0 or end > len(accepted_text) or start >= end:
        raise ValueError(
            f"Invalid offsets: start={start}, end={end} for text of length {len(accepted_text)}"
        )

    spans = _paragraph_accepted_spans(paragraph)
    if not spans:
        raise ValueError("Paragraph has no accepted-view text")

    affected: list[tuple[_AcceptedSpan, int, int]] = []
    for span in spans:
        overlap_start = max(start, span.start)
        overlap_end = min(end, span.end)
        if overlap_start >= overlap_end:
            continue
        local_start = overlap_start - span.start
        local_end = overlap_end - span.start
        affected.append((span, local_start, local_end))

    if not affected:
        raise ValueError(
            f"Invalid offsets: start={start}, end={end} for text of length {len(accepted_text)}"
        )

    unsupported = {span.kind for span, _, _ in affected if span.kind != "run"}
    if unsupported:
        raise ValueError(
            "comment ranges currently support only plain paragraph runs; "
            f"encountered: {', '.join(sorted(unsupported))}"
        )

    first_span, first_local_start, _ = affected[0]
    last_span, _, last_local_end = affected[-1]
    parent = first_span.element.getparent()
    if parent is None:
        raise ValueError("Paragraph element has no parent")

    index = list(parent).index(first_span.element)
    before_text = first_span.text[:first_local_start]
    after_text = last_span.text[last_local_end:]

    removed_elements: set[int] = set()
    for span, _, _ in affected:
        if id(span.element) in removed_elements:
            continue
        if span.element.getparent() is parent:
            parent.remove(span.element)
            removed_elements.add(id(span.element))

    insert_idx = index
    if before_text:
        parent.insert(insert_idx, _make_visible_fragment(first_span, before_text))
        insert_idx += 1

    selected_runs = []
    for span, local_start, local_end in affected:
        selected_text = span.text[local_start:local_end]
        if not selected_text:
            continue
        fragment = _make_visible_fragment(span, selected_text)
        parent.insert(insert_idx, fragment)
        selected_runs.append(fragment)
        insert_idx += 1

    if after_text:
        parent.insert(insert_idx, _make_visible_fragment(last_span, after_text))

    if not selected_runs:
        raise ValueError("Comment range resolved to no runs")

    return (
        Run(selected_runs[0], paragraph),
        Run(selected_runs[-1], paragraph),
    )


def _edit_accepted_range(
    paragraph,
    start: int,
    end: int,
    *,
    author: str,
    replace_text: str | None,
    revision_id: int | None = None,
) -> TrackedDeletion | None:
    accepted_text = paragraph.accepted_text
    if start < 0 or end > len(accepted_text) or start >= end:
        raise ValueError(
            f"Invalid offsets: start={start}, end={end} for text of length {len(accepted_text)}"
        )

    spans = _paragraph_accepted_spans(paragraph)
    if not spans:
        raise ValueError("Paragraph has no accepted-view text")

    affected: list[tuple[_AcceptedSpan, int, int]] = []
    deleted_text_parts: list[str] = []

    for span in spans:
        overlap_start = max(start, span.start)
        overlap_end = min(end, span.end)
        if overlap_start >= overlap_end:
            continue
        local_start = overlap_start - span.start
        local_end = overlap_end - span.start
        affected.append((span, local_start, local_end))
        if span.kind in {"run", "hyperlink"}:
            deleted_text_parts.append(span.text[local_start:local_end])

    if not affected:
        raise ValueError(
            f"Invalid offsets: start={start}, end={end} for text of length {len(accepted_text)}"
        )

    first_span, first_local_start, _ = affected[0]
    last_span, _, last_local_end = affected[-1]
    parent = first_span.element.getparent()
    if parent is None:
        raise ValueError("Paragraph element has no parent")

    index = list(parent).index(first_span.element)
    before_text = first_span.text[:first_local_start]
    after_text = last_span.text[last_local_end:]
    deleted_text = "".join(deleted_text_parts)
    now = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    removed_elements: set[int] = set()
    for span, _, _ in affected:
        if id(span.element) in removed_elements:
            continue
        if span.element.getparent() is parent:
            parent.remove(span.element)
            removed_elements.add(id(span.element))

    insert_idx = index
    if before_text:
        parent.insert(insert_idx, _make_visible_fragment(first_span, before_text))
        insert_idx += 1

    tracked_deletion: TrackedDeletion | None = None
    if deleted_text:
        deletion_id = revision_id if revision_id is not None else next_revision_id(paragraph._p)
        del_elem = make_del_element(deleted_text, author, deletion_id, now)
        parent.insert(insert_idx, del_elem)
        insert_idx += 1
        tracked_deletion = TrackedDeletion(del_elem, paragraph)  # pyright: ignore[reportArgumentType]

    if replace_text:
        ins_elem = make_ins_element(replace_text, author, next_revision_id(paragraph._p), now)
        parent.insert(insert_idx, ins_elem)
        insert_idx += 1

    if after_text:
        parent.insert(insert_idx, _make_visible_fragment(last_span, after_text))

    return tracked_deletion


def paragraph_add_tracked_deletion(
    paragraph,
    start: int,
    end: int,
    author: str = "",
    revision_id: int | None = None,
) -> TrackedDeletion | None:
    return _edit_accepted_range(
        paragraph,
        start,
        end,
        author=author,
        replace_text=None,
        revision_id=revision_id,
    )


def paragraph_replace_tracked_at(
    paragraph, start: int, end: int, replace_text: str, author: str = ""
) -> None:
    _edit_accepted_range(paragraph, start, end, author=author, replace_text=replace_text)


def paragraph_replace_tracked(
    paragraph, search_text: str, replace_text: str, author: str = ""
) -> int:
    count = 0
    full_text = paragraph.accepted_text
    search_len = len(search_text)
    positions: list[int] = []
    start = 0
    while True:
        idx = full_text.find(search_text, start)
        if idx == -1:
            break
        positions.append(idx)
        start = idx + search_len
    for pos in reversed(positions):
        paragraph_replace_tracked_at(paragraph, pos, pos + search_len, replace_text, author=author)
        count += 1
    return count


def run_delete_tracked(run, author: str = "", revision_id: int | None = None) -> TrackedDeletion:
    if revision_id is None:
        revision_id = next_revision_id(run._r)
    parent = run._r.getparent()
    if parent is None:
        raise ValueError("Run has no parent element")
    del_elem = OxmlElement(
        "w:del",
        attrs=revision_attrs(
            revision_id,
            author,
            dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        ),
    )
    for t_elem in run._r.findall(qn("w:t")):
        del_text = OxmlElement("w:delText")
        del_text.text = t_elem.text
        if t_elem.get(qn("xml:space")) == "preserve":
            del_text.set(qn("xml:space"), "preserve")
        t_elem.getparent().replace(t_elem, del_text)  # pyright: ignore[reportOptionalMemberAccess]
    index = list(parent).index(run._r)
    parent.insert(index, del_elem)
    del_elem.append(run._r)
    return TrackedDeletion(del_elem, run._parent)  # pyright: ignore[reportArgumentType]


def run_replace_tracked_at(run, start: int, end: int, replace_text: str, author: str = "") -> None:
    text = run.text
    if start < 0 or end > len(text) or start >= end:
        raise ValueError(
            f"Invalid offsets: start={start}, end={end} for text of length {len(text)}"
        )
    before_text = text[:start] or None
    deleted_text = text[start:end]
    after_text = text[end:] or None
    now = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    parent = run._r.getparent()
    if parent is None:
        raise ValueError("Run has no parent element")
    index = list(parent).index(run._r)
    parent.remove(run._r)
    insert_idx = index
    if before_text:
        parent.insert(insert_idx, make_text_run(before_text))
        insert_idx += 1
    parent.insert(insert_idx, make_del_element(deleted_text, author, next_revision_id(run._r), now))
    insert_idx += 1
    parent.insert(insert_idx, make_ins_element(replace_text, author, next_revision_id(run._r), now))
    insert_idx += 1
    if after_text:
        parent.insert(insert_idx, make_text_run(after_text))


__all__ = [
    "TrackedChange",
    "TrackedDeletion",
    "TrackedInsertion",
    "next_revision_id",
    "paragraph_accepted_text",
    "paragraph_add_tracked_deletion",
    "paragraph_add_tracked_insertion",
    "paragraph_add_tracked_insertion_after",
    "paragraph_add_tracked_insertion_at",
    "paragraph_add_tracked_insertion_before",
    "paragraph_comment_range_runs",
    "paragraph_deleted_text",
    "paragraph_deletions",
    "paragraph_has_track_changes",
    "paragraph_insertions",
    "paragraph_iter_inner_content",
    "paragraph_replace_tracked",
    "paragraph_replace_tracked_at",
    "paragraph_track_changes",
    "run_delete_tracked",
    "run_replace_tracked_at",
]
