from __future__ import annotations

from io import BytesIO

import pytest

from docx import Document
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.revisions import TrackedDeletion, TrackedInsertion


def _revision_attrs(revision_id: int = 1, author: str = "TestAuthor") -> dict[str, str]:
    return {
        qn("w:id"): str(revision_id),
        qn("w:author"): author,
        qn("w:date"): "2026-03-17T00:00:00Z",
    }


def _insert_body_child(document: Document, element) -> None:
    body = document._element.body
    insert_idx = len(body)
    if len(body) and body[-1].tag == qn("w:sectPr"):
        insert_idx -= 1
    body.insert(insert_idx, element)


def _new_block_change(change_tag: str, text: str) -> object:
    change = OxmlElement(change_tag, attrs=_revision_attrs())
    paragraph = OxmlElement("w:p")
    run = OxmlElement("w:r")
    text_elm = OxmlElement("w:delText" if change_tag == "w:del" else "w:t")
    text_elm.text = text
    run.append(text_elm)
    paragraph.append(run)
    change.append(paragraph)
    return change


class DescribeRevisions:
    def it_excludes_inserted_text_from_paragraph_text_but_includes_it_in_accepted_text(self):
        document = Document()
        paragraph = document.add_paragraph("Alpha")

        insertion = paragraph.add_tracked_insertion("Beta", author="TestAuthor")

        assert isinstance(insertion, TrackedInsertion)
        assert paragraph.text == "Alpha"
        assert paragraph.accepted_text == "AlphaBeta"
        assert paragraph.deleted_text == ""

    def it_includes_deleted_text_in_paragraph_text_but_not_in_accepted_text(self):
        document = Document()
        paragraph = document.add_paragraph("Alpha")

        deletion = paragraph.add_tracked_deletion(1, 4, author="TestAuthor")

        assert isinstance(deletion, TrackedDeletion)
        assert paragraph.text == "Alpha"
        assert paragraph.accepted_text == "Aa"
        assert paragraph.deleted_text == "lph"

    def it_can_iterate_revisions_when_requested(self):
        document = Document()
        paragraph = document.add_paragraph("Alpha")
        paragraph.add_tracked_insertion("Beta", author="TestAuthor")
        paragraph.add_tracked_deletion(1, 3, author="TestAuthor")

        items = list(paragraph.iter_inner_content(include_revisions=True))

        assert any(isinstance(item, TrackedInsertion) for item in items)
        assert any(isinstance(item, TrackedDeletion) for item in items)

    def it_can_accept_and_reject_tracked_changes(self):
        document = Document()
        paragraph = document.add_paragraph("Alpha")
        insertion = paragraph.add_tracked_insertion("Beta", author="TestAuthor")
        deletion = paragraph.add_tracked_deletion(1, 4, author="TestAuthor")

        assert deletion is not None
        deletion.reject()
        insertion.reject()

        assert paragraph.text == "Alpha"
        assert paragraph.accepted_text == "Alpha"
        assert paragraph.track_changes == []

    def it_can_find_and_replace_with_tracking(self):
        document = Document()
        paragraph = document.add_paragraph("Alpha")

        count = document.find_and_replace_tracked("ph", "XY", author="TestAuthor")

        assert count == 1
        assert paragraph.text == "Alpha"
        assert paragraph.accepted_text == "AlXYa"
        assert len(document.track_changes) == 2

    def it_can_reject_all_tracked_changes_in_the_document(self):
        document = Document()
        first = document.add_paragraph("Alpha")
        second = document.add_paragraph("Gamma")
        first.add_tracked_insertion("Beta", author="TestAuthor")
        second.add_tracked_deletion(1, 3, author="TestAuthor")

        document.reject_all()

        assert document.track_changes == []
        assert first.text == "Alpha"
        assert first.accepted_text == "Alpha"
        assert second.text == "Gamma"
        assert second.accepted_text == "Gamma"

    def it_ignores_comment_markers_when_computing_paragraph_text_views(self):
        document = Document()
        paragraph = document.add_paragraph("Alpha")

        paragraph.add_comment("Comment", author="TestAuthor")
        paragraph.add_tracked_insertion("Beta", author="TestAuthor")
        paragraph.add_tracked_deletion(1, 4, author="TestAuthor")

        assert paragraph.text == "Alpha"
        assert paragraph.accepted_text == "AaBeta"
        assert paragraph.deleted_text == "lph"

    def it_preserves_revision_text_semantics_on_save_and_reload(self):
        document = Document()
        paragraph = document.add_paragraph("Alpha")
        paragraph.add_tracked_insertion("Beta", author="TestAuthor")
        paragraph.add_tracked_deletion(1, 4, author="TestAuthor")

        stream = BytesIO()
        document.save(stream)
        stream.seek(0)

        reloaded = Document(stream)
        paragraph = reloaded.paragraphs[0]

        assert paragraph.text == "Alpha"
        assert paragraph.accepted_text == "AaBeta"
        assert paragraph.deleted_text == "lph"

    def it_can_add_a_comment_to_a_run_level_tracked_insertion(self):
        document = Document()
        paragraph = document.add_paragraph("Alpha")
        insertion = paragraph.add_tracked_insertion("Beta", author="Editor")

        comment = insertion.add_comment("Review insertion", author="Reviewer")

        children = list(paragraph._p)
        assert comment.text == "Review insertion"
        assert [child.tag for child in children] == [
            qn("w:r"),
            qn("w:commentRangeStart"),
            qn("w:ins"),
            qn("w:commentRangeEnd"),
            qn("w:r"),
        ]
        assert children[1].get(qn("w:id")) == str(comment.comment_id)
        assert children[3].get(qn("w:id")) == str(comment.comment_id)
        assert children[4].xpath("./w:commentReference/@w:id") == [str(comment.comment_id)]

    def it_can_add_a_comment_to_a_run_level_tracked_deletion(self):
        document = Document()
        paragraph = document.add_paragraph("Alpha")
        deletion = paragraph.add_tracked_deletion(1, 4, author="Editor")

        comment = deletion.add_comment("Review deletion", author="Reviewer")

        children = list(paragraph._p)
        assert comment.text == "Review deletion"
        assert [child.tag for child in children] == [
            qn("w:r"),
            qn("w:commentRangeStart"),
            qn("w:del"),
            qn("w:commentRangeEnd"),
            qn("w:r"),
            qn("w:r"),
        ]
        assert children[1].get(qn("w:id")) == str(comment.comment_id)
        assert children[3].get(qn("w:id")) == str(comment.comment_id)
        assert children[4].xpath("./w:commentReference/@w:id") == [str(comment.comment_id)]

    def it_can_add_multiple_comments_to_the_same_tracked_change_span(self):
        document = Document()
        paragraph = document.add_paragraph("Alpha")
        insertion = paragraph.add_tracked_insertion("Beta", author="Editor")

        first = insertion.add_comment("First", author="Reviewer")
        second = insertion.add_comment("Second", author="Reviewer")

        children = list(paragraph._p)
        assert [child.tag for child in children] == [
            qn("w:r"),
            qn("w:commentRangeStart"),
            qn("w:commentRangeStart"),
            qn("w:ins"),
            qn("w:commentRangeEnd"),
            qn("w:r"),
            qn("w:commentRangeEnd"),
            qn("w:r"),
        ]
        assert children[1].get(qn("w:id")) == str(first.comment_id)
        assert children[2].get(qn("w:id")) == str(second.comment_id)
        assert children[4].get(qn("w:id")) == str(first.comment_id)
        assert children[6].get(qn("w:id")) == str(second.comment_id)
        assert children[5].xpath("./w:commentReference/@w:id") == [str(first.comment_id)]
        assert children[7].xpath("./w:commentReference/@w:id") == [str(second.comment_id)]

    def it_can_add_a_comment_to_a_block_level_tracked_paragraph(self):
        document = Document()
        insertion_elm = _new_block_change("w:ins", "Inserted paragraph")
        _insert_body_child(document, insertion_elm)
        insertion = TrackedInsertion(insertion_elm, document)

        comment = insertion.add_comment("Review block insertion", author="Reviewer")

        paragraph = insertion_elm.xpath("./w:p")[0]
        children = list(paragraph)
        assert comment.text == "Review block insertion"
        assert [child.tag for child in children] == [
            qn("w:commentRangeStart"),
            qn("w:r"),
            qn("w:commentRangeEnd"),
            qn("w:r"),
        ]
        assert insertion_elm.getparent()[0].tag == qn("w:ins")

    def it_can_add_a_comment_to_a_block_level_tracked_table(self):
        document = Document()
        table = document.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "Cell text"
        tbl = table._tbl
        parent = tbl.getparent()
        index = list(parent).index(tbl)
        parent.remove(tbl)
        deletion_elm = OxmlElement("w:del", attrs=_revision_attrs())
        deletion_elm.append(tbl)
        parent.insert(index, deletion_elm)
        deletion = TrackedDeletion(deletion_elm, document)

        comment = deletion.add_comment("Review table deletion", author="Reviewer")

        first_paragraph = deletion_elm.xpath(".//w:tc[1]/w:p[1]")[0]
        children = list(first_paragraph)
        assert comment.text == "Review table deletion"
        assert children[0].tag == qn("w:commentRangeStart")
        assert children[1].tag == qn("w:r")
        assert deletion_elm.xpath("./w:commentRangeStart") == []
        assert deletion_elm.xpath("./w:commentRangeEnd") == []

    def it_preserves_tracked_change_comments_on_save_and_reload(self):
        document = Document()
        paragraph = document.add_paragraph("Alpha")
        insertion = paragraph.add_tracked_insertion("Beta", author="Editor")
        comment = insertion.add_comment("Review insertion", author="Reviewer")

        stream = BytesIO()
        document.save(stream)
        stream.seek(0)

        reloaded = Document(stream)
        paragraph = reloaded.paragraphs[0]
        reloaded_comment = next(iter(reloaded.comments))

        assert reloaded_comment.comment_id == comment.comment_id
        assert reloaded_comment.text == "Review insertion"
        assert paragraph._p.xpath("./w:commentRangeStart/@w:id") == [str(comment.comment_id)]
        assert paragraph._p.xpath("./w:commentRangeEnd/@w:id") == [str(comment.comment_id)]

    def it_rejects_tracked_change_comments_outside_the_main_document_story(self):
        document = Document()
        root_comment = document.add_paragraph("Alpha").add_comment("Root", author="Reviewer")
        insertion = root_comment.paragraphs[0].add_tracked_insertion("Beta", author="Editor")

        with pytest.raises(ValueError, match="main document story"):
            insertion.add_comment("Nested nope", author="Reviewer")

    def it_rejects_tracked_changes_with_no_anchorable_content(self):
        document = Document()
        paragraph = document.add_paragraph("Alpha")
        empty_insertion = OxmlElement("w:ins", attrs=_revision_attrs())
        paragraph._p.append(empty_insertion)
        insertion = TrackedInsertion(empty_insertion, paragraph)

        with pytest.raises(ValueError, match="no anchorable content"):
            insertion.add_comment("Nope", author="Reviewer")
