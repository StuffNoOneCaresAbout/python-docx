.. _comments:

Working with Comments
=====================

Word allows *comments* to be added to a document. This is an aspect of the *reviewing*
feature-set and is typically used by a second party to provide feedback to the author
without changing the document itself.

The procedure is simple:

- You select some range of text with the mouse or Shift+Arrow keys
- You press the *New Comment* button (Review toolbar)
- You type or paste in your comment

.. image:: /_static/img/comment-parts.png

A comment can only be added to the main document. A comment cannot be added in a header,
a footer, or within a comment. A comment _can_ be added to a footnote or endnote, but
those are not yet supported by *python-docx*.

**Comment Anatomy.** Each comment has two parts, the *comment-reference* and the
*comment-content*:

The **comment-refererence**, sometimes *comment-anchor*, is the text in the main
document you selected before pressing the *New Comment* button. It is a so-called
*range* in the main document that starts at the first selected character and ends after
the last one.

The **comment-content**, sometimes just *comment*, is whatever content you typed or
pasted in. The content for each comment is stored in a separate comment object, and
these comment objects are stored in a separate *comments-part* (part-name
``word/comments.xml``), not in the main document. Each comment is assigned a unique id
when it is created, allowing the comment reference to be associated with its content and
vice versa.

**Comment Reference.** The comment-reference is a *range*. A range must both start and
end at an even *run* boundary. Intuitively, a range corresponds to a *selection* of text
in the Word UI, one formed by dragging with the mouse or using the *Shift-Arrow* keys.

In the XML, this range is delimited by a start marker `<w:commentRangeStart/>` and an
end marker `<w:commentRangeEnd/>`, both of which contain the *id* of the comment they
delimit. The start marker appears before the run starting with the first character of
the range and the end marker appears immediately after the run ending with the last
character of the range. Adding a comment that references an arbitrary range of text in
an existing document may require splitting runs on the desired character boundaries.

In general a range can span paragraphs, such that the range begins in one paragraph and
ends in a later paragraph. However, a range must enclose *contiguous* runs, such that a
range that contains only two vertically adjacent cells in a multi-column table is not
possible (even though Word allows such a selection with the mouse).

**Comment Content.** Interestingly, although commonly used to contain a single line of
plain text, the comment-content can contain essentially any content that can appear in
the document body. This includes rich text with emphasis, runs with a different typeface
and size, both paragraph and character styles, hyperlinks, images, and tables. Note that
tables do not appear in the comment as displayed in the *comment-sidebar* although they
do apper in the *reviewing-pane*.

**Comment Metadata.** Each comment can be assigned *author*, *initals*, and *date*
metadata. In Word, these fields are assigned automatically based on values in ``Settings
> User`` of the installed Word application. These might be configured automatically in
an enterprise installation, based on the user account, but by default they are empty.

*author* metadata is required, although silently assigned the empty string by Word if
the user name is not configured. *initials* is optional, but always set by Word, to the
empty string if not configured. *date* is also optional, but always set by Word to the
UTC date and time the comment was added, with seconds resolution (no milliseconds or
microseconds).

**Additional Features.** Later versions of Word allow a top-level comment to be
*resolved*. A comment in this state will appear grayed-out in the Word UI. Later
versions of Word also allow a comment to be *replied to*, forming a *comment thread*.
*python-docx* supports both replies and resolved-state metadata, but only for the
top-level comment in a thread.

**Applicability.** Note that comments cannot be added to a header or footer and cannot
be nested inside a comment itself. In general the *python-docx* API will not allow these
operations but if you outsmart it then the resulting comment will either be silently
removed or trigger a repair error when the document is loaded by Word.


Adding a Comment
----------------

A simple example is adding a comment to a paragraph::

    >>> from docx import Document
    >>> document = Document()
    >>> paragraph = document.add_paragraph("Hello, world!")

    >>> comment = document.add_comment(
    ...    runs=paragraph.runs,
    ...    text="I have this to say about that"
    ...    author="Steve Canny",
    ...    initials="SC",
    ... )
    >>> comment
    <docx.comments.Comment object at 0x02468ACE>
    >>> comment.id
    0
    >>> comment.author
    'Steve Canny'
    >>> comment.initials
    'SC'
    >>> comment.date
    datetime.datetime(2025, 6, 11, 20, 42, 30, 0, tzinfo=datetime.timezone.utc)
    >>> comment.text
    'I have this to say about that'

The API documentation for :meth:`.Document.add_comment` provides further details.

For convenience, a comment can also be added directly from a paragraph, run, or
table cell::

    >>> paragraph = document.add_paragraph("Hello, world!")
    >>> comment = paragraph.add_comment("Please reword this.", author="Steve Canny")
    >>> run = paragraph.runs[0]
    >>> comment = run.add_comment("Comment on just this run.", author="Steve Canny")
    >>> table = document.add_table(rows=1, cols=1)
    >>> cell = table.cell(0, 0)
    >>> cell.text = "Cell text"
    >>> comment = cell.add_comment("Comment on this cell.", author="Steve Canny")

Tracked changes also provide a convenience API. A comment can be anchored directly to
an insertion or deletion::

    >>> paragraph = document.add_paragraph("Hello")
    >>> insertion = paragraph.add_tracked_insertion(" world", author="Editor")
    >>> comment = insertion.add_comment("Please justify this insertion.", author="Reviewer")

When added from a cell, the comment is anchored from the first run in the first
paragraph of the cell to the last run in the last paragraph of the cell. This matches
Word's XML model, where a so-called "cell comment" is really a comment range anchored
inside the cell's paragraph content rather than on the cell element itself.

For run-level tracked changes, Word stores the comment as an ordinary comment range
that brackets the ``<w:ins>`` or ``<w:del>`` wrapper itself. For block-level tracked
changes such as inserted or deleted paragraphs and tables, the comment is anchored to
the first and last paragraph content inside the tracked block because comment markers
still have to live on paragraph/run boundaries.

When you need to anchor a comment to only part of a paragraph's text, use
``Paragraph.add_comment_range(start, end, ...)`` with offsets measured against
``paragraph.accepted_text``::

    >>> paragraph = document.add_paragraph("South")
    >>> comment = paragraph.add_comment_range(1, 3, "Comment on just 'ou'.")

The method will split runs as needed so the comment range lands on proper run
boundaries. In this first pass, range comments are limited to plain paragraph runs;
selections that include deleted text, hyperlinks, or other non-run content raise
``ValueError`` rather than guessing.


Accessing and using the Comments collection
-------------------------------------------

The comments collection is accessed via the :attr:`.Document.comments` property::

    >>> comments = document.comments
    >>> comments
    <docx.parts.comments.Comments object at 0x02468ACE>
    >>> len(comments)
    1

The comments collection supports random access to a comment by its id::

    >>> comment = comments.get(0)
    >>> comment
    <docx.comments.Comment object at 0x02468ACE>


Adding rich content to a comment
--------------------------------

A comment is a *block-item container*, just like the document body or a table cell, so
it can contain any content that can appear in those places. It does not contain
page-layout sections and cannot contain a comment reference, but it can contain multiple
paragraphs and/or tables, and runs within paragraphs can have emphasis such as bold or
italic, and have images or hyperlinks.

A comment created with `text=""` will contain a single paragraph with a single empty run
containing the so-called *annotation reference* but no text. It's probably best to leave
this run as it is but you can freely add additional runs to the paragraph that contain
whatever content you like.

The methods for adding this content are the same as those used for the document and
table cells::

    >>> paragraph = document.add_paragraph("The rain in Spain.")
    >>> comment = document.add_comment(
    ...     runs=paragraph.runs,
    ...     text="",
    ... )
    >>> cmt_para = comment.paragraphs[0]
    >>> cmt_para.add_run("Please finish this thought. I believe it should be ")
    >>> cmt_para.add_run("falls mainly in the plain.").bold = True


Updating comment metadata
-------------------------

The author and initials metadata can be updated as desired::

    >>> comment.author = "John Smith"
    >>> comment.initials = "JS"
    >>> comment.author
    'John Smith'
    >>> comment.initials
    'JS'


Resolving and reopening comments
--------------------------------

A top-level comment can be marked resolved or reopened using either the ``resolved``
property or the convenience methods ``resolve()`` and ``reopen()``::

    >>> comment.resolved
    False
    >>> comment.resolve()
    >>> comment.resolved
    True
    >>> comment.resolved_at is not None
    True
    >>> comment.reopen()
    >>> comment.resolved
    False

The ``resolved_at`` value records the UTC timestamp associated with the resolved-state
metadata when that information is available in the document.

Reply comments do not support independent resolved-state operations. This matches Word's
review UI, which treats resolution as a property of the thread root rather than each
individual reply.
