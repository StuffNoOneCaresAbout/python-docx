Working with Revisions
======================

Word's *track changes* feature stores inserted and deleted content in revision wrappers
such as ``<w:ins>`` and ``<w:del>``. *python-docx* exposes those content revisions so
they can be read, created, accepted, rejected, and used in tracked find-and-replace
operations.


Reading paragraph text with revisions
-------------------------------------

Paragraph text now has two primary views:

- ``Paragraph.text`` returns the paragraph's original reading view: normal text plus
  deleted text, excluding inserted text.
- ``Paragraph.accepted_text`` returns the visible/accepted view: normal text plus
  inserted text, excluding deleted text.

Deleted-only content is also available directly via ``Paragraph.deleted_text``::

    >>> from docx import Document
    >>> document = Document()
    >>> paragraph = document.add_paragraph("Alpha")
    >>> paragraph.add_tracked_insertion("Beta", author="Editor")
    >>> paragraph.add_tracked_deletion(1, 4, author="Editor")

    >>> paragraph.text
    'Alpha'
    >>> paragraph.accepted_text
    'AaBeta'
    >>> paragraph.deleted_text
    'lph'


Inspecting tracked changes
--------------------------

Tracked insertions and deletions can be accessed from a paragraph using:

- ``paragraph.has_track_changes``
- ``paragraph.insertions``
- ``paragraph.deletions``
- ``paragraph.track_changes``

Each tracked change provides metadata such as author, date, revision id, and text::

    >>> change = paragraph.track_changes[0]
    >>> change.author
    'Editor'
    >>> change.text
    'Beta'


Adding tracked changes
----------------------

Tracked insertions and deletions can be created directly on ``Paragraph`` and ``Run``
objects::

    >>> paragraph = document.add_paragraph("Alpha")
    >>> paragraph.add_tracked_insertion("Beta", author="Editor")
    >>> paragraph.add_tracked_deletion(1, 4, author="Editor")

Tracked insertion can also target a visible-text position or a unique accepted-text
match::

    >>> paragraph = document.add_paragraph("Hello World")
    >>> paragraph.add_tracked_insertion_at(6, "Brave ", author="Editor")
    >>> paragraph.add_tracked_insertion_before("World", "Really ", author="Editor")
    >>> paragraph.add_tracked_insertion_after("Hello", ", there", author="Editor")

These insertion helpers use ``accepted_text`` coordinates and search semantics, so
existing insertions are counted as visible text and deleted text is ignored. The
search-based methods require exactly one match and raise ``ValueError`` when the text
is missing or ambiguous.

Run-level deletion and replacement are also available::

    >>> run = paragraph.runs[0]
    >>> run.delete_tracked(author="Editor")


Accepting and rejecting changes
-------------------------------

Individual changes can be accepted or rejected using the tracked-change proxy object::

    >>> change = paragraph.track_changes[0]
    >>> change.accept()

Document-wide operations are also available::

    >>> document.accept_all()
    >>> document.reject_all()


Tracked find-and-replace
------------------------

Tracked replacement operates on the accepted/visible text view so that inserted text is
searchable and deleted text is ignored::

    >>> document.find_and_replace_tracked("Acme", "NewCo", author="Editor")


Interaction with comments
-------------------------

Comment markers do not contribute text to either ``Paragraph.text`` or
``Paragraph.accepted_text``. A paragraph can contain both comments and tracked changes;
text extraction remains based on document text and revision wrappers, not on comment
anchor markers.

Tracked changes can also be commented directly using the tracked-change proxy::

    >>> paragraph = document.add_paragraph("Hello")
    >>> insertion = paragraph.add_tracked_insertion(" world", author="Editor")
    >>> comment = insertion.add_comment("Why was this added?", author="Reviewer")

For run-level revisions, the comment range brackets the ``<w:ins>`` or ``<w:del>``
element externally, matching Word-authored documents. For block-level revisions, such
as inserted paragraphs or deleted tables, the comment is anchored to the first and last
paragraph content inside the tracked block so the markers still land on valid run
boundaries.
