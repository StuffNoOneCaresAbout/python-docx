# python-docx

_python-docx_ is a Python library for reading, creating, and updating Microsoft Word 2007+ (.docx) files.

## Installation

```
pip install python-docx
```

## Development

This project is managed with [uv](https://github.com/astral-sh/uv).
Supported Python versions are 3.13+.

```
uv sync
uv run pytest -x
uv run --group test behave --stop
uv run --group docs sphinx-build -b html docs docs/.build/html
```

## Example

```python
>>> from docx import Document

>>> document = Document()
>>> document.add_paragraph("It was a dark and stormy night.")
<docx.text.paragraph.Paragraph object at 0x10f19e760>
>>> document.save("dark-and-stormy.docx")

>>> document = Document("dark-and-stormy.docx")
>>> document.paragraphs[0].text
'It was a dark and stormy night.'
```

More information is available in the [python-docx documentation](https://python-docx.readthedocs.org/en/latest/)
