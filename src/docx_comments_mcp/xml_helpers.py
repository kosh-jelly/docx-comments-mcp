"""Low-level XML parsing utilities for OOXML documents."""

from __future__ import annotations

import re
from typing import Iterator
from lxml import etree

# OOXML namespaces
NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

# Inverse mapping for creating elements
NS_MAP = {v: k for k, v in NAMESPACES.items()}


_TYPOGRAPHY_TABLE = str.maketrans({
    "\u2018": "'",   # left single quote
    "\u2019": "'",   # right single quote / apostrophe
    "\u201A": "'",   # single low-9 quote
    "\u201C": '"',   # left double quote
    "\u201D": '"',   # right double quote
    "\u201E": '"',   # double low-9 quote
    "\u2013": "-",   # en-dash
    "\u2014": "-",   # em-dash
    "\u2011": "-",   # non-breaking hyphen
    "\u00A0": " ",   # non-breaking space
})


def normalize_typography(text: str) -> str:
    """Replace smart/fancy Unicode characters with their plain ASCII equivalents.

    All replacements are 1:1 character mappings, so string length is preserved.
    """
    return text.translate(_TYPOGRAPHY_TABLE)


def qn(tag: str) -> str:
    """Convert a prefixed tag name to Clark notation.

    Example: qn('w:p') -> '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'
    """
    if ":" not in tag:
        return tag
    prefix, local = tag.split(":", 1)
    if prefix not in NAMESPACES:
        raise ValueError(f"Unknown namespace prefix: {prefix}")
    return f"{{{NAMESPACES[prefix]}}}{local}"


def local_name(tag: str) -> str:
    """Extract local name from a Clark notation tag.

    Example: '{http://...}p' -> 'p'
    """
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag


def get_text_content(element: etree._Element) -> str:
    """Extract all text content from an element and its descendants."""
    texts = []
    for text_elem in element.iter(qn("w:t"), qn("w:delText")):
        if text_elem.text:
            texts.append(text_elem.text)
    return "".join(texts)


def find_text_in_paragraph(paragraph: etree._Element, search_text: str) -> list[tuple[etree._Element, int, int]]:
    """Find occurrences of text within a paragraph's runs.

    Returns list of (run_element, start_offset, end_offset) tuples.
    This handles text split across multiple runs.
    """
    # Build a map of character positions to runs
    runs = list(paragraph.iter(qn("w:r")))
    char_map: list[tuple[etree._Element, etree._Element, int]] = []  # (run, text_elem, char_index_in_text)

    for run in runs:
        for text_elem in run.iter(qn("w:t")):
            if text_elem.text:
                for i, _ in enumerate(text_elem.text):
                    char_map.append((run, text_elem, i))

    # Build full text
    full_text = "".join(text_elem.text[idx] for _, text_elem, idx in char_map) if char_map else ""

    # Find all occurrences (normalize both sides for smart-quote tolerance)
    norm_full = normalize_typography(full_text)
    norm_search = normalize_typography(search_text)
    matches = []
    start = 0
    while True:
        idx = norm_full.find(norm_search, start)
        if idx == -1:
            break
        matches.append((idx, idx + len(search_text)))
        start = idx + 1

    return [(char_map, match) for match in matches] if char_map else []


def iter_paragraphs(document: etree._Element) -> Iterator[tuple[int, etree._Element]]:
    """Iterate over paragraphs in document body, yielding (index, element)."""
    body = document.find(qn("w:body"))
    if body is None:
        return

    idx = 0
    for elem in body.iter(qn("w:p")):
        yield idx, elem
        idx += 1


def get_paragraph_style(paragraph: etree._Element) -> str | None:
    """Get the style name of a paragraph."""
    pPr = paragraph.find(qn("w:pPr"))
    if pPr is not None:
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is not None:
            return pStyle.get(qn("w:val"))
    return None


def parse_datetime(date_str: str | None) -> str | None:
    """Parse and normalize an OOXML datetime string."""
    if not date_str:
        return None
    # OOXML uses ISO 8601 format, just return as-is
    return date_str


def create_element(tag: str, attribs: dict[str, str] | None = None, nsmap: dict[str, str] | None = None) -> etree._Element:
    """Create an element with the given tag and attributes.

    Tag should be in Clark notation or prefixed form (e.g., 'w:comment').
    """
    if ":" in tag and not tag.startswith("{"):
        tag = qn(tag)

    # Build nsmap with only needed namespaces
    if nsmap is None:
        nsmap = {}

    elem = etree.Element(tag, nsmap=nsmap)

    if attribs:
        for key, value in attribs.items():
            if ":" in key and not key.startswith("{"):
                key = qn(key)
            elem.set(key, value)

    return elem


def get_max_id(root: etree._Element, id_attr: str = "w:id") -> int:
    """Find the maximum ID value used in the document for a given attribute."""
    max_id = -1
    attr_name = qn(id_attr) if ":" in id_attr else id_attr

    for elem in root.iter():
        id_val = elem.get(attr_name)
        if id_val is not None:
            try:
                max_id = max(max_id, int(id_val))
            except ValueError:
                pass

    return max_id


def serialize_xml(root: etree._Element, xml_declaration: bool = True) -> bytes:
    """Serialize an XML tree to bytes, preserving namespaces."""
    return etree.tostring(
        root,
        xml_declaration=xml_declaration,
        encoding="UTF-8",
        standalone=True,
    )
