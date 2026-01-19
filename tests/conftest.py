"""Pytest configuration and fixtures."""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Callable

import pytest
from lxml import etree

FIXTURES_DIR = Path(__file__).parent / "fixtures"


# OOXML namespace constants
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def create_content_types() -> bytes:
    """Create [Content_Types].xml."""
    root = etree.Element(
        f"{{{CT_NS}}}Types",
        nsmap={None: CT_NS},
    )

    # Default extensions
    etree.SubElement(
        root,
        f"{{{CT_NS}}}Default",
        Extension="rels",
        ContentType="application/vnd.openxmlformats-package.relationships+xml",
    )
    etree.SubElement(
        root,
        f"{{{CT_NS}}}Default",
        Extension="xml",
        ContentType="application/xml",
    )

    # Override for specific parts
    etree.SubElement(
        root,
        f"{{{CT_NS}}}Override",
        PartName="/word/document.xml",
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    )
    etree.SubElement(
        root,
        f"{{{CT_NS}}}Override",
        PartName="/docProps/core.xml",
        ContentType="application/vnd.openxmlformats-package.core-properties+xml",
    )

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def create_rels() -> bytes:
    """Create _rels/.rels."""
    root = etree.Element(
        f"{{{PR_NS}}}Relationships",
        nsmap={None: PR_NS},
    )

    etree.SubElement(
        root,
        f"{{{PR_NS}}}Relationship",
        Id="rId1",
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        Target="word/document.xml",
    )
    etree.SubElement(
        root,
        f"{{{PR_NS}}}Relationship",
        Id="rId2",
        Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        Target="docProps/core.xml",
    )

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def create_document_rels(has_comments: bool = False) -> bytes:
    """Create word/_rels/document.xml.rels."""
    root = etree.Element(
        f"{{{PR_NS}}}Relationships",
        nsmap={None: PR_NS},
    )

    if has_comments:
        etree.SubElement(
            root,
            f"{{{PR_NS}}}Relationship",
            Id="rId1",
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
            Target="comments.xml",
        )

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def create_core_props(author: str = "Test Author") -> bytes:
    """Create docProps/core.xml."""
    nsmap = {
        "cp": CP_NS,
        "dc": DC_NS,
        "dcterms": DCTERMS_NS,
    }

    root = etree.Element(f"{{{CP_NS}}}coreProperties", nsmap=nsmap)

    creator = etree.SubElement(root, f"{{{DC_NS}}}creator")
    creator.text = author

    created = etree.SubElement(
        root,
        f"{{{DCTERMS_NS}}}created",
        {"{http://www.w3.org/2001/XMLSchema-instance}type": "dcterms:W3CDTF"},
    )
    created.text = "2025-01-15T10:30:00Z"

    modified = etree.SubElement(
        root,
        f"{{{DCTERMS_NS}}}modified",
        {"{http://www.w3.org/2001/XMLSchema-instance}type": "dcterms:W3CDTF"},
    )
    modified.text = "2025-01-18T14:22:00Z"

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def create_simple_document(paragraphs: list[str]) -> bytes:
    """Create a simple word/document.xml with given paragraphs."""
    nsmap = {
        "w": W_NS,
        "r": R_NS,
    }

    root = etree.Element(f"{{{W_NS}}}document", nsmap=nsmap)
    body = etree.SubElement(root, f"{{{W_NS}}}body")

    for para_text in paragraphs:
        p = etree.SubElement(body, f"{{{W_NS}}}p")
        r = etree.SubElement(p, f"{{{W_NS}}}r")
        t = etree.SubElement(r, f"{{{W_NS}}}t")
        t.text = para_text

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def create_document_with_comments(
    paragraphs: list[str],
    comment_anchors: list[tuple[int, str, int]],  # (para_idx, anchor_text, comment_id)
) -> bytes:
    """Create document.xml with comment range markers."""
    nsmap = {
        "w": W_NS,
        "w14": W14_NS,
        "r": R_NS,
    }

    root = etree.Element(f"{{{W_NS}}}document", nsmap=nsmap)
    body = etree.SubElement(root, f"{{{W_NS}}}body")

    for para_idx, para_text in enumerate(paragraphs):
        p = etree.SubElement(body, f"{{{W_NS}}}p")

        # Check if this paragraph has a comment anchor
        anchors_in_para = [a for a in comment_anchors if a[0] == para_idx]

        if not anchors_in_para:
            # Simple paragraph
            r = etree.SubElement(p, f"{{{W_NS}}}r")
            t = etree.SubElement(r, f"{{{W_NS}}}t")
            t.text = para_text
        else:
            # Paragraph with comment anchors
            remaining_text = para_text
            for anchor_para_idx, anchor_text, comment_id in anchors_in_para:
                if anchor_text in remaining_text:
                    before, after = remaining_text.split(anchor_text, 1)

                    # Text before anchor
                    if before:
                        r = etree.SubElement(p, f"{{{W_NS}}}r")
                        t = etree.SubElement(r, f"{{{W_NS}}}t")
                        t.text = before

                    # Comment range start
                    etree.SubElement(
                        p,
                        f"{{{W_NS}}}commentRangeStart",
                        {f"{{{W_NS}}}id": str(comment_id)},
                    )

                    # Anchor text
                    r = etree.SubElement(p, f"{{{W_NS}}}r")
                    t = etree.SubElement(r, f"{{{W_NS}}}t")
                    t.text = anchor_text

                    # Comment range end
                    etree.SubElement(
                        p,
                        f"{{{W_NS}}}commentRangeEnd",
                        {f"{{{W_NS}}}id": str(comment_id)},
                    )

                    # Comment reference
                    r = etree.SubElement(p, f"{{{W_NS}}}r")
                    etree.SubElement(
                        r,
                        f"{{{W_NS}}}commentReference",
                        {f"{{{W_NS}}}id": str(comment_id)},
                    )

                    remaining_text = after

            # Any remaining text
            if remaining_text:
                r = etree.SubElement(p, f"{{{W_NS}}}r")
                t = etree.SubElement(r, f"{{{W_NS}}}t")
                t.text = remaining_text

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def create_comments_xml(
    comments: list[tuple[int, str, str, str]],  # (id, author, date, text)
) -> bytes:
    """Create word/comments.xml."""
    nsmap = {
        "w": W_NS,
        "w14": W14_NS,
    }

    root = etree.Element(f"{{{W_NS}}}comments", nsmap=nsmap)

    for comment_id, author, date, text in comments:
        comment = etree.SubElement(
            root,
            f"{{{W_NS}}}comment",
            {
                f"{{{W_NS}}}id": str(comment_id),
                f"{{{W_NS}}}author": author,
                f"{{{W_NS}}}date": date,
            },
        )
        # Add paragraph with text
        p = etree.SubElement(
            comment,
            f"{{{W_NS}}}p",
            {f"{{{W14_NS}}}paraId": f"para{comment_id:08X}"},
        )
        r = etree.SubElement(p, f"{{{W_NS}}}r")
        t = etree.SubElement(r, f"{{{W_NS}}}t")
        t.text = text

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def create_comments_extended_xml(
    threading: list[tuple[str, str | None, bool]],  # (paraId, parentParaId or None, resolved)
) -> bytes:
    """Create word/commentsExtended.xml for reply threading and resolved status."""
    nsmap = {
        "w15": W15_NS,
    }

    root = etree.Element(f"{{{W15_NS}}}commentsEx", nsmap=nsmap)

    for para_id, parent_para_id, resolved in threading:
        attrs = {
            f"{{{W15_NS}}}paraId": para_id,
            f"{{{W15_NS}}}done": "1" if resolved else "0",
        }
        if parent_para_id:
            attrs[f"{{{W15_NS}}}paraIdParent"] = parent_para_id

        etree.SubElement(root, f"{{{W15_NS}}}commentEx", attrs)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def create_document_with_track_changes(
    paragraphs: list[str],
    insertions: list[tuple[int, str, int, str, str]],  # (para_idx, text, id, author, date)
    deletions: list[tuple[int, str, int, str, str]],  # (para_idx, text, id, author, date)
) -> bytes:
    """Create document.xml with track changes."""
    nsmap = {
        "w": W_NS,
        "r": R_NS,
    }

    root = etree.Element(f"{{{W_NS}}}document", nsmap=nsmap)
    body = etree.SubElement(root, f"{{{W_NS}}}body")

    for para_idx, para_text in enumerate(paragraphs):
        p = etree.SubElement(body, f"{{{W_NS}}}p")

        # Get changes for this paragraph
        para_insertions = [i for i in insertions if i[0] == para_idx]
        para_deletions = [d for d in deletions if d[0] == para_idx]

        if not para_insertions and not para_deletions:
            # Simple paragraph
            r = etree.SubElement(p, f"{{{W_NS}}}r")
            t = etree.SubElement(r, f"{{{W_NS}}}t")
            t.text = para_text
        else:
            # Handle insertions
            for _, ins_text, ins_id, author, date in para_insertions:
                ins = etree.SubElement(
                    p,
                    f"{{{W_NS}}}ins",
                    {
                        f"{{{W_NS}}}id": str(ins_id),
                        f"{{{W_NS}}}author": author,
                        f"{{{W_NS}}}date": date,
                    },
                )
                r = etree.SubElement(ins, f"{{{W_NS}}}r")
                t = etree.SubElement(r, f"{{{W_NS}}}t")
                t.text = ins_text

            # Handle deletions
            for _, del_text, del_id, author, date in para_deletions:
                deletion = etree.SubElement(
                    p,
                    f"{{{W_NS}}}del",
                    {
                        f"{{{W_NS}}}id": str(del_id),
                        f"{{{W_NS}}}author": author,
                        f"{{{W_NS}}}date": date,
                    },
                )
                r = etree.SubElement(deletion, f"{{{W_NS}}}r")
                t = etree.SubElement(r, f"{{{W_NS}}}delText")
                t.text = del_text

            # Add remaining text if any (simplified - puts it after changes)
            if para_text:
                r = etree.SubElement(p, f"{{{W_NS}}}r")
                t = etree.SubElement(r, f"{{{W_NS}}}t")
                t.text = para_text

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def write_docx(
    path: Path,
    document_xml: bytes,
    comments_xml: bytes | None = None,
    comments_extended_xml: bytes | None = None,
    author: str = "Test Author",
) -> None:
    """Write a complete docx file."""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        # Required parts
        zf.writestr("[Content_Types].xml", create_content_types())
        zf.writestr("_rels/.rels", create_rels())
        zf.writestr("word/document.xml", document_xml)
        zf.writestr(
            "word/_rels/document.xml.rels",
            create_document_rels(has_comments=comments_xml is not None),
        )
        zf.writestr("docProps/core.xml", create_core_props(author))

        # Optional parts
        if comments_xml:
            zf.writestr("word/comments.xml", comments_xml)

        if comments_extended_xml:
            zf.writestr("word/commentsExtended.xml", comments_extended_xml)


@pytest.fixture
def fixtures_dir() -> Path:
    """Return the fixtures directory path."""
    FIXTURES_DIR.mkdir(parents=True, exist_ok=True)
    return FIXTURES_DIR


@pytest.fixture
def simple_docx(fixtures_dir: Path) -> Path:
    """Create a simple document with no comments or track changes."""
    path = fixtures_dir / "simple.docx"

    paragraphs = [
        "This is the first paragraph of the document.",
        "The second paragraph contains more text for testing purposes.",
        "Finally, the third paragraph concludes our simple test document.",
    ]

    document_xml = create_simple_document(paragraphs)
    write_docx(path, document_xml)

    return path


@pytest.fixture
def docx_with_comments(fixtures_dir: Path) -> Path:
    """Create a document with comments and one reply."""
    path = fixtures_dir / "with_comments.docx"

    paragraphs = [
        "This is the introduction to our research paper.",
        "The study examines disorganized attachment patterns in early childhood.",
        "Our methodology follows established protocols from previous research.",
    ]

    # Comment anchors: (para_idx, anchor_text, comment_id)
    comment_anchors = [
        (1, "disorganized attachment patterns", 0),
        (2, "established protocols", 1),
    ]

    # Comments: (id, author, date, text)
    comments = [
        (0, "Dr. Smith", "2025-01-16T09:15:00Z", "Consider citing Main & Hesse here"),
        (1, "Dr. Jones", "2025-01-17T10:00:00Z", "Which protocols specifically?"),
        (2, "Josh", "2025-01-17T11:00:00Z", "Added citation - see revision"),  # Reply to comment 0
    ]

    document_xml = create_document_with_comments(paragraphs, comment_anchors)
    comments_xml = create_comments_xml(comments)

    # Threading: comment 2 is a reply to comment 0
    # Format: (paraId, parentParaId, resolved)
    threading = [
        ("para00000000", None, False),  # Comment 0, no parent, not resolved
        ("para00000001", None, False),  # Comment 1, no parent, not resolved
        ("para00000002", "para00000000", False),  # Comment 2, parent is comment 0, not resolved
    ]
    comments_extended_xml = create_comments_extended_xml(threading)

    write_docx(
        path,
        document_xml,
        comments_xml=comments_xml,
        comments_extended_xml=comments_extended_xml,
    )

    return path


@pytest.fixture
def docx_with_track_changes(fixtures_dir: Path) -> Path:
    """Create a document with track changes (insertions and deletions)."""
    path = fixtures_dir / "with_track_changes.docx"

    paragraphs = [
        "The children showed attachment behaviors.",
        "",  # This paragraph will contain the track changes
        "The study concludes with recommendations.",
    ]

    # Insertions and deletions in paragraph 1
    insertions = [
        (1, "frequently", 6, "Dr. Smith", "2025-01-16T09:20:00Z"),
    ]
    deletions = [
        (1, "invariably", 5, "Dr. Smith", "2025-01-16T09:20:00Z"),
    ]

    document_xml = create_document_with_track_changes(paragraphs, insertions, deletions)
    write_docx(path, document_xml)

    return path


@pytest.fixture
def docx_with_resolved_comment(fixtures_dir: Path) -> Path:
    """Create a document with one resolved and one open comment."""
    path = fixtures_dir / "with_resolved_comment.docx"

    paragraphs = [
        "This is the first paragraph with some content.",
        "This is the second paragraph with more content.",
    ]

    # Comment anchors: (para_idx, anchor_text, comment_id)
    comment_anchors = [
        (0, "first paragraph", 0),
        (1, "second paragraph", 1),
    ]

    # Comments: (id, author, date, text)
    comments = [
        (0, "Reviewer", "2025-01-16T09:00:00Z", "This comment is resolved"),
        (1, "Reviewer", "2025-01-16T10:00:00Z", "This comment is still open"),
    ]

    document_xml = create_document_with_comments(paragraphs, comment_anchors)
    comments_xml = create_comments_xml(comments)

    # Format: (paraId, parentParaId, resolved)
    threading = [
        ("para00000000", None, True),   # Comment 0 is resolved
        ("para00000001", None, False),  # Comment 1 is not resolved
    ]
    comments_extended_xml = create_comments_extended_xml(threading)

    write_docx(
        path,
        document_xml,
        comments_xml=comments_xml,
        comments_extended_xml=comments_extended_xml,
    )

    return path


@pytest.fixture
def complex_docx(fixtures_dir: Path) -> Path:
    """Create a document with comments, replies, and track changes."""
    path = fixtures_dir / "complex.docx"

    paragraphs = [
        "Introduction to the research study.",
        "The methodology section describes our approach to studying attachment.",
        "Results indicate significant findings in the data analysis.",
    ]

    # Comment in paragraph 1
    comment_anchors = [
        (1, "attachment", 0),
    ]

    # Comments with replies
    comments = [
        (0, "Dr. Smith", "2025-01-16T09:00:00Z", "Good methodology description"),
        (1, "Josh", "2025-01-16T10:00:00Z", "Thank you for the feedback"),
    ]

    # Track changes in paragraph 2
    insertions = [
        (2, "highly ", 10, "Dr. Smith", "2025-01-16T11:00:00Z"),
    ]
    deletions = []

    # Build document with both comments and track changes
    nsmap = {
        "w": W_NS,
        "w14": W14_NS,
        "r": R_NS,
    }

    root = etree.Element(f"{{{W_NS}}}document", nsmap=nsmap)
    body = etree.SubElement(root, f"{{{W_NS}}}body")

    # Paragraph 0: simple
    p0 = etree.SubElement(body, f"{{{W_NS}}}p")
    r0 = etree.SubElement(p0, f"{{{W_NS}}}r")
    t0 = etree.SubElement(r0, f"{{{W_NS}}}t")
    t0.text = paragraphs[0]

    # Paragraph 1: with comment
    p1 = etree.SubElement(body, f"{{{W_NS}}}p")
    r1a = etree.SubElement(p1, f"{{{W_NS}}}r")
    t1a = etree.SubElement(r1a, f"{{{W_NS}}}t")
    t1a.text = "The methodology section describes our approach to studying "

    etree.SubElement(p1, f"{{{W_NS}}}commentRangeStart", {f"{{{W_NS}}}id": "0"})
    r1b = etree.SubElement(p1, f"{{{W_NS}}}r")
    t1b = etree.SubElement(r1b, f"{{{W_NS}}}t")
    t1b.text = "attachment"
    etree.SubElement(p1, f"{{{W_NS}}}commentRangeEnd", {f"{{{W_NS}}}id": "0"})
    r1c = etree.SubElement(p1, f"{{{W_NS}}}r")
    etree.SubElement(r1c, f"{{{W_NS}}}commentReference", {f"{{{W_NS}}}id": "0"})

    r1d = etree.SubElement(p1, f"{{{W_NS}}}r")
    t1d = etree.SubElement(r1d, f"{{{W_NS}}}t")
    t1d.text = "."

    # Paragraph 2: with track change
    p2 = etree.SubElement(body, f"{{{W_NS}}}p")
    r2a = etree.SubElement(p2, f"{{{W_NS}}}r")
    t2a = etree.SubElement(r2a, f"{{{W_NS}}}t")
    t2a.text = "Results indicate "

    ins = etree.SubElement(
        p2,
        f"{{{W_NS}}}ins",
        {
            f"{{{W_NS}}}id": "10",
            f"{{{W_NS}}}author": "Dr. Smith",
            f"{{{W_NS}}}date": "2025-01-16T11:00:00Z",
        },
    )
    r_ins = etree.SubElement(ins, f"{{{W_NS}}}r")
    t_ins = etree.SubElement(r_ins, f"{{{W_NS}}}t")
    t_ins.text = "highly "

    r2b = etree.SubElement(p2, f"{{{W_NS}}}r")
    t2b = etree.SubElement(r2b, f"{{{W_NS}}}t")
    t2b.text = "significant findings in the data analysis."

    document_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    comments_xml = create_comments_xml(comments)

    # Threading: comment 1 is a reply to comment 0
    # Format: (paraId, parentParaId, resolved)
    threading = [
        ("para00000000", None, False),  # Not resolved
        ("para00000001", "para00000000", False),  # Reply, not resolved
    ]
    comments_extended_xml = create_comments_extended_xml(threading)

    write_docx(
        path,
        document_xml,
        comments_xml=comments_xml,
        comments_extended_xml=comments_extended_xml,
    )

    return path
