"""Read operations for Word documents."""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from lxml import etree

from .xml_helpers import (
    NAMESPACES,
    get_paragraph_style,
    get_text_content,
    iter_paragraphs,
    normalize_typography,
    parse_datetime,
    qn,
)


@dataclass
class CommentReply:
    """A reply to a comment."""

    id: int
    parent_id: int
    author: str
    date: str | None
    text: str


@dataclass
class Comment:
    """A document comment with optional replies."""

    id: int
    author: str
    date: str | None
    text: str
    anchor_text: str | None = None
    anchor_paragraph: int | None = None
    resolved: bool = False
    replies: list[CommentReply] = field(default_factory=list)


@dataclass
class TrackChange:
    """A tracked change (insertion or deletion)."""

    id: int
    type: str  # "insertion" or "deletion"
    author: str
    date: str | None
    text: str
    paragraph: int | None = None


@dataclass
class Paragraph:
    """A document paragraph."""

    index: int
    text: str
    style: str | None = None


@dataclass
class DocumentMetadata:
    """Document metadata."""

    path: str
    author: str | None = None
    created: str | None = None
    modified: str | None = None
    word_count: int = 0


@dataclass
class SearchMatch:
    """A search match with surrounding context."""

    paragraph_index: int
    paragraph_text: str
    paragraph_style: str | None
    match_start: int  # Character offset within paragraph
    match_end: int  # Character offset within paragraph
    context_before: list[Paragraph]
    context_after: list[Paragraph]
    comments: list[Comment] | None = None
    track_changes: list[TrackChange] | None = None

    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        result: dict[str, Any] = {
            "paragraph_index": self.paragraph_index,
            "paragraph_text": self.paragraph_text,
            "paragraph_style": self.paragraph_style,
            "match_start": self.match_start,
            "match_end": self.match_end,
            "context_before": [
                {"index": p.index, "text": p.text, "style": p.style}
                for p in self.context_before
            ],
            "context_after": [
                {"index": p.index, "text": p.text, "style": p.style}
                for p in self.context_after
            ],
        }
        if self.comments is not None:
            result["comments"] = [
                {
                    "id": c.id,
                    "author": c.author,
                    "date": c.date,
                    "text": c.text,
                    "anchor_text": c.anchor_text,
                    "anchor_paragraph": c.anchor_paragraph,
                    "resolved": c.resolved,
                }
                for c in self.comments
            ]
        if self.track_changes is not None:
            result["track_changes"] = [
                {
                    "id": tc.id,
                    "type": tc.type,
                    "author": tc.author,
                    "date": tc.date,
                    "text": tc.text,
                    "paragraph": tc.paragraph,
                }
                for tc in self.track_changes
            ]
        return result


@dataclass
class DocumentContent:
    """Complete document content with all extracted data."""

    metadata: DocumentMetadata
    paragraphs: list[Paragraph]
    comments: list[Comment]
    track_changes: list[TrackChange]

    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return {
            "metadata": {
                "path": self.metadata.path,
                "author": self.metadata.author,
                "created": self.metadata.created,
                "modified": self.metadata.modified,
                "word_count": self.metadata.word_count,
            },
            "paragraphs": [
                {"index": p.index, "text": p.text, "style": p.style}
                for p in self.paragraphs
            ],
            "comments": [
                {
                    "id": c.id,
                    "author": c.author,
                    "date": c.date,
                    "text": c.text,
                    "anchor_text": c.anchor_text,
                    "anchor_paragraph": c.anchor_paragraph,
                    "resolved": c.resolved,
                    "replies": [
                        {
                            "id": r.id,
                            "parent_id": r.parent_id,
                            "author": r.author,
                            "date": r.date,
                            "text": r.text,
                        }
                        for r in c.replies
                    ],
                }
                for c in self.comments
            ],
            "track_changes": [
                {
                    "id": tc.id,
                    "type": tc.type,
                    "author": tc.author,
                    "date": tc.date,
                    "text": tc.text,
                    "paragraph": tc.paragraph,
                }
                for tc in self.track_changes
            ],
        }


class DocxReader:
    """Reader for Word documents with comments and track changes support."""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        self._zip: zipfile.ZipFile | None = None
        self._document: etree._Element | None = None
        self._comments_xml: etree._Element | None = None
        self._comments_extended_xml: etree._Element | None = None
        self._core_props: etree._Element | None = None

    def __enter__(self) -> "DocxReader":
        self._zip = zipfile.ZipFile(self.path, "r")
        self._load_parts()
        return self

    def __exit__(self, *args) -> None:
        if self._zip:
            self._zip.close()
            self._zip = None

    def _load_parts(self) -> None:
        """Load XML parts from the docx archive."""
        assert self._zip is not None

        # Main document
        with self._zip.open("word/document.xml") as f:
            self._document = etree.parse(f).getroot()

        # Comments (optional)
        if "word/comments.xml" in self._zip.namelist():
            with self._zip.open("word/comments.xml") as f:
                self._comments_xml = etree.parse(f).getroot()

        # Extended comments for threading (optional)
        if "word/commentsExtended.xml" in self._zip.namelist():
            with self._zip.open("word/commentsExtended.xml") as f:
                self._comments_extended_xml = etree.parse(f).getroot()

        # Core properties (optional)
        if "docProps/core.xml" in self._zip.namelist():
            with self._zip.open("docProps/core.xml") as f:
                self._core_props = etree.parse(f).getroot()

    def read(
        self,
        include_text: bool = True,
        include_comments: bool = True,
        include_track_changes: bool = True,
    ) -> DocumentContent:
        """Read the document and extract requested content."""
        metadata = self._read_metadata()
        paragraphs = self._read_paragraphs() if include_text else []
        comments = self._read_comments() if include_comments else []
        track_changes = self._read_track_changes() if include_track_changes else []

        # Calculate word count from paragraphs
        if paragraphs:
            all_text = " ".join(p.text for p in paragraphs)
            metadata.word_count = len(all_text.split())

        return DocumentContent(
            metadata=metadata,
            paragraphs=paragraphs,
            comments=comments,
            track_changes=track_changes,
        )

    def search(
        self,
        query: str,
        case_sensitive: bool = False,
        context_paragraphs: int = 1,
        max_results: int = 20,
        include_annotations: bool = False,
    ) -> tuple[list[SearchMatch], int]:
        """Search for text in document paragraphs.

        Args:
            query: Text to search for
            case_sensitive: Match case exactly (default: False)
            context_paragraphs: Paragraphs to include before/after each match (default: 1)
            max_results: Maximum matches to return (default: 20)
            include_annotations: Include comments/track changes on matched paragraphs (default: False)

        Returns:
            Tuple of (matches list, total_matches count)
        """
        if not query:
            return [], 0

        paragraphs = self._read_paragraphs()

        # Load annotations if needed
        comments = self._read_comments() if include_annotations else None
        track_changes = self._read_track_changes() if include_annotations else None

        matches = []
        search_text = normalize_typography(query if case_sensitive else query.lower())

        for para in paragraphs:
            para_text = normalize_typography(para.text if case_sensitive else para.text.lower())
            pos = para_text.find(search_text)

            if pos != -1:
                # Build context
                start_idx = max(0, para.index - context_paragraphs)
                end_idx = min(len(paragraphs) - 1, para.index + context_paragraphs)

                context_before = [p for p in paragraphs[start_idx:para.index]]
                context_after = [p for p in paragraphs[para.index + 1:end_idx + 1]]

                # Filter annotations to this paragraph if requested
                para_comments = None
                para_changes = None
                if include_annotations:
                    para_comments = [c for c in (comments or []) if c.anchor_paragraph == para.index]
                    para_changes = [tc for tc in (track_changes or []) if tc.paragraph == para.index]

                matches.append(SearchMatch(
                    paragraph_index=para.index,
                    paragraph_text=para.text,
                    paragraph_style=para.style,
                    match_start=pos,
                    match_end=pos + len(query),
                    context_before=context_before,
                    context_after=context_after,
                    comments=para_comments,
                    track_changes=para_changes,
                ))

        total = len(matches)
        return matches[:max_results], total

    def get_paragraph_range(
        self,
        start_index: int,
        end_index: int,
        include_annotations: bool = False,
    ) -> dict[str, Any]:
        """Get a range of paragraphs.

        Args:
            start_index: First paragraph index (0-based, inclusive)
            end_index: Last paragraph index (0-based, inclusive)
            include_annotations: Include comments/track changes in range (default: False)

        Returns:
            Dictionary containing paragraphs and optionally annotations
        """
        paragraphs = self._read_paragraphs()

        # Clamp to valid range
        start = max(0, start_index)
        end = min(len(paragraphs) - 1, end_index) if paragraphs else -1

        result_paragraphs = paragraphs[start:end + 1] if end >= start else []

        result: dict[str, Any] = {
            "start_index": start,
            "end_index": end if end >= 0 else 0,
            "total_paragraphs": len(paragraphs),
            "paragraphs": [
                {"index": p.index, "text": p.text, "style": p.style}
                for p in result_paragraphs
            ],
        }

        if include_annotations:
            indices = set(range(start, end + 1)) if end >= start else set()
            comments = self._read_comments()
            track_changes = self._read_track_changes()
            result["comments"] = [
                {
                    "id": c.id,
                    "author": c.author,
                    "date": c.date,
                    "text": c.text,
                    "anchor_text": c.anchor_text,
                    "anchor_paragraph": c.anchor_paragraph,
                    "resolved": c.resolved,
                }
                for c in comments if c.anchor_paragraph in indices
            ]
            result["track_changes"] = [
                {
                    "id": tc.id,
                    "type": tc.type,
                    "author": tc.author,
                    "date": tc.date,
                    "text": tc.text,
                    "paragraph": tc.paragraph,
                }
                for tc in track_changes if tc.paragraph in indices
            ]

        return result

    def _read_metadata(self) -> DocumentMetadata:
        """Read document metadata from core properties."""
        metadata = DocumentMetadata(path=str(self.path))

        if self._core_props is not None:
            # Author
            creator = self._core_props.find("dc:creator", NAMESPACES)
            if creator is not None and creator.text:
                metadata.author = creator.text

            # Created date
            created = self._core_props.find("dcterms:created", NAMESPACES)
            if created is not None and created.text:
                metadata.created = parse_datetime(created.text)

            # Modified date
            modified = self._core_props.find("dcterms:modified", NAMESPACES)
            if modified is not None and modified.text:
                metadata.modified = parse_datetime(modified.text)

        return metadata

    def _read_paragraphs(self) -> list[Paragraph]:
        """Read all paragraphs from the document (cached after first call)."""
        if self._document is None:
            return []

        if hasattr(self, "_paragraphs_cache"):
            return self._paragraphs_cache

        paragraphs = []
        for idx, para_elem in iter_paragraphs(self._document):
            text = get_text_content(para_elem)
            style = get_paragraph_style(para_elem)
            paragraphs.append(Paragraph(index=idx, text=text, style=style))

        self._paragraphs_cache = paragraphs
        return paragraphs

    def _read_comments(self) -> list[Comment]:
        """Read all comments and their replies."""
        if self._comments_xml is None:
            return []

        # First, build a map of comment IDs to their data
        comments_map: dict[int, Comment] = {}

        # Build a mapping from comment ID to paragraph ID for status lookup
        comment_id_to_para_id: dict[int, str] = {}

        for comment_elem in self._comments_xml.iter(qn("w:comment")):
            comment_id = int(comment_elem.get(qn("w:id"), "0"))
            author = comment_elem.get(qn("w:author"), "Unknown")
            date = parse_datetime(comment_elem.get(qn("w:date")))
            text = get_text_content(comment_elem)

            # Get paragraph ID for status lookup
            first_para = comment_elem.find(qn("w:p"))
            if first_para is not None:
                para_id = first_para.get(qn("w14:paraId"))
                if para_id:
                    comment_id_to_para_id[comment_id] = para_id

            comments_map[comment_id] = Comment(
                id=comment_id,
                author=author,
                date=date,
                text=text,
            )

        # Find anchor text and paragraph for each comment
        self._find_comment_anchors(comments_map)

        # Process reply threading and resolved status from commentsExtended.xml
        self._process_comment_extended(comments_map, comment_id_to_para_id)

        # Return top-level comments (those without a parent)
        top_level_ids = set(comments_map.keys())
        for comment in comments_map.values():
            for reply in comment.replies:
                top_level_ids.discard(reply.id)

        return [comments_map[cid] for cid in sorted(top_level_ids) if cid in comments_map]

    def _get_para_map(self) -> dict[etree._Element, int]:
        """Get paragraph element to index map (cached after first call)."""
        if hasattr(self, "_para_map_cache"):
            return self._para_map_cache

        para_map: dict[etree._Element, int] = {}
        if self._document is not None:
            for idx, para_elem in iter_paragraphs(self._document):
                para_map[para_elem] = idx
        self._para_map_cache = para_map
        return para_map

    def _find_comment_anchors(self, comments_map: dict[int, Comment]) -> None:
        """Find the anchor text and paragraph for each comment."""
        if self._document is None:
            return

        para_map = self._get_para_map()

        # Find comment ranges in document
        for comment_id, comment in comments_map.items():
            range_start = self._document.find(
                f".//{qn('w:commentRangeStart')}[@{qn('w:id')}='{comment_id}']"
            )
            range_end = self._document.find(
                f".//{qn('w:commentRangeEnd')}[@{qn('w:id')}='{comment_id}']"
            )

            if range_start is not None:
                # Find the paragraph containing the range start
                parent = range_start.getparent()
                while parent is not None and parent.tag != qn("w:p"):
                    parent = parent.getparent()

                if parent is not None and parent in para_map:
                    comment.anchor_paragraph = para_map[parent]

                # Extract anchor text between range start and end
                anchor_text = self._extract_range_text(range_start, range_end)
                if anchor_text:
                    comment.anchor_text = anchor_text

    def _extract_range_text(
        self, range_start: etree._Element, range_end: etree._Element | None
    ) -> str:
        """Extract text between comment range start and end markers."""
        if range_end is None:
            return ""

        # Get IDs to match
        start_id = range_start.get(qn("w:id"))

        # Walk through siblings and collect text
        texts = []
        current = range_start.getnext()
        collecting = True

        while current is not None and collecting:
            # Check if we hit the range end
            if current.tag == qn("w:commentRangeEnd"):
                if current.get(qn("w:id")) == start_id:
                    break

            # Collect text from this element
            if current.tag == qn("w:r"):
                for t in current.iter(qn("w:t")):
                    if t.text:
                        texts.append(t.text)

            current = current.getnext()

        return "".join(texts)

    def _process_comment_extended(
        self, comments_map: dict[int, Comment], comment_id_to_para_id: dict[int, str]
    ) -> None:
        """Process comment threading and resolved status from commentsExtended.xml."""
        if self._comments_extended_xml is None:
            return

        # Build reverse mapping: paraId -> comment ID
        para_to_comment: dict[str, int] = {
            para_id: comment_id for comment_id, para_id in comment_id_to_para_id.items()
        }

        # Process extended comments for threading and resolved status
        for comment_ex in self._comments_extended_xml.iter(qn("w15:commentEx")):
            para_id = comment_ex.get(qn("w15:paraId"))
            parent_para_id = comment_ex.get(qn("w15:paraIdParent"))
            done = comment_ex.get(qn("w15:done"))

            if para_id:
                comment_id = para_to_comment.get(para_id)
                if comment_id is not None and comment_id in comments_map:
                    # Set resolved status (done="1" means resolved)
                    if done == "1":
                        comments_map[comment_id].resolved = True

            # Handle threading
            if para_id and parent_para_id:
                child_id = para_to_comment.get(para_id)
                parent_id = para_to_comment.get(parent_para_id)

                if child_id is not None and parent_id is not None:
                    if child_id in comments_map and parent_id in comments_map:
                        child = comments_map[child_id]
                        parent = comments_map[parent_id]

                        # Convert child to reply and add to parent
                        reply = CommentReply(
                            id=child.id,
                            parent_id=parent.id,
                            author=child.author,
                            date=child.date,
                            text=child.text,
                        )
                        parent.replies.append(reply)

    def _read_track_changes(self) -> list[TrackChange]:
        """Read all track changes (insertions and deletions)."""
        if self._document is None:
            return []

        changes = []

        para_map = self._get_para_map()

        # Find insertions
        for ins_elem in self._document.iter(qn("w:ins")):
            change_id = int(ins_elem.get(qn("w:id"), "0"))
            author = ins_elem.get(qn("w:author"), "Unknown")
            date = parse_datetime(ins_elem.get(qn("w:date")))
            text = get_text_content(ins_elem)

            # Find containing paragraph
            para_idx = None
            parent = ins_elem.getparent()
            while parent is not None:
                if parent in para_map:
                    para_idx = para_map[parent]
                    break
                parent = parent.getparent()

            changes.append(
                TrackChange(
                    id=change_id,
                    type="insertion",
                    author=author,
                    date=date,
                    text=text,
                    paragraph=para_idx,
                )
            )

        # Find deletions
        for del_elem in self._document.iter(qn("w:del")):
            change_id = int(del_elem.get(qn("w:id"), "0"))
            author = del_elem.get(qn("w:author"), "Unknown")
            date = parse_datetime(del_elem.get(qn("w:date")))
            text = get_text_content(del_elem)

            # Find containing paragraph
            para_idx = None
            parent = del_elem.getparent()
            while parent is not None:
                if parent in para_map:
                    para_idx = para_map[parent]
                    break
                parent = parent.getparent()

            changes.append(
                TrackChange(
                    id=change_id,
                    type="deletion",
                    author=author,
                    date=date,
                    text=text,
                    paragraph=para_idx,
                )
            )

        return sorted(changes, key=lambda c: c.id)


def read_docx(
    path: str,
    include_text: bool = True,
    include_comments: bool = True,
    include_track_changes: bool = True,
) -> dict[str, Any]:
    """Read a Word document and extract content.

    Args:
        path: Path to the .docx file
        include_text: Include full document text
        include_comments: Include comments with anchors
        include_track_changes: Include insertions/deletions

    Returns:
        Dictionary containing metadata, paragraphs, comments, and track_changes
    """
    with DocxReader(path) as reader:
        content = reader.read(
            include_text=include_text,
            include_comments=include_comments,
            include_track_changes=include_track_changes,
        )
        return content.to_dict()


def search_docx(
    path: str,
    query: str,
    case_sensitive: bool = False,
    context_paragraphs: int = 1,
    max_results: int = 20,
    include_annotations: bool = False,
) -> dict[str, Any]:
    """Search for text in a Word document.

    Args:
        path: Path to the .docx file
        query: Text to search for
        case_sensitive: Match case exactly (default: False)
        context_paragraphs: Paragraphs to include before/after each match (default: 1)
        max_results: Maximum matches to return (default: 20)
        include_annotations: Include comments/track changes on matched paragraphs (default: False)

    Returns:
        Dictionary containing query, total_matches, matches_returned, and matches list
    """
    with DocxReader(path) as reader:
        matches, total = reader.search(
            query=query,
            case_sensitive=case_sensitive,
            context_paragraphs=context_paragraphs,
            max_results=max_results,
            include_annotations=include_annotations,
        )
        return {
            "query": query,
            "case_sensitive": case_sensitive,
            "total_matches": total,
            "matches_returned": len(matches),
            "matches": [m.to_dict() for m in matches],
        }


def get_paragraph_range_docx(
    path: str,
    start_index: int,
    end_index: int,
    include_annotations: bool = False,
) -> dict[str, Any]:
    """Get a range of paragraphs from a Word document.

    Args:
        path: Path to the .docx file
        start_index: First paragraph index (0-based, inclusive)
        end_index: Last paragraph index (0-based, inclusive)
        include_annotations: Include comments/track changes in range (default: False)

    Returns:
        Dictionary containing paragraphs and optionally annotations
    """
    with DocxReader(path) as reader:
        return reader.get_paragraph_range(
            start_index=start_index,
            end_index=end_index,
            include_annotations=include_annotations,
        )
