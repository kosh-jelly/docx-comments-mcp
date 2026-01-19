"""Tests for the reader module."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_comments_mcp.reader import read_docx, DocxReader


class TestReadSimpleDocument:
    """Tests for reading basic documents without comments or track changes."""

    def test_read_paragraphs(self, simple_docx: Path) -> None:
        """Should extract paragraphs from basic document."""
        result = read_docx(str(simple_docx))

        assert len(result["paragraphs"]) == 3
        assert result["paragraphs"][0]["text"] == "This is the first paragraph of the document."
        assert result["paragraphs"][1]["text"] == "The second paragraph contains more text for testing purposes."
        assert result["paragraphs"][2]["text"] == "Finally, the third paragraph concludes our simple test document."

    def test_read_metadata(self, simple_docx: Path) -> None:
        """Should extract document metadata."""
        result = read_docx(str(simple_docx))

        assert result["metadata"]["path"] == str(simple_docx)
        assert result["metadata"]["author"] == "Test Author"
        assert result["metadata"]["created"] == "2025-01-15T10:30:00Z"
        assert result["metadata"]["modified"] == "2025-01-18T14:22:00Z"

    def test_word_count(self, simple_docx: Path) -> None:
        """Should calculate word count."""
        result = read_docx(str(simple_docx))

        # Count words in our test paragraphs
        assert result["metadata"]["word_count"] > 0

    def test_no_comments_in_simple_doc(self, simple_docx: Path) -> None:
        """Simple document should have no comments."""
        result = read_docx(str(simple_docx))

        assert result["comments"] == []

    def test_no_track_changes_in_simple_doc(self, simple_docx: Path) -> None:
        """Simple document should have no track changes."""
        result = read_docx(str(simple_docx))

        assert result["track_changes"] == []


class TestReadComments:
    """Tests for reading comments and replies."""

    def test_read_comments(self, docx_with_comments: Path) -> None:
        """Should extract comments with author, date, text, and anchor."""
        result = read_docx(str(docx_with_comments))

        # Should have 2 top-level comments (comment 2 is a reply)
        assert len(result["comments"]) >= 2

        # Find the "Main & Hesse" comment
        main_comment = next(
            (c for c in result["comments"] if "Main & Hesse" in c["text"]),
            None,
        )
        assert main_comment is not None
        assert main_comment["author"] == "Dr. Smith"
        assert main_comment["date"] == "2025-01-16T09:15:00Z"

    def test_comment_anchor_text(self, docx_with_comments: Path) -> None:
        """Should correctly identify the text anchored by each comment."""
        result = read_docx(str(docx_with_comments))

        # Find the comment about attachment patterns
        attachment_comment = next(
            (c for c in result["comments"] if "Main & Hesse" in c["text"]),
            None,
        )
        assert attachment_comment is not None
        assert attachment_comment["anchor_text"] == "disorganized attachment patterns"
        assert attachment_comment["anchor_paragraph"] == 1

    def test_comment_replies(self, docx_with_comments: Path) -> None:
        """Should correctly thread replies to parent comments."""
        result = read_docx(str(docx_with_comments))

        # Find comment with replies
        comment_with_reply = next(
            (c for c in result["comments"] if c["replies"]),
            None,
        )

        if comment_with_reply:
            assert len(comment_with_reply["replies"]) >= 1
            reply = comment_with_reply["replies"][0]
            assert reply["parent_id"] == comment_with_reply["id"]
            assert reply["author"] is not None


class TestReadTrackChanges:
    """Tests for reading track changes."""

    def test_read_insertions(self, docx_with_track_changes: Path) -> None:
        """Should identify inserted text with author attribution."""
        result = read_docx(str(docx_with_track_changes))

        insertions = [tc for tc in result["track_changes"] if tc["type"] == "insertion"]
        assert len(insertions) >= 1

        insertion = insertions[0]
        assert insertion["text"] == "frequently"
        assert insertion["author"] == "Dr. Smith"

    def test_read_deletions(self, docx_with_track_changes: Path) -> None:
        """Should identify deleted text with author attribution."""
        result = read_docx(str(docx_with_track_changes))

        deletions = [tc for tc in result["track_changes"] if tc["type"] == "deletion"]
        assert len(deletions) >= 1

        deletion = deletions[0]
        assert deletion["text"] == "invariably"
        assert deletion["author"] == "Dr. Smith"

    def test_track_change_dates(self, docx_with_track_changes: Path) -> None:
        """Should include dates for track changes."""
        result = read_docx(str(docx_with_track_changes))

        for change in result["track_changes"]:
            assert change["date"] is not None


class TestReadComplexDocument:
    """Tests for documents with all features combined."""

    def test_read_complex_document(self, complex_docx: Path) -> None:
        """Should handle document with comments, replies, and track changes."""
        result = read_docx(str(complex_docx))

        # Should have paragraphs
        assert len(result["paragraphs"]) == 3

        # Should have comments
        assert len(result["comments"]) >= 1

        # Should have track changes
        assert len(result["track_changes"]) >= 1

    def test_complex_comments_and_replies(self, complex_docx: Path) -> None:
        """Should correctly parse comments with replies in complex document."""
        result = read_docx(str(complex_docx))

        # Check that we have comments
        assert len(result["comments"]) >= 1

        # At least one comment should have the methodology description
        methodology_comment = next(
            (c for c in result["comments"] if "methodology" in c["text"].lower()),
            None,
        )
        assert methodology_comment is not None

    def test_complex_track_changes(self, complex_docx: Path) -> None:
        """Should correctly parse track changes in complex document."""
        result = read_docx(str(complex_docx))

        insertions = [tc for tc in result["track_changes"] if tc["type"] == "insertion"]
        assert len(insertions) >= 1

        # Should have the "highly " insertion
        highly_insertion = next(
            (tc for tc in insertions if "highly" in tc["text"]),
            None,
        )
        assert highly_insertion is not None


class TestCommentResolvedStatus:
    """Tests for reading comment resolved/done status."""

    def test_read_resolved_comment(self, docx_with_resolved_comment: Path) -> None:
        """Should correctly read the resolved status of comments."""
        result = read_docx(str(docx_with_resolved_comment))

        assert len(result["comments"]) == 2

        # Find the resolved comment
        resolved_comment = next(
            (c for c in result["comments"] if "resolved" in c["text"].lower()),
            None,
        )
        assert resolved_comment is not None
        assert resolved_comment["resolved"] is True

        # Find the open comment
        open_comment = next(
            (c for c in result["comments"] if "still open" in c["text"].lower()),
            None,
        )
        assert open_comment is not None
        assert open_comment["resolved"] is False

    def test_default_resolved_is_false(self, simple_docx: Path, tmp_path: Path) -> None:
        """Comments without commentsExtended info should default to resolved=False."""
        from docx_comments_mcp.server import create_comment

        output_path = tmp_path / "with_comment.docx"
        result = create_comment(
            str(simple_docx),
            anchor_text="first paragraph",
            comment_text="Test comment",
            output_path=str(output_path),
        )
        assert result["success"] is True

        # Read the document and verify resolved defaults to False
        doc = read_docx(str(output_path))
        assert len(doc["comments"]) == 1
        assert doc["comments"][0]["resolved"] is False

    def test_resolve_comment_updates_status(self, simple_docx: Path, tmp_path: Path) -> None:
        """Resolving a comment should update its resolved status to True."""
        from docx_comments_mcp.server import create_comment, mark_comment_resolved

        # Create a comment
        step1 = tmp_path / "step1.docx"
        result = create_comment(
            str(simple_docx),
            anchor_text="first paragraph",
            comment_text="Test comment to resolve",
            output_path=str(step1),
        )
        assert result["success"] is True
        comment_id = result["comment_id"]

        # Verify it starts as not resolved
        doc = read_docx(str(step1))
        assert doc["comments"][0]["resolved"] is False

        # Resolve the comment
        step2 = tmp_path / "step2.docx"
        result = mark_comment_resolved(
            str(step1),
            comment_id=comment_id,
            output_path=str(step2),
        )
        assert result["success"] is True

        # Verify it is now resolved
        doc = read_docx(str(step2))
        assert len(doc["comments"]) == 1
        assert doc["comments"][0]["resolved"] is True


class TestReadOptions:
    """Tests for read options (include_text, include_comments, etc.)."""

    def test_exclude_text(self, simple_docx: Path) -> None:
        """Should not include paragraphs when include_text=False."""
        result = read_docx(str(simple_docx), include_text=False)

        assert result["paragraphs"] == []

    def test_exclude_comments(self, docx_with_comments: Path) -> None:
        """Should not include comments when include_comments=False."""
        result = read_docx(str(docx_with_comments), include_comments=False)

        assert result["comments"] == []

    def test_exclude_track_changes(self, docx_with_track_changes: Path) -> None:
        """Should not include track changes when include_track_changes=False."""
        result = read_docx(str(docx_with_track_changes), include_track_changes=False)

        assert result["track_changes"] == []


class TestDocxReaderContextManager:
    """Tests for the DocxReader context manager."""

    def test_context_manager(self, simple_docx: Path) -> None:
        """Should properly open and close file with context manager."""
        with DocxReader(simple_docx) as reader:
            content = reader.read()
            assert len(content.paragraphs) == 3

    def test_file_not_found(self, fixtures_dir: Path) -> None:
        """Should raise error for non-existent file."""
        with pytest.raises(FileNotFoundError):
            with DocxReader(fixtures_dir / "nonexistent.docx") as reader:
                reader.read()
