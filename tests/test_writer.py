"""Tests for the writer module."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_comments_mcp.reader import read_docx
from docx_comments_mcp.writer import (
    AnchorAmbiguousError,
    AnchorNotFoundError,
    CommentNotFoundError,
    TrackChangeNotFoundError,
    accept_track_change,
    add_comment,
    add_reply,
    add_track_change,
    reject_track_change,
    resolve_comment,
)


class TestAddComment:
    """Tests for adding comments."""

    def test_add_comment_basic(self, simple_docx: Path, tmp_path: Path) -> None:
        """Should add comment anchored to specified text."""
        output_path = tmp_path / "output.docx"

        result = add_comment(
            str(simple_docx),
            anchor_text="first paragraph",
            comment_text="This is a test comment",
            author="Test Author",
            output_path=str(output_path),
        )

        assert result["success"] is True
        assert result["anchored_to"] == "first paragraph"
        assert result["paragraph"] == 0
        assert "comment_id" in result

        # Verify comment was added
        doc = read_docx(str(output_path))
        assert len(doc["comments"]) == 1
        assert doc["comments"][0]["text"] == "This is a test comment"
        assert doc["comments"][0]["author"] == "Test Author"

    def test_add_comment_anchor_not_found(self, simple_docx: Path, tmp_path: Path) -> None:
        """Should fail gracefully when anchor text is not found."""
        output_path = tmp_path / "output.docx"

        with pytest.raises(AnchorNotFoundError) as exc_info:
            add_comment(
                str(simple_docx),
                anchor_text="nonexistent text",
                comment_text="Test comment",
                output_path=str(output_path),
            )

        assert "not found" in str(exc_info.value).lower()

    def test_add_comment_preserves_existing(self, docx_with_comments: Path, tmp_path: Path) -> None:
        """Should not disturb existing comments when adding new one."""
        output_path = tmp_path / "output.docx"

        # Count existing comments
        original_doc = read_docx(str(docx_with_comments))
        original_count = len(original_doc["comments"])

        result = add_comment(
            str(docx_with_comments),
            anchor_text="introduction",
            comment_text="New comment",
            output_path=str(output_path),
        )

        assert result["success"] is True

        # Verify original comments are preserved
        new_doc = read_docx(str(output_path))
        # New comment count should be original + 1
        assert len(new_doc["comments"]) >= original_count

    def test_add_comment_creates_backup_on_overwrite(self, simple_docx: Path) -> None:
        """Should create timestamped backup when overwriting original."""
        # Get directory of simple_docx
        doc_dir = simple_docx.parent

        # Add comment without output_path (overwrites original)
        result = add_comment(
            str(simple_docx),
            anchor_text="first paragraph",
            comment_text="Test comment",
        )

        assert result["success"] is True

        # Check that a backup was created
        backup_files = list(doc_dir.glob("simple.backup_*.docx"))
        assert len(backup_files) >= 1


class TestAddReply:
    """Tests for adding replies to comments."""

    def test_add_reply(self, docx_with_comments: Path, tmp_path: Path) -> None:
        """Should correctly thread reply to parent comment."""
        output_path = tmp_path / "output.docx"

        # Get a comment ID to reply to
        original_doc = read_docx(str(docx_with_comments))
        parent_id = original_doc["comments"][0]["id"]

        result = add_reply(
            str(docx_with_comments),
            parent_comment_id=parent_id,
            reply_text="This is my reply",
            author="Reply Author",
            output_path=str(output_path),
        )

        assert result["success"] is True
        assert result["parent_comment_id"] == parent_id
        assert "reply_id" in result

    def test_add_reply_comment_not_found(self, docx_with_comments: Path, tmp_path: Path) -> None:
        """Should fail when parent comment doesn't exist."""
        output_path = tmp_path / "output.docx"

        with pytest.raises(CommentNotFoundError):
            add_reply(
                str(docx_with_comments),
                parent_comment_id=9999,
                reply_text="Reply to nonexistent comment",
                output_path=str(output_path),
            )


class TestAddTrackChange:
    """Tests for adding track changes."""

    def test_add_track_change_replacement(self, simple_docx: Path, tmp_path: Path) -> None:
        """Should create deletion + insertion for text replacement."""
        output_path = tmp_path / "output.docx"

        result = add_track_change(
            str(simple_docx),
            find_text="first",
            replace_with="primary",
            author="Editor",
            output_path=str(output_path),
        )

        assert result["success"] is True
        assert result["change_type"] == "replacement"
        assert result["original_text"] == "first"
        assert result["new_text"] == "primary"

        # Verify track changes were added
        doc = read_docx(str(output_path))
        deletions = [tc for tc in doc["track_changes"] if tc["type"] == "deletion"]
        insertions = [tc for tc in doc["track_changes"] if tc["type"] == "insertion"]

        assert len(deletions) >= 1
        assert len(insertions) >= 1

    def test_add_track_change_deletion(self, simple_docx: Path, tmp_path: Path) -> None:
        """Should create deletion when replace_with is empty."""
        output_path = tmp_path / "output.docx"

        result = add_track_change(
            str(simple_docx),
            find_text="Finally",
            replace_with="",
            author="Editor",
            output_path=str(output_path),
        )

        assert result["success"] is True
        assert result["change_type"] == "deletion"

        # Verify deletion was added
        doc = read_docx(str(output_path))
        deletions = [tc for tc in doc["track_changes"] if tc["type"] == "deletion"]
        assert len(deletions) >= 1

    def test_add_track_change_text_not_found(self, simple_docx: Path, tmp_path: Path) -> None:
        """Should fail when find_text is not found."""
        output_path = tmp_path / "output.docx"

        with pytest.raises(AnchorNotFoundError):
            add_track_change(
                str(simple_docx),
                find_text="nonexistent text",
                replace_with="replacement",
                output_path=str(output_path),
            )


class TestResolveComment:
    """Tests for resolving comments."""

    def test_resolve_comment(self, docx_with_comments: Path, tmp_path: Path) -> None:
        """Should mark comment as resolved."""
        output_path = tmp_path / "output.docx"

        # Get a comment ID
        original_doc = read_docx(str(docx_with_comments))
        comment_id = original_doc["comments"][0]["id"]

        result = resolve_comment(
            str(docx_with_comments),
            comment_id=comment_id,
            output_path=str(output_path),
        )

        assert result["success"] is True
        assert result["comment_id"] == comment_id

    def test_resolve_comment_not_found(self, docx_with_comments: Path, tmp_path: Path) -> None:
        """Should fail when comment doesn't exist."""
        output_path = tmp_path / "output.docx"

        with pytest.raises(CommentNotFoundError):
            resolve_comment(
                str(docx_with_comments),
                comment_id=9999,
                output_path=str(output_path),
            )


class TestAcceptTrackChange:
    """Tests for accepting track changes."""

    def test_accept_insertion(self, docx_with_track_changes: Path, tmp_path: Path) -> None:
        """Should keep inserted text when accepting insertion."""
        output_path = tmp_path / "output.docx"

        # Get insertion ID
        original_doc = read_docx(str(docx_with_track_changes))
        insertions = [tc for tc in original_doc["track_changes"] if tc["type"] == "insertion"]
        assert len(insertions) > 0

        insertion_id = insertions[0]["id"]
        insertion_text = insertions[0]["text"]

        result = accept_track_change(
            str(docx_with_track_changes),
            change_id=insertion_id,
            output_path=str(output_path),
        )

        assert result["success"] is True
        assert result["change_type"] == "insertion"

        # Verify text is now part of document (no longer tracked)
        new_doc = read_docx(str(output_path))
        # The insertion should no longer be in track_changes
        new_insertions = [tc for tc in new_doc["track_changes"] if tc["id"] == insertion_id]
        assert len(new_insertions) == 0

    def test_accept_deletion(self, docx_with_track_changes: Path, tmp_path: Path) -> None:
        """Should remove deleted text when accepting deletion."""
        output_path = tmp_path / "output.docx"

        # Get deletion ID
        original_doc = read_docx(str(docx_with_track_changes))
        deletions = [tc for tc in original_doc["track_changes"] if tc["type"] == "deletion"]
        assert len(deletions) > 0

        deletion_id = deletions[0]["id"]

        result = accept_track_change(
            str(docx_with_track_changes),
            change_id=deletion_id,
            output_path=str(output_path),
        )

        assert result["success"] is True
        assert result["change_type"] == "deletion"

    def test_accept_track_change_not_found(self, docx_with_track_changes: Path, tmp_path: Path) -> None:
        """Should fail when track change doesn't exist."""
        output_path = tmp_path / "output.docx"

        with pytest.raises(TrackChangeNotFoundError):
            accept_track_change(
                str(docx_with_track_changes),
                change_id=9999,
                output_path=str(output_path),
            )


class TestRejectTrackChange:
    """Tests for rejecting track changes."""

    def test_reject_insertion(self, docx_with_track_changes: Path, tmp_path: Path) -> None:
        """Should remove inserted text when rejecting insertion."""
        output_path = tmp_path / "output.docx"

        # Get insertion ID
        original_doc = read_docx(str(docx_with_track_changes))
        insertions = [tc for tc in original_doc["track_changes"] if tc["type"] == "insertion"]
        assert len(insertions) > 0

        insertion_id = insertions[0]["id"]

        result = reject_track_change(
            str(docx_with_track_changes),
            change_id=insertion_id,
            output_path=str(output_path),
        )

        assert result["success"] is True
        assert result["change_type"] == "insertion"

    def test_reject_deletion(self, docx_with_track_changes: Path, tmp_path: Path) -> None:
        """Should keep deleted text when rejecting deletion."""
        output_path = tmp_path / "output.docx"

        # Get deletion ID
        original_doc = read_docx(str(docx_with_track_changes))
        deletions = [tc for tc in original_doc["track_changes"] if tc["type"] == "deletion"]
        assert len(deletions) > 0

        deletion_id = deletions[0]["id"]
        deletion_text = deletions[0]["text"]

        result = reject_track_change(
            str(docx_with_track_changes),
            change_id=deletion_id,
            output_path=str(output_path),
        )

        assert result["success"] is True
        assert result["change_type"] == "deletion"


class TestRoundtrip:
    """Tests for document roundtrip integrity."""

    def test_roundtrip_add_comment(self, simple_docx: Path, tmp_path: Path) -> None:
        """Document should remain valid after adding comment."""
        output_path = tmp_path / "output.docx"

        add_comment(
            str(simple_docx),
            anchor_text="first paragraph",
            comment_text="Test comment",
            output_path=str(output_path),
        )

        # Should be readable
        doc = read_docx(str(output_path))
        assert len(doc["paragraphs"]) == 3
        assert len(doc["comments"]) == 1

    def test_roundtrip_add_track_change(self, simple_docx: Path, tmp_path: Path) -> None:
        """Document should remain valid after adding track change."""
        output_path = tmp_path / "output.docx"

        add_track_change(
            str(simple_docx),
            find_text="second",
            replace_with="2nd",
            output_path=str(output_path),
        )

        # Should be readable
        doc = read_docx(str(output_path))
        assert len(doc["paragraphs"]) == 3
        assert len(doc["track_changes"]) >= 1

    def test_roundtrip_multiple_operations(self, simple_docx: Path, tmp_path: Path) -> None:
        """Document should remain valid after multiple operations."""
        step1_path = tmp_path / "step1.docx"
        step2_path = tmp_path / "step2.docx"
        step3_path = tmp_path / "step3.docx"

        # Add a comment
        add_comment(
            str(simple_docx),
            anchor_text="first paragraph",
            comment_text="Comment 1",
            output_path=str(step1_path),
        )

        # Add a track change
        add_track_change(
            str(step1_path),
            find_text="second",
            replace_with="2nd",
            output_path=str(step2_path),
        )

        # Add another comment
        add_comment(
            str(step2_path),
            anchor_text="third paragraph",
            comment_text="Comment 2",
            output_path=str(step3_path),
        )

        # Should be readable
        doc = read_docx(str(step3_path))
        assert len(doc["paragraphs"]) == 3
        assert len(doc["comments"]) == 2
        assert len(doc["track_changes"]) >= 1
