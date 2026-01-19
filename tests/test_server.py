"""Tests for the MCP server tools."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_comments_mcp.server import (
    accept_change,
    create_comment,
    create_reply,
    create_track_change,
    mark_comment_resolved,
    read_document,
    reject_change,
)


class TestReadDocument:
    """Tests for read_document tool."""

    def test_read_simple_document(self, simple_docx: Path) -> None:
        """Should read a simple document."""
        result = read_document(str(simple_docx))

        assert "metadata" in result
        assert "paragraphs" in result
        assert "comments" in result
        assert "track_changes" in result
        assert len(result["paragraphs"]) == 3

    def test_read_nonexistent_file(self) -> None:
        """Should return error for non-existent file."""
        result = read_document("/nonexistent/path/file.docx")

        assert result["success"] is False
        assert "error" in result


class TestCreateComment:
    """Tests for create_comment tool."""

    def test_create_comment_success(self, simple_docx: Path, tmp_path: Path) -> None:
        """Should create a comment successfully."""
        output_path = tmp_path / "output.docx"

        result = create_comment(
            str(simple_docx),
            anchor_text="first paragraph",
            comment_text="Test comment",
            output_path=str(output_path),
        )

        assert result["success"] is True
        assert result["comment_id"] is not None

    def test_create_comment_anchor_not_found(self, simple_docx: Path, tmp_path: Path) -> None:
        """Should return error when anchor not found."""
        output_path = tmp_path / "output.docx"

        result = create_comment(
            str(simple_docx),
            anchor_text="nonexistent text",
            comment_text="Test comment",
            output_path=str(output_path),
        )

        assert result["success"] is False
        assert "not found" in result["error"].lower()


class TestCreateReply:
    """Tests for create_reply tool."""

    def test_create_reply_success(self, docx_with_comments: Path, tmp_path: Path) -> None:
        """Should create a reply successfully."""
        output_path = tmp_path / "output.docx"

        # Get a comment ID first
        doc = read_document(str(docx_with_comments))
        comment_id = doc["comments"][0]["id"]

        result = create_reply(
            str(docx_with_comments),
            parent_comment_id=comment_id,
            reply_text="Test reply",
            output_path=str(output_path),
        )

        assert result["success"] is True
        assert result["parent_comment_id"] == comment_id

    def test_create_reply_comment_not_found(self, docx_with_comments: Path, tmp_path: Path) -> None:
        """Should return error when parent comment not found."""
        output_path = tmp_path / "output.docx"

        result = create_reply(
            str(docx_with_comments),
            parent_comment_id=9999,
            reply_text="Test reply",
            output_path=str(output_path),
        )

        assert result["success"] is False
        assert "not found" in result["error"].lower()


class TestCreateTrackChange:
    """Tests for create_track_change tool."""

    def test_create_track_change_success(self, simple_docx: Path, tmp_path: Path) -> None:
        """Should create a track change successfully."""
        output_path = tmp_path / "output.docx"

        result = create_track_change(
            str(simple_docx),
            find_text="first",
            replace_with="primary",
            output_path=str(output_path),
        )

        assert result["success"] is True
        assert result["change_type"] == "replacement"


class TestMarkCommentResolved:
    """Tests for mark_comment_resolved tool."""

    def test_mark_resolved_success(self, docx_with_comments: Path, tmp_path: Path) -> None:
        """Should mark a comment as resolved."""
        output_path = tmp_path / "output.docx"

        # Get a comment ID first
        doc = read_document(str(docx_with_comments))
        comment_id = doc["comments"][0]["id"]

        result = mark_comment_resolved(
            str(docx_with_comments),
            comment_id=comment_id,
            output_path=str(output_path),
        )

        assert result["success"] is True


class TestAcceptChange:
    """Tests for accept_change tool."""

    def test_accept_change_success(self, docx_with_track_changes: Path, tmp_path: Path) -> None:
        """Should accept a track change."""
        output_path = tmp_path / "output.docx"

        # Get a change ID first
        doc = read_document(str(docx_with_track_changes))
        change_id = doc["track_changes"][0]["id"]

        result = accept_change(
            str(docx_with_track_changes),
            change_id=change_id,
            output_path=str(output_path),
        )

        assert result["success"] is True


class TestRejectChange:
    """Tests for reject_change tool."""

    def test_reject_change_success(self, docx_with_track_changes: Path, tmp_path: Path) -> None:
        """Should reject a track change."""
        output_path = tmp_path / "output.docx"

        # Get a change ID first
        doc = read_document(str(docx_with_track_changes))
        change_id = doc["track_changes"][0]["id"]

        result = reject_change(
            str(docx_with_track_changes),
            change_id=change_id,
            output_path=str(output_path),
        )

        assert result["success"] is True
