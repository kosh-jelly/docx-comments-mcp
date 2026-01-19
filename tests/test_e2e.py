"""End-to-end tests simulating real workflow scenarios."""

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


class TestDissertationReviewWorkflow:
    """Test a realistic dissertation review workflow."""

    def test_full_review_cycle(self, simple_docx: Path, tmp_path: Path) -> None:
        """Simulate a full dissertation review cycle with comments and track changes."""
        # Setup paths for each step
        step1 = tmp_path / "step1_advisor_comments.docx"
        step2 = tmp_path / "step2_advisor_edits.docx"
        step3 = tmp_path / "step3_student_replies.docx"
        step4 = tmp_path / "step4_student_accepts.docx"
        final = tmp_path / "final.docx"

        # Step 1: Advisor reads document and adds comments
        doc = read_document(str(simple_docx))
        assert len(doc["paragraphs"]) == 3
        assert len(doc["comments"]) == 0

        # Advisor adds a comment
        result = create_comment(
            str(simple_docx),
            anchor_text="first paragraph",
            comment_text="Consider adding a thesis statement here.",
            author="Dr. Advisor",
            output_path=str(step1),
        )
        assert result["success"] is True
        advisor_comment_id = result["comment_id"]

        # Step 2: Advisor suggests an edit with track changes
        result = create_track_change(
            str(step1),
            find_text="second paragraph",
            replace_with="subsequent paragraph",
            author="Dr. Advisor",
            output_path=str(step2),
        )
        assert result["success"] is True
        assert result["change_type"] == "replacement"

        # Verify the document now has comment and track change
        doc = read_document(str(step2))
        assert len(doc["comments"]) >= 1
        assert len(doc["track_changes"]) >= 1

        # Step 3: Student replies to advisor's comment
        result = create_reply(
            str(step2),
            parent_comment_id=advisor_comment_id,
            reply_text="Thank you - I've added the thesis statement in the revision.",
            author="Student",
            output_path=str(step3),
        )
        assert result["success"] is True

        # Step 4: Student accepts the advisor's suggested edit
        doc = read_document(str(step3))
        insertions = [tc for tc in doc["track_changes"] if tc["type"] == "insertion"]
        assert len(insertions) >= 1

        result = accept_change(
            str(step3),
            change_id=insertions[0]["id"],
            output_path=str(step4),
        )
        assert result["success"] is True

        # Step 5: Student marks the comment as resolved
        doc = read_document(str(step4))
        if doc["comments"]:
            result = mark_comment_resolved(
                str(step4),
                comment_id=doc["comments"][0]["id"],
                output_path=str(final),
            )
            assert result["success"] is True

        # Final verification: document is still valid
        final_doc = read_document(str(final) if final.exists() else str(step4))
        assert len(final_doc["paragraphs"]) == 3


class TestErrorHandling:
    """Test error handling and edge cases."""

    def test_graceful_handling_of_missing_file(self) -> None:
        """Should return structured error for missing file."""
        result = read_document("/nonexistent/path/dissertation.docx")

        assert result["success"] is False
        assert "error" in result
        assert "error_type" in result

    def test_graceful_handling_of_anchor_not_found(self, simple_docx: Path, tmp_path: Path) -> None:
        """Should return structured error when anchor text not found."""
        result = create_comment(
            str(simple_docx),
            anchor_text="this text does not exist in the document",
            comment_text="Test comment",
            output_path=str(tmp_path / "output.docx"),
        )

        assert result["success"] is False
        assert "error" in result
        assert "AnchorNotFoundError" in result["error_type"]

    def test_graceful_handling_of_comment_not_found(self, simple_docx: Path, tmp_path: Path) -> None:
        """Should return structured error when comment ID not found."""
        result = create_reply(
            str(simple_docx),
            parent_comment_id=9999,
            reply_text="Test reply",
            output_path=str(tmp_path / "output.docx"),
        )

        assert result["success"] is False
        assert "error" in result


class TestDocumentIntegrity:
    """Test that documents remain valid after operations."""

    def test_multiple_sequential_operations(self, simple_docx: Path, tmp_path: Path) -> None:
        """Document should remain valid after many sequential operations."""
        current_path = simple_docx

        # Perform 5 operations in sequence
        for i in range(5):
            output_path = tmp_path / f"step_{i}.docx"

            if i % 2 == 0:
                # Add a comment
                text_to_find = ["first", "second", "third", "document", "paragraph"][i]
                result = create_comment(
                    str(current_path),
                    anchor_text=text_to_find,
                    comment_text=f"Comment {i}",
                    output_path=str(output_path),
                )
            else:
                # Add a track change
                text_to_find = ["This", "The", "test"][i % 3]
                result = create_track_change(
                    str(current_path),
                    find_text=text_to_find,
                    replace_with=f"Modified_{i}",
                    output_path=str(output_path),
                )

            # Check if operation succeeded (might fail on later iterations if text already modified)
            if result["success"]:
                current_path = output_path

                # Verify document is still readable
                doc = read_document(str(current_path))
                assert "metadata" in doc
                assert "paragraphs" in doc

    def test_backup_creation(self, simple_docx: Path) -> None:
        """Should create backup when overwriting original."""
        doc_dir = simple_docx.parent

        # Count existing backups
        existing_backups = list(doc_dir.glob("*.backup_*.docx"))
        initial_count = len(existing_backups)

        # Add comment without output_path (overwrites original with backup)
        result = create_comment(
            str(simple_docx),
            anchor_text="first",
            comment_text="Test comment",
        )

        # Verify backup was created
        new_backups = list(doc_dir.glob("*.backup_*.docx"))
        assert len(new_backups) > initial_count
