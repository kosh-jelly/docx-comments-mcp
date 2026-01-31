"""MCP server for Word document comments and track changes."""

from __future__ import annotations

import json
from typing import Any

from mcp.server.fastmcp import FastMCP

from .reader import get_paragraph_range_docx, read_docx, search_docx
from .writer import (
    AnchorAmbiguousError,
    AnchorNotFoundError,
    CommentNotFoundError,
    DocxWriteError,
    TrackChangeNotFoundError,
    accept_track_change,
    add_comment,
    add_reply,
    add_track_change,
    reject_track_change,
    resolve_comment,
)

# Create the MCP server
mcp = FastMCP(name="docx-comments-mcp")


def _format_error(error: Exception) -> dict[str, Any]:
    """Format an exception as a structured error response."""
    return {
        "success": False,
        "error": str(error),
        "error_type": type(error).__name__,
    }


@mcp.tool()
def read_document(
    path: str,
    include_text: bool = True,
    include_comments: bool = True,
    include_track_changes: bool = True,
) -> dict[str, Any]:
    """Read a Word document and extract content, comments, and track changes.

    Args:
        path: Path to the .docx file
        include_text: Include full document text (default: True)
        include_comments: Include comments with anchors (default: True)
        include_track_changes: Include insertions/deletions (default: True)

    Returns:
        Dictionary containing:
        - metadata: Document metadata (path, author, created, modified, word_count)
        - paragraphs: List of paragraphs with index, text, and style
        - comments: List of comments with id, author, date, text, anchor_text, anchor_paragraph, resolved (boolean), and replies
        - track_changes: List of track changes with id, type (insertion/deletion), author, date, text, paragraph
    """
    try:
        return read_docx(
            path,
            include_text=include_text,
            include_comments=include_comments,
            include_track_changes=include_track_changes,
        )
    except FileNotFoundError as e:
        return _format_error(e)
    except Exception as e:
        return _format_error(e)


@mcp.tool()
def search_document(
    path: str,
    query: str,
    case_sensitive: bool = False,
    context_paragraphs: int = 1,
    max_results: int = 20,
    include_annotations: bool = False,
) -> dict[str, Any]:
    """Search for text in a Word document.

    Use this to find specific content without loading the entire document.
    Returns matching paragraphs with surrounding context.

    Args:
        path: Path to the .docx file
        query: Text to search for
        case_sensitive: Match case exactly (default: False)
        context_paragraphs: Paragraphs to include before/after each match (default: 1)
        max_results: Maximum matches to return (default: 20)
        include_annotations: Include comments/track changes on matched paragraphs (default: False)

    Returns:
        Dictionary containing:
        - query: The search query
        - case_sensitive: Whether search was case-sensitive
        - total_matches: Total matches found
        - matches_returned: Number returned (may be limited)
        - matches: List with paragraph_index, paragraph_text, paragraph_style,
          match_start, match_end, context_before, context_after, and optionally
          comments and track_changes
    """
    if not query:
        return {
            "success": False,
            "error": "Query cannot be empty",
            "error_type": "ValueError",
        }
    try:
        return search_docx(
            path=path,
            query=query,
            case_sensitive=case_sensitive,
            context_paragraphs=context_paragraphs,
            max_results=max_results,
            include_annotations=include_annotations,
        )
    except FileNotFoundError as e:
        return _format_error(e)
    except Exception as e:
        return _format_error(e)


@mcp.tool()
def get_paragraph_range(
    path: str,
    start_index: int,
    end_index: int,
    include_annotations: bool = False,
) -> dict[str, Any]:
    """Get a specific range of paragraphs from a Word document.

    Use after search_document to get more context around matches,
    or to read a specific section without loading the full document.

    Args:
        path: Path to the .docx file
        start_index: First paragraph index (0-based, inclusive)
        end_index: Last paragraph index (0-based, inclusive)
        include_annotations: Include comments/track changes in range (default: False)

    Returns:
        Dictionary containing:
        - start_index: Actual start (may be clamped)
        - end_index: Actual end (may be clamped)
        - total_paragraphs: Total paragraphs in document
        - paragraphs: List with index, text, style
        - comments: Comments in range (if include_annotations=True)
        - track_changes: Track changes in range (if include_annotations=True)
    """
    try:
        return get_paragraph_range_docx(
            path=path,
            start_index=start_index,
            end_index=end_index,
            include_annotations=include_annotations,
        )
    except FileNotFoundError as e:
        return _format_error(e)
    except Exception as e:
        return _format_error(e)


@mcp.tool()
def create_comment(
    path: str,
    anchor_text: str,
    comment_text: str,
    author: str = "Claude",
    output_path: str | None = None,
) -> dict[str, Any]:
    """Add a comment anchored to specific text in a Word document.

    Args:
        path: Path to the .docx file
        anchor_text: Text to anchor the comment to (must exist and be unique in document)
        comment_text: The comment content
        author: Comment author name (default: "Claude")
        output_path: Save to new file; if omitted, creates timestamped backup and overwrites original

    Returns:
        Dictionary containing:
        - success: True if successful
        - comment_id: ID of the created comment
        - anchored_to: The text the comment is anchored to
        - paragraph: Index of the paragraph containing the anchor
        - output_path: Path where the file was saved

    Errors:
        - If anchor text is not found: {"success": false, "error": "Anchor text not found in document"}
        - If anchor text appears multiple times: {"success": false, "error": "Anchor text appears N times; provide more context for unique match"}
    """
    try:
        return add_comment(
            path=path,
            anchor_text=anchor_text,
            comment_text=comment_text,
            author=author,
            output_path=output_path,
        )
    except AnchorNotFoundError as e:
        return _format_error(e)
    except AnchorAmbiguousError as e:
        return _format_error(e)
    except DocxWriteError as e:
        return _format_error(e)
    except Exception as e:
        return _format_error(e)


@mcp.tool()
def create_reply(
    path: str,
    parent_comment_id: int,
    reply_text: str,
    author: str = "Claude",
    output_path: str | None = None,
) -> dict[str, Any]:
    """Add a reply to an existing comment in a Word document.

    Args:
        path: Path to the .docx file
        parent_comment_id: ID of the comment to reply to
        reply_text: The reply content
        author: Reply author name (default: "Claude")
        output_path: Save to new file; if omitted, creates timestamped backup and overwrites original

    Returns:
        Dictionary containing:
        - success: True if successful
        - reply_id: ID of the created reply
        - parent_comment_id: ID of the parent comment
        - output_path: Path where the file was saved
    """
    try:
        return add_reply(
            path=path,
            parent_comment_id=parent_comment_id,
            reply_text=reply_text,
            author=author,
            output_path=output_path,
        )
    except CommentNotFoundError as e:
        return _format_error(e)
    except DocxWriteError as e:
        return _format_error(e)
    except Exception as e:
        return _format_error(e)


@mcp.tool()
def create_track_change(
    path: str,
    find_text: str,
    replace_with: str,
    author: str = "Claude",
    output_path: str | None = None,
) -> dict[str, Any]:
    """Make an edit with track changes enabled (insertion, deletion, or replacement).

    Args:
        path: Path to the .docx file
        find_text: Text to find and modify (must exist and be unique)
        replace_with: Replacement text (use empty string for deletion-only)
        author: Change author name (default: "Claude")
        output_path: Save to new file; if omitted, creates timestamped backup and overwrites original

    Returns:
        Dictionary containing:
        - success: True if successful
        - change_type: "replacement", "deletion", or "insertion"
        - original_text: The text that was changed
        - new_text: The replacement text
        - paragraph: Index of the paragraph containing the change
        - output_path: Path where the file was saved
    """
    try:
        return add_track_change(
            path=path,
            find_text=find_text,
            replace_with=replace_with,
            author=author,
            output_path=output_path,
        )
    except AnchorNotFoundError as e:
        return _format_error(e)
    except AnchorAmbiguousError as e:
        return _format_error(e)
    except DocxWriteError as e:
        return _format_error(e)
    except Exception as e:
        return _format_error(e)


@mcp.tool()
def mark_comment_resolved(
    path: str,
    comment_id: int,
    output_path: str | None = None,
) -> dict[str, Any]:
    """Mark a comment as resolved/done.

    Args:
        path: Path to the .docx file
        comment_id: ID of the comment to resolve
        output_path: Save to new file; if omitted, creates timestamped backup and overwrites original

    Returns:
        Dictionary containing:
        - success: True if successful
        - comment_id: ID of the resolved comment
        - output_path: Path where the file was saved
    """
    try:
        return resolve_comment(
            path=path,
            comment_id=comment_id,
            output_path=output_path,
        )
    except CommentNotFoundError as e:
        return _format_error(e)
    except DocxWriteError as e:
        return _format_error(e)
    except Exception as e:
        return _format_error(e)


@mcp.tool()
def accept_change(
    path: str,
    change_id: int,
    output_path: str | None = None,
) -> dict[str, Any]:
    """Accept a tracked change (apply the change permanently).

    For insertions: The inserted text becomes part of the document.
    For deletions: The deleted text is permanently removed.

    Args:
        path: Path to the .docx file
        change_id: ID of the track change to accept
        output_path: Save to new file; if omitted, creates timestamped backup and overwrites original

    Returns:
        Dictionary containing:
        - success: True if successful
        - change_id: ID of the accepted change
        - change_type: "insertion" or "deletion"
        - output_path: Path where the file was saved
    """
    try:
        return accept_track_change(
            path=path,
            change_id=change_id,
            output_path=output_path,
        )
    except TrackChangeNotFoundError as e:
        return _format_error(e)
    except DocxWriteError as e:
        return _format_error(e)
    except Exception as e:
        return _format_error(e)


@mcp.tool()
def reject_change(
    path: str,
    change_id: int,
    output_path: str | None = None,
) -> dict[str, Any]:
    """Reject a tracked change (undo the change).

    For insertions: The inserted text is removed.
    For deletions: The deleted text is restored.

    Args:
        path: Path to the .docx file
        change_id: ID of the track change to reject
        output_path: Save to new file; if omitted, creates timestamped backup and overwrites original

    Returns:
        Dictionary containing:
        - success: True if successful
        - change_id: ID of the rejected change
        - change_type: "insertion" or "deletion"
        - output_path: Path where the file was saved
    """
    try:
        return reject_track_change(
            path=path,
            change_id=change_id,
            output_path=output_path,
        )
    except TrackChangeNotFoundError as e:
        return _format_error(e)
    except DocxWriteError as e:
        return _format_error(e)
    except Exception as e:
        return _format_error(e)


def main():
    """Run the MCP server."""
    mcp.run()


if __name__ == "__main__":
    main()
