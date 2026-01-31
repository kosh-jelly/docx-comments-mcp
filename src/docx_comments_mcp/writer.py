"""Write operations for Word documents (comments, track changes)."""

from __future__ import annotations

import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from lxml import etree

from .xml_helpers import (
    NAMESPACES,
    create_element,
    get_max_id,
    get_text_content,
    iter_paragraphs,
    normalize_typography,
    qn,
    serialize_xml,
)


class DocxWriteError(Exception):
    """Error during document write operations."""

    pass


class AnchorNotFoundError(DocxWriteError):
    """Anchor text was not found in the document."""

    pass


class AnchorAmbiguousError(DocxWriteError):
    """Anchor text appears multiple times in the document."""

    pass


class CommentNotFoundError(DocxWriteError):
    """Comment with the given ID was not found."""

    pass


class TrackChangeNotFoundError(DocxWriteError):
    """Track change with the given ID was not found."""

    pass


def _create_backup(path: Path) -> Path:
    """Create a timestamped backup of the file."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = path.with_suffix(f".backup_{timestamp}.docx")
    shutil.copy2(path, backup_path)
    return backup_path


def _get_output_path(input_path: Path, output_path: str | None) -> Path:
    """Determine the output path, creating backup if needed."""
    if output_path:
        return Path(output_path)
    else:
        # Create backup before overwriting
        _create_backup(input_path)
        return input_path


def _get_current_datetime() -> str:
    """Get current datetime in OOXML format."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _read_all_zip_contents(zip_path: Path) -> dict[str, bytes]:
    """Read all contents from a zip file into memory."""
    contents = {}
    with zipfile.ZipFile(zip_path, "r") as zf:
        for item in zf.namelist():
            contents[item] = zf.read(item)
    return contents


def _write_docx_with_modifications(
    input_path: Path,
    output_path: Path,
    modifications: dict[str, bytes],
    additions: dict[str, bytes] | None = None,
    skip_items: set[str] | None = None,
) -> None:
    """Write a docx file with specified modifications.

    Args:
        input_path: Source docx file
        output_path: Destination docx file
        modifications: Dict of {item_name: new_content} for items to modify or add/replace
        additions: Dict of {item_name: content} for items to add (will also replace existing)
        skip_items: Set of item names to skip entirely (don't copy from original)
    """
    # Read all contents first to avoid issues with same input/output
    original_contents = _read_all_zip_contents(input_path)

    if skip_items is None:
        skip_items = set()
    if additions is None:
        additions = {}

    # Write to a temp file first, then move (atomic on same filesystem)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp_path = Path(tmp.name)

    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
            # Write original/modified items
            for item, content in original_contents.items():
                # Skip items we don't want to copy
                if item in skip_items:
                    continue
                if item in modifications:
                    zf_out.writestr(item, modifications[item])
                else:
                    zf_out.writestr(item, content)

            # Write additions (new items or replacements for skipped items)
            for item, content in additions.items():
                # Add if not in original, or if it was skipped from original
                if item not in original_contents or item in skip_items:
                    zf_out.writestr(item, content)

        # Move temp file to final destination
        shutil.move(str(tmp_path), str(output_path))
    except Exception:
        # Clean up temp file on error
        tmp_path.unlink(missing_ok=True)
        raise


def _find_anchor_in_document(
    document: etree._Element, anchor_text: str
) -> list[tuple[int, etree._Element, int, int]]:
    """Find all occurrences of anchor text in the document.

    Returns list of (paragraph_index, paragraph_element, start_run_idx, end_run_idx).
    """
    occurrences = []

    for para_idx, para in iter_paragraphs(document):
        para_text = get_text_content(para)

        # Check if anchor text is in this paragraph (normalize for smart quotes)
        norm_para = normalize_typography(para_text)
        norm_anchor = normalize_typography(anchor_text)
        if norm_anchor in norm_para:
            # Find the runs containing this text
            runs = list(para.iter(qn("w:r")))

            # Build character-to-run mapping
            char_positions: list[tuple[int, etree._Element]] = []  # (run_idx, run_elem)
            current_pos = 0

            for run_idx, run in enumerate(runs):
                for t_elem in run.iter(qn("w:t")):
                    if t_elem.text:
                        for _ in t_elem.text:
                            char_positions.append((run_idx, run))

            # Find anchor position in paragraph (using normalized text)
            start_pos = norm_para.find(norm_anchor)
            while start_pos != -1:
                end_pos = start_pos + len(anchor_text)

                if char_positions and start_pos < len(char_positions) and end_pos <= len(char_positions):
                    start_run_idx = char_positions[start_pos][0]
                    end_run_idx = char_positions[end_pos - 1][0]
                    occurrences.append((para_idx, para, start_run_idx, end_run_idx))

                start_pos = norm_para.find(norm_anchor, start_pos + 1)

    return occurrences


def _ensure_comments_xml_exists(zip_path: Path) -> etree._Element:
    """Ensure comments.xml exists and return its root element."""
    with zipfile.ZipFile(zip_path, "r") as zf:
        if "word/comments.xml" in zf.namelist():
            with zf.open("word/comments.xml") as f:
                return etree.parse(f).getroot()

    # Create new comments.xml
    nsmap = {
        "w": NAMESPACES["w"],
        "w14": NAMESPACES["w14"],
        "w15": NAMESPACES["w15"],
    }
    return etree.Element(qn("w:comments"), nsmap=nsmap)


def _ensure_comments_extended_exists(zip_path: Path) -> etree._Element:
    """Ensure commentsExtended.xml exists and return its root element."""
    with zipfile.ZipFile(zip_path, "r") as zf:
        if "word/commentsExtended.xml" in zf.namelist():
            with zf.open("word/commentsExtended.xml") as f:
                return etree.parse(f).getroot()

    # Create new commentsExtended.xml
    nsmap = {"w15": NAMESPACES["w15"]}
    return etree.Element(qn("w15:commentsEx"), nsmap=nsmap)


def _update_content_types_for_comments(zip_path: Path) -> bytes:
    """Update [Content_Types].xml to include comments part if not present."""
    with zipfile.ZipFile(zip_path, "r") as zf:
        with zf.open("[Content_Types].xml") as f:
            content_types = etree.parse(f).getroot()

    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"

    # Check if comments override exists
    has_comments = any(
        elem.get("PartName") == "/word/comments.xml"
        for elem in content_types.iter(f"{{{ct_ns}}}Override")
    )

    if not has_comments:
        override = etree.SubElement(
            content_types,
            f"{{{ct_ns}}}Override",
            PartName="/word/comments.xml",
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        )

    return serialize_xml(content_types)


def _update_document_rels_for_comments(zip_path: Path) -> bytes:
    """Update word/_rels/document.xml.rels to include comments relationship."""
    pr_ns = "http://schemas.openxmlformats.org/package/2006/relationships"

    with zipfile.ZipFile(zip_path, "r") as zf:
        if "word/_rels/document.xml.rels" in zf.namelist():
            with zf.open("word/_rels/document.xml.rels") as f:
                rels = etree.parse(f).getroot()
        else:
            rels = etree.Element(f"{{{pr_ns}}}Relationships", nsmap={None: pr_ns})

    # Check if comments relationship exists
    comments_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    has_comments_rel = any(
        elem.get("Type") == comments_type for elem in rels.iter(f"{{{pr_ns}}}Relationship")
    )

    if not has_comments_rel:
        # Find next available rId
        existing_ids = [
            int(elem.get("Id", "rId0")[3:])
            for elem in rels.iter(f"{{{pr_ns}}}Relationship")
            if elem.get("Id", "").startswith("rId")
        ]
        next_id = max(existing_ids, default=0) + 1

        etree.SubElement(
            rels,
            f"{{{pr_ns}}}Relationship",
            Id=f"rId{next_id}",
            Type=comments_type,
            Target="comments.xml",
        )

    return serialize_xml(rels)


def _insert_comment_markers(
    para: etree._Element,
    anchor_text: str,
    comment_id: int,
) -> bool:
    """Insert comment range markers around anchor text in a paragraph.

    Returns True if successful, False otherwise.
    """
    # Get all text content and build mapping
    runs = list(para.iter(qn("w:r")))
    full_text = get_text_content(para)

    if anchor_text not in full_text:
        return False

    # Find position of anchor in full text
    anchor_start = full_text.find(anchor_text)
    anchor_end = anchor_start + len(anchor_text)

    # Build character position to element mapping
    char_to_elem: list[tuple[etree._Element, etree._Element, int]] = []  # (run, t_elem, char_idx_in_t)

    for run in runs:
        for t_elem in run.iter(qn("w:t")):
            if t_elem.text:
                for i, _ in enumerate(t_elem.text):
                    char_to_elem.append((run, t_elem, i))

    if anchor_end > len(char_to_elem):
        return False

    # Get the runs at start and end
    start_run, start_t, start_idx = char_to_elem[anchor_start]
    end_run, end_t, end_idx = char_to_elem[anchor_end - 1]

    # Simple case: anchor is within a single run's single text element
    # We'll use a simplified approach: insert markers before/after the runs

    # Find run index for start and end
    run_list = list(para)
    start_run_idx = None
    end_run_idx = None

    for i, elem in enumerate(run_list):
        if elem is start_run or (elem.tag == qn("w:r") and start_run in list(elem.iter())):
            if start_run_idx is None:
                start_run_idx = i
        if elem is end_run or (elem.tag == qn("w:r") and end_run in list(elem.iter())):
            end_run_idx = i

    if start_run_idx is None:
        # Find the actual run element
        for i, child in enumerate(para):
            if child.tag == qn("w:r"):
                if get_text_content(child) and anchor_text[:1] in get_text_content(child):
                    start_run_idx = i
                    break

    if end_run_idx is None:
        end_run_idx = start_run_idx

    if start_run_idx is None:
        return False

    # Create marker elements
    range_start = create_element("w:commentRangeStart", {qn("w:id"): str(comment_id)})
    range_end = create_element("w:commentRangeEnd", {qn("w:id"): str(comment_id)})

    # Create comment reference run
    ref_run = create_element("w:r")
    ref = create_element("w:commentReference", {qn("w:id"): str(comment_id)})
    ref_run.append(ref)

    # Insert markers
    para.insert(start_run_idx, range_start)
    # Adjust for inserted element
    para.insert(end_run_idx + 2, range_end)
    para.insert(end_run_idx + 3, ref_run)

    return True


def add_comment(
    path: str,
    anchor_text: str,
    comment_text: str,
    author: str = "Claude",
    output_path: str | None = None,
) -> dict[str, Any]:
    """Add a comment anchored to specific text in the document.

    Args:
        path: Path to the .docx file
        anchor_text: Text to anchor the comment to (must exist in document)
        comment_text: The comment content
        author: Comment author name
        output_path: Save to new file; if omitted, creates backup and overwrites original

    Returns:
        Dictionary with success status, comment_id, anchored_to, paragraph, output_path

    Raises:
        AnchorNotFoundError: If anchor text is not found
        AnchorAmbiguousError: If anchor text appears multiple times
    """
    input_path = Path(path)
    final_output_path = _get_output_path(input_path, output_path)

    # Read document
    with zipfile.ZipFile(input_path, "r") as zf:
        with zf.open("word/document.xml") as f:
            document = etree.parse(f).getroot()

    # Find anchor text
    occurrences = _find_anchor_in_document(document, anchor_text)

    if not occurrences:
        raise AnchorNotFoundError(f"Anchor text not found in document: {anchor_text}")

    if len(occurrences) > 1:
        raise AnchorAmbiguousError(
            f"Anchor text appears {len(occurrences)} times; provide more context for unique match"
        )

    para_idx, para, _, _ = occurrences[0]

    # Get or create comments.xml
    comments_root = _ensure_comments_xml_exists(input_path)

    # Find next comment ID
    max_id = get_max_id(document, "w:id")
    comments_max_id = get_max_id(comments_root, "w:id")
    comment_id = max(max_id, comments_max_id) + 1

    # Create comment element
    comment_date = _get_current_datetime()
    comment_elem = create_element(
        "w:comment",
        {
            qn("w:id"): str(comment_id),
            qn("w:author"): author,
            qn("w:date"): comment_date,
        },
    )

    # Add paragraph with text to comment (include paraId for threading/resolve support)
    para_id = f"para{comment_id:08X}"
    comment_para = create_element("w:p", {qn("w14:paraId"): para_id})
    comment_run = create_element("w:r")
    comment_t = create_element("w:t")
    comment_t.text = comment_text
    comment_run.append(comment_t)
    comment_para.append(comment_run)
    comment_elem.append(comment_para)

    comments_root.append(comment_elem)

    # Insert comment markers in document
    if not _insert_comment_markers(para, anchor_text, comment_id):
        raise DocxWriteError("Failed to insert comment markers")

    # Update content types and relationships
    content_types_xml = _update_content_types_for_comments(input_path)
    doc_rels_xml = _update_document_rels_for_comments(input_path)

    # Write output file using helper
    _write_docx_with_modifications(
        input_path=input_path,
        output_path=final_output_path,
        modifications={
            "word/document.xml": serialize_xml(document),
            "[Content_Types].xml": content_types_xml,
            "word/_rels/document.xml.rels": doc_rels_xml,
        },
        additions={
            "word/comments.xml": serialize_xml(comments_root),
        },
        skip_items={"word/comments.xml"},  # Skip original if exists, we add fresh
    )

    return {
        "success": True,
        "comment_id": comment_id,
        "anchored_to": anchor_text,
        "paragraph": para_idx,
        "output_path": str(final_output_path),
    }


def add_reply(
    path: str,
    parent_comment_id: int,
    reply_text: str,
    author: str = "Claude",
    output_path: str | None = None,
) -> dict[str, Any]:
    """Add a reply to an existing comment.

    Args:
        path: Path to the .docx file
        parent_comment_id: ID of comment to reply to
        reply_text: The reply content
        author: Reply author name
        output_path: Save to new file; if omitted, creates backup and overwrites original

    Returns:
        Dictionary with success status, reply_id, parent_comment_id, output_path

    Raises:
        CommentNotFoundError: If parent comment is not found
    """
    input_path = Path(path)
    final_output_path = _get_output_path(input_path, output_path)

    # Read document and comments
    with zipfile.ZipFile(input_path, "r") as zf:
        with zf.open("word/document.xml") as f:
            document = etree.parse(f).getroot()

        if "word/comments.xml" not in zf.namelist():
            raise CommentNotFoundError(f"No comments found in document")

        with zf.open("word/comments.xml") as f:
            comments_root = etree.parse(f).getroot()

    # Find parent comment
    parent_comment = None
    parent_para_id = None

    for comment in comments_root.iter(qn("w:comment")):
        if comment.get(qn("w:id")) == str(parent_comment_id):
            parent_comment = comment
            # Get paragraph ID for threading
            first_para = comment.find(qn("w:p"))
            if first_para is not None:
                parent_para_id = first_para.get(qn("w14:paraId"))
            break

    if parent_comment is None:
        raise CommentNotFoundError(f"Comment with ID {parent_comment_id} not found")

    # Find next comment ID
    max_id = get_max_id(document, "w:id")
    comments_max_id = get_max_id(comments_root, "w:id")
    reply_id = max(max_id, comments_max_id) + 1

    # Create reply paragraph ID
    reply_para_id = f"para{reply_id:08X}"

    # Create reply comment element
    reply_date = _get_current_datetime()
    reply_elem = create_element(
        "w:comment",
        {
            qn("w:id"): str(reply_id),
            qn("w:author"): author,
            qn("w:date"): reply_date,
        },
    )

    # Add paragraph with text
    reply_para = create_element("w:p", {qn("w14:paraId"): reply_para_id})
    reply_run = create_element("w:r")
    reply_t = create_element("w:t")
    reply_t.text = reply_text
    reply_run.append(reply_t)
    reply_para.append(reply_run)
    reply_elem.append(reply_para)

    comments_root.append(reply_elem)

    # Update commentsExtended.xml for threading
    comments_extended = _ensure_comments_extended_exists(input_path)

    # Add commentEx element for threading
    if parent_para_id:
        comment_ex = create_element(
            "w15:commentEx",
            {
                qn("w15:paraId"): reply_para_id,
                qn("w15:paraIdParent"): parent_para_id,
                qn("w15:done"): "0",
            },
        )
        comments_extended.append(comment_ex)

    # Update content types for commentsExtended if needed
    content_types_xml = _update_content_types_for_comments(input_path)

    # Add commentsExtended to content types if not present
    ct_root = etree.fromstring(content_types_xml)
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    has_ext = any(
        elem.get("PartName") == "/word/commentsExtended.xml"
        for elem in ct_root.iter(f"{{{ct_ns}}}Override")
    )
    if not has_ext:
        etree.SubElement(
            ct_root,
            f"{{{ct_ns}}}Override",
            PartName="/word/commentsExtended.xml",
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
        )
    content_types_xml = serialize_xml(ct_root)

    # Write output file using helper
    _write_docx_with_modifications(
        input_path=input_path,
        output_path=final_output_path,
        modifications={
            "word/comments.xml": serialize_xml(comments_root),
            "[Content_Types].xml": content_types_xml,
        },
        additions={
            "word/commentsExtended.xml": serialize_xml(comments_extended),
        },
        skip_items={"word/commentsExtended.xml"},
    )

    return {
        "success": True,
        "reply_id": reply_id,
        "parent_comment_id": parent_comment_id,
        "output_path": str(final_output_path),
    }


def add_track_change(
    path: str,
    find_text: str,
    replace_with: str,
    author: str = "Claude",
    output_path: str | None = None,
) -> dict[str, Any]:
    """Make an edit with track changes enabled.

    Args:
        path: Path to the .docx file
        find_text: Text to find and modify
        replace_with: Replacement text (empty string for deletion)
        author: Change author name
        output_path: Save to new file; if omitted, creates backup and overwrites original

    Returns:
        Dictionary with success status, change_type, original_text, new_text, paragraph, output_path

    Raises:
        AnchorNotFoundError: If find_text is not found
        AnchorAmbiguousError: If find_text appears multiple times
    """
    input_path = Path(path)
    final_output_path = _get_output_path(input_path, output_path)

    # Read document
    with zipfile.ZipFile(input_path, "r") as zf:
        with zf.open("word/document.xml") as f:
            document = etree.parse(f).getroot()

    # Find text in document
    occurrences = _find_anchor_in_document(document, find_text)

    if not occurrences:
        raise AnchorNotFoundError(f"Text not found in document: {find_text}")

    if len(occurrences) > 1:
        raise AnchorAmbiguousError(
            f"Text appears {len(occurrences)} times; provide more context for unique match"
        )

    para_idx, para, start_run_idx, end_run_idx = occurrences[0]

    # Get next change ID
    max_id = get_max_id(document, "w:id")
    change_date = _get_current_datetime()

    # Determine change type
    if not replace_with:
        change_type = "deletion"
    elif not find_text:
        change_type = "insertion"
    else:
        change_type = "replacement"

    # For simplicity, we'll implement a basic approach:
    # 1. Find the run containing the text
    # 2. Create deletion element for original text
    # 3. Create insertion element for new text (if any)

    # Find the run with the text
    norm_find = normalize_typography(find_text)
    target_run = None
    for run in para.iter(qn("w:r")):
        run_text = get_text_content(run)
        if norm_find in normalize_typography(run_text):
            target_run = run
            break

    if target_run is None:
        raise DocxWriteError("Could not locate target run for modification")

    # Get the t element with the text
    target_t = None
    for t_elem in target_run.iter(qn("w:t")):
        if t_elem.text and norm_find in normalize_typography(t_elem.text):
            target_t = t_elem
            break

    if target_t is None:
        raise DocxWriteError("Could not locate text element for modification")

    # Split the text (find position via normalized text, split actual text)
    full_text = target_t.text
    split_pos = normalize_typography(full_text).find(norm_find)
    before = full_text[:split_pos]
    after = full_text[split_pos + len(find_text):]

    # Get parent of target_run to insert new elements
    run_parent = target_run.getparent()
    run_idx = list(run_parent).index(target_run)

    # Clear original text and rebuild
    target_t.text = before if before else None

    # Create elements to insert after target_run
    elements_to_insert = []

    # Deletion element
    if find_text:
        del_id = max_id + 1
        del_elem = create_element(
            "w:del",
            {
                qn("w:id"): str(del_id),
                qn("w:author"): author,
                qn("w:date"): change_date,
            },
        )
        del_run = create_element("w:r")
        del_text = create_element("w:delText")
        del_text.text = find_text
        del_run.append(del_text)
        del_elem.append(del_run)
        elements_to_insert.append(del_elem)
        max_id = del_id

    # Insertion element
    if replace_with:
        ins_id = max_id + 1
        ins_elem = create_element(
            "w:ins",
            {
                qn("w:id"): str(ins_id),
                qn("w:author"): author,
                qn("w:date"): change_date,
            },
        )
        ins_run = create_element("w:r")
        ins_text = create_element("w:t")
        ins_text.text = replace_with
        ins_run.append(ins_text)
        ins_elem.append(ins_run)
        elements_to_insert.append(ins_elem)

    # After text run
    if after:
        after_run = create_element("w:r")
        after_t = create_element("w:t")
        after_t.text = after
        after_run.append(after_t)
        elements_to_insert.append(after_run)

    # Insert elements
    for i, elem in enumerate(elements_to_insert):
        run_parent.insert(run_idx + 1 + i, elem)

    # Remove original run if it's now empty
    if not before and target_t.text is None:
        # Check if run has any remaining content
        remaining_text = get_text_content(target_run)
        if not remaining_text:
            run_parent.remove(target_run)

    # Write output file using helper
    _write_docx_with_modifications(
        input_path=input_path,
        output_path=final_output_path,
        modifications={
            "word/document.xml": serialize_xml(document),
        },
    )

    return {
        "success": True,
        "change_type": change_type,
        "original_text": find_text,
        "new_text": replace_with,
        "paragraph": para_idx,
        "output_path": str(final_output_path),
    }


def resolve_comment(
    path: str,
    comment_id: int,
    output_path: str | None = None,
) -> dict[str, Any]:
    """Mark a comment as resolved/done.

    Args:
        path: Path to the .docx file
        comment_id: ID of comment to resolve
        output_path: Save to new file; if omitted, creates backup and overwrites original

    Returns:
        Dictionary with success status and output_path

    Raises:
        CommentNotFoundError: If comment is not found
    """
    input_path = Path(path)
    final_output_path = _get_output_path(input_path, output_path)

    # Check comment exists
    with zipfile.ZipFile(input_path, "r") as zf:
        if "word/comments.xml" not in zf.namelist():
            raise CommentNotFoundError("No comments found in document")

        with zf.open("word/comments.xml") as f:
            comments_root = etree.parse(f).getroot()

    # Find comment
    found = False
    para_id = None
    for comment in comments_root.iter(qn("w:comment")):
        if comment.get(qn("w:id")) == str(comment_id):
            found = True
            first_para = comment.find(qn("w:p"))
            if first_para is not None:
                para_id = first_para.get(qn("w14:paraId"))
            break

    if not found:
        raise CommentNotFoundError(f"Comment with ID {comment_id} not found")

    # Update commentsExtended.xml
    comments_extended = _ensure_comments_extended_exists(input_path)

    # Find or create commentEx for this comment
    found_ex = False
    for comment_ex in comments_extended.iter(qn("w15:commentEx")):
        if comment_ex.get(qn("w15:paraId")) == para_id:
            comment_ex.set(qn("w15:done"), "1")
            found_ex = True
            break

    if not found_ex and para_id:
        # Create new commentEx
        comment_ex = create_element(
            "w15:commentEx",
            {
                qn("w15:paraId"): para_id,
                qn("w15:done"): "1",
            },
        )
        comments_extended.append(comment_ex)

    # Update content types
    content_types_xml = _update_content_types_for_comments(input_path)
    ct_root = etree.fromstring(content_types_xml)
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    has_ext = any(
        elem.get("PartName") == "/word/commentsExtended.xml"
        for elem in ct_root.iter(f"{{{ct_ns}}}Override")
    )
    if not has_ext:
        etree.SubElement(
            ct_root,
            f"{{{ct_ns}}}Override",
            PartName="/word/commentsExtended.xml",
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
        )
    content_types_xml = serialize_xml(ct_root)

    # Write output file using helper
    _write_docx_with_modifications(
        input_path=input_path,
        output_path=final_output_path,
        modifications={
            "[Content_Types].xml": content_types_xml,
        },
        additions={
            "word/commentsExtended.xml": serialize_xml(comments_extended),
        },
        skip_items={"word/commentsExtended.xml"},
    )

    return {
        "success": True,
        "comment_id": comment_id,
        "output_path": str(final_output_path),
    }


def accept_track_change(
    path: str,
    change_id: int,
    output_path: str | None = None,
) -> dict[str, Any]:
    """Accept a tracked change (apply the change permanently).

    Args:
        path: Path to the .docx file
        change_id: ID of the track change
        output_path: Save to new file; if omitted, creates backup and overwrites original

    Returns:
        Dictionary with success status and output_path

    Raises:
        TrackChangeNotFoundError: If track change is not found
    """
    input_path = Path(path)
    final_output_path = _get_output_path(input_path, output_path)

    # Read document
    with zipfile.ZipFile(input_path, "r") as zf:
        with zf.open("word/document.xml") as f:
            document = etree.parse(f).getroot()

    # Find the track change
    change_elem = None
    change_type = None

    # Check insertions
    for ins in document.iter(qn("w:ins")):
        if ins.get(qn("w:id")) == str(change_id):
            change_elem = ins
            change_type = "insertion"
            break

    # Check deletions
    if change_elem is None:
        for del_elem in document.iter(qn("w:del")):
            if del_elem.get(qn("w:id")) == str(change_id):
                change_elem = del_elem
                change_type = "deletion"
                break

    if change_elem is None:
        raise TrackChangeNotFoundError(f"Track change with ID {change_id} not found")

    parent = change_elem.getparent()
    idx = list(parent).index(change_elem)

    if change_type == "insertion":
        # Accept insertion: unwrap the content (remove ins wrapper, keep content)
        children = list(change_elem)
        parent.remove(change_elem)
        for i, child in enumerate(children):
            parent.insert(idx + i, child)

    elif change_type == "deletion":
        # Accept deletion: remove the entire del element and its content
        parent.remove(change_elem)

    # Write output file using helper
    _write_docx_with_modifications(
        input_path=input_path,
        output_path=final_output_path,
        modifications={
            "word/document.xml": serialize_xml(document),
        },
    )

    return {
        "success": True,
        "change_id": change_id,
        "change_type": change_type,
        "output_path": str(final_output_path),
    }


def reject_track_change(
    path: str,
    change_id: int,
    output_path: str | None = None,
) -> dict[str, Any]:
    """Reject a tracked change (undo the change).

    Args:
        path: Path to the .docx file
        change_id: ID of the track change
        output_path: Save to new file; if omitted, creates backup and overwrites original

    Returns:
        Dictionary with success status and output_path

    Raises:
        TrackChangeNotFoundError: If track change is not found
    """
    input_path = Path(path)
    final_output_path = _get_output_path(input_path, output_path)

    # Read document
    with zipfile.ZipFile(input_path, "r") as zf:
        with zf.open("word/document.xml") as f:
            document = etree.parse(f).getroot()

    # Find the track change
    change_elem = None
    change_type = None

    # Check insertions
    for ins in document.iter(qn("w:ins")):
        if ins.get(qn("w:id")) == str(change_id):
            change_elem = ins
            change_type = "insertion"
            break

    # Check deletions
    if change_elem is None:
        for del_elem in document.iter(qn("w:del")):
            if del_elem.get(qn("w:id")) == str(change_id):
                change_elem = del_elem
                change_type = "deletion"
                break

    if change_elem is None:
        raise TrackChangeNotFoundError(f"Track change with ID {change_id} not found")

    parent = change_elem.getparent()
    idx = list(parent).index(change_elem)

    if change_type == "insertion":
        # Reject insertion: remove the entire ins element and its content
        parent.remove(change_elem)

    elif change_type == "deletion":
        # Reject deletion: unwrap the content (keep the deleted text, remove del wrapper)
        # Convert delText to regular text
        for del_text in change_elem.iter(qn("w:delText")):
            del_text.tag = qn("w:t")

        children = list(change_elem)
        parent.remove(change_elem)
        for i, child in enumerate(children):
            parent.insert(idx + i, child)

    # Write output file using helper
    _write_docx_with_modifications(
        input_path=input_path,
        output_path=final_output_path,
        modifications={
            "word/document.xml": serialize_xml(document),
        },
    )

    return {
        "success": True,
        "change_id": change_id,
        "change_type": change_type,
        "output_path": str(final_output_path),
    }
