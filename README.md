# docx-comments-mcp

An MCP server for Claude Desktop that provides comprehensive read/write access to Word documents, including comments, track changes, and reply threads — features that `python-docx` doesn't fully expose.

## Features

- **Read documents**: Extract text, comments (with reply threads), and track changes
- **Add comments**: Anchor comments to specific text in the document
- **Reply to comments**: Create threaded replies on existing comments
- **Track changes**: Make edits with insertions and deletions tracked
- **Resolve comments**: Mark comments as done
- **Accept/reject changes**: Apply or undo tracked changes

## Installation

```bash
# Clone the repository
git clone https://github.com/your-username/docx-comments-mcp.git
cd docx-comments-mcp

# Install with uv
uv sync
```

## Usage with Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "docx-comments": {
      "command": "uv",
      "args": ["--directory", "/path/to/docx-comments-mcp", "run", "docx-comments-mcp"]
    }
  }
}
```

## Available Tools

### `read_document`

Read a Word document and extract content, comments, and track changes.

**Parameters:**
- `path` (required): Path to the .docx file
- `include_text` (default: true): Include full document text
- `include_comments` (default: true): Include comments with anchors
- `include_track_changes` (default: true): Include insertions/deletions

**Returns:**
```json
{
  "metadata": {
    "path": "/path/to/file.docx",
    "author": "Original Author",
    "created": "2025-01-15T10:30:00Z",
    "modified": "2025-01-18T14:22:00Z",
    "word_count": 4523
  },
  "paragraphs": [
    {"index": 0, "text": "The paragraph content...", "style": "Heading 1"}
  ],
  "comments": [
    {
      "id": 0,
      "author": "Dr. Smith",
      "date": "2025-01-16T09:15:00Z",
      "text": "Consider citing Main & Hesse here",
      "anchor_text": "disorganized attachment patterns",
      "anchor_paragraph": 12,
      "resolved": false,
      "replies": [
        {
          "id": 1,
          "parent_id": 0,
          "author": "Josh",
          "date": "2025-01-17T11:00:00Z",
          "text": "Added citation — see revision"
        }
      ]
    }
  ],
  "track_changes": [
    {
      "id": 5,
      "type": "deletion",
      "author": "Dr. Smith",
      "date": "2025-01-16T09:20:00Z",
      "text": "invariably",
      "paragraph": 8
    }
  ]
}
```

### `create_comment`

Add a comment anchored to specific text in a Word document.

**Parameters:**
- `path` (required): Path to the .docx file
- `anchor_text` (required): Text to anchor the comment to (must exist and be unique)
- `comment_text` (required): The comment content
- `author` (default: "Claude"): Comment author name
- `output_path` (optional): Save to new file; if omitted, creates timestamped backup and overwrites

**Returns:**
```json
{
  "success": true,
  "comment_id": 3,
  "anchored_to": "the exact text that was matched",
  "paragraph": 15,
  "output_path": "/path/to/output.docx"
}
```

### `create_reply`

Add a reply to an existing comment.

**Parameters:**
- `path` (required): Path to the .docx file
- `parent_comment_id` (required): ID of comment to reply to
- `reply_text` (required): The reply content
- `author` (default: "Claude"): Reply author name
- `output_path` (optional): Save to new file; if omitted, creates backup

### `create_track_change`

Make an edit with track changes enabled (insertion, deletion, or replacement).

**Parameters:**
- `path` (required): Path to the .docx file
- `find_text` (required): Text to find and modify
- `replace_with` (required): Replacement text (empty string for deletion)
- `author` (default: "Claude"): Change author name
- `output_path` (optional): Save to new file; if omitted, creates backup

### `mark_comment_resolved`

Mark a comment as resolved/done.

**Parameters:**
- `path` (required): Path to the .docx file
- `comment_id` (required): ID of comment to resolve
- `output_path` (optional): Save to new file; if omitted, creates backup

### `accept_change`

Accept a tracked change (apply permanently).

**Parameters:**
- `path` (required): Path to the .docx file
- `change_id` (required): ID of the track change to accept
- `output_path` (optional): Save to new file; if omitted, creates backup

### `reject_change`

Reject a tracked change (undo the change).

**Parameters:**
- `path` (required): Path to the .docx file
- `change_id` (required): ID of the track change to reject
- `output_path` (optional): Save to new file; if omitted, creates backup

## Safety Features

- **Automatic backups**: When modifying a file without specifying `output_path`, a timestamped backup is created (e.g., `document.backup_20250119_143022.docx`)
- **Atomic writes**: Uses temporary files and atomic moves to prevent corruption
- **Unique anchor matching**: Comments require unique anchor text to prevent ambiguity

## Development

```bash
# Install dev dependencies
uv sync
uv pip install pytest pytest-asyncio

# Run tests
uv run pytest -v

# Run specific test file
uv run pytest tests/test_reader.py -v
```

## Architecture

```
src/docx_comments_mcp/
├── __init__.py
├── server.py          # MCP server with tool definitions
├── reader.py          # Read operations (document, comments, track changes)
├── writer.py          # Write operations (add comments, track changes)
└── xml_helpers.py     # Low-level OOXML parsing utilities
```

## License

MIT
