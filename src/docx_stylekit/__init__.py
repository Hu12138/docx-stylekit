from .api import (
    observe_docx,
    merge_yaml,
    diff_yaml,
    render_from_json,
    render_from_markdown,
    fix_image_paragraphs,
)

__all__ = [
    "__version__",
    "observe_docx",
    "merge_yaml",
    "diff_yaml",
    "render_from_json",
    "render_from_markdown",
    "fix_image_paragraphs",
]

__version__ = "0.2.0"
