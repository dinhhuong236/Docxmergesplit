"""DocxTool – gộp / tách file Word (.docx) giữ nguyên định dạng."""
from .merge import MergeOptions, merge_files
from .package import DocxError
from .split import MODES, SplitOptions, analyze, split_file

__version__ = "1.0.0"
__all__ = ["merge_files", "MergeOptions", "split_file", "SplitOptions", "analyze", "MODES", "DocxError"]
