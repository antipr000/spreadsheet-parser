"""
Block grouping — associates heading blocks with the content block(s)
immediately below them to form semantic chunks.

A chunk is a list of related blocks (e.g. [HeadingBlock, TableBlock]).
"""

from __future__ import annotations

from typing import Dict, List, Tuple

from openpyxl.utils import column_index_from_string

from dto.blocks import Block, HeadingBlock


def _parse_coord(coord: str) -> Tuple[int, int]:
    """Parse an A1-style coordinate into (row, col) — both 1-based."""
    col_str = "".join(c for c in coord if c.isalpha())
    row_num = int("".join(c for c in coord if c.isdigit()) or "0")
    try:
        col_num = column_index_from_string(col_str) if col_str else 0
    except Exception:
        col_num = 0
    return row_num, col_num


def _col_overlap(block_a: Block, block_b: Block) -> float:
    """
    Return the fraction of block_a's column span that overlaps with block_b.
    0.0 = no overlap, 1.0 = full overlap.
    """
    _, a_min_col = _parse_coord(block_a.bounding_box.top_left)
    _, a_max_col = _parse_coord(block_a.bounding_box.bottom_right)
    _, b_min_col = _parse_coord(block_b.bounding_box.top_left)
    _, b_max_col = _parse_coord(block_b.bounding_box.bottom_right)

    overlap_start = max(a_min_col, b_min_col)
    overlap_end = min(a_max_col, b_max_col)
    overlap = max(0, overlap_end - overlap_start + 1)

    a_span = a_max_col - a_min_col + 1
    return overlap / a_span if a_span > 0 else 0.0


# Maximum row gap between a heading's bottom row and the next block's top row
# for them to be considered related.
_MAX_ROW_GAP = 3

# Minimum column overlap fraction for heading association.
_MIN_COL_OVERLAP = 0.4


def group_blocks_into_chunks(blocks: List[Block]) -> Dict[str, List[Block]]:
    """
    Walk the reading-order-sorted block list and merge heading blocks
    with the content block(s) directly below them.

    Returns an ordered dict  ``{"block0": [...], "block1": [...], ...}``.
    """
    if not blocks:
        return {}

    chunks: Dict[str, List[Block]] = {}
    chunk_idx = 0
    i = 0

    while i < len(blocks):
        block = blocks[i]

        if isinstance(block, HeadingBlock) and i + 1 < len(blocks):
            next_block = blocks[i + 1]

            heading_bottom_row, _ = _parse_coord(block.bounding_box.bottom_right)
            next_top_row, _ = _parse_coord(next_block.bounding_box.top_left)
            row_gap = next_top_row - heading_bottom_row

            col_ovlp = _col_overlap(block, next_block)

            if (
                0 < row_gap <= _MAX_ROW_GAP
                and col_ovlp >= _MIN_COL_OVERLAP
                and not isinstance(next_block, HeadingBlock)
            ):
                # Associate heading with the next content block
                # If next block is a TableBlock, also set its title
                from dto.blocks import TableBlock

                if isinstance(next_block, TableBlock) and not next_block.title:
                    next_block.title = block.text

                chunks[f"block{chunk_idx}"] = [block, next_block]
                chunk_idx += 1
                i += 2
                continue

        # Standalone block (no heading association)
        chunks[f"block{chunk_idx}"] = [block]
        chunk_idx += 1
        i += 1

    return chunks
