from typing import List, Tuple
from docx.text.paragraph import Paragraph


def _pos_to_run_offset(run_texts: List[str], pos: int) -> Tuple[int, int]:
    """Map a flat string position to (run_index, offset_in_run)."""
    acc = 0
    for idx, text in enumerate(run_texts):
        next_acc = acc + len(text)
        if pos < next_acc:
            return idx, pos - acc
        acc = next_acc
    # If position equals total length, place at end of last run
    return len(run_texts) - 1, len(run_texts[-1])


def replace_text_in_paragraph(paragraph: Paragraph, old_text: str, new_text: str) -> bool:
    """
    Replace text in a paragraph while preserving existing run formatting.

    - If replacements are contained within single runs, replace in place.
    - If a match spans multiple runs, rewrite only the affected runs' text
      without recreating runs, so formatting stays with the original runs.
    """

    # 1) Fast path: replace within single runs
    replaced_in_run = False
    for run in paragraph.runs:
        if old_text in run.text:
            run.text = run.text.replace(old_text, new_text)
            replaced_in_run = True

    # Continue searching for cross-run matches even if we already replaced some runs
    run_texts: List[str] = [r.text for r in paragraph.runs]
    flat = "".join(run_texts)
    if old_text not in flat:
        return replaced_in_run

    changed = False
    search_from = 0

    while True:
        flat = "".join(run_texts)
        idx = flat.find(old_text, search_from)
        if idx == -1:
            break

        start_run, start_off = _pos_to_run_offset(run_texts, idx)
        end_run, end_off = _pos_to_run_offset(run_texts, idx + len(old_text))

        # Trim text in start and end runs, blank intermediates
        start_text = run_texts[start_run]
        end_text = run_texts[end_run]

        prefix = start_text[:start_off]
        suffix = end_text[end_off:] if start_run == end_run else end_text[end_off:]

        run_texts[start_run] = prefix + new_text + suffix

        for i in range(start_run + 1, end_run + 1):
            run_texts[i] = "" if i != start_run else run_texts[i]

        changed = True
        search_from = idx + len(new_text)

    if changed:
        for run, text in zip(paragraph.runs, run_texts):
            run.text = text

    return replaced_in_run or changed
