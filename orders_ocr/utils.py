from __future__ import annotations

def pop_word(string: str, word_no: int, row: dict, col_name: str) -> str:
    """Remove the *word_no*-th word from *string* and place it in *row* under *col_name*."""
    words = string.split()
    if len(words) > word_no:
        row[col_name] = words[word_no]
        return " ".join(words[:word_no] + words[word_no + 1 :])
    row[col_name] = ""
    return " ".join(words[:word_no])


def is_valid_float(value: str) -> bool:
    """True if *value* can be turned into float, comma or dot accepted."""
    try:
        float(value.replace(",", "."))
        return True
    except ValueError:
        return False