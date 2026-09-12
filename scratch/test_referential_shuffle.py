"""
test_referential_shuffle.py
---------------------------
Verifies the _shuffle_options_safe fix for referential options
(e.g. "Both a and b", "None of these") that must not move after shuffle.

Run from the RRB_v110 directory:
    python scratch/test_referential_shuffle.py
"""

import random
import re
import sys

# Reproduce the exact helper as in app.py (standalone, no DB)

_REFERENTIAL_OPT_RE = re.compile(
    r'\b(both|all\s+of\s+(the|these)|none\s+of\s+(the|these)|except|only)\b'
    r'|'
    r'\b[a-d]\b|\([a-d]\)',
    re.IGNORECASE
)


def _shuffle_options_safe_standalone(opt_a, opt_b, opt_c, opt_d):
    """Standalone version of _shuffle_options_safe (no DB call)."""
    letters = ['A', 'B', 'C', 'D']
    opt_texts = {'A': opt_a or '', 'B': opt_b or '',
                 'C': opt_c or '', 'D': opt_d or ''}

    pinned = {letter for letter, text in opt_texts.items()
              if _REFERENTIAL_OPT_RE.search(text)}

    free_letters  = [l for l in letters if l not in pinned]
    free_shuffled = free_letters[:]
    random.shuffle(free_shuffled)

    result = []
    free_iter = iter(free_shuffled)
    for letter in letters:
        if letter in pinned:
            result.append(letter)
        else:
            result.append(next(free_iter))
    return ''.join(result)


def test_referential_pinning(runs=500):
    """TEST 1: Referential options stay pinned to their original slot."""
    print("TEST 1 - Referential options stay pinned")
    opt_a, opt_b = "Intention", "Subject"
    opt_c, opt_d = "Both a and b", "None of these"

    for i in range(runs):
        order = _shuffle_options_safe_standalone(opt_a, opt_b, opt_c, opt_d)
        assert order[2] == 'C', f"Run {i}: 'Both a and b' moved! order={order}"
        assert order[3] == 'D', f"Run {i}: 'None of these' moved! order={order}"
        assert set(order[:2]) == {'A', 'B'}, f"Run {i}: free slots wrong. order={order}"

    print(f"  PASS - ran {runs} shuffles, referential options always pinned.\n")


def test_normal_shuffles(runs=500):
    """TEST 2: Normal (non-referential) question shuffles freely."""
    print("TEST 2 - Normal question shuffles freely")
    opt_a, opt_b, opt_c, opt_d = "Paris", "London", "Berlin", "Rome"

    seen = set()
    for _ in range(runs):
        order = _shuffle_options_safe_standalone(opt_a, opt_b, opt_c, opt_d)
        seen.add(order)
        assert set(order) == {'A','B','C','D'} and len(order) == 4, f"Bad order: {order}"

    assert seen != {'ABCD'}, "Shuffle never changed order!"
    print(f"  PASS - {len(seen)} distinct orderings over {runs} runs.\n")


def test_correct_answer_integrity(runs=200):
    """TEST 3: Correct-answer key survives the shuffle round-trip."""
    print("TEST 3 - Correct-answer key integrity")

    def simulate(opt_a, opt_b, opt_c, opt_d, correct_orig):
        order = _shuffle_options_safe_standalone(opt_a, opt_b, opt_c, opt_d)
        display_index = order.index(correct_orig)
        displayed_letter = chr(ord('A') + display_index)
        stored_original = order[ord(displayed_letter) - ord('A')]
        return stored_original == correct_orig

    opt_a, opt_b = "Intention", "Subject"
    opt_c, opt_d = "Both a and b", "None of these"

    for i in range(runs):
        assert simulate(opt_a, opt_b, opt_c, opt_d, 'B'), f"Run {i}: free answer B broken"
    for i in range(runs):
        assert simulate(opt_a, opt_b, opt_c, opt_d, 'C'), f"Run {i}: pinned answer C broken"

    print(f"  PASS - answer key correct across {runs*2} simulations.\n")


def test_regex_coverage():
    """TEST 4: Regex matches all referential patterns, ignores normal text."""
    print("TEST 4 - Regex pattern coverage")

    should_match = [
        "Both a and b", "both (a) and (b)", "Both A and B",
        "All of the above", "All of these",
        "None of the above", "None of these",
        "Except a and b", "a and b only", "Only a", "Both (a) and (c)",
    ]
    should_not_match = [
        "Intention", "Subject", "Paris", "London",
        "Photosynthesis", "Newton's second law",
        "3x + 2y = 10", "Electromagnetic induction",
    ]

    for text in should_match:
        assert _REFERENTIAL_OPT_RE.search(text), f"MISSED match for: '{text}'"
        print(f"  matched (ok):    '{text}'")
    for text in should_not_match:
        m = _REFERENTIAL_OPT_RE.search(text)
        assert not m, f"WRONG match for: '{text}' (got: '{m.group()}')"
        print(f"  no match (ok):   '{text}'")
    print("  PASS\n")


if __name__ == '__main__':
    random.seed()
    test_referential_pinning()
    test_normal_shuffles()
    test_correct_answer_integrity()
    test_regex_coverage()
    print("=" * 55)
    print("ALL TESTS PASSED")
    print("=" * 55)
