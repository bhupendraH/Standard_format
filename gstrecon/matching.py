"""Pair books documents with GSTR-2B documents for one month."""
from __future__ import annotations

from collections import defaultdict

from .model import Pair
from .normalise import is_gstin

EXACT = "EXACT MATCH"
MISMATCH = "MISMATCH"
PROBABLE = "PROBABLE MATCH"
PR_ONLY = "PR+"
G2B_ONLY = "GSTR2B+"


def _key(d):
    return (d.gstin.strip().upper(), d.doc_type, str(d.norm))


def _close(a, b, tolerance):
    return abs(a.tax - b.tax) <= tolerance + 1e-9


def match(pr_docs, g2b_docs, tolerance: float = 5.0):
    """Return (pairs, warnings). pr_docs must already exclude RCM documents.

    Order: paired rows (books order), then PR+ (books order), then GSTR2B+ (2B order).
    A document whose key is duplicated on either side is never paired, at any step:
    when in doubt, do not pair.
    """
    warnings = []
    pr_ok = [i for i, d in enumerate(pr_docs) if is_gstin(d.gstin)]
    g2b_ok = list(range(len(g2b_docs)))

    pr_by_key, g2b_by_key = defaultdict(list), defaultdict(list)
    for i in pr_ok:
        pr_by_key[_key(pr_docs[i])].append(i)
    for j in g2b_ok:
        g2b_by_key[_key(g2b_docs[j])].append(j)

    blocked_pr, blocked_g2b = set(), set()
    for side, by_key, blocked, docs in (
        ("books", pr_by_key, blocked_pr, pr_docs),
        ("GSTR-2B", g2b_by_key, blocked_g2b, g2b_docs),
    ):
        for key, idx in by_key.items():
            if len(idx) > 1:
                blocked.update(idx)
                blocked_pr.update(pr_by_key.get(key, ()))
                blocked_g2b.update(g2b_by_key.get(key, ()))
                warnings.append(
                    f"duplicate invoice key {key[0]} {docs[idx[0]].number} in {side}; not paired")

    pr_match, used_g2b = {}, set()
    # Step 1: exact key
    for i in pr_ok:
        if i in blocked_pr:
            continue
        cands = [j for j in g2b_by_key.get(_key(pr_docs[i]), ()) if j not in blocked_g2b]
        if len(cands) == 1:
            j = cands[0]
            status = EXACT if _close(pr_docs[i], g2b_docs[j], tolerance) else MISMATCH
            pr_match[i] = (j, status)
            used_g2b.add(j)

    # Step 2: probable - same GSTIN, type, date; tax within tolerance; unique both ways
    free_pr = [i for i in pr_ok if i not in pr_match and i not in blocked_pr]
    free_g2b = [j for j in g2b_ok if j not in used_g2b and j not in blocked_g2b]

    def group(d):
        return (d.gstin.strip().upper(), d.doc_type, d.doc_date)

    for i in free_pr:
        a = pr_docs[i]
        if a.doc_date is None:
            continue
        cands = [j for j in free_g2b if j not in used_g2b
                 and group(g2b_docs[j]) == group(a) and _close(a, g2b_docs[j], tolerance)]
        if len(cands) != 1:
            continue
        j = cands[0]
        rivals = [k for k in free_pr if k not in pr_match and k != i
                  and group(pr_docs[k]) == group(a) and _close(pr_docs[k], g2b_docs[j], tolerance)]
        if rivals:
            continue
        pr_match[i] = (j, PROBABLE)
        used_g2b.add(j)

    pairs = [Pair(pr_docs[i], g2b_docs[pr_match[i][0]], pr_match[i][1])
             for i in range(len(pr_docs)) if i in pr_match]
    pairs += [Pair(d, None, PR_ONLY) for i, d in enumerate(pr_docs) if i not in pr_match]
    pairs += [Pair(None, d, G2B_ONLY) for j, d in enumerate(g2b_docs) if j not in used_g2b]
    return pairs, warnings
