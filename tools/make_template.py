"""Build the reconciliation template from a filled sample workbook.

Clears every data row of the six data sheets but keeps row 2 as the formula
prototype, so table formulas (Remarks, Reco Decision, ...) stay in the template.

    python tools/make_template.py "Sample Output.xlsx" Output_template_fixed.xlsx
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from gstrecon.xlsx_writer import fill_template  # noqa: E402

DATA_SHEETS = ("Books_CurrMnth", "Books_PrevMnth_NextMnth", "GSTR2B", "GSTR2B_PrevMonth",
               "EXCLUSIONS", "BOT_ReconData")


def make_template(sample: bytes, sheets=DATA_SHEETS) -> bytes:
    return fill_template(sample, {name: [] for name in sheets}, header={})


def main(argv) -> int:
    if len(argv) != 2:
        print(__doc__.strip().splitlines()[-1].strip())
        return 1
    with open(argv[0], "rb") as f:
        data = make_template(f.read())
    with open(argv[1], "wb") as f:
        f.write(data)
    print(f"WROTE {argv[1]}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
