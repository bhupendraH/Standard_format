"""Data shared by the readers, the matcher and the sheet builder."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Optional, Union


class InputError(Exception):
    """A problem with an input file. The message is safe to print."""


@dataclass
class Doc:
    """One invoice / credit note / debit note, from books (PR) or GSTR-2B."""

    gstin: str
    vendor_name: str
    number: str  # as written in the source
    norm: Union[str, int]  # normalise_number(number)
    doc_type: str  # "Invoice" | "Credit Note" | "Debit Note"
    doc_date: Optional[date]
    value: float
    taxable: float
    igst: float
    cgst: float
    sgst: float
    cess: float = 0.0
    pos_name: str = ""  # books: Party Region; GSTR-2B: Place of supply
    reverse_charge: str = "N"  # books: as in file; GSTR-2B: "Yes"/"No"
    rcm: bool = False
    extra: dict = field(default_factory=dict)

    @property
    def tax(self) -> float:
        return self.igst + self.cgst + self.sgst


@dataclass
class Pair:
    pr: Optional[Doc]
    g2b: Optional[Doc]
    status: str


@dataclass
class PrFile:
    month: tuple  # (year, month)
    docs: list


@dataclass
class G2bFile:
    month: tuple  # (year, month)
    gstin: str
    financial_year: str
    generated: Optional[date]
    docs: list
    ignored: dict  # sheet name -> rows not reconciled
