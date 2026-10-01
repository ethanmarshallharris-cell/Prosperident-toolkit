"""
Registry of available document-type plugins. See base.py for the contract
each plugin module implements, and dentrix_audit_trail.py for a complete
reference implementation.

To add a new document type: implement it as a new module in this package
(base.py explains the required contents), import it below, and add it to
DOCUMENT_TYPES. That's the only change main_gui.py needs -- it drives the
dropdown and everything else entirely from this list.
"""
from . import dentrix_audit_trail
from . import excel_spreadsheet

DOCUMENT_TYPES = [
    dentrix_audit_trail,
    excel_spreadsheet,
]


def get_by_id(doc_type_id):
    for dt in DOCUMENT_TYPES:
        if dt.ID == doc_type_id:
            return dt
    raise KeyError(f"Unknown document type id: {doc_type_id!r}")
