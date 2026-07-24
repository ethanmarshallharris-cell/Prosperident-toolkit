"""
Generates a synthetic PDF that mimics the layout of a real Dentrix
"Audit Trail Report" (based on the structure observed in a real, sanitized
sample: header lines beginning with "DATE:", followed by one or more detail
lines that repeat the date and then show "Last, First" before further
tab-separated columns).

This is used ONLY to test the redaction pipeline end-to-end without needing
a real patient file. Names used here are fake/placeholder test data.
"""
import fitz  # PyMuPDF

FONT = "helv"
FONT_SIZE = 9
LINE_HEIGHT = 12

# Column x-positions (points), mimicking tab stops in the real report.
X_LEFT = 50
X_DESC = 260
X_AMOUNT = 430
X_CODE = 480

PAGE_W, PAGE_H = 612, 792  # US Letter


def draw_report_header(page, y, page_no, run_date="07/14/2026"):
    page.insert_text((X_LEFT, y), run_date, fontname=FONT, fontsize=FONT_SIZE)
    y += LINE_HEIGHT
    page.insert_text((X_LEFT, y), "AUDIT TRAIL REPORT", fontname=FONT, fontsize=FONT_SIZE)
    y += LINE_HEIGHT
    page.insert_text((X_LEFT, y), "Gallery Dental", fontname=FONT, fontsize=FONT_SIZE)
    y += LINE_HEIGHT
    page.insert_text((X_LEFT, y), f"Page:\t{page_no}", fontname=FONT, fontsize=FONT_SIZE)
    y += LINE_HEIGHT * 2
    return y


def draw_entry(page, y, date, time_, user, type_, details):
    """details: list of (date, name, desc, amount, code)"""
    header = f"DATE: {date}    TIME: {time_}    USER: {user}    TYPE: {type_}"
    page.insert_text((X_LEFT, y), header, fontname=FONT, fontsize=FONT_SIZE)
    y += LINE_HEIGHT
    for (d, name, desc, amount, code) in details:
        page.insert_text((X_LEFT, y), f"{d} {name}", fontname=FONT, fontsize=FONT_SIZE)
        page.insert_text((X_DESC, y), desc, fontname=FONT, fontsize=FONT_SIZE)
        page.insert_text((X_AMOUNT, y), amount, fontname=FONT, fontsize=FONT_SIZE)
        page.insert_text((X_CODE, y), code, fontname=FONT, fontsize=FONT_SIZE)
        y += LINE_HEIGHT
    y += LINE_HEIGHT * 0.5
    return y


def build():
    doc = fitz.open()
    page = doc.new_page(width=PAGE_W, height=PAGE_H)
    y = 50
    y = draw_report_header(page, y, page_no=3)

    entries = [
        ("01/07/2016", "12:03:05", "5115", "Guar Payment", [
            ("07/01/2016", "Smith, John", "Visa Payment - Thank You", "-189.00", "HYGO"),
            ("07/01/2016", "Smith, John", "MasterCard Payment - Thank You", "-189.00", "HYGO"),
        ]),
        ("01/07/2016", "12:25:10", "5115", "Guar Payment", [
            ("07/01/2016", "Jones, Amy", "Visa Payment - Thank You", "763.00", "[SPLIT]"),
        ]),
        ("01/07/2016", "13:54:18", "5115", "Guar Payment", [
            ("07/01/2016", "Brunson, Jalen", "MasterCard Payment - Thank You", "2952.00", "[SPLIT]"),
        ]),
        ("01/07/2016", "14:02:00", "5115", "Guar Payment", [
            ("07/01/2016", "Jones, Jim", "Cash Payment - Thank You", "400.00", "[SPLIT]"),
        ]),
        ("01/07/2016", "14:16:19", "5115", "Guar Payment", [
            ("07/01/2016", "Jones, Jim", "Debit Card Payment - Thank You", "5.38", "[SPLIT]"),
        ]),
        # Edge cases: hyphenated last name, apostrophe, multi-word last name, suffix
        ("01/08/2016", "09:12:00", "5115", "Guar Payment", [
            ("07/08/2016", "Smith-Jones, Mary", "Cash Payment - Thank You", "50.00", "PROP"),
        ]),
        ("01/08/2016", "09:30:00", "5115", "Guar Payment", [
            ("07/08/2016", "O'Brien, Patrick", "Visa Payment - Thank You", "120.00", "PROP"),
        ]),
        ("01/08/2016", "10:00:00", "5115", "Guar Payment", [
            ("07/08/2016", "Van Der Berg, Kim", "Check Payment - Thank You", "80.00", "PROP"),
        ]),
        ("01/08/2016", "10:15:00", "5115", "Guar Payment", [
            ("07/08/2016", "Adams Jr, Robert", "Cash Payment - Thank You", "60.00", "PROP"),
        ]),
    ]

    for (date, time_, user, type_, details) in entries:
        if y > PAGE_H - 80:
            page = doc.new_page(width=PAGE_W, height=PAGE_H)
            y = 50
        y = draw_entry(page, y, date, time_, user, type_, details)

    # Force a second page and repeat "Jones, Jim" to test cross-page numbering consistency
    page2 = doc.new_page(width=PAGE_W, height=PAGE_H)
    y = 50
    y = draw_report_header(page2, y, page_no=4)
    y = draw_entry(page2, y, "01/09/2016", "08:00:00", "5115", "Guar Payment", [
        ("07/09/2016", "Jones, Jim", "Cash Payment - Thank You", "25.00", "PROP"),
    ])
    page2.insert_text((X_LEFT, y + 20), "Audit #: 50", fontname=FONT, fontsize=FONT_SIZE)

    out_path = "/home/claude/dentrix_redactor/test/synthetic_audit_trail.pdf"
    doc.save(out_path)
    doc.close()
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    build()
