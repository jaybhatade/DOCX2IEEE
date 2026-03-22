from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 🔢 Global counter for tables
table_counter = 1


def to_roman(num):
    val = [
        1000, 900, 500, 400,
        100, 90, 50, 40,
        10, 9, 5, 4, 1
    ]
    syms = [
        "M", "CM", "D", "CD",
        "C", "XC", "L", "XL",
        "X", "IX", "V", "IV", "I"
    ]
    roman = ''
    i = 0
    while num > 0:
        for _ in range(num // val[i]):
            roman += syms[i]
            num -= val[i]
        i += 1
    return roman


def add_table_caption(doc, caption_text):
    global table_counter

    roman = to_roman(table_counter)

    # TABLE I
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = p1.add_run(f"TABLE {roman}")
    run1.bold = True

    # Caption text
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(caption_text.upper())
    run2.bold = True

    table_counter += 1


def add_table(doc, table_data, caption=None):
    # 👉 Default caption if not provided
    if not caption:
        caption = "Sample Table Showing Data"

    add_table_caption(doc, caption)

    rows = len(table_data)
    cols = len(table_data[0]) if rows > 0 else 0

    table = doc.add_table(rows=rows, cols=cols)

    # Fill data
    for i, row in enumerate(table_data):
        for j, cell in enumerate(row):
            table.rows[i].cells[j].text = cell

    # Center table
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER


def add_section(doc, section, level=1):
    doc.add_heading(section["heading"], level=level)

    for item in section["content"]:
        if isinstance(item, dict):
            if item["type"] == "text":
                doc.add_paragraph(item["value"])

            elif item["type"] == "table":
                add_table(
                    doc,
                    item["value"],
                    item.get("caption", None)
                )
        else:
            doc.add_paragraph(item)

    for sub in section["subsections"]:
        add_section(doc, sub, level=level + 1)


def generate_doc(data, output_path):
    global table_counter
    table_counter = 1  # reset for each document

    doc = Document("templates/ieee_template.docx")

    for para in doc.paragraphs:
        if "{{TITLE}}" in para.text:
            para.text = data["title"]

        elif "{{ABSTRACT}}" in para.text:
            para.text = " ".join(data["abstract"])

        elif "{{KEYWORDS}}" in para.text:
            para.text = ", ".join(data["keywords"])

        elif "{{CONTENT}}" in para.text:
            para.text = ""

            for section in data["sections"]:
                add_section(doc, section, level=1)

    doc.save(output_path)