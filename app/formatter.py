from docx import Document

def add_section(doc, section, level=1):
    doc.add_heading(section["heading"], level=level)

    for para in section["content"]:
        doc.add_paragraph(para)

    for sub in section["subsections"]:
        add_section(doc, sub, level=level+1)

def generate_doc(data, output_path):
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