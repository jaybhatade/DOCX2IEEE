from docx import Document

def get_heading_level(style_name):
    if "Heading" in style_name:
        try:
            return int(style_name.split()[-1])
        except:
            return None
    return None

def parse_docx(file_path):
    doc = Document(file_path)

    data = {
        "title": "",
        "abstract": [],
        "keywords": [],
        "sections": []
    }

    stack = []
    mode = None  # tracks where we are

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Title
        if not data["title"]:
            data["title"] = text
            continue

        lower = text.lower()

        # Detect Abstract
        if lower == "abstract":
            mode = "abstract"
            continue

        # Detect Keywords
        if "keywords" in lower:
            mode = "keywords"
            continue

        level = get_heading_level(para.style.name)

        # If heading → section
        if level:
            mode = "section"

            section = {
                "heading": text,
                "level": level,
                "content": [],
                "subsections": []
            }

            if level == 1:
                data["sections"].append(section)
                stack = [section]
            else:
                parent = stack[level - 2]
                parent["subsections"].append(section)

                if len(stack) >= level:
                    stack[level - 1] = section
                else:
                    stack.append(section)

        else:
            if mode == "abstract":
                data["abstract"].append(text)

            elif mode == "keywords":
                data["keywords"].append(text)

            elif mode == "section" and stack:
                stack[-1]["content"].append(text)

    return data