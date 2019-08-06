import os
import re
from docx import Document

# Install python-docx package to convert MS Word documents
# pip install python-docx


def convert_zero(txt):
    return re.sub(r'([\u0660]|[\u06F0])', "0", txt)


def convert_one(txt):
    return re.sub(r'([\u0661]|[\u06F1])', "1", txt)


def convert_two(txt):
    return re.sub(r'([\u0662]|[\u06F2])', "2", txt)


def convert_three(txt):
    return re.sub(r'([\u0663]|[\u06F3])', "3", txt)


def convert_four(txt):
    return re.sub(r'([\u0664]|[\u06F4])', "4", txt)


def convert_five(txt):
    return re.sub(r'([\u0665]|[\u06F5])', "5", txt)


def convert_six(txt):
    return re.sub(r'([\u0666]|[\u06F6])', "6", txt)


def convert_seven(txt):
    return re.sub(r'([\u0667]|[\u06F7])', "7", txt)


def convert_eight(txt):
    return re.sub(r'([\u0668]|[\u06F8])', "8", txt)


def convert_nine(txt):
    return re.sub(r'([\u0669]|[\u06F9])', "9", txt)


def convert_numbers(txt):
    txt = convert_zero(txt)
    txt = convert_one(txt)
    txt = convert_two(txt)
    txt = convert_three(txt)
    txt = convert_four(txt)
    txt = convert_five(txt)
    txt = convert_six(txt)
    txt = convert_seven(txt)
    txt = convert_eight(txt)
    txt = convert_nine(txt)

    return txt


for f in os.listdir("."):
    if not f.endswith(".docx") and not f.endswith(".txt"):
        continue

    is_docx = f.endswith(".docx")
    with open(f, 'rb' if is_docx else 'r') as doc:
        if is_docx:
            document = Document(doc)
            lines = document.paragraphs
        else:
            lines = doc.readlines()

        with open(os.path.splitext(doc.name)[0] + ".usfm", 'w+') as usfm:
            usfm.write('\\id BOOK_CODE Unlocked Literal Bible \n')
            usfm.write('\\ide UTF-8 \n')
            usfm.write('\\h Book_Name \n')
            usfm.write('\\toc1 Book_Name_Extended \n')
            usfm.write('\\toc2 Book_Name \n')
            usfm.write('\\toc3 Book_Code \n')
            usfm.write('\\mt Book_Name \n\n')

            is_first_paragraph = False

            for line in lines:
                if is_docx:
                    line = line.text
                line = line.strip()
                if len(line) == 0:
                    continue

                reg = re.compile(r'([0-9]+)')

                # Convert arabic numbers if there any
                p_line = convert_numbers(line)
                m = reg.search(p_line)

                if m is not None:
                    if len(line) <= 3:
                        # chapter marker
                        usfm.write('\\s5\n')
                        usfm.write('\\c ' + m.group() + '\n\n')
                        is_first_paragraph = True
                    else:
                        # verse markers
                        v_line = reg.sub(r'\n\\v \1 ', p_line)
                        if not is_first_paragraph:
                            usfm.write('\\s5\n')
                        usfm.write('\\p')
                        usfm.write(v_line + '\n\n')
                        is_first_paragraph = False
