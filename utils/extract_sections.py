import pandas as pd
import docx
import numpy as np

from utils.PathManager import load_path_manager as lpm


def iter_block_items(parent):
    body_elements = document._body._body
    # extract those wrapped in <w:r> tag
    rs = body_elements.xpath('.//w:r')
    # check if style is hyperlink (toc)
    table_of_content = np.array([r.text for r in rs if r.style == "Hyperlink"])

    result = []
    current_section = [0]
    section_header = ""
    for each in table_of_content:
        try:
            section_number = float(each)

            if section_header != "":
                current_section.append(section_header)
                if section_header.startswith("FS-"):
                    current_section.append(section_header.split(' ')[0])
                    current_section.append('')
                elif section_header.startswith("AA-API-") \
                        or section_header.startswith("AA-BAT-") \
                        or section_header.startswith("AA-API-"):
                    current_section.append('')
                    current_section.append(section_header.split(' ')[0])
                else:
                    current_section.append('')
                    current_section.append('')
                result.append(current_section)

                current_section = [section_number]
                section_header = ""
        except Exception as e:
            section_header += each

    current_section.append(section_header)
    if section_header.startswith("FS-"):
        current_section.append(section_header.split(' ')[0])
        current_section.append('')
    elif section_header.startswith("AA-API-") \
            or section_header.startswith("AA-BAT-") \
            or section_header.startswith("AA-API-"):
        current_section.append('')
        current_section.append(section_header.split(' ')[0])
    else:
        current_section.append('')
        current_section.append('')
    result.append(current_section)

    current_section = [section_number]
    section_header = ""

    df = pd.DataFrame(result, columns=["Section Number", "Section Name", "FS ID", "API"])
    df.to_excel(writer, index=False)


if __name__ == '__main__':
    file_refer = lpm.refer("Program Specification.docx")
    document = docx.Document(file_refer)

    folder_output = str(lpm.output)
    writer = pd.ExcelWriter('{}/Program Spec Inventory.xlsx'.format(folder_output), engine="xlsxwriter")

    iter_block_items(document)

    writer._save()
