from docx import Document
import os


def get_list_files(dir_name):
    res = []
    for root, dirs, files in os.walk(dir_name):
        for file in files:
            if file.endswith(".docx"):
                res.append(os.path.join(root, file))
    return res


def combine_word_documents(files):
    merged_document = Document()

    for index, file in enumerate(files):
        sub_doc = Document(file)

        # Don't add a page break if you've reached the last file.
        if index < len(files)-1:
           sub_doc.add_page_break()

        for element in sub_doc.element.body:
            merged_document.element.body.append(element)

    merged_document.save('merged.docx')


combine_word_documents(get_list_files('to_merge'))
