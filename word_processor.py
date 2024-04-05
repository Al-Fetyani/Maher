from typing import List
from docx import Document
from pathlib import Path
import os


def read_docx(file_path):
    return Document(file_path)


def launch_word(file_path):
    path_file_name = Path(file_path).absolute()
    os.startfile(path_file_name)


def replace_placeholder(doc: Document, data):
    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if key not in paragraph.text:
                continue
            paragraph.text = paragraph.text.replace(key, value)


def save_word(doc: Document, output_file_path):
    doc.save(output_file_path)
    while not Path(output_file_path).exists():
        pass


def merge_documents(docs: List[Document]) -> Document:
    for doc in docs[:-1]:
        doc.add_page_break()

    for doc in docs[1:]:
        for element in doc.element.body:
            docs[0].element.body.append(element)
    return docs[0]
