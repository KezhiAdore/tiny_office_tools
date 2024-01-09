import os
import pandas as pd
from docx import Document

def get_title(doc_path):
    document = Document(doc_path)
    return [document.paragraphs[i].text for i in range(4)]

def list_title(dir_path):
    title_list = []
    for file in os.listdir(dir_path):
        if file.endswith('.docx'):
            doc_path = os.path.join(dir_path, file)
            title_list.append(get_title(doc_path))
            print(title_list)
    return title_list
            
if __name__ == '__main__':
    dir_path = './data'
    list_title(dir_path)