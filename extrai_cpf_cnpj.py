# -*- coding: utf-8 -*-
from modules.utils import procura_cnpj, procura_cpf, select_files, choose_type
from pdfminer.high_level import extract_text
import docx2txt
import os
import sys
from tqdm import tqdm
import win32com.client


class Doc:
    def __init__(self, file_path, file_type):
        self.type = file_type
        self.path = file_path
        self.filename = os.path.basename(file_path)
        self.text = ''

    # Get the file content
    def get_text(self):
        if self.type == '*.pdf':
            try:
                self.text = extract_text(self.path)
            except TypeError:
                raise f"Erro ao ler o arquivo {0}. É um 'pdf' pesquisável?".format(self.filename)
        elif self.type == '*.docx':
            try:
                self.text = docx2txt.process(self.path)
            except TypeError:
                raise f"Erro ao ler o arquivo {0}. É um '.docx'?".format(self.filename)
        elif self.type == '*.doc':
            try:
                word = win32com.client.Dispatch("Word.Application")
                word.visible = False
                word.Documents.Open(self.path)
                doc = word.ActiveDocument
                self.text = doc.Range().Text
                word.Application.Quit()
            except TypeError:
                raise f"Erro ao ler o arquivo {0}. É um '.doc'?".format(self.filename)

    # Find CPFs on text
    def get_cpfs(self):
        cpfs = procura_cpf(self.text)
        return [(cpf, self.filename) for cpf in cpfs]

    # Find cnpjs on text
    def get_cnpjs(self):
        cpfs = procura_cpf(self.text)
        return [(cpf, self.filename) for cpf in cpfs]


if __name__ == '__main__':
    # Get the type of the files from user according to its extensions
    file_type = {'1': '*.pdf', '2': '*.docx', '3': '*.doc'}
    input_type = choose_type()

    # Check the input [3 - leave program]
    if input_type == '4':
        sys.exit()
    # If the input option is not valid, ask again
    elif input_type not in ['1', '2', '3', '4']:
        choose_type()

    # Set the working directory to the current
    BASE_PATH = os.getcwd()

    # Lists to store final results
    cpf_results = []
    cnpj_results = []

    # Open GUI window to user select the sources he wants to get CPFs/CNPJs
    files = select_files(file_type[input_type])

    print('\nLendo arquivos...\n')
    for file in tqdm(files):  # To each file selected
        doc = Doc(file, file_type[input_type])  # Instantiate a Doc object

        # Store the file's text
        # text = doc.get_text()
        doc.get_text()
        # Get file's valid CPFs/CNPJs
        cpfs = procura_cpf(doc.text)
        cnpjs = procura_cnpj(doc.text)

        # Save each CPF/CNPJ as a string (Ex: "81781726255;source_file.docx")
        for cpf in cpfs:
            cpf_results.append(';'.join([cpf, doc.filename]))
        for cnpj in cnpjs:
            cnpj_results.append(';'.join([cnpj, doc.filename]))

    # Create a .csv to store CPFs.
    print('\nGravando CPFs...')
    with open("cpfs_encontrados.csv", mode="a", newline='\n') as new_file:
        for result in tqdm(cpf_results):
            new_file.write(f'{result}\n')

    # Create a .csv to store CNPJs.
    print('\nGravando CNPJs...')
    with open("cnpjs_encontrados.csv", mode="a", newline='\n') as new_file:
        for result in tqdm(cnpj_results):
            new_file.write(f'{result}\n')
