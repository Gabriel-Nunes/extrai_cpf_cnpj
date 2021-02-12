# -*- coding: utf-8 -*-

from modules.utils import procura_cnpj, procura_cpf, select_files, choose_type, to_table, show_exception_and_exit
from pdfminer.high_level import extract_text
import docx2txt
import os
import re
import sys
from tqdm import tqdm
import win32com.client
from time import sleep


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
                text = extract_text(self.path)
                self.text = re.sub(r'[\r\n]+', r' ', text)
            except TypeError:
                raise f"Erro ao ler o arquivo {0}. É um 'pdf' pesquisável?".format(self.filename)
        elif self.type == '*.docx':
            try:
                text = docx2txt.process(self.path)
                self.text = re.sub(r'[\n\r]+', ' ', text)
            except TypeError:
                raise f"Erro ao ler o arquivo {0}. É um '.docx'?".format(self.filename)
        elif self.type == '*.doc':
            try:
                word = win32com.client.Dispatch("Word.Application")
                word.visible = False
                word.Documents.Open(self.path)
                doc = word.ActiveDocument
                text = doc.Range().Text
                self.text = re.sub(r'[\n\r]+', ' ', text)
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
    input_type = choose_type()
    while True:
        # Hook any exception avoiding close console
        sys.excepthook = show_exception_and_exit

        # Get the type of the files from user according to its extensions
        file_type = {'1': '*.pdf', '2': '*.docx', '3': '*.doc'}

        # If the input option is not valid, ask again
        if input_type not in ['1', '2', '3']:
            choose_type()

        # Lists to store final results
        cpf_results = []
        cnpj_results = []

        # Open GUI window to user select the sources he wants to get CPFs/CNPJs
        files = select_files(file_type[input_type])

        # Set the working directory to the sources folder
        BASE_PATH = os.path.dirname(files[0])

        print('\nLendo arquivos...\n')
        for file in tqdm(files):  # To each file selected
            doc = Doc(os.path.abspath(file), file_type[input_type])  # Instantiate a Doc object

            # Store the file's text in one line
            doc.get_text()

            # Get file's valid CPFs/CNPJs
            cpfs = procura_cpf(doc.text)
            cnpjs = procura_cnpj(doc.text)

            # Save each CPF/CNPJ as a string (Ex: "81781726255|source_file.docx")
            for cpf in cpfs:
                cpf_results.append('|'.join([cpf, doc.filename, doc.path, f"\"{doc.text}\""]))
            for cnpj in cnpjs:
                cnpj_results.append('|'.join([cnpj, doc.filename, doc.path, f"\"{doc.text}\""]))

        # Create a .txt to store CPFs.
        print('\nGravando CPFs em .txt ...')
        with open(os.path.join(f"{BASE_PATH}", "cpfs_encontrados.txt"), mode="a", newline='\n', encoding="utf-8") as new_file:
            # Write headers
            # new_file.write('cpf\tarquivo\tcaminho\ttexto\n')
            for result in tqdm(cpf_results):
                new_file.write(f'{result}\n')

        # Create a .txt to store CNPJs.
        print('\nGravando CNPJs ...')
        with open(os.path.join(f"{BASE_PATH}", "cnpjs_encontrados.txt"), mode="a", newline='\n', encoding="utf-8") as new_file:
            # Write headers
            # new_file.write('cnpj\tarquivo\tcaminho\ttexto\n')
            for result in tqdm(cnpj_results):
                new_file.write(f'{result}\n')

        sleep(3)

        # Generate html tables
        print('\nGerando tabelas...')
        to_table(os.path.join(f"{BASE_PATH}", "cpfs_encontrados.txt"))
        to_table(os.path.join(f"{BASE_PATH}", "cnpjs_encontrados.txt"))

        input_type = choose_type()
