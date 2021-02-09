# -*- coding: utf-8 -*-
from modules.utils import procura_cnpj, procura_cpf, select_files, choose_type
from pdfminer.high_level import extract_text
import docx2txt
import os
import sys
from tqdm import tqdm


class Doc:
    def __init__(self, file_path, file_type):
        self.type = file_type
        self.path = file_path
        self.filename = os.path.basename(file_path)
        self._text = ''

    # @property
    # def type(self):
    #     mime = magic.Magic(mime=True)
    #     self._type = mime.from_file(self.path)
    #     return self._type

    # TODO get the file content
    def get_text(self):
        if self.type == '*.pdf':
            try:
                text = extract_text(self.path)
                return text
            except TypeError:
                raise f"Erro ao ler o arquivo {0}. É um 'pdf' pesquisável?".format(self.filename)
        elif self.type == '*.docx':
            try:
                text = docx2txt.process(self.path)
                return text
            except TypeError:
                raise f"Erro ao ler o arquivo {0}. É um '.docx'?".format(self.filename)

    # TODO find cpfs on text
    def get_cpfs(self):
        cpfs = procura_cpf(self.text)
        return [(cpf, self.filename) for cpf in cpfs]

    # TODO find cnpjs on text
    def get_cnpjs(self):
        cpfs = procura_cpf(self.text)
        return [(cpf, self.filename) for cpf in cpfs]


if __name__ == '__main__':
    file_type = {'1': '*.pdf', '2': '*.docx'}
    input_type = choose_type()
    if input_type == '3':
        sys.exit()
    elif input_type not in ['1', '2', '3']:
        choose_type()
    BASE_PATH = os.getcwd()
    cpf_results = []
    cnpj_results = []

    files = select_files(file_type[input_type])
    print('\nLendo arquivos...\n')
    for file in tqdm(files):
        doc = Doc(file, file_type[input_type])
        text = doc.get_text()
        cpfs = procura_cpf(text)
        cnpjs = procura_cnpj(text)
        for cpf in cpfs:
            cpf_results.append(';'.join([cpf, doc.filename]))
        for cnpj in cnpjs:
            cnpj_results.append(';'.join([cnpj, doc.filename]))

    # Create a .csv to store cpfs.
    print('\nGravando CPFs...')
    with open("cpfs_encontrados.csv", mode="a", newline='\n') as new_file:
        for result in tqdm(cpf_results):
            new_file.write(f'{result}\n')

    # Create a .csv to store cnpjs.
    print('\nGravando CNPJs...')
    with open("cnpjs_encontrados.csv", mode="a", newline='\n') as new_file:
        for result in tqdm(cnpj_results):
            new_file.write(f'{result}\n')
