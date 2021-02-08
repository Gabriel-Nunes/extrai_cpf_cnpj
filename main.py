# -*- coding: utf-8 -*-
from modules.utils import select_files
from modules.utils import procura_cnpj, procura_cpf, select_files, choose_type
from pdfminer.high_level import extract_text
import docx2txt
import os
import sys
import xlsxwriter


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

    for file in select_files(file_type[input_type]):
        doc = Doc(file, file_type[input_type])
        cpf_results.append([procura_cpf(doc.get_text()), doc.filename])
        cnpj_results.append([procura_cnpj(doc.get_text()), doc.filename])

    # Create a workbook and add a worksheet to store cpfs.
    workbook = xlsxwriter.Workbook('cpfs_encontrados.xlsx')
    worksheet = workbook.add_worksheet()

    # Iterate over the data and write it out row by row.
    row = 0
    col = 0
    for cpf, file_name in (cpf_results):
        worksheet.write(row, col, cpf)
        worksheet.write(row, col + 1, file_name)
        row += 1

    row = 0
    col = 0
    for cnpj, file_name in (cnpj_results):
        worksheet.write(row, col, cnpj)
        worksheet.write(row, col + 1, file_name)
        row += 1