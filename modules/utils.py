# -*- coding: utf-8 -*-

from tkinter import *
from tkinter import filedialog
from validate_docbr import CPF, CNPJ
import re


def choose_type():
    return input('''
        \nQual o tipo de arquivo que deseja extrair os CPFs/CNPJs?

        1 - .pdf
        2 - .docx
        3 - .doc
        4 - sair

        Digite a opcao desejada: 

        ''')


def select_files(doc_type):
    '''
    Abre janela gráfica para seleção de arquivos
    :return: lista com os 'paths' absolutos dos arquivos
    '''
    root = Tk()
    root.withdraw()
    root.filenames = filedialog.askopenfilenames(initialdir="/", title="Selecione os arquivos...",
                                                filetypes=((f"{doc_type} files", f"{doc_type}"), ("all files", "*.*")))
    return list(root.filenames)


def procura_cpf(text: str) -> list:
    '''
    Retorna uma lista com CPFs encontrados, somente com números.
    :param text: str
    :return: list
    '''
    regexCPF = re.compile(r'\b\d{11,11}\b|\b\d\d\d.\d\d\d.\d\d\d-\d\d\b')
    cpfs = set([''.join([num for num in x if num.isalnum()]) for x in regexCPF.findall(text)])
    valid_cpfs = []
    for cpf in cpfs:
        cpf_num = CPF()
        if cpf_num.validate(cpf):
            valid_cpfs.append(cpf)
    return valid_cpfs


def procura_cnpj(text: str) -> list:
    '''
    Retorna uma lista com CNPJs válidos encontrados, somente com números.
    :param text: str
    :return: list
    '''
    regexCPF = re.compile(r'\b\d{14,14}\b|\b\d\d.\d\d\d.\d\d\d\/\d\d\d\d-\d\d\b')
    cnpjs = set([''.join([num for num in x if num.isalnum()]) for x in regexCPF.findall(text)])
    valid_cnpjs = []
    for cnpj in cnpjs:
        cnpj_num = CNPJ()
        if cnpj_num.validate(cnpj):
            valid_cnpjs.append(cnpj)
    return valid_cnpjs
