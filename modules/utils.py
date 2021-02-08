# -*- coding: utf-8 -*-

import csv
import os
import re
from tkinter import *
from tkinter import filedialog
from unicodedata import normalize


def choose_type():
    return input('''
        \nQual o tipo de arquivo que deseja extrair os CPFs/CNPJs?

        1 - .pdf
        2 - .docx
        3 - sair

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
                                                filetypes=((f"{doc_type}", "*.*"), ("all files", "*.*")))
    return list(root.filenames)


def procura_cpf(text: str) -> list:
    '''
    Retorna uma lista com CPFs encontrados em uma string, somente com números.
    :param text: str
    :return: list
    '''
    regexCPF = re.compile(r'\D?\d\d\d.?\d\d\d\.?\d\d\d-?\d\d\D?')
    return [''.join([num for num in x if num.isalnum()]) for x in regexCPF.findall(text)]


def procura_cnpj(text: str) -> list:
    '''
    Retorna uma lista com CNPJs encontrados em um texto, sem caracteres especiais.
    :param text: str
    :return: list
    '''
    regexCNPJ = re.compile(r'\D?\d\d.?\d\d\d\.?\d\d\d\/?\d\d\d\d-?\d\d\D?')
    return [''.join([num for num in x if num.isalnum()]) for x in regexCNPJ.findall(text)]
