# -*- coding: utf-8 -*-

import csv
import os
import re
from tkinter import *
from tkinter import filedialog
from unicodedata import normalize


def select_files():
    '''
    Abre janela gráfica para seleção de arquivos
    :return: lista com os 'paths' absolutos dos arquivos
    '''
    root = Tk()
    root.withdraw()
    root.filenames = filedialog.askopenfilenames(initialdir="/", title="Selecione os arquivos...",
                                                filetypes=(("all files", "*.*"), ("all files", "*.*")))
    return list(root.filenames)


def procura_cpf(text: str) -> list:
    '''
    Retorna uma lista com CPFs encontrados em uma string, somente com números.
    :param text: str
    :return: list
    '''
    regexCPF = re.compile(r'\d\d\d.?\d\d\d\.?\d\d\d-?\d\d')
    return [''.join([num for num in x if num.isalnum()]) for x in regexCPF.findall(text)]


def procura_cnpj(text: str) -> list:
    '''
    Retorna uma lista com CNPJs encontrados em um texto, sem caracteres especiais.
    :param text: str
    :return: list
    '''
    regexCNPJ = re.compile(r'\d\d.?\d\d\d\.?\d\d\d\/?\d\d\d\d-?\d\d')
    return [''.join([num for num in x if num.isalnum()]) for x in regexCNPJ.findall(text)]
