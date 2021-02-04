# -*- coding: utf-8 -*-

import csv
import os
from tkinter import *
from tkinter import filedialog
from unicodedata import normalize


def select_files():
    root = Tk()
    root.withdraw()
    root.filenames = filedialog.askopenfilenames(initialdir="/", title="Selecione os arquivos CSV...",
                                                filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))
    return list(root.filenames)


def text_files_on_folder(folder):
    """
    Retorna uma lista com os arquivos CSV de 'folder'.
    Importante: o arquivo deve conter a extensão '.csv'
    """
    # mime = magic.Magic(mime=True)
    # text_files = []
    # for folder, subfolders, files in os.walk(folder):
    #     for file in files:
    #         file_url = os.path.join(folder, file)
    #         if ('text/plain' or 'text/x-Algol68') in mime.from_file(file_url):
    #             text_files.append(file_url)
    # return text_files
    text_files = []
    for folder, subfolders, files in os.walk(folder):
        for file in files:
            if not (('.csv' in file) or ('.txt' in file)):
                continue
            text_files.append(file)
    return text_files


def limpa_cpf_cnpj(cpf):
    return ''.join([num for num in cpf if num.isalnum()])


def normaliza(txt):
    """
    Devolve cópia de uma str substituindo os caracteres
    acentuados pelos seus equivalentes não acentuados.

    Remove também espaços anteriores, posteriores e duplicados.

    ATENÇÃO: caracteres gráficos não ASCII e não alfa-numéricos,
    tais como bullets, travessões, aspas assimétricas, etc.
    são simplesmente removidos!

        >>> normaliza('[ACENTUAÇÃO] ç: áàãâä! éèêë? íìîï, óòõôö; úùûü.')
        '[ACENTUACAO] c: aaaaa! eeee? iiii, ooooo; uuuu.'
    """
    result = normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')
    result.strip()
    result = re.sub('\s{2,}', ' ', result)
    return result


def abrevia(frase, max):
    """
    Abrevia uma frase para o máximo de caracteres definidos em 'max'
    """
    listaFrase = frase.split(' ')
    maxLetras = 4
    novaFrase = frase

    while len(novaFrase) > max:
        for i in range(len(listaFrase)):
            palavra = listaFrase[i]
            if len(palavra) > 1:
                listaFrase[i] = palavra[:maxLetras]
            novaFrase = ' '.join(listaFrase)
            if len(novaFrase) <= max:
                break
        maxLetras -= 1
    return novaFrase


def numero_da_coluna(letra):
    """
    Retorna o número correspondente à coluna da planilha.
    :param letra: letra correspondente à coluna da planilha ('a', 'b', 'c', ...)
    :return: retorna o número correspondente à letra ('a' = 0, 'b' = 1, 'c' = 2, ...)
    """
    letter = letra.lower()
    alfabeto = ('a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u',
                'v', 'w', 'x', 'y', 'z', 'aa', 'ab', 'ac', 'ad', 'ae', 'af', 'ag', 'ah', 'ai', 'aj')
    return alfabeto.index(letter)


def procura_pessoa_fisica(text):
    regexPF = re.compile(r'\S\S\S.\d\d\d\.\d\d\d-\S\S|PESSOA FÍSICA')
    if regexPF.search(text):
        return True
    else:
        return False


def procura_pessoa_juridica(text):
    regexPJ = re.compile(r'\d\d\d\d\d\d\d\d\d\d\d\d\d\d|CNPJ - PESSOA JURÍDICA')
    regexIdentificacaoEspecial = re.compile(r'IDENTIFICAÇÃO ESPECIAL', re.I)
    regexInscricaoGenerica = re.compile(r'\WOUTROS\W', re.I)
    if regexIdentificacaoEspecial.search(text) or regexInscricaoGenerica.search(text):  # para linhas de Identificação especial, não é PF nem PJ
        return False
    if regexPJ.search(text):
        return True
    else:
        return False


def iniciais_maiusculas(nome):
    name = nome.lower()
    p = ['a', 'o', 'e', 'da', 'do', 'dos', 'das', 'de', 'di', 'em', 'para', 'com', 'por']
    return ' '.join(list(map(lambda w: w.capitalize() if w not in p else w, name.split())))


def moeda(x):
    """
    retorna uma string no formato "00,00" (entre aspas duplas)
    :param x: string de valor monetário no formato 00.00
    :return: string de valor monetário no formato "00,00"
    """
    # regex para moedas no formato $ XX.XX
    regex2 = re.compile(r'\d+\.\d\b')
    regex3 = re.compile(r'\d+\.\d\d\b')

    if regex3.search(x):
        return x.replace('.', ',')
    elif regex2.search(x):
        x = x + '0'
        return x.replace('.', ',')
    else:
        return x + ',00'


def acha_cep(text):
    try:
        cep = re.findall(r'\D?(\d\d\d\d\d-\d\d\d)\D?', text)[0]
        return cep
    except IndexError:
        print('''
        Erro na fonte de dados.
        Verifique se há linhas de cabeçalho repetidas na planilha e
        tente novamente.
        ''')


def zeros_a_esquerda(casas_decimais, num):
    """
    Adiciona zeros à esquerda da string 'num' até ficar com 'x' casas decimais.
    :param casas_decimais: número de casas decimais
    :param num: string com texto a ser modificado
    :return: 'num' com 'x' casas decimais
    """
    while len(num) < casas_decimais:
        num = '0' + num
    return num


def get_csv_data(file, separator):
    """
    usage: get_csv_data(file_path, separator)

    filepath: path do arquivo CSV que deseja extrair os dados
    separator: ex: ',' ou ';'

    return: lista com os dados de 'file'
    """
    return list(csv.reader(open(file), delimiter=separator))


def save_data_to_csv(list, filename):
    """
    usage: save_data_to_csv(list, filename)

    list: lista com os dados a serem salvos no CSV
    filename: path absoluto do novo CSV a ser salvo
    """
    openedFile = open(filename, mode='w', encoding='latin-1', newline='')
    writer = csv.writer(openedFile)
    writer.writerows(list)
    openedFile.close()


def tel_i2(fone):
    """

    Esta função edita numeros de telefone para o padrão do i2 (XXXXXXXXXX)

    fone: número de telefone a ser editado (ex: 11-9362 2313)
    :return Retorna o 'fone' no padrão i2 (1193622513), sem os caracteres especiais e espaços

    OBS: a função não insere dígitos faltantes.
    """
    return ''.join([num for num in fone if num.isalnum()])


def join_csv(csvs_to_join, final_CSV, sep):
    """
    Junta o conteúdo de vários CSV com o mesmo separador ('sep') em um único arquivo 'final_CSV'.
    Retorna o conteúdo do arquivo criado em uma lista 'finalCsvData'
    :param 'csvs_to_join' -> lista com arquivos .csv
    :param 'final_CSV' -> caminho do arquivo final a ser salvo
    :param 'sep' -> separador (normalmente ',' ou ';')
    """
    newData = []
    folder = os.path.dirname(csvs_to_join[0])
    newData.append(get_csv_data(csvs_to_join[0], sep)[0])  # Salva o cabeçalho do 1º arquivo em newData
    for file in csvs_to_join:
        # Recebe uma lista a partir das linhas de cada CSV
        # e salva em newData, exceto o cabeçalho
        fileData = get_csv_data(file, sep)
        for row in fileData[1:]:
            if any(row):
                newData.append(row)
    save_data_to_csv(newData, os.path.join(folder, final_CSV) + '.csv')
    return get_csv_data(os.path.join(folder, final_CSV) + '.csv', ',')
