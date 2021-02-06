# -*- coding: utf-8 -*-
from modules.utils import select_files
from winmagic import magic
import os


class Doc:
    def __init__(self, file_path):
        self._type = ''
        self.path = file_path
        self.filename = os.path.basename(file_path)
        self._text = ''

    @property
    def type(self):
        mime = magic.Magic(mime=True)
        self._type = mime.from_file(self.path)
        return self._type

    # TODO get the file content
    @property
    def text(self):
        return

    # TODO find cpfs on text

    # TODO find cnpjs on text


if __name__ == '__main__':
    doc_teste = Doc('test\\test.pdf')

