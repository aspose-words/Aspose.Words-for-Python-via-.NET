# -*- coding: utf-8 -*-



class DocumentTestClass:
    @property
    def document(self):
        pass

    @document.setter
    def document(self, value):
        pass

    @property
    def document_stream(self):
        pass

    @document_stream.setter
    def document_stream(self, value):
        pass

    @property
    def document_bytes(self):
        pass

    @document_bytes.setter
    def document_bytes(self, value):
        pass

    @property
    def document_string(self):
        pass

    @document_string.setter
    def document_string(self, value):
        pass

    def __init__(self, doc, doc_stream, doc_bytes, doc_string):
        document = doc
        document_stream = doc_stream
        document_bytes = doc_bytes
        document_string = doc_string
