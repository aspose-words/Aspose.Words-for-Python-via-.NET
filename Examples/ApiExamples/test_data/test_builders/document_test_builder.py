# -*- coding: utf-8 -*-

import aspose.words as aw
from api_example_base import ApiExampleBase


class DocumentTestBuilder(ApiExampleBase):
    def __init__(self):
        self.m_document_stream = None
        self.m_document = aw.Document()
        m_document_stream = io.BytesIO(b'')
        self.m_document_bytes = []
        self.m_document_string = ""

    def with_document(self, doc):
        self.m_document = doc
        return self

    def with_document_stream(self, stream):
        self.m_document_stream = stream
        return self

    def with_document_bytes(self, doc_bytes):
        self.m_document_bytes = doc_bytes
        return self

    def with_document_string(self, doc_string):
        self.m_document_string = doc_string
        return self

    def build(self):
        return DocumentTestClass(self.m_document,self.m_document_stream,self.m_document_bytes,self.m_document_string)
