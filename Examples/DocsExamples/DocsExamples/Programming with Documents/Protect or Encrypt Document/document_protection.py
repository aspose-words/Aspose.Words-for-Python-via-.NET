import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class DocumentProtection(docs_base.DocsExamplesBase):
    
        def test_protect(self) :
        
            #ExStart:ProtectDocument
            doc = aw.Document(docs_base.my_dir + "Document.docx")
            doc.protect(aw.ProtectionType.ALLOW_ONLY_FORM_FIELDS, "password")
            #ExEnd:ProtectDocument
        

        def test_unprotect(self) :
        
            #ExStart:UnprotectDocument
            doc = aw.Document(docs_base.my_dir + "Document.docx")
            doc.unprotect()
            #ExEnd:UnprotectDocument
        

        def test_get_protection_type(self) :
        
            #ExStart:GetProtectionType
            doc = aw.Document(docs_base.my_dir + "Document.docx")
            protectionType = doc.protection_type
            #ExEnd:GetProtectionType
        
    

if __name__ == '__main__':
        unittest.main()