import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class BaseOperations(docs_base.DocsExamplesBase):
    
    def test_simple_mail_merge(self) :
        
        #ExStart:SimpleMailMerge
        # Include the code for our template.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create Merge Fields.
        builder.insert_field(" MERGEFIELD CustomerName ")
        builder.insert_paragraph()
        builder.insert_field(" MERGEFIELD Item ")
        builder.insert_paragraph()
        builder.insert_field(" MERGEFIELD Quantity ")

        # Fill the fields in the document with user data.
        doc.mail_merge.execute([ "CustomerName", "Item", "Quantity" ],
            [ "John Doe", "Hawaiian", "2" ])

        doc.save(docs_base.artifacts_dir + "BaseOperations.simple_mail_merge.docx")
        #ExEnd:SimpleMailMerge
        

    def test_use_if_else_mustache(self) :
        
        #ExStart:UseOfifelseMustacheSyntax
        doc = aw.Document(docs_base.my_dir + "Mail merge destinations - Mustache syntax.docx")

        doc.mail_merge.use_non_merge_fields = True
        doc.mail_merge.execute([ "GENDER" ], [ "MALE" ])

        doc.save(docs_base.artifacts_dir + "BaseOperations.if_else_mustache.docx")
        #ExEnd:UseOfifelseMustacheSyntax
        

if __name__ == '__main__':
    unittest.main()