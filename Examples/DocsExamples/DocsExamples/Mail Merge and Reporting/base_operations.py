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

    def test_create_mail_merge_template(self) :

        doc = self.create_mail_merge_template()
        doc.save(docs_base.artifacts_dir + "BaseOperations.create_mail_merge_template.docx")

    #ExStart:CreateMailMergeTemplate
    @staticmethod
    def create_mail_merge_template() :

        builder = aw.DocumentBuilder()

        # Insert a text input field the unique name of this field is "Hello", the other parameters define
        # what type of FormField it is, the format of the text, the field result and the maximum text length (0 = no limit)
        builder.insert_text_input("TextInput", aw.fields.TextFormFieldType.REGULAR, "", "Hello", 0)
        builder.insert_field("MERGEFIELD CustomerFirstName \\* MERGEFORMAT")

        builder.insert_text_input("TextInput1", aw.fields.TextFormFieldType.REGULAR, "", " ", 0)
        builder.insert_field("MERGEFIELD CustomerLastName \\* MERGEFORMAT")

        builder.insert_text_input("TextInput1", aw.fields.TextFormFieldType.REGULAR, "", " , ", 0)

        # Inserts a paragraph break into the document
        builder.insert_paragraph()

        # Insert mail body
        builder.insert_text_input("TextInput", aw.fields.TextFormFieldType.REGULAR, "", "Thanks for purchasing our ", 0)
        builder.insert_field("MERGEFIELD ProductName \\* MERGEFORMAT")

        builder.insert_text_input("TextInput", aw.fields.TextFormFieldType.REGULAR, "", ", please download your Invoice at ", 0)
        builder.insert_field("MERGEFIELD InvoiceURL \\* MERGEFORMAT")

        builder.insert_text_input("TextInput", aw.fields.TextFormFieldType.REGULAR, "", ". If you have any questions please call ", 0)
        builder.insert_field("MERGEFIELD Supportphone \\* MERGEFORMAT")

        builder.insert_text_input("TextInput", aw.fields.TextFormFieldType.REGULAR, "", ", or email us at ", 0)
        builder.insert_field("MERGEFIELD SupportEmail \\* MERGEFORMAT")

        builder.insert_text_input("TextInput", aw.fields.TextFormFieldType.REGULAR, "", ".", 0)

        builder.insert_paragraph()

        # Insert mail ending
        builder.insert_text_input("TextInput", aw.fields.TextFormFieldType.REGULAR, "", "Best regards,", 0)
        builder.insert_break(aw.BreakType.LINE_BREAK)
        builder.insert_field("MERGEFIELD EmployeeFullname \\* MERGEFORMAT")

        builder.insert_text_input("TextInput1", aw.fields.TextFormFieldType.REGULAR, "", " ", 0)
        builder.insert_field("MERGEFIELD EmployeeDepartment \\* MERGEFORMAT")

        return builder.document
    #ExEnd:CreateMailMergeTemplate


if __name__ == '__main__':
    unittest.main()