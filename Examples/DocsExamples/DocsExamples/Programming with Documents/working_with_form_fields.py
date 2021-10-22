import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithFormFields(docs_base.DocsExamplesBase):

    def test_insert_form_fields(self) :

        #ExStart:InsertFormFields
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        items =  ["One", "Two", "Three"]
        builder.insert_combo_box("DropDown", items, 0)
        #ExEnd:InsertFormFields

    def test_document_builder_insert_text_input_form_field(self) :

        #ExStart:DocumentBuilderInsertTextInputFormField
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_text_input("TextInput", aw.fields.TextFormFieldType.REGULAR, "", "Hello", 0)

        doc.save(docs_base.artifacts_dir + "WorkingWithFormFields.document_builder_insert_text_input_form_field.docx")
        #ExEnd:DocumentBuilderInsertTextInputFormField

    def test_document_builder_insert_check_box_form_field(self) :

        #ExStart:DocumentBuilderInsertCheckBoxFormField
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_check_box("CheckBox", True, True, 0)

        doc.save(docs_base.artifacts_dir + "WorkingWithFormFields.document_builder_insert_check_box_form_field.docx")
        #ExEnd:DocumentBuilderInsertCheckBoxFormField

    def test_document_builder_insert_combo_box_form_field(self) :

        #ExStart:DocumentBuilderInsertComboBoxFormField
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        items =  ["One", "Two", "Three"]
        builder.insert_combo_box("DropDown", items, 0)

        doc.save(docs_base.artifacts_dir + "WorkingWithFormFields.document_builder_insert_combo_box_form_field.docx")
        #ExEnd:DocumentBuilderInsertComboBoxFormField

    def test_form_fields_work_with_properties(self) :

        #ExStart:FormFieldsWorkWithProperties
        doc = aw.Document(docs_base.my_dir + "Form fields.docx")
        form_field = doc.range.form_fields[3]

        if form_field.type == aw.fields.FieldType.FIELD_FORM_TEXT_INPUT :
            form_field.result = "My name is " + form_field.name
        #ExEnd:FormFieldsWorkWithProperties


    def test_form_fields_get_form_fields_collection(self) :

        #ExStart:FormFieldsGetFormFieldsCollection
        doc = aw.Document(docs_base.my_dir + "Form fields.docx")

        form_fields = doc.range.form_fields
        #ExEnd:FormFieldsGetFormFieldsCollection


    def test_form_fields_get_by_name(self) :

        #ExStart:FormFieldsFontFormatting
        #ExStart:FormFieldsGetByName
        doc = aw.Document(docs_base.my_dir + "Form fields.docx")

        document_form_fields = doc.range.form_fields

        form_field1 = document_form_fields[3]
        #formField2 = documentFormFields["Text2"]
        #ExEnd:FormFieldsGetByName

        form_field1.font.size = 20
        #formField2.font.color = Color.red
        #ExEnd:FormFieldsFontFormatting


if __name__ == '__main__':
    unittest.main()
