import unittest
import os
import sys
from datetime import date, datetime

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class DocumentPropertiesAndVariables(docs_base.DocsExamplesBase):

    def test_get_variables(self):

        #ExStart:GetVariables
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        doc.variables.add("my_var", "test")

        print(doc.variables.get_by_name("my_var"))

        #ExEnd:GetVariables


    def test_enumerate_properties(self):

        #ExStart:EnumerateProperties
        doc = aw.Document(docs_base.my_dir + "Properties.docx")

        print("1. Document name: 0", doc.original_file_name)
        print("2. Built-in Properties")

        for prop in doc.built_in_document_properties:
            print("0: 1", prop.name, prop.value)

        print("3. Custom Properties")

        for prop in doc.custom_document_properties:
            print("0: 1", prop.name, prop.value)
        #ExEnd:EnumerateProperties


    def test_add_custom_document_properties(self):

        #ExStart:AddCustomDocumentProperties
        doc = aw.Document(docs_base.my_dir + "Properties.docx")

        custom_document_properties = doc.custom_document_properties

        if custom_document_properties.get_by_name("Authorized") is not None:
            return

        custom_document_properties.add("Authorized", True)
        custom_document_properties.add("Authorized By", "John Smith")
        custom_document_properties.add("Authorized Date", datetime.today())
        custom_document_properties.add("Authorized Revision", doc.built_in_document_properties.revision_number)
        custom_document_properties.add("Authorized Amount", 123.45)
        #ExEnd:AddCustomDocumentProperties


    def test_remove_custom_document_properties(self):

        #ExStart:CustomRemove
        doc = aw.Document(docs_base.my_dir + "Properties.docx")
        doc.custom_document_properties.remove("Authorized Date")
        #ExEnd:CustomRemove


    def test_remove_personal_information(self):

        #ExStart:RemovePersonalInformation
        doc = aw.Document(docs_base.my_dir + "Properties.docx")
        doc.remove_personal_information = True

        doc.save(docs_base.artifacts_dir + "DocumentPropertiesAndVariables.remove_personal_information.docx")
        #ExEnd:RemovePersonalInformation


    def test_configuring_link_to_content(self):

        #ExStart:ConfiguringLinkToContent
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_bookmark("MyBookmark")
        builder.writeln("Text inside a bookmark.")
        builder.end_bookmark("MyBookmark")

        # Retrieve a list of all custom document properties from the file.
        custom_properties = doc.custom_document_properties
        # Add linked to content property.
        custom_property = custom_properties.add_link_to_content("Bookmark", "MyBookmark")
        custom_property = custom_properties.get_by_name("Bookmark")

        is_linked_to_content = custom_property.is_link_to_content

        link_source = custom_property.link_source

        custom_property_value = custom_property.value
        #ExEnd:ConfiguringLinkToContent


    def test_convert_between_measurement_units(self):

        #ExStart:ConvertBetweenMeasurementUnits
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        page_setup = builder.page_setup
        page_setup.top_margin = aw.ConvertUtil.inch_to_point(1.0)
        page_setup.bottom_margin = aw.ConvertUtil.inch_to_point(1.0)
        page_setup.left_margin = aw.ConvertUtil.inch_to_point(1.5)
        page_setup.right_margin = aw.ConvertUtil.inch_to_point(1.5)
        page_setup.header_distance = aw.ConvertUtil.inch_to_point(0.2)
        page_setup.footer_distance = aw.ConvertUtil.inch_to_point(0.2)
        #ExEnd:ConvertBetweenMeasurementUnits


    def test_use_control_characters(self):

        #ExStart:UseControlCharacters
        text = "test\r"
        # Replace "\r" control character with "\r\n".
        replace = text.replace(aw.ControlChar.CR, aw.ControlChar.CR_LF)
        #ExEnd:UseControlCharacters


if __name__ == '__main__':
    unittest.main()
