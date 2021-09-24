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
    
        @unittest.skip("cannot iterate though variables for now.")
        def test_get_variables(self) :
        
            #ExStart:GetVariables
            doc = aw.Document(docs_base.my_dir + "Document.docx")
            
            doc.variables.add("test", "test")

            variables = ""
            for entry in doc.variables :
            
                name = entry.key
                value = entry.value
                if (variables == "") :
                    variables = "Name: " + name + "," + "Value: " + value
                else :
                    variables = variables + "Name: " + name + "," + "Value: " + value
                
            #ExEnd:GetVariables

            print("\nDocument have following variables " + variables)
        

        def test_enumerate_properties(self) :
        
            #ExStart:EnumerateProperties            
            doc = aw.Document(docs_base.my_dir + "Properties.docx")
            
            print("1. Document name: 0", doc.original_file_name)
            print("2. Built-in Properties")
            
            for prop in doc.built_in_document_properties :
                print("0 : 1", prop.name, prop.value)

            print("3. Custom Properties")
            
            for prop in doc.custom_document_properties :
                print("0 : 1", prop.name, prop.value)
            #ExEnd:EnumerateProperties
        

        def test_add_custom_document_properties(self) :
        
            #ExStart:AddCustomDocumentProperties            
            doc = aw.Document(docs_base.my_dir + "Properties.docx")

            customDocumentProperties = doc.custom_document_properties
            
            if (customDocumentProperties.get_by_name("Authorized") != None) :
                return
            
            customDocumentProperties.add("Authorized", True)
            customDocumentProperties.add("Authorized By", "John Smith")
            customDocumentProperties.add("Authorized Date", datetime.today())
            customDocumentProperties.add("Authorized Revision", doc.built_in_document_properties.revision_number)
            customDocumentProperties.add("Authorized Amount", 123.45)
            #ExEnd:AddCustomDocumentProperties
        

        def test_remove_custom_document_properties(self) :
        
            #ExStart:CustomRemove            
            doc = aw.Document(docs_base.my_dir + "Properties.docx")
            doc.custom_document_properties.remove("Authorized Date")
            #ExEnd:CustomRemove
        

        def test_remove_personal_information(self) :
        
            #ExStart:RemovePersonalInformation            
            doc = aw.Document(docs_base.my_dir + "Properties.docx")  
            doc.remove_personal_information = True 

            doc.save(docs_base.artifacts_dir + "DocumentPropertiesAndVariables.remove_personal_information.docx")
            #ExEnd:RemovePersonalInformation
        

        def test_configuring_link_to_content(self) :
        
            #ExStart:ConfiguringLinkToContent            
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc)
            
            builder.start_bookmark("MyBookmark")
            builder.writeln("Text inside a bookmark.")
            builder.end_bookmark("MyBookmark")

            # Retrieve a list of all custom document properties from the file.
            customProperties = doc.custom_document_properties
            # Add linked to content property.
            customProperty = customProperties.add_link_to_content("Bookmark", "MyBookmark")
            customProperty = customProperties.get_by_name("Bookmark")

            isLinkedToContent = customProperty.is_link_to_content
            
            linkSource = customProperty.link_source
            
            customPropertyValue = customProperty.value
            #ExEnd:ConfiguringLinkToContent
        

        def test_convert_between_measurement_units(self) :
        
            #ExStart:ConvertBetweenMeasurementUnits
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc)

            pageSetup = builder.page_setup
            pageSetup.top_margin = aw.ConvertUtil.inch_to_point(1.0)
            pageSetup.bottom_margin = aw.ConvertUtil.inch_to_point(1.0)
            pageSetup.left_margin = aw.ConvertUtil.inch_to_point(1.5)
            pageSetup.right_margin = aw.ConvertUtil.inch_to_point(1.5)
            pageSetup.header_distance = aw.ConvertUtil.inch_to_point(0.2)
            pageSetup.footer_distance = aw.ConvertUtil.inch_to_point(0.2)
            #ExEnd:ConvertBetweenMeasurementUnits
        

        def test_use_control_characters(self) :
        
            #ExStart:UseControlCharacters
            text = "test\r"
            # Replace "\r" control character with "\r\n".
            replace = text.replace(aw.ControlChar.CR, aw.ControlChar.CR_LF)
            #ExEnd:UseControlCharacters
        
    

if __name__ == '__main__':
        unittest.main()