import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithTxtSaveOptions(docs_base.DocsExamplesBase):
    
    def test_add_bidi_marks(self) :
        
        #ExStart:AddBidiMarks
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello world!")
        builder.paragraph_format.bidi = True
        builder.writeln("שלום עולם!")
        builder.writeln("مرحبا بالعالم!")

        saveOptions = aw.saving.TxtSaveOptions()
        saveOptions.add_bidi_marks = True 

        doc.save(docs_base.artifacts_dir + "WorkingWithTxtSaveOptions.add_bidi_marks.txt", saveOptions)
        #ExEnd:AddBidiMarks
        

    def test_use_tab_character_per_level_for_list_indentation(self) :
        
        #ExStart:UseTabCharacterPerLevelForListIndentation
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create a list with three levels of indentation.
        builder.list_format.apply_number_default()
        builder.writeln("Item 1")
        builder.list_format.list_indent()
        builder.writeln("Item 2")
        builder.list_format.list_indent() 
        builder.write("Item 3")

        saveOptions = aw.saving.TxtSaveOptions()
        saveOptions.list_indentation.count = 1
        #saveOptions.list_indentation.character = '\t'

        doc.save(docs_base.artifacts_dir + "WorkingWithTxtSaveOptions.use_tab_character_per_level_for_list_indentation.txt", saveOptions)
        #ExEnd:UseTabCharacterPerLevelForListIndentation
        

    def test_use_space_character_per_level_for_list_indentation(self) :
        
        #ExStart:UseSpaceCharacterPerLevelForListIndentation
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create a list with three levels of indentation.
        builder.list_format.apply_number_default()
        builder.writeln("Item 1")
        builder.list_format.list_indent()
        builder.writeln("Item 2")
        builder.list_format.list_indent() 
        builder.write("Item 3")

        saveOptions = aw.saving.TxtSaveOptions()
        saveOptions.list_indentation.count = 3
        #saveOptions.list_indentation.character = ' '

        doc.save(docs_base.artifacts_dir + "WorkingWithTxtSaveOptions.use_space_character_per_level_for_list_indentation.txt", saveOptions)
        #ExEnd:UseSpaceCharacterPerLevelForListIndentation
        
    def test_export_headers_footers_mode(self) :

        #ExStart:ExportHeadersFootersMode
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        options = aw.saving.TxtSaveOptions()
        options.save_format = aw.SaveFormat.TEXT

        # All headers and footers are placed at the very end of the output document.
        options.export_headers_footers_mode = aw.saving.TxtExportHeadersFootersMode.ALL_AT_END
        doc.save(docs_base.artifacts_dir + "WorkingWithTxtSaveOptions.export_headers_footers_mode_A.txt", options)

        # Only primary headers and footers are exported at the beginning and end of each section.
        options.export_headers_footers_mode = aw.saving.TxtExportHeadersFootersMode.PRIMARY_ONLY
        doc.save(docs_base.artifacts_dir + "WorkingWithTxtSaveOptions.export_headers_footers_mode_B.txt", options)

        # No headers and footers are exported.
        options.export_headers_footers_mode = aw.saving.TxtExportHeadersFootersMode.NONE
        doc.save(docs_base.artifacts_dir + "WorkingWithTxtSaveOptions.export_headers_footers_mode_C.txt", options)
        #ExEnd:ExportHeadersFootersMode
    

if __name__ == '__main__':
    unittest.main()