import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw
#import system.text

class WorkingWithLoadOptions(docs_base.DocsExamplesBase):
    
    def test_update_dirty_fields(self) :
        
        #ExStart:UpdateDirtyFields
        loadOptions = aw.loading.LoadOptions()
        loadOptions.update_dirty_fields = True 

        doc = aw.Document(docs_base.my_dir + "Dirty field.docx", loadOptions)

        doc.save(docs_base.artifacts_dir + "WorkingWithLoadOptions.update_dirty_fields.docx")
        #ExEnd:UpdateDirtyFields
        

    def test_load_encrypted_document(self) :
        
        #ExStart:LoadSaveEncryptedDoc
        #ExStart:OpenEncryptedDocument
        doc = aw.Document(docs_base.my_dir + "Encrypted.docx", aw.loading.LoadOptions("docPassword"))
        #ExEnd:OpenEncryptedDocument

        doc.save(docs_base.artifacts_dir + "WorkingWithLoadOptions.load_and_save_encrypted_odt.odt", aw.saving.OdtSaveOptions("newPassword"))
        #ExEnd:LoadSaveEncryptedDoc
        

    def test_convert_shape_to_office_math(self) :
        
        #ExStart:ConvertShapeToOfficeMath
        loadOptions = aw.loading.LoadOptions()
        loadOptions.convert_shape_to_office_math = True 

        doc = aw.Document(docs_base.my_dir + "Office math.docx", loadOptions)

        doc.save(docs_base.artifacts_dir + "WorkingWithLoadOptions.convert_shape_to_office_math.docx", aw.SaveFormat.DOCX)
        #ExEnd:ConvertShapeToOfficeMath
        

    def test_set_ms_word_version(self) :
        
        #ExStart:SetMSWordVersion
        # Create a new LoadOptions object, which will load documents according to MS Word 2019 specification by default
        # and change the loading version to Microsoft Word 2010.
        loadOptions = aw.loading.LoadOptions()
        loadOptions.msw_version = aw.settings.MsWordVersion.WORD2010 
            
        doc = aw.Document(docs_base.my_dir + "Document.docx", loadOptions)

        doc.save(docs_base.artifacts_dir + "WorkingWithLoadOptions.set_ms_word_version.docx")
        #ExEnd:SetMSWordVersion
        

    def test_use_temp_folder(self) :
        
        #ExStart:UseTempFolder  
        loadOptions = aw.loading.LoadOptions()
        loadOptions.temp_folder = docs_base.artifacts_dir 

        doc = aw.Document(docs_base.my_dir + "Document.docx", loadOptions)
        #ExEnd:UseTempFolder  
        

    def test_load_with_encoding(self) :
        
        #ExStart:LoadWithEncoding
        loadOptions = aw.loading.LoadOptions()
        loadOptions.encoding = "utf-7"

        doc = aw.Document(docs_base.my_dir + "Encoded in UTF-7.txt", loadOptions)
        #ExEnd:LoadWithEncoding
        

    def test_skip_pdf_images(self) :
        
        #ExStart:SkipPdfImages
        loadOptions = aw.loading.PdfLoadOptions()
        loadOptions.skip_pdf_images = True 

        doc = aw.Document(docs_base.my_dir + "Pdf Document.pdf", loadOptions)
        #ExEnd:SkipPdfImages
        

    def test_convert_metafiles_to_png(self) :
        
        #ExStart:ConvertMetafilesToPng
        loadOptions = aw.loading.LoadOptions()
        loadOptions.convert_metafiles_to_png = True 

        doc = aw.Document(docs_base.my_dir + "WMF with image.docx", loadOptions)
        #ExEnd:ConvertMetafilesToPng
        

    def test_load_chm(self) :
        
        #ExStart:LoadCHM
        loadOptions = aw.loading.LoadOptions()
        loadOptions.encoding = "windows-1251" 

        doc = aw.Document(docs_base.my_dir + "HTML help.chm", loadOptions)
        #ExEnd:LoadCHM
        
    

if __name__ == '__main__':
    unittest.main()