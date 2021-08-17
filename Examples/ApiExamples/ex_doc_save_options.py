import unittest
from datetime import date, datetime
import os

import api_example_base as aeb
from document_helper import DocumentHelper

import aspose.words as aw

class ExDocSaveOptions(aeb.ApiExampleBase):
    
    def test_save_as_doc(self) :
        
        #ExStart
        #ExFor:DocSaveOptions
        #ExFor:DocSaveOptions.#ctor
        #ExFor:DocSaveOptions.#ctor(SaveFormat)
        #ExFor:DocSaveOptions.password
        #ExFor:DocSaveOptions.save_format
        #ExFor:DocSaveOptions.save_routing_slip
        #ExSummary:Shows how to set save options for older Microsoft Word formats.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.write("Hello world!")

        options = aw.saving.DocSaveOptions(aw.SaveFormat.DOC)
            
        # Set a password which will protect the loading of the document by Microsoft Word or Aspose.words.
        # Note that this does not encrypt the contents of the document in any way.
        options.password = "MyPassword"

        # If the document contains a routing slip, we can preserve it while saving by setting this flag to True.
        options.save_routing_slip = True

        doc.save(aeb.artifacts_dir + "DocSaveOptions.save_as_doc.doc", options)

        # To be able to load the document,
        # we will need to apply the password we specified in the DocSaveOptions object in a LoadOptions object.
        with self.assertRaises(RuntimeError) as ex:
            doc = aw.Document(aeb.artifacts_dir + "DocSaveOptions.save_as_doc.doc")

        loadOptions = aw.loading.LoadOptions("MyPassword")
        doc = aw.Document(aeb.artifacts_dir + "DocSaveOptions.save_as_doc.doc", loadOptions)

        self.assertEqual("Hello world!", doc.get_text().strip())
        #ExEnd
        

    def test_temp_folder(self) :
        
        #ExStart
        #ExFor:SaveOptions.temp_folder
        #ExSummary:Shows how to use the hard drive instead of memory when saving a document.
        doc = aw.Document(aeb.my_dir + "Rendering.docx")

        # When we save a document, various elements are temporarily stored in memory as the save operation is taking place.
        # We can use this option to use a temporary folder in the local file system instead,
        # which will reduce our application's memory overhead.
        options = aw.saving.DocSaveOptions()
        options.temp_folder = aeb.artifacts_dir + "TempFiles"

        # The specified temporary folder must exist in the local file system before the save operation.
        if not os.path.exists(options.temp_folder):
            os.makedirs(options.temp_folder)

        doc.save(aeb.artifacts_dir + "DocSaveOptions.temp_folder.doc", options)

        # The folder will persist with no residual contents from the load operation.
        self.assertTrue(len(os.listdir(options.temp_folder) ) == 0)
        #ExEnd
        

    def test_picture_bullets(self) :
        
        #ExStart
        #ExFor:DocSaveOptions.save_picture_bullet
        #ExSummary:Shows how to omit PictureBullet data from the document when saving.
        doc = aw.Document(aeb.my_dir + "Image bullet points.docx")
        self.assertNotEqual(doc.lists[0].list_levels[0].image_data, None) #ExSkip

        # Some word processors, such as Microsoft Word 97, are incompatible with PictureBullet data.
        # By setting a flag in the SaveOptions object,
        # we can convert all image bullet points to ordinary bullet points while saving.
        saveOptions = aw.saving.DocSaveOptions(aw.SaveFormat.DOC)
        saveOptions.save_picture_bullet = False

        doc.save(aeb.artifacts_dir + "DocSaveOptions.picture_bullets.doc", saveOptions)
        #ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocSaveOptions.picture_bullets.doc")

        self.assertIsNone(doc.lists[0].list_levels[0].image_data)
        

    def test_update_last_printed_property(self) :
        for isUpdateLastPrintedProperty in [True, False]:
            with self.subTest(isUpdateLastPrintedProperty=isUpdateLastPrintedProperty):    
       
                #ExStart
                #ExFor:SaveOptions.update_last_printed_property
                #ExSummary:Shows how to update a document's "Last printed" property when saving.
                doc = aw.Document()
                doc.built_in_document_properties.last_printed = date(2019, 12, 20)

                # This flag determines whether the last printed date, which is a built-in property, is updated.
                # If so, then the date of the document's most recent save operation
                # with this SaveOptions object passed as a parameter is used as the print date.
                saveOptions = aw.saving.DocSaveOptions()
                saveOptions.update_last_printed_property = isUpdateLastPrintedProperty

                # In Microsoft Word 2003, this property can be found via File -> Properties -> Statistics -> Printed.
                # It can also be displayed in the document's body by using a PRINTDATE field.
                doc.save(aeb.artifacts_dir + "DocSaveOptions.update_last_printed_property.doc", saveOptions)

                # Open the saved document, then verify the value of the property.
                doc = aw.Document(aeb.artifacts_dir + "DocSaveOptions.update_last_printed_property.doc")

                self.assertNotEqual(isUpdateLastPrintedProperty, date(2019, 12, 20) == doc.built_in_document_properties.last_printed.date())
                #ExEnd
        

    def test_update_created_time_property(self):
        for isUpdateCreatedTimeProperty in [True, False]:
            with self.subTest(isUpdateCreatedTimeProperty=isUpdateCreatedTimeProperty):    
        
                #ExStart
                #ExFor:SaveOptions.update_last_printed_property
                #ExSummary:Shows how to update a document's "CreatedTime" property when saving.
                doc = aw.Document()
                doc.built_in_document_properties.created_time = date(2019, 12, 20)

                # This flag determines whether the created time, which is a built-in property, is updated.
                # If so, then the date of the document's most recent save operation
                # with this SaveOptions object passed as a parameter is used as the created time.
                saveOptions = aw.saving.DocSaveOptions()
                saveOptions.update_created_time_property = isUpdateCreatedTimeProperty

                doc.save(aeb.artifacts_dir + "DocSaveOptions.update_created_time_property.docx", saveOptions)

                # Open the saved document, then verify the value of the property.
                doc = aw.Document(aeb.artifacts_dir + "DocSaveOptions.update_created_time_property.docx")

                self.assertNotEqual(isUpdateCreatedTimeProperty, date(2019, 12, 20) == doc.built_in_document_properties.created_time.date())
                #ExEnd
        

    def test_always_compress_metafiles(self):
        for compressAllMetafiles in [True, False]:
            with self.subTest(compressAllMetafiles=compressAllMetafiles): 
                
                #ExStart
                #ExFor:DocSaveOptions.always_compress_metafiles
                #ExSummary:Shows how to change metafiles compression in a document while saving.
                # Open a document that contains a Microsoft Equation 3.0 formula.
                doc = aw.Document(aeb.my_dir + "Microsoft equation object.docx")

                # When we save a document, smaller metafiles are not compressed for performance reasons.
                # We can set a flag in a SaveOptions object to compress every metafile when saving.
                # Some editors such as LibreOffice cannot read uncompressed metafiles.
                saveOptions = aw.saving.DocSaveOptions()
                saveOptions.always_compress_metafiles = compressAllMetafiles

                doc.save(aeb.artifacts_dir + "DocSaveOptions.always_compress_metafiles.docx", saveOptions)

                print(os.path.getsize(aeb.artifacts_dir + "DocSaveOptions.always_compress_metafiles.docx"))
                if compressAllMetafiles :
                    self.assertTrue(10000 > os.path.getsize(aeb.artifacts_dir + "DocSaveOptions.always_compress_metafiles.docx"))
                else :
                    self.assertTrue(30000 > os.path.getsize(aeb.artifacts_dir + "DocSaveOptions.always_compress_metafiles.docx"))
                #ExEnd
        
    
if __name__ == '__main__':
    unittest.main() 