import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithDocSaveOptions(DocsExamplesBase):

    def test_encrypt_document_with_password(self):

        #ExStart:EncryptDocumentWithPassword
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Hello world!")

        save_options = aw.saving.DocSaveOptions()
        save_options.password = "password"

        doc.save(ARTIFACTS_DIR + "WorkingWithDocSaveOptions.encrypt_document_with_password.doc", save_options)
        #ExEnd:EncryptDocumentWithPassword


    def test_do_not_compress_small_metafiles(self):

        #ExStart:DoNotCompressSmallMetafiles
        doc = aw.Document(MY_DIR + "Microsoft equation object.docx")

        save_options = aw.saving.DocSaveOptions()
        save_options.always_compress_metafiles = False

        doc.save(ARTIFACTS_DIR + "WorkingWithDocSaveOptions.not_compress_small_metafiles.doc", save_options)
        #ExEnd:DoNotCompressSmallMetafiles


    def test_do_not_save_picture_bullet(self):

        #ExStart:DoNotSavePictureBullet
        doc = aw.Document(MY_DIR + "Image bullet points.docx")

        save_options = aw.saving.DocSaveOptions()
        save_options.save_picture_bullet = False

        doc.save(ARTIFACTS_DIR + "WorkingWithDocSaveOptions.do_not_save_picture_bullet.doc", save_options)
        #ExEnd:DoNotSavePictureBullet

if __name__ == '__main__':
    unittest.main()
