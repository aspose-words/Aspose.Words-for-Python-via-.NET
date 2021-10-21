import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithDocSaveOptions(docs_base.DocsExamplesBase):

    def test_encrypt_document_with_password(self) :

        #ExStart:EncryptDocumentWithPassword
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Hello world!")

        save_options = aw.saving.DocSaveOptions()
        save_options.password = "password"

        doc.save(docs_base.artifacts_dir + "WorkingWithDocSaveOptions.encrypt_document_with_password.doc", save_options)
        #ExEnd:EncryptDocumentWithPassword


    def test_do_not_compress_small_metafiles(self) :

        #ExStart:DoNotCompressSmallMetafiles
        doc = aw.Document(docs_base.my_dir + "Microsoft equation object.docx")

        save_options = aw.saving.DocSaveOptions()
        save_options.always_compress_metafiles = False

        doc.save(docs_base.artifacts_dir + "WorkingWithDocSaveOptions.not_compress_small_metafiles.doc", save_options)
        #ExEnd:DoNotCompressSmallMetafiles


    def test_do_not_save_picture_bullet(self) :

        #ExStart:DoNotSavePictureBullet
        doc = aw.Document(docs_base.my_dir + "Image bullet points.docx")

        save_options = aw.saving.DocSaveOptions()
        save_options.save_picture_bullet = False

        doc.save(docs_base.artifacts_dir + "WorkingWithDocSaveOptions.do_not_save_picture_bullet.doc", save_options)
        #ExEnd:DoNotSavePictureBullet

if __name__ == '__main__':
    unittest.main()


