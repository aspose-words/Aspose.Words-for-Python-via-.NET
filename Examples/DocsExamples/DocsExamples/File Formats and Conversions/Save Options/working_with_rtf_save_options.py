import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithRtfSaveOptions(docs_base.DocsExamplesBase):

    def test_saving_images_as_wmf(self) :

        #ExStart:SavingImagesAsWmf
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        saveOptions = aw.saving.RtfSaveOptions()
        saveOptions.save_images_as_wmf = True

        doc.save(docs_base.artifacts_dir + "WorkingWithRtfSaveOptions.saving_images_as_wmf.rtf", saveOptions)
        #ExEnd:SavingImagesAsWmf



if __name__ == '__main__':
    unittest.main()