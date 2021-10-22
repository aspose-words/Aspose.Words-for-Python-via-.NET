import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithPclSaveOptions(docs_base.DocsExamplesBase):

    def test_rasterize_transformed_elements(self):

        #ExStart:RasterizeTransformedElements
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        save_options = aw.saving.PclSaveOptions()
        save_options.save_format = aw.SaveFormat.PCL
        save_options.rasterize_transformed_elements = False

        doc.save(docs_base.artifacts_dir + "WorkingWithPclSaveOptions.rasterize_transformed_elements.pcl", save_options)
        #ExEnd:RasterizeTransformedElements


if __name__ == '__main__':
    unittest.main()
