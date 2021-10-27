import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithPclSaveOptions(DocsExamplesBase):

    def test_rasterize_transformed_elements(self):

        #ExStart:RasterizeTransformedElements
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.PclSaveOptions()
        save_options.save_format = aw.SaveFormat.PCL
        save_options.rasterize_transformed_elements = False

        doc.save(ARTIFACTS_DIR + "WorkingWithPclSaveOptions.rasterize_transformed_elements.pcl", save_options)
        #ExEnd:RasterizeTransformedElements


if __name__ == '__main__':
    unittest.main()
