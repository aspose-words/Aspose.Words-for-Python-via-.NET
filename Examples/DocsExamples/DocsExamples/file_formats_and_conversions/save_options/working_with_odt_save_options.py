import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithOdtSaveOptions(DocsExamplesBase):

    def test_measure_unit(self):

        #ExStart:MeasureUnit
        doc = aw.Document(MY_DIR + "Document.docx")

        # Open Office uses centimeters when specifying lengths, widths and other measurable formatting
        # and content properties in documents whereas MS Office uses inches.
        save_options = aw.saving.OdtSaveOptions()
        save_options.measure_unit = aw.saving.OdtSaveMeasureUnit.INCHES

        doc.save(ARTIFACTS_DIR + "WorkingWithOdtSaveOptions.measure_unit.odt", save_options)
        #ExEnd:MeasureUnit


if __name__ == '__main__':
    unittest.main()
