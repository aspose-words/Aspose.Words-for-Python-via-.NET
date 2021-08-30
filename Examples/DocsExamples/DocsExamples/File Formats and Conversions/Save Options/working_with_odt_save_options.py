import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithOdtSaveOptions(docs_base.DocsExamplesBase):
    
    def test_measure_unit(self) :
        
        #ExStart:MeasureUnit
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        # Open Office uses centimeters when specifying lengths, widths and other measurable formatting
        # and content properties in documents whereas MS Office uses inches.
        saveOptions = aw.saving.OdtSaveOptions()
        saveOptions.measure_unit = aw.saving.OdtSaveMeasureUnit.INCHES 

        doc.save(docs_base.artifacts_dir + "WorkingWithOdtSaveOptions.measure_unit.odt", saveOptions)
        #ExEnd:MeasureUnit
        
    

if __name__ == '__main__':
    unittest.main()