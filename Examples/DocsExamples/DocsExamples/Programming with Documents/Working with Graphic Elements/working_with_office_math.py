import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithOfficeMath(docs_base.DocsExamplesBase):

    
    def test_math_equations(self) :
        
        #ExStart:MathEquations
        doc = aw.Document(docs_base.my_dir + "Office math.docx")
        officeMath = doc.get_child(aw.NodeType.OFFICE_MATH, 0, True).as_office_math()

        # OfficeMath display type represents whether an equation is displayed inline with the text or displayed on its line.
        officeMath.display_type = aw.math.OfficeMathDisplayType.DISPLAY
        officeMath.justification = aw.math.OfficeMathJustification.LEFT

        doc.save(docs_base.artifacts_dir + "WorkingWithOfficeMath.math_equations.docx")
        #ExEnd:MathEquations
        
    

if __name__ == '__main__':
    unittest.main()