import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithOfficeMath(DocsExamplesBase):

    def test_math_equations(self):

        #ExStart:MathEquations
        doc = aw.Document(MY_DIR + "Office math.docx")
        office_math = doc.get_child(aw.NodeType.OFFICE_MATH, 0, True).as_office_math()

        # OfficeMath display type represents whether an equation is displayed inline with the text or displayed on its line.
        office_math.display_type = aw.math.OfficeMathDisplayType.DISPLAY
        office_math.justification = aw.math.OfficeMathJustification.LEFT

        doc.save(ARTIFACTS_DIR + "WorkingWithOfficeMath.math_equations.docx")
        #ExEnd:MathEquations
