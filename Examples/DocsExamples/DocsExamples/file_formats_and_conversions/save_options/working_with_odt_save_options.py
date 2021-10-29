import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

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
