import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR

class WorkingWithRanges(DocsExamplesBase):

    def test_ranges_delete_text(self):

        #ExStart:RangesDeleteText
        doc = aw.Document(MY_DIR + "Document.docx")
        doc.sections[0].range.delete()
        #ExEnd:RangesDeleteText

    def test_ranges_get_text(self):

        #ExStart:RangesGetText
        doc = aw.Document(MY_DIR + "Document.docx")
        text = doc.range.text
        #ExEnd:RangesGetText
