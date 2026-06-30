import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR

class WorkingWithRanges(DocsExamplesBase):

    def test_ranges_delete_text(self):

        #ExStart:RangesDeleteText
        #GistId:9164e9c0658006e51db723b0742c12fc
        doc = aw.Document(MY_DIR + "Document.docx")
        doc.sections[0].range.delete()
        #ExEnd:RangesDeleteText

    def test_ranges_get_text(self):

        #ExStart:RangesGetText
        #GistId:9164e9c0658006e51db723b0742c12fc
        doc = aw.Document(MY_DIR + "Document.docx")
        text = doc.range.text
        #ExEnd:RangesGetText
