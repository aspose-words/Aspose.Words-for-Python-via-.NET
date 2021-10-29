import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithRtfLoadOptions(DocsExamplesBase):

    def test_recognize_utf_8_text(self):

        #ExStart:RecognizeUtf8Text
        load_options = aw.loading.RtfLoadOptions()
        load_options.recognize_utf8_text = True

        doc = aw.Document(MY_DIR + "UTF-8 characters.rtf", load_options)

        doc.save(ARTIFACTS_DIR + "WorkingWithRtfLoadOptions.recognize_utf_8_text.rtf")
        #ExEnd:RecognizeUtf8Text
