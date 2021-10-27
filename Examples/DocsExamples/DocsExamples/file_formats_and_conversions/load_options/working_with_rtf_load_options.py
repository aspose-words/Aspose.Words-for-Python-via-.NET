import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithRtfLoadOptions(DocsExamplesBase):

    def test_recognize_utf_8_text(self):

        #ExStart:RecognizeUtf8Text
        load_options = aw.loading.RtfLoadOptions()
        load_options.recognize_utf8_text = True

        doc = aw.Document(MY_DIR + "UTF-8 characters.rtf", load_options)

        doc.save(ARTIFACTS_DIR + "WorkingWithRtfLoadOptions.recognize_utf_8_text.rtf")
        #ExEnd:RecognizeUtf8Text


if __name__ == '__main__':
    unittest.main()
