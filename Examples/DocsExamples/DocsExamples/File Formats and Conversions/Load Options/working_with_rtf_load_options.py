import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithRtfLoadOptions(docs_base.DocsExamplesBase):

    def test_recognize_utf_8_text(self) :

        #ExStart:RecognizeUtf8Text
        load_options = aw.loading.RtfLoadOptions()
        load_options.recognize_utf8_text = True

        doc = aw.Document(docs_base.my_dir + "UTF-8 characters.rtf", load_options)

        doc.save(docs_base.artifacts_dir + "WorkingWithRtfLoadOptions.recognize_utf_8_text.rtf")
        #ExEnd:RecognizeUtf8Text



if __name__ == '__main__':
    unittest.main()
