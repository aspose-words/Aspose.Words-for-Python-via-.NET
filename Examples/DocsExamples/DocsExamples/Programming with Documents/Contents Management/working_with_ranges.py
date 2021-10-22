import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithRanges(docs_base.DocsExamplesBase):

    def test_ranges_delete_text(self) :

        #ExStart:RangesDeleteText
        doc = aw.Document(docs_base.my_dir + "Document.docx")
        doc.sections[0].range.delete()
        #ExEnd:RangesDeleteText


    def test_ranges_get_text(self) :

        #ExStart:RangesGetText
        doc = aw.Document(docs_base.my_dir + "Document.docx")
        text = doc.range.text
        #ExEnd:RangesGetText


if __name__ == '__main__':
    unittest.main()
