import unittest
import os
import sys
import io

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class BuildOptions(docs_base.DocsExamplesBase):

    def test_remove_empty_paragraphs(self) :

        #ExStart:RemoveEmptyParagraphs
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Remove empty paragraphs.docx")

        engine = aw.reporting.ReportingEngine()
        engine.options = aw.reporting.ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS
        engine.build_report(doc,  aw.reporting.JsonDataSource(docs_base.json_dir + "managers.json"), "Managers")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.remove_empty_paragraphs.docx")
        #ExEnd:RemoveEmptyParagraphs



if __name__ == '__main__':
    unittest.main()