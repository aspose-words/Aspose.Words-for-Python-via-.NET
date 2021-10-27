from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR, JSON_DIR

import aspose.words as aw

class BuildOptions(DocsExamplesBase):

    def test_remove_empty_paragraphs(self):

        #ExStart:RemoveEmptyParagraphs
        doc = aw.Document(MY_DIR + "Reporting engine template - Remove empty paragraphs.docx")

        engine = aw.reporting.ReportingEngine()
        engine.options = aw.reporting.ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS
        engine.build_report(doc, aw.reporting.JsonDataSource(JSON_DIR + "managers.json"), "Managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.remove_empty_paragraphs.docx")
        #ExEnd:RemoveEmptyParagraphs
