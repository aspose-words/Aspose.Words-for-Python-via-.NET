import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR, JSON_DIR

class Tables(DocsExamplesBase):

    def test_in_table_alternate_content(self):

        #ExStart:InTableAlternateContent
        doc = aw.Document(MY_DIR + "Reporting engine template - Total.docx")

        opt = aw.reporting.JsonDataLoadOptions()
        opt.always_generate_root_object = True

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(JSON_DIR + "contracts.json", opt), "contracts")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.in_table_alternate_content.docx")
        #ExEnd:InTableAlternateContent

    def test_in_table_master_detail(self):

        #ExStart:InTableMasterDetail
        doc = aw.Document(MY_DIR + "Reporting engine template - Nested data table.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(JSON_DIR + "managers.json"), "Managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.in_table_master_detail.docx")
        #ExEnd:InTableMasterDetail

    def test_in_table_with_filtering_grouping_sorting(self):

        #ExStart:InTableWithFilteringGroupingSorting
        doc = aw.Document(MY_DIR + "Reporting engine template - Table with filtering.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(JSON_DIR + "contracts.json"), "contracts")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.in_table_with_filtering_grouping_sorting.docx")
        #ExEnd:InTableWithFilteringGroupingSorting
