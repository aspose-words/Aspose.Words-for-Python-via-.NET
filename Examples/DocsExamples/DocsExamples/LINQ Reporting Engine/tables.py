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

class Tables(docs_base.DocsExamplesBase):

    def test_in_table_alternate_content(self) :

        #ExStart:InTableAlternateContent
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Total.docx")

        opt = aw.reporting.JsonDataLoadOptions()
        opt.always_generate_root_object = True
        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "contracts.json", opt), "contracts")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.in_table_alternate_content.docx")
        #ExEnd:InTableAlternateContent


    def test_in_table_master_detail(self) :

        #ExStart:InTableMasterDetail
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Nested data table.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "managers.json"), "Managers")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.in_table_master_detail.docx")
        #ExEnd:InTableMasterDetail


    def test_in_table_with_filtering_grouping_sorting(self) :

        #ExStart:InTableWithFilteringGroupingSorting
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Table with filtering.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "contracts.json"), "contracts")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.in_table_with_filtering_grouping_sorting.docx")
        #ExEnd:InTableWithFilteringGroupingSorting



if __name__ == '__main__':
    unittest.main()