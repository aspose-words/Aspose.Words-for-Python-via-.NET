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

class Charts(docs_base.DocsExamplesBase):


    def test_create_bubble_chart(self) :

        #ExStart:BubbleChart
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Bubble chart.docx")

        opt = aw.reporting.JsonDataLoadOptions()
        opt.always_generate_root_object = True
        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "managers.json", opt), "Managers")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.create_bubble_chart.docx")
        #ExEnd:BubbleChart


    def test_set_chart_series_name_dynamically(self) :

        #ExStart:SetChartSeriesNameDynamically
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Chart.docx")

        opt = aw.reporting.JsonDataLoadOptions()
        opt.always_generate_root_object = True
        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "managers.json", opt), "Managers")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.set_chart_series_name_dynamically.docx")
        #ExEnd:SetChartSeriesNameDynamically


    def test_chart_with_filtering_grouping_ordering(self) :

        #ExStart:ChartWithFilteringGroupingOrdering
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Chart with filtering.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "contracts.json"), "contracts")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.chart_with_filtering_grouping_ordering.docx")
        #ExEnd:ChartWithFilteringGroupingOrdering


    def test_pie_chart(self) :

        #ExStart:PieChart
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Pie chart.docx")

        opt = aw.reporting.JsonDataLoadOptions()
        opt.always_generate_root_object = True
        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "managers.json", opt), "Managers")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.pie_chart.docx")
        #ExEnd:PieChart


    def test_scatter_chart(self) :

        #ExStart:ScatterChart
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Scatter chart.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "contracts.json"), "contracts")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.scatter_chart.docx")
        #ExEnd:ScatterChart



if __name__ == '__main__':
    unittest.main()