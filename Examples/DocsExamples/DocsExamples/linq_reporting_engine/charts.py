from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR, JSON_DIR

import aspose.words as aw

class Charts(DocsExamplesBase):

    def test_create_bubble_chart(self):

        #ExStart:BubbleChart
        doc = aw.Document(MY_DIR + "Reporting engine template - Bubble chart.docx")

        opt = aw.reporting.JsonDataLoadOptions()
        opt.always_generate_root_object = True

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(JSON_DIR + "managers.json", opt), "Managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.create_bubble_chart.docx")
        #ExEnd:BubbleChart

    def test_set_chart_series_name_dynamically(self):

        #ExStart:SetChartSeriesNameDynamically
        doc = aw.Document(MY_DIR + "Reporting engine template - Chart.docx")

        opt = aw.reporting.JsonDataLoadOptions()
        opt.always_generate_root_object = True

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(JSON_DIR + "managers.json", opt), "Managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.set_chart_series_name_dynamically.docx")
        #ExEnd:SetChartSeriesNameDynamically

    def test_chart_with_filtering_grouping_ordering(self):

        #ExStart:ChartWithFilteringGroupingOrdering
        doc = aw.Document(MY_DIR + "Reporting engine template - Chart with filtering.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(JSON_DIR + "contracts.json"), "contracts")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.chart_with_filtering_grouping_ordering.docx")
        #ExEnd:ChartWithFilteringGroupingOrdering

    def test_pie_chart(self):

        #ExStart:PieChart
        doc = aw.Document(MY_DIR + "Reporting engine template - Pie chart.docx")

        opt = aw.reporting.JsonDataLoadOptions()
        opt.always_generate_root_object = True

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(JSON_DIR + "managers.json", opt), "Managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.pie_chart.docx")
        #ExEnd:PieChart

    def test_scatter_chart(self):

        #ExStart:ScatterChart
        doc = aw.Document(MY_DIR + "Reporting engine template - Scatter chart.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(JSON_DIR + "contracts.json"), "contracts")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.scatter_chart.docx")
        #ExEnd:ScatterChart
