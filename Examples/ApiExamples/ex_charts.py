import unittest
from datetime import date, datetime

import api_example_base as aeb
from document_helper import DocumentHelper

import aspose.words as aw

class ExCharts(aeb.ApiExampleBase):
    
    def test_chart_title(self) :
        
        #ExStart
        #ExFor:Chart
        #ExFor:Chart.title
        #ExFor:ChartTitle
        #ExFor:ChartTitle.overlay
        #ExFor:ChartTitle.show
        #ExFor:ChartTitle.text
        #ExSummary:Shows how to insert a chart and set a title.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a chart shape with a document builder and get its chart.
        chartShape = builder.insert_chart(aw.drawing.charts.ChartType.BAR, 400, 300)
        chart = chartShape.chart

        # Use the "Title" property to give our chart a title, which appears at the top center of the chart area.
        title = chart.title
        title.text = "My Chart"

        # Set the "Show" property to "True" to make the title visible. 
        title.show = True

        # Set the "Overlay" property to "True" Give other chart elements more room by allowing them to overlap the title
        title.overlay = True

        doc.save(aeb.artifacts_dir + "Charts.chart_title.docx")
        #ExEnd

# no type casting yet.
#        doc = aw.Document(aeb.artifacts_dir + "Charts.chart_title.docx")
#        chartShape = (Shape)doc.get_child(NodeType.shape, 0, True)
#
#        self.assertEqual(ShapeType.non_primitive, chartShape.shape_type)
#        self.assertTrue(chartShape.has_chart)
#
#        title = chartShape.chart.title
#
#        self.assertEqual("My Chart", title.text)
#        self.assertTrue(title.overlay)
#        self.assertTrue(title.show)
        

    def test_data_label_number_format(self) :
        
        #ExStart
        #ExFor:ChartDataLabelCollection.number_format
        #ExFor:ChartNumberFormat.format_code
        #ExSummary:Shows how to enable and configure data labels for a chart series.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Add a line chart, then clear its demo data series to start with a clean chart,
        # and then set a title.
        shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 500, 300)
        chart = shape.chart
        chart.series.clear()
        chart.title.text = "Monthly sales report"
            
        # Insert a custom chart series with months as categories for the X-axis,
        # and respective decimal amounts for the Y-axis.
        series = chart.series.add("Revenue", ["January", "February", "March"], [25.611, 21.439, 33.750] )

        # Enable data labels, and then apply a custom number format for values displayed in the data labels.
        # This format will treat displayed decimal values as millions of US Dollars.
        series.has_data_labels = True
        dataLabels = series.data_labels
        dataLabels.show_value = True
        dataLabels.number_format.format_code = "\"US$\" #,##0.000\"M\""

        doc.save(aeb.artifacts_dir + "Charts.data_label_number_format.docx")
        #ExEnd

# no type casting yet.
#        doc = new Document(aeb.artifacts_dir + "Charts.data_label_number_format.docx")
#        series = ((Shape)doc.get_child(NodeType.shape, 0, True)).chart.series[0]
#
#        self.assertTrue(series.has_data_labels)
#        self.assertTrue(series.data_labels.show_value)
#        self.assertEqual("\"US$\" #,##0.000\"M\"", series.data_labels.number_format.format_code)
        

    def test_data_arrays_wrong_size(self) :
        
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 500, 300)
        chart = shape.chart

        seriesColl = chart.series
        seriesColl.clear()

        categories =  ["Cat1", None, "Cat3", "Cat4", "Cat5", None]
        seriesColl.add("AW Series 1", categories, [ 1, 2, None, 4, 5, 6] )
        seriesColl.add("AW Series 2", categories, [ 2, 3, None, 5, 6, 7] )

        with self.assertRaises(RuntimeError) as ex:
            seriesColl.add("AW Series 3", categories, [None, 4, 5, None, None])

        with self.assertRaises(RuntimeError) as ex:
           riesColl.add("AW Series 4", categories, [None, None, None, None, None])

        

    def test_empty_values_in_chart_data(self) :
        
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 500, 300)
        chart = shape.chart

        seriesColl = chart.series
        seriesColl.clear()

        categories =  ["Cat1", None, "Cat3", "Cat4", "Cat5", None]
        seriesColl.add("AW Series 1", categories, [1, 2, None, 4, 5, 6])
        seriesColl.add("AW Series 2", categories, [2, 3, None, 5, 6, 7])
        seriesColl.add("AW Series 3", categories, [None, 4, 5, None, 7, 8])
        seriesColl.add("AW Series 4", categories, [None, None, None, None, None, 9])

        doc.save(aeb.artifacts_dir + "Charts.empty_values_in_chart_data.docx")
        

    def test_axis_properties(self) :
        
        #ExStart
        #ExFor:ChartAxis
        #ExFor:ChartAxis.category_type
        #ExFor:ChartAxis.crosses
        #ExFor:ChartAxis.reverse_order
        #ExFor:ChartAxis.major_tick_mark
        #ExFor:ChartAxis.minor_tick_mark
        #ExFor:ChartAxis.major_unit
        #ExFor:ChartAxis.minor_unit
        #ExFor:ChartAxis.tick_label_offset
        #ExFor:ChartAxis.tick_label_position
        #ExFor:ChartAxis.tick_label_spacing_is_auto
        #ExFor:ChartAxis.tick_mark_spacing
        #ExFor:Charts.axis_category_type
        #ExFor:Charts.axis_crosses
        #ExFor:Charts.chart.axis_x
        #ExFor:Charts.chart.axis_y
        #ExFor:Charts.chart.axis_z
        #ExSummary:Shows how to insert a chart and modify the appearance of its axes.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.COLUMN, 500, 300)
        chart = shape.chart

        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()

        # Insert a chart series with categories for the X-axis and respective numeric values for the Y-axis.
        chart.series.add("Aspose Test Series", ["Word", "PDF", "Excel", "GoogleDocs", "Note"], [640, 320, 280, 120, 150])
            
        # Chart axes have various options that can change their appearance,
        # such as their direction, major/minor unit ticks, and tick marks.
        xAxis = chart.axis_x
        xAxis.category_type = aw.drawing.charts.AxisCategoryType.CATEGORY
        xAxis.crosses = aw.drawing.charts.AxisCrosses.MINIMUM
        xAxis.reverse_order = False
        xAxis.major_tick_mark = aw.drawing.charts.AxisTickMark.INSIDE
        xAxis.minor_tick_mark = aw.drawing.charts.AxisTickMark.CROSS
        xAxis.major_unit = 10.0
        xAxis.minor_unit = 15.0
        xAxis.tick_label_offset = 50
        xAxis.tick_label_position = aw.drawing.charts.AxisTickLabelPosition.LOW
        xAxis.tick_label_spacing_is_auto = False
        xAxis.tick_mark_spacing = 1

        yAxis = chart.axis_y
        yAxis.category_type = aw.drawing.charts.AxisCategoryType.AUTOMATIC
        yAxis.crosses = aw.drawing.charts.AxisCrosses.MAXIMUM
        yAxis.reverse_order = True
        yAxis.major_tick_mark = aw.drawing.charts.AxisTickMark.INSIDE
        yAxis.minor_tick_mark = aw.drawing.charts.AxisTickMark.CROSS
        yAxis.major_unit = 100.0
        yAxis.minor_unit = 20.0
        yAxis.tick_label_position = aw.drawing.charts.AxisTickLabelPosition.NEXT_TO_AXIS

        # Column charts do not have a Z-axis.
        self.assertIsNone(chart.axis_z)

        doc.save(aeb.artifacts_dir + "Charts.axis_properties.docx")
        #ExEnd

# no type casting yet.
#        doc = new Document(aeb.artifacts_dir + "Charts.axis_properties.docx")
#        chart = ((Shape)doc.get_child(NodeType.shape, 0, True)).chart
#
#        self.assertEqual(AxisCategoryType.category, chart.axis_x.category_type)
#        self.assertEqual(AxisCrosses.minimum, chart.axis_x.crosses)
#        self.assertFalse(chart.axis_x.reverse_order)
#        self.assertEqual(aw.drawing.charts.AxisTickMark.INSIDE, chart.axis_x.major_tick_mark)
#        self.assertEqual(aw.drawing.charts.AxisTickMark.CROSS, chart.axis_x.minor_tick_mark)
#        self.assertEqual(1.0d, chart.axis_x.major_unit)
#        self.assertEqual(0.5d, chart.axis_x.minor_unit)
#        self.assertEqual(50, chart.axis_x.tick_label_offset)
#        self.assertEqual(AxisTickLabelPosition.low, chart.axis_x.tick_label_position)
#        self.assertFalse(chart.axis_x.tick_label_spacing_is_auto)
#        self.assertEqual(1, chart.axis_x.tick_mark_spacing)
#
#        self.assertEqual(AxisCategoryType.category, chart.axis_y.category_type)
#        self.assertEqual(AxisCrosses.maximum, chart.axis_y.crosses)
#        self.assertTrue(chart.axis_y.reverse_order)
#        self.assertEqual(aw.drawing.charts.AxisTickMark.INSIDE, chart.axis_y.major_tick_mark)
#        self.assertEqual(aw.drawing.charts.AxisTickMark.CROSS, chart.axis_y.minor_tick_mark)
#        self.assertEqual(100.0d, chart.axis_y.major_unit)
#        self.assertEqual(20.0d, chart.axis_y.minor_unit)
#        self.assertEqual(AxisTickLabelPosition.next_to_axis, chart.axis_y.tick_label_position)
        

    def test_date_time_values(self) :
        
        #ExStart
        #ExFor:AxisBound
        #ExFor:AxisBound.#ctor(Double)
        #ExFor:AxisBound.#ctor(DateTime)
        #ExFor:AxisScaling.minimum
        #ExFor:AxisScaling.maximum
        #ExFor:ChartAxis.scaling
        #ExFor:Charts.axis_tick_mark
        #ExFor:Charts.axis_tick_label_position
        #ExFor:Charts.axis_time_unit
        #ExFor:Charts.chart_axis.base_time_unit
        #ExSummary:Shows how to insert chart with date/time values.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 500, 300)
        chart = shape.chart

        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()

        # Add a custom series containing date/time values for the X-axis, and respective decimal values for the Y-axis.
        chart.series.add("Aspose Test Series", 
            [date(2017, 11, 6), date(2017, 11, 9), date(2017, 11, 15), date(2017, 11, 21), date(2017, 11, 25), date(2017, 11, 29)],
            [1.2, 0.3, 2.1, 2.9, 4.2, 5.3])


        # Set lower and upper bounds for the X-axis.
        xAxis = chart.axis_x
        xAxis.scaling.minimum = aw.drawing.charts.AxisBound(date(2017, 11, 5))
        xAxis.scaling.maximum = aw.drawing.charts.AxisBound(date(2017, 12, 3))

        # Set the major units of the X-axis to a week, and the minor units to a day.
        xAxis.base_time_unit = aw.drawing.charts.AxisTimeUnit.DAYS
        xAxis.major_unit = 7.0
        xAxis.major_tick_mark = aw.drawing.charts.AxisTickMark.CROSS
        xAxis.minor_unit = 1.0
        xAxis.minor_tick_mark = aw.drawing.charts.AxisTickMark.OUTSIDE

        # Define Y-axis properties for decimal values.
        yAxis = chart.axis_y
        yAxis.tick_label_position = aw.drawing.charts.AxisTickLabelPosition.HIGH
        yAxis.major_unit = 100.0
        yAxis.minor_unit = 50.0
        yAxis.display_unit.unit = aw.drawing.charts.AxisBuiltInUnit.HUNDREDS
        yAxis.scaling.minimum = aw.drawing.charts.AxisBound(100)
        yAxis.scaling.maximum = aw.drawing.charts.AxisBound(700)

        doc.save(aeb.artifacts_dir + "Charts.date_time_values.docx")
        #ExEnd

# no type casting yet.
#        doc = new Document(aeb.artifacts_dir + "Charts.date_time_values.docx")
#        chart = ((Shape)doc.get_child(NodeType.shape, 0, True)).chart
#
#        self.assertEqual(aw.drawing.charts.AxisBound(date(2017, 11, 05).to_oa_date()), chart.axis_x.scaling.minimum)
#        self.assertEqual(aw.drawing.charts.AxisBound(date(2017, 12, 03)), chart.axis_x.scaling.maximum)
#        self.assertEqual(AxisTimeUnit.days, chart.axis_x.base_time_unit)
#        self.assertEqual(7.0d, chart.axis_x.major_unit)
#        self.assertEqual(1.0d, chart.axis_x.minor_unit)
#        self.assertEqual(aw.drawing.charts.AxisTickMark.CROSS, chart.axis_x.major_tick_mark)
#        self.assertEqual(aw.drawing.charts.AxisTickMark.OUTSIDE, chart.axis_x.minor_tick_mark)
#
#        self.assertEqual(AxisTickLabelPosition.high, chart.axis_y.tick_label_position)
#        self.assertEqual(100.0d, chart.axis_y.major_unit)
#        self.assertEqual(50.0d, chart.axis_y.minor_unit)
#        self.assertEqual(AxisBuiltInUnit.hundreds, chart.axis_y.display_unit.unit)
#        self.assertEqual(aw.drawing.charts.AxisBound(100), chart.axis_y.scaling.minimum)
#        self.assertEqual(aw.drawing.charts.AxisBound(700), chart.axis_y.scaling.maximum)
        

    def test_hide_chart_axis(self) :
        
        #ExStart
        #ExFor:ChartAxis.hidden
        #ExSummary:Shows how to hide chart axes.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 500, 300)
        chart = shape.chart

        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()

        # Add a custom series with categories for the X-axis, and respective decimal values for the Y-axis.
        chart.series.add("AW Series 1",
            ["Item 1", "Item 2", "Item 3", "Item 4", "Item 5"],
            [1.2, 0.3, 2.1, 2.9, 4.2])

        # Hide the chart axes to simplify the appearance of the chart. 
        chart.axis_x.hidden = True
        chart.axis_y.hidden = True

        doc.save(aeb.artifacts_dir + "Charts.hide_chart_axis.docx")
        #ExEnd

# no type casting yet.
#        doc = new Document(aeb.artifacts_dir + "Charts.hide_chart_axis.docx")
#        chart = ((Shape)doc.get_child(NodeType.shape, 0, True)).chart
#
#        self.assertTrue(chart.axis_x.hidden)
#        self.assertTrue(chart.axis_y.hidden)
        

    def test_set_number_format_to_chart_axis(self) :
        
        #ExStart
        #ExFor:ChartAxis.number_format
        #ExFor:Charts.chart_number_format
        #ExFor:ChartNumberFormat.format_code
        #ExFor:Charts.chart_number_format.is_linked_to_source
        #ExSummary:Shows how to set formatting for chart values.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.COLUMN, 500, 300)
        chart = shape.chart

        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()

        # Add a custom series to the chart with categories for the X-axis,
        # and large respective numeric values for the Y-axis. 
        chart.series.add("Aspose Test Series",
            ["Word", "PDF", "Excel", "GoogleDocs", "Note"],
            [1900000, 850000, 2100000, 600000, 1500000])

        # Set the number format of the Y-axis tick labels to not group digits with commas. 
        chart.axis_y.number_format.format_code = "#,##0"

        # This flag can override the above value and draw the number format from the source cell.
        self.assertFalse(chart.axis_y.number_format.is_linked_to_source)

        doc.save(aeb.artifacts_dir + "Charts.set_number_format_to_chart_axis.docx")
        #ExEnd

# no type casting yet.
#        doc = new Document(aeb.artifacts_dir + "Charts.set_number_format_to_chart_axis.docx")
#        chart = ((Shape)doc.get_child(NodeType.shape, 0, True)).chart
#
#        self.assertEqual("#,##0", chart.axis_y.number_format.format_code)
        

    def test_display_charts_with_conversion(self) :
        for chartType in [aw.drawing.charts.ChartType.COLUMN, aw.drawing.charts.ChartType.LINE, aw.drawing.charts.ChartType.PIE, aw.drawing.charts.ChartType.BAR, aw.drawing.charts.ChartType.AREA]:
        
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc)

            shape = builder.insert_chart(chartType, 500, 300)
            chart = shape.chart
            chart.series.clear()
                
            chart.series.add("Aspose Test Series",
                ["Word", "PDF", "Excel", "GoogleDocs", "Note"],
                [1900000, 850000, 2100000, 600000, 1500000])

            doc.save(aeb.artifacts_dir + "Charts.test_display_charts_with_conversion.docx")
            doc.save(aeb.artifacts_dir + "Charts.test_display_charts_with_conversion.pdf")
        

    def test_surface_3_d_chart(self) :
        
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.SURFACE3_D, 500, 300)
        chart = shape.chart
        chart.series.clear()
            
        chart.series.add("Aspose Test Series 1",
            ["Word", "PDF", "Excel", "GoogleDocs", "Note"],
            [1900000, 850000, 2100000, 600000, 1500000])
            
        chart.series.add("Aspose Test Series 2",
            ["Word", "PDF", "Excel", "GoogleDocs", "Note"],
            [900000, 50000, 1100000, 400000, 2500000])
            
        chart.series.add("Aspose Test Series 3",
            ["Word", "PDF", "Excel", "GoogleDocs", "Note"],
            [500000, 820000, 1500000, 400000, 100000])

        doc.save(aeb.artifacts_dir + "Charts.surface_chart.docx")
        doc.save(aeb.artifacts_dir + "Charts.surface_chart.pdf")
        

    def test_data_labels_bubble_chart(self) :
        
        #ExStart
        #ExFor:ChartDataLabelCollection.separator
        #ExFor:ChartDataLabelCollection.show_bubble_size
        #ExFor:ChartDataLabelCollection.show_category_name
        #ExFor:ChartDataLabelCollection.show_series_name
        #ExSummary:Shows how to work with data labels of a bubble chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        chart = builder.insert_chart(aw.drawing.charts.ChartType.BUBBLE, 500, 300).chart

        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()

        # Add a custom series with X/Y coordinates and diameter of each of the bubbles. 
        series = chart.series.add("Aspose Test Series",
            [2.9, 3.5, 1.1, 4.0, 4.0],
            [1.9, 8.5, 2.1, 6.0, 1.5],
            [9.0, 4.5, 2.5, 8.0, 5.0])

        # Enable data labels, and then modify their appearance.
        series.has_data_labels = True
        dataLabels = series.data_labels
        dataLabels.show_bubble_size = True
        dataLabels.show_category_name = True
        dataLabels.show_series_name = True
        dataLabels.separator = " & "

        doc.save(aeb.artifacts_dir + "Charts.data_labels_bubble_chart.docx")
        #ExEnd

# no type casting yet.
#        doc = new Document(aeb.artifacts_dir + "Charts.data_labels_bubble_chart.docx")
#        dataLabels = ((Shape)doc.get_child(NodeType.shape, 0, True)).chart.series[0].data_labels
#
#        self.assertTrue(dataLabels.show_bubble_size)
#        self.assertTrue(dataLabels.show_category_name)
#        self.assertTrue(dataLabels.show_series_name)
#        self.assertEqual(" & ", dataLabels.separator)
        

    def test_data_labels_pie_chart(self) :
        
        #ExStart
        #ExFor:ChartDataLabelCollection.separator
        #ExFor:ChartDataLabelCollection.show_leader_lines
        #ExFor:ChartDataLabelCollection.show_legend_key
        #ExFor:ChartDataLabelCollection.show_percentage
        #ExFor:ChartDataLabelCollection.show_value
        #ExSummary:Shows how to work with data labels of a pie chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        chart = builder.insert_chart(aw.drawing.charts.ChartType.PIE, 500, 300).chart

        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()

        # Insert a custom chart series with a category name for each of the sectors, and their frequency table.
        series = chart.series.add("Aspose Test Series",
            ["Word", "PDF", "Excel"] ,
            [2.7, 3.2, 0.8])

        # Enable data labels that will display both percentage and frequency of each sector, and modify their appearance.
        series.has_data_labels = True
        dataLabels = series.data_labels
        dataLabels.show_leader_lines = True
        dataLabels.show_legend_key = True
        dataLabels.show_percentage = True
        dataLabels.show_value = True
        dataLabels.separator = " "

        doc.save(aeb.artifacts_dir + "Charts.data_labels_pie_chart.docx")
        #ExEnd

# no type casting yet.
#        doc = new Document(aeb.artifacts_dir + "Charts.data_labels_pie_chart.docx")
#        dataLabels = ((Shape)doc.get_child(NodeType.shape, 0, True)).chart.series[0].data_labels
#
#        self.assertTrue(dataLabels.show_leader_lines)
#        self.assertTrue(dataLabels.show_legend_key)
#        self.assertTrue(dataLabels.show_percentage)
#        self.assertTrue(dataLabels.show_value)
#        self.assertEqual(" ", dataLabels.separator)
        

    #ExStart
    #ExFor:ChartSeries
    #ExFor:ChartSeries.data_labels
    #ExFor:ChartSeries.data_points
    #ExFor:ChartSeries.name
    #ExFor:ChartDataLabel
    #ExFor:ChartDataLabel.index
    #ExFor:ChartDataLabel.is_visible
    #ExFor:ChartDataLabel.number_format
    #ExFor:ChartDataLabel.separator
    #ExFor:ChartDataLabel.show_category_name
    #ExFor:ChartDataLabel.show_data_labels_range
    #ExFor:ChartDataLabel.show_leader_lines
    #ExFor:ChartDataLabel.show_legend_key
    #ExFor:ChartDataLabel.show_percentage
    #ExFor:ChartDataLabel.show_series_name
    #ExFor:ChartDataLabel.show_value
    #ExFor:ChartDataLabel.is_hidden
    #ExFor:ChartDataLabelCollection
    #ExFor:ChartDataLabelCollection.add(System.int_32)
    #ExFor:ChartDataLabelCollection.clear
    #ExFor:ChartDataLabelCollection.count
    #ExFor:ChartDataLabelCollection.get_enumerator
    #ExFor:ChartDataLabelCollection.item(System.int_32)
    #ExFor:ChartDataLabelCollection.remove_at(System.int_32)
    #ExSummary:Shows how to apply labels to data points in a line chart.
    def test_data_labels(self) :
        
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
            
        chartShape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 400, 300)
        chart = chartShape.chart

        self.assertEqual(3, chart.series.count)
        self.assertEqual("Series 1", chart.series[0].name)
        self.assertEqual("Series 2", chart.series[1].name)
        self.assertEqual("Series 3", chart.series[2].name)

        # Apply data labels to every series in the chart.
        # These labels will appear next to each data point in the graph and display its value.
        for series in chart.series :
            
            self.apply_data_labels(series, 4, "000.0", ", ")
            self.assertEqual(4, series.data_labels.count)
            

        # Change the separator string for every data label in a series.
        for data_label in chart.series[0].data_labels :
            self.assertEqual(", ", data_label.separator)
            data_label.separator = " & "
       
        # For a cleaner looking graph, we can remove data labels individually.
        chart.series[1].data_labels[2].clear_format()

        # We can also strip an entire series of its data labels at once.
        chart.series[2].data_labels.clear_format()

        doc.save(aeb.artifacts_dir + "Charts.data_labels.docx")
        

    # <summary>
    # Apply data labels with custom number format and separator to several data points in a series.
    # </summary>
    def apply_data_labels(self, series, labelsCount : int, numberFormat : str, separator : str) :
        
        for i in range(0, labelsCount) :
            
            series.has_data_labels = True

            self.assertFalse(series.data_labels[i].is_visible)

            series.data_labels[i].show_category_name = True
            series.data_labels[i].show_series_name = True
            series.data_labels[i].show_value = True
            series.data_labels[i].show_leader_lines = True
            series.data_labels[i].show_legend_key = True
            series.data_labels[i].show_percentage = False
            series.data_labels[i].is_hidden = False
            self.assertFalse(series.data_labels[i].show_data_labels_range)

            series.data_labels[i].number_format.format_code = numberFormat
            series.data_labels[i].separator = separator

            self.assertFalse(series.data_labels[i].show_data_labels_range)
            self.assertTrue(series.data_labels[i].is_visible)
            self.assertFalse(series.data_labels[i].is_hidden)
            
        
    #ExEnd

    #ExStart
    #ExFor:ChartSeries.smooth
    #ExFor:ChartDataPoint
    #ExFor:ChartDataPoint.index
    #ExFor:ChartDataPointCollection
    #ExFor:ChartDataPointCollection.clear_format
    #ExFor:ChartDataPointCollection.count
    #ExFor:ChartDataPointCollection.get_enumerator
    #ExFor:ChartDataPointCollection.item(System.int_32)
    #ExFor:ChartMarker
    #ExFor:ChartMarker.size
    #ExFor:ChartMarker.symbol
    #ExFor:IChartDataPoint
    #ExFor:IChartDataPoint.invert_if_negative
    #ExFor:IChartDataPoint.marker
    #ExFor:MarkerSymbol
    #ExSummary:Shows how to work with data points on a line chart.
    def test_chart_data_point(self) :
        
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 500, 350)
        chart = shape.chart

        self.assertEqual(3, chart.series.count)
        self.assertEqual("Series 1", chart.series[0].name)
        self.assertEqual("Series 2", chart.series[1].name)
        self.assertEqual("Series 3", chart.series[2].name)

        # Emphasize the chart's data points by making them appear as diamond shapes.
        for series in chart.series : 
            self.apply_data_points(series, 4, aw.drawing.charts.MarkerSymbol.DIAMOND, 15)

        # Smooth out the line that represents the first data series.
        chart.series[0].smooth = True

        # Verify that data points for the first series will not invert their colors if the value is negative.
        for data_point in chart.series[0].data_points :
            self.assertFalse(data_point.invert_if_negative)
                
            

        # For a cleaner looking graph, we can clear format individually.
        chart.series[1].data_points[2].clear_format()

        # We can also strip an entire series of data points at once.
        chart.series[2].data_points.clear_format()

        doc.save(aeb.artifacts_dir + "Charts.chart_data_point.docx")
        

    # <summary>
    # Applies a number of data points to a series.
    # </summary>
    def apply_data_points(self, series, dataPointsCount : int, markerSymbol, dataPointSize : int) :
        
        for i in range(0, dataPointsCount) :
            
            point = series.data_points[i]
            point.marker.symbol = markerSymbol
            point.marker.size = dataPointSize

            self.assertEqual(i, point.index)
            
        
    #ExEnd

    def test_pie_chart_explosion(self) :
        
        #ExStart
        #ExFor:Charts.i_chart_data_point.explosion
        #ExSummary:Shows how to move the slices of a pie chart away from the center.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.PIE, 500, 350)
        chart = shape.chart

        self.assertEqual(1, chart.series.count)
        self.assertEqual("Sales", chart.series[0].name)

        # "Slices" of a pie chart may be moved away from the center by a distance via the respective data point's Explosion attribute.
        # Add a data point to the first portion of the pie chart and move it away from the center by 10 points.
        # Aspose.words create data points automatically if them does not exist.
        dataPoint = chart.series[0].data_points[0]
        dataPoint.explosion = 10

        # Displace the second portion by a greater distance.
        dataPoint = chart.series[0].data_points[1]
        dataPoint.explosion = 40

        doc.save(aeb.artifacts_dir + "Charts.pie_chart_explosion.docx")
        #ExEnd

# no type casting yet.
#        doc = new Document(aeb.artifacts_dir + "Charts.pie_chart_explosion.docx")
#        ChartSeries series = ((Shape)doc.get_child(NodeType.shape, 0, True)).chart.series[0]
#
#        self.assertEqual(10, series.data_points[0].explosion)
#        self.assertEqual(40, series.data_points[1].explosion)
        

    def test_bubble_3_d(self) :
        
        #ExStart
        #ExFor:Charts.chart_data_label.show_bubble_size
        #ExFor:Charts.i_chart_data_point.bubble_3_d
        #ExSummary:Shows how to use 3D effects with bubble charts.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.BUBBLE3_D, 500, 350)
        chart = shape.chart

        self.assertEqual(1, chart.series.count)
        self.assertEqual("Y-Values", chart.series[0].name)
        self.assertTrue(chart.series[0].bubble3_d)

        # Apply a data label to each bubble that displays its diameter.
        for i in range (0, 3) :
            
            chart.series[0].has_data_labels = True
            chart.series[0].data_labels[i].show_bubble_size = True
            
            
        doc.save(aeb.artifacts_dir + "Charts.bubble_3_d.docx")
        #ExEnd

# no type casting yet.
#        doc = new Document(aeb.artifacts_dir + "Charts.bubble_3_d.docx")
#        ChartSeries series = ((Shape)doc.get_child(NodeType.shape, 0, True)).chart.series[0]
#
#        for (int i = 0 i < 3 i++)
#            
#            self.assertTrue(series.data_labels[i].show_bubble_size)
            
        

    #ExStart
    #ExFor:ChartAxis.type
    #ExFor:ChartAxisType
    #ExFor:ChartType
    #ExFor:Chart.series
    #ExFor:ChartSeriesCollection.add(String,DateTime[],Double[])
    #ExFor:ChartSeriesCollection.add(String,Double[],Double[])
    #ExFor:ChartSeriesCollection.add(String,Double[],Double[],Double[])
    #ExFor:ChartSeriesCollection.add(String,String[],Double[])
    #ExSummary:Shows how to create an appropriate type of chart series for a graph type.
    def test_chart_series_collection(self) :
        
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
            
        # There are several ways of populating a chart's series collection.
        # Different series schemas are intended for different chart types.
        # 1 -  Column chart with columns grouped and banded along the X-axis by category:
        chart = self.append_chart(builder, aw.drawing.charts.ChartType.COLUMN, 500, 300)

        categories =  ["Category 1", "Category 2", "Category 3"] 

        # Insert two series of decimal values containing a value for each respective category.
        # This column chart will have three groups, each with two columns.
        chart.series.add("Series 1", categories, [76.6, 82.1, 91.6])
        chart.series.add("Series 2", categories, [64.2, 79.5, 94.0])

        # Categories are distributed along the X-axis, and values are distributed along the Y-axis.
        self.assertEqual(aw.drawing.charts.ChartAxisType.CATEGORY, chart.axis_x.type)
        self.assertEqual(aw.drawing.charts.ChartAxisType.VALUE, chart.axis_y.type)

        # 2 -  Area chart with dates distributed along the X-axis:
        chart = self.append_chart(builder, aw.drawing.charts.ChartType.AREA, 500, 300)

        dates =  [datetime(2014, 3, 31), datetime(2017, 1, 23), datetime(2017, 6, 18), datetime(2019, 11, 22), datetime(2020, 9, 7)]
            
        # Insert a series with a decimal value for each respective date.
        # The dates will be distributed along a linear X-axis,
        # and the values added to this series will create data points.
        chart.series.add("Series 1", dates, [15.8, 21.5, 22.9, 28.7, 33.1])

        self.assertEqual(aw.drawing.charts.ChartAxisType.CATEGORY, chart.axis_x.type)
        self.assertEqual(aw.drawing.charts.ChartAxisType.VALUE, chart.axis_y.type)

        # 3 -  2D scatter plot:
        chart = self.append_chart(builder, aw.drawing.charts.ChartType.SCATTER, 500, 300)

        # Each series will need two decimal arrays of equal length.
        # The first array contains X-values, and the second contains corresponding Y-values
        # of data points on the chart's graph.
        chart.series.add("Series 1", 
            [3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6], 
            [3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9])
        chart.series.add("Series 2", 
            [2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3], 
            [7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6])

        self.assertEqual(aw.drawing.charts.ChartAxisType.VALUE, chart.axis_x.type)
        self.assertEqual(aw.drawing.charts.ChartAxisType.VALUE, chart.axis_y.type)

        # 4 -  Bubble chart:
        chart = self.append_chart(builder, ChartType.bubble, 500, 300)

        # Each series will need three decimal arrays of equal length.
        # The first array contains X-values, the second contains corresponding Y-values,
        # and the third contains diameters for each of the graph's data points.
        chart.series.add("Series 1", 
            [1.1, 5.0, 9.8], 
            [1.2, 4.9, 9.9], 
            [2.0, 4.0, 8.0])

        doc.save(aeb.artifacts_dir + "Charts.chart_series_collection.docx")
        
        
    # <summary>
    # Insert a chart using a document builder of a specified ChartType, width and height, and remove its demo data.
    # </summary>
    def append_chart(self, builder, chartType, width, height) :
        
        chartShape = builder.insert_chart(chartType, width, height)
        chart = chartShape.chart
        chart.series.clear()
        self.assertEqual(0, chart.series.count) #ExSkip

        return chart
        
    #ExEnd

    def test_chart_series_collection_modify(self) :
        
        #ExStart
        #ExFor:ChartSeriesCollection
        #ExFor:ChartSeriesCollection.clear
        #ExFor:ChartSeriesCollection.count
        #ExFor:ChartSeriesCollection.get_enumerator
        #ExFor:ChartSeriesCollection.item(Int32)
        #ExFor:ChartSeriesCollection.remove_at(Int32)
        #ExSummary:Shows how to add and remove series data in a chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a column chart that will contain three series of demo data by default.
        chartShape = builder.insert_chart(aw.drawing.charts.ChartType.COLUMN, 400, 300)
        chart = chartShape.chart

        # Each series has four decimal values: one for each of the four categories.
        # Four clusters of three columns will represent this data.
        chartData = chart.series

        self.assertEqual(3, chartData.count)

        # Print the name of every series in the chart.
        for series in chart.series :
            print(series.name)

        # These are the names of the categories in the chart.
        categories =  ["Category 1", "Category 2", "Category 3", "Category 4"] 

        # We can add a series with new values for existing categories.
        # This chart will now contain four clusters of four columns.
        chart.series.add("Series 4", categories, [4.4, 7.0, 3.5, 2.1])
        self.assertEqual(4, chartData.count) #ExSkip
        self.assertEqual("Series 4", chartData[3].name) #ExSkip
            
        # A chart series can also be removed by index, like this.
        # This will remove one of the three demo series that came with the chart.
        chartData.remove_at(2)

#        self.assertFalse(chartData.any(s => s.name == "Series 3"))
        self.assertEqual(3, chartData.count) #ExSkip
        self.assertEqual("Series 4", chartData[2].name) #ExSkip

        # We can also clear all the chart's data at once with this method.
        # When creating a new chart, this is the way to wipe all the demo data
        # before we can begin working on a blank chart.
        chartData.clear()
        self.assertEqual(0, chartData.count) #ExSkip
        #ExEnd
        

    def test_axis_scaling(self) :
        
        #ExStart
        #ExFor:AxisScaleType
        #ExFor:AxisScaling
        #ExFor:AxisScaling.log_base
        #ExFor:AxisScaling.type
        #ExSummary:Shows how to apply logarithmic scaling to a chart axis.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        chartShape = builder.insert_chart(aw.drawing.charts.ChartType.SCATTER, 450, 300)
        chart = chartShape.chart

        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()

        # Insert a series with X/Y coordinates for five points.
        chart.series.add("Series 1", 
            [1.0, 2.0, 3.0, 4.0, 5.0], 
            [1.0, 20.0, 400.0, 8000.0, 160000.0])

        # The scaling of the X-axis is linear by default,
        # displaying evenly incrementing values that cover our X-value range (0, 1, 2, 3...).
        # A linear axis is not ideal for our Y-values
        # since the points with the smaller Y-values will be harder to read.
        # A logarithmic scaling with a base of 20 (1, 20, 400, 8000...)
        # will spread the plotted points, allowing us to read their values on the chart more easily.
        chart.axis_y.scaling.type = aw.drawing.charts.AxisScaleType.LOGARITHMIC
        chart.axis_y.scaling.log_base = 20

        doc.save(aeb.artifacts_dir + "Charts.axis_scaling.docx")
        #ExEnd

# no type casting yet.
#        doc = new Document(aeb.artifacts_dir + "Charts.axis_scaling.docx")
#        chart = ((Shape)doc.get_child(NodeType.shape, 0, True)).chart
#
#        self.assertEqual(aw.drawing.charts.ChartType.LINEAR, chart.axis_x.scaling.type)
#        self.assertEqual(aw.drawing.charts.AxisScaleType.LOGARITHMIC, chart.axis_y.scaling.type)
#        self.assertEqual(20.0d, chart.axis_y.scaling.log_base)
        

    def test_axis_bound(self) :
        
        #ExStart
        #ExFor:AxisBound.#ctor
        #ExFor:AxisBound.is_auto
        #ExFor:AxisBound.value
        #ExFor:AxisBound.value_as_date
        #ExSummary:Shows how to set custom axis bounds.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        chartShape = builder.insert_chart(aw.drawing.charts.ChartType.SCATTER, 450, 300)
        chart = chartShape.chart

        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()

        # Add a series with two decimal arrays. The first array contains the X-values,
        # and the second contains corresponding Y-values for points in the scatter chart.
        chart.series.add("Series 1", 
            [1.1, 5.4, 7.9, 3.5, 2.1, 9.7],
            [2.1, 0.3, 0.6, 3.3, 1.4, 1.9])

        # By default, default scaling is applied to the graph's X and Y-axes,
        # so that both their ranges are big enough to encompass every X and Y-value of every series.
        self.assertTrue(chart.axis_x.scaling.minimum.is_auto)

        # We can define our own axis bounds.
        # In this case, we will make both the X and Y-axis rulers show a range of 0 to 10.
        chart.axis_x.scaling.minimum = aw.drawing.charts.AxisBound(0)
        chart.axis_x.scaling.maximum = aw.drawing.charts.AxisBound(10)
        chart.axis_y.scaling.minimum = aw.drawing.charts.AxisBound(0)
        chart.axis_y.scaling.maximum = aw.drawing.charts.AxisBound(10)

        self.assertFalse(chart.axis_x.scaling.minimum.is_auto)
        self.assertFalse(chart.axis_y.scaling.minimum.is_auto)

        # Create a line chart with a series requiring a range of dates on the X-axis, and decimal values for the Y-axis.
        chartShape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 450, 300)
        chart = chartShape.chart
        chart.series.clear()

        dates =  [date(1973, 5, 11), date(1981, 2, 4), date(1985, 9, 23), date(1989, 6, 28), date(1994, 12, 15)]
            

        chart.series.add("Series 1", dates, [3.0, 4.7, 5.9, 7.1, 8.9])

        # We can set axis bounds in the form of dates as well, limiting the chart to a period.
        # Setting the range to 1980-1990 will omit the two of the series values
        # that are outside of the range from the graph.
        chart.axis_x.scaling.minimum = aw.drawing.charts.AxisBound(date(1980, 1, 1))
        chart.axis_x.scaling.maximum = aw.drawing.charts.AxisBound(date(1990, 1, 1))

        doc.save(aeb.artifacts_dir + "Charts.axis_bound.docx")
        #ExEnd
# no type casting yet.
#        doc = new Document(aeb.artifacts_dir + "Charts.axis_bound.docx")
#        chart = ((Shape)doc.get_child(NodeType.shape, 0, True)).chart
#
#        self.assertFalse(chart.axis_x.scaling.minimum.is_auto)
#        self.assertEqual(0.0d, chart.axis_x.scaling.minimum.value)
#        self.assertEqual(10.0d, chart.axis_x.scaling.maximum.value)
#
#        self.assertFalse(chart.axis_y.scaling.minimum.is_auto)
#        self.assertEqual(0.0d, chart.axis_y.scaling.minimum.value)
#        self.assertEqual(10.0d, chart.axis_y.scaling.maximum.value)
#
#        chart = ((Shape)doc.get_child(NodeType.shape, 1, True)).chart
#
#        self.assertFalse(chart.axis_x.scaling.minimum.is_auto)
#        self.assertEqual(aw.drawing.charts.AxisBound(date(1980, 1, 1)), chart.axis_x.scaling.minimum)
#        self.assertEqual(aw.drawing.charts.AxisBound(date(1990, 1, 1)), chart.axis_x.scaling.maximum)
#
#        self.assertTrue(chart.axis_y.scaling.minimum.is_auto)
        

    def test_chart_legend(self) :
        
        #ExStart
        #ExFor:Chart.legend
        #ExFor:ChartLegend
        #ExFor:ChartLegend.overlay
        #ExFor:ChartLegend.position
        #ExFor:LegendPosition
        #ExSummary:Shows how to edit the appearance of a chart's legend.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 450, 300)
        chart = shape.chart

        self.assertEqual(3, chart.series.count)
        self.assertEqual("Series 1", chart.series[0].name)
        self.assertEqual("Series 2", chart.series[1].name)
        self.assertEqual("Series 3", chart.series[2].name)

        # Move the chart's legend to the top right corner.
        legend = chart.legend
        legend.position = aw.drawing.charts.LegendPosition.TOP_RIGHT

        # Give other chart elements, such as the graph, more room by allowing them to overlap the legend.
        legend.overlay = True

        doc.save(aeb.artifacts_dir + "Charts.chart_legend.docx")
        #ExEnd

# no type casting yet.
#        doc = new Document(aeb.artifacts_dir + "Charts.chart_legend.docx")
#
#        legend = ((Shape)doc.get_child(NodeType.shape, 0, True)).chart.legend
#
#        self.assertTrue(legend.overlay)
#        self.assertEqual(LegendPosition.TOP_RIGHT, legend.position)
        

    def test_axis_cross(self) :
        
        #ExStart
        #ExFor:ChartAxis.axis_between_categories
        #ExFor:ChartAxis.crosses_at
        #ExSummary:Shows how to get a graph axis to cross at a custom location.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.COLUMN, 450, 250)
        chart = shape.chart

        self.assertEqual(3, chart.series.count)
        self.assertEqual("Series 1", chart.series[0].name)
        self.assertEqual("Series 2", chart.series[1].name)
        self.assertEqual("Series 3", chart.series[2].name)

        # For column charts, the Y-axis crosses at zero by default,
        # which means that columns for all values below zero point down to represent negative values.
        # We can set a different value for the Y-axis crossing. In this case, we will set it to 3.
        axis = chart.axis_x
        axis.crosses = aw.drawing.charts.AxisCrosses.CUSTOM
        axis.crosses_at = 3
        axis.axis_between_categories = True

        doc.save(aeb.artifacts_dir + "Charts.axis_cross.docx")
        #ExEnd

# no type casting yet.
#        doc = new Document(aeb.artifacts_dir + "Charts.axis_cross.docx")
#        axis = ((Shape)doc.get_child(NodeType.shape, 0, True)).chart.axis_x
#
#        self.assertTrue(axis.axis_between_categories)
#        self.assertEqual(aw.drawing.charts.AxisCrosses.CUSTOM, axis.crosses)
#        self.assertEqual(3.0d, axis.crosses_at)
        

    def test_axis_display_unit(self) :
        
        #ExStart
        #ExFor:AxisBuiltInUnit
        #ExFor:ChartAxis.display_unit
        #ExFor:ChartAxis.major_unit_is_auto
        #ExFor:ChartAxis.major_unit_scale
        #ExFor:ChartAxis.minor_unit_is_auto
        #ExFor:ChartAxis.minor_unit_scale
        #ExFor:ChartAxis.tick_label_spacing
        #ExFor:ChartAxis.tick_label_alignment
        #ExFor:AxisDisplayUnit
        #ExFor:AxisDisplayUnit.custom_unit
        #ExFor:AxisDisplayUnit.unit
        #ExSummary:Shows how to manipulate the tick marks and displayed values of a chart axis.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.SCATTER, 450, 250)
        chart = shape.chart

        self.assertEqual(1, chart.series.count)
        self.assertEqual("Y-Values", chart.series[0].name)

        # Set the minor tick marks of the Y-axis to point away from the plot area,
        # and the major tick marks to cross the axis.
        axis = chart.axis_y
        axis.major_tick_mark = aw.drawing.charts.AxisTickMark.CROSS
        axis.minor_tick_mark = aw.drawing.charts.AxisTickMark.OUTSIDE

        # Set they Y-axis to show a major tick every 10 units, and a minor tick every 1 unit.
        axis.major_unit = 10
        axis.minor_unit = 1
            
        # Set the Y-axis bounds to -10 and 20.
        # This Y-axis will now display 4 major tick marks and 27 minor tick marks.
        axis.scaling.minimum = aw.drawing.charts.AxisBound(-10)
        axis.scaling.maximum = aw.drawing.charts.AxisBound(20)

        # For the X-axis, set the major tick marks at every 10 units,
        # every minor tick mark at 2.5 units.
        axis = chart.axis_x
        axis.major_unit = 10
        axis.minor_unit = 2.5

        # Configure both types of tick marks to appear inside the graph plot area.
        axis.major_tick_mark = aw.drawing.charts.AxisTickMark.INSIDE
        axis.minor_tick_mark = aw.drawing.charts.AxisTickMark.INSIDE

        # Set the X-axis bounds so that the X-axis spans 5 major tick marks and 12 minor tick marks.
        axis.scaling.minimum = aw.drawing.charts.AxisBound(-10)
        axis.scaling.maximum = aw.drawing.charts.AxisBound(30)
        axis.tick_label_alignment = aw.ParagraphAlignment.RIGHT

        self.assertEqual(1, axis.tick_label_spacing)
            
        # Set the tick labels to display their value in millions.
        axis.display_unit.unit = aw.drawing.charts.AxisBuiltInUnit.MILLIONS

        # We can set a more specific value by which tick labels will display their values.
        # This statement is equivalent to the one above.
        axis.display_unit.custom_unit = 1000000
        self.assertEqual(aw.drawing.charts.AxisBuiltInUnit.CUSTOM, axis.display_unit.unit) #ExSkip

        doc.save(aeb.artifacts_dir + "Charts.axis_display_unit.docx")
        #ExEnd

# no type casting yet.
#        doc = new Document(aeb.artifacts_dir + "Charts.axis_display_unit.docx")
#        shape = (Shape)doc.get_child(NodeType.shape, 0, True)
#
#        self.assertEqual(450.0d, shape.width)
#        self.assertEqual(250.0d, shape.height)
#
#        axis = shape.chart.axis_x
#
#        self.assertEqual(aw.drawing.charts.AxisTickMark.INSIDE, axis.major_tick_mark)
#        self.assertEqual(aw.drawing.charts.AxisTickMark.INSIDE, axis.minor_tick_mark)
#        self.assertEqual(10.0d, axis.major_unit)
#        self.assertEqual(-10.0d, axis.scaling.minimum.value)
#        self.assertEqual(30.0d, axis.scaling.maximum.value)
#        self.assertEqual(1, axis.tick_label_spacing)
#        self.assertEqual(aw.ParagraphAlignment.RIGHT, axis.tick_label_alignment)
#        self.assertEqual(aw.drawing.charts.AxisBuiltInUnit.CUSTOM, axis.display_unit.unit)
#        self.assertEqual(1000000.0d, axis.display_unit.custom_unit)
#
#        axis = shape.chart.axis_y
#
#        self.assertEqual(aw.drawing.charts.AxisTickMark.CROSS, axis.major_tick_mark)
#        self.assertEqual(aw.drawing.charts.AxisTickMark.OUTSIDE, axis.minor_tick_mark)
#        self.assertEqual(10.0d, axis.major_unit)
#        self.assertEqual(1.0d, axis.minor_unit)
#        self.assertEqual(-10.0d, axis.scaling.minimum.value)
#        self.assertEqual(20.0d, axis.scaling.maximum.value)
        

    def test_marker_formatting(self) :
        
        #ExStart
        #ExFor:ChartMarker.format
        #ExFor:ChartFormat.fill
        #ExFor:ChartFormat.stroke
        #ExFor:Stroke.fore_color
        #ExFor:Stroke.back_color
        #ExFor:Stroke.visible
        #ExFor:Stroke.transparency
        #ExFor:Fill.preset_textured(PresetTexture)
        #ExSummary:Show how to set marker formatting.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.SCATTER, 432, 252)
        chart = shape.chart
            
        # Delete default generated series.
        chart.series.clear()
        series = chart.series.add("AW Series 1", [0.7, 1.8, 2.6, 3.9], [2.7, 3.2, 0.8, 1.7])

        # Set marker formatting.
        series.marker.size = 40
        series.marker.symbol = aw.drawing.charts.MarkerSymbol.SQUARE
        dataPoints = series.data_points
        dataPoints[0].marker.format.fill.preset_textured(aw.drawing.PresetTexture.DENIM)
#        dataPoints[0].marker.format.stroke.fore_color = Color.yellow
#        dataPoints[0].marker.format.stroke.back_color = Color.red
        dataPoints[1].marker.format.fill.preset_textured(aw.drawing.PresetTexture.WATER_DROPLETS)
#        dataPoints[1].marker.format.stroke.fore_color = Color.yellow
        dataPoints[1].marker.format.stroke.visible = False
        dataPoints[2].marker.format.fill.preset_textured(aw.drawing.PresetTexture.GREEN_MARBLE)
#        dataPoints[2].marker.format.stroke.fore_color = Color.yellow
        dataPoints[3].marker.format.fill.preset_textured(aw.drawing.PresetTexture.OAK)
#        dataPoints[3].marker.format.stroke.fore_color = Color.yellow
        dataPoints[3].marker.format.stroke.transparency = 0.5

        doc.save(aeb.artifacts_dir + "Charts.marker_formatting.docx")
        #ExEnd
        

    def test_series_color(self) :
        
        #ExStart
        #ExFor:ChartSeries.format
        #ExSummary:Sows how to set series color.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.COLUMN, 432, 252)

        chart = shape.chart
        seriesColl = chart.series

        # Delete default generated series.
        seriesColl.clear()

        # Create category names array.
        categories = ["Category 1", "Category 2"]

        # Adding new series. Value and category arrays must be the same size.
        series1 = seriesColl.add("Series 1", categories, [1, 2])
        series2 = seriesColl.add("Series 2", categories, [3, 4])
        series3 = seriesColl.add("Series 3", categories, [5, 6])

        # Set series color.
#        series1.format.fill.fore_color = Color.red
#        series2.format.fill.fore_color = Color.yellow
#        series3.format.fill.fore_color = Color.blue

        doc.save(aeb.artifacts_dir + "Charts.series_color.docx")
        #ExEnd
        

    def test_data_points_formatting(self) :
        
        #ExStart
        #ExFor:ChartDataPoint.format
        #ExSummary:Shows how to set individual formatting for categories of a column chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.COLUMN, 432, 252)
        chart = shape.chart

        # Delete default generated series.
        chart.series.clear()

        # Adding new series.
        series = chart.series.add("Series 1",
            ["Category 1", "Category 2", "Category 3", "Category 4"],
            [1, 2, 3, 4])

        # Set column formatting.
        dataPoints = series.data_points
        dataPoints[0].format.fill.preset_textured(aw.drawing.PresetTexture.DENIM)
#        dataPoints[1].format.fill.fore_color = Color.red
#        dataPoints[2].format.fill.fore_color = Color.yellow
#        dataPoints[3].format.fill.fore_color = Color.blue

        doc.save(aeb.artifacts_dir + "Charts.data_points_formatting.docx")
        #ExEnd
        
    
if __name__ == '__main__':
    unittest.main() 