# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

from math import nan
from datetime import date

import aspose.words as aw
import aspose.words.drawing as awd
import aspose.words.drawing.charts as awdc
import aspose.pydrawing as drawing
from aspose.words import Document, DocumentBuilder, NodeType
from aspose.pydrawing import Color
from aspose.words.drawing.charts import ChartXValue, ChartYValue, ChartSeriesType, ChartType, ChartShapeType
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExCharts(ApiExampleBase):

    def test_chart_title(self):
        #ExStart
        #ExFor:Chart
        #ExFor:Chart.title
        #ExFor:ChartTitle
        #ExFor:ChartTitle.overlay
        #ExFor:ChartTitle.show
        #ExFor:ChartTitle.text
        #ExFor:ChartTitle.font
        #ExSummary:Shows how to insert a chart and set a title.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a chart shape with a document builder and get its chart.
        chart_shape = builder.insert_chart(aw.drawing.charts.ChartType.BAR, 400, 300)
        chart = chart_shape.chart

        # Use the "title" property to give our chart a title, which appears at the top center of the chart area.
        title = chart.title
        title.text = "My Chart"
        title.font.size = 15
        title.font.color = drawing.Color.blue

        # Set the "show" property to "True" to make the title visible.
        title.show = True

        # Set the "overlay" property to "True" Give other chart elements more room by allowing them to overlap the title
        title.overlay = True

        doc.save(ARTIFACTS_DIR + "Charts.chart_title.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Charts.chart_title.docx")
        chart_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.assertEqual(aw.drawing.ShapeType.NON_PRIMITIVE, chart_shape.shape_type)
        self.assertTrue(chart_shape.has_chart)

        title = chart_shape.chart.title

        self.assertEqual("My Chart", title.text)
        self.assertTrue(title.overlay)
        self.assertTrue(title.show)

    def test_data_label_number_format(self):

        #ExStart
        #ExFor:ChartDataLabelCollection.number_format
        #ExFor:ChartDataLabelCollection.font
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
        series = chart.series.add("Revenue", ["January", "February", "March"], [25.611, 21.439, 33.750])

        # Enable data labels, and then apply a custom number format for values displayed in the data labels.
        # This format will treat displayed decimal values as millions of US Dollars.
        series.has_data_labels = True
        data_labels = series.data_labels
        data_labels.show_value = True
        data_labels.number_format.format_code = "\"US$\" #,##0.000\"M\""
        data_labels.font.size = 12

        doc.save(ARTIFACTS_DIR + "Charts.data_label_number_format.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Charts.data_label_number_format.docx")
        series = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart.series[0]

        self.assertTrue(series.has_data_labels)
        self.assertTrue(series.data_labels.show_value)
        self.assertEqual("\"US$\" #,##0.000\"M\"", series.data_labels.number_format.format_code)

    def test_data_arrays_wrong_size(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 500, 300)
        chart = shape.chart

        series_coll = chart.series
        series_coll.clear()

        categories = ["Cat1", None, "Cat3", "Cat4", "Cat5", None]
        series_coll.add("AW Series 1", categories, [1, 2, nan, 4, 5, 6])
        series_coll.add("AW Series 2", categories, [2, 3, nan, 5, 6, 7])

        with self.assertRaises(Exception):
            series_coll.add("AW Series 3", categories, [nan, 4, 5, nan, nan])

        with self.assertRaises(Exception):
            series_coll.add("AW Series 4", categories, [nan, nan, nan, nan, nan])

    def test_empty_values_in_chart_data(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 500, 300)
        chart = shape.chart

        series_coll = chart.series
        series_coll.clear()

        categories = ["Cat1", None, "Cat3", "Cat4", "Cat5", None]
        series_coll.add("AW Series 1", categories, [1, 2, nan, 4, 5, 6])
        series_coll.add("AW Series 2", categories, [2, 3, nan, 5, 6, 7])
        series_coll.add("AW Series 3", categories, [nan, 4, 5, nan, 7, 8])
        series_coll.add("AW Series 4", categories, [nan, nan, nan, nan, nan, 9])

        doc.save(ARTIFACTS_DIR + "Charts.empty_values_in_chart_data.docx")

    def test_axis_properties(self):

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
        #ExFor:AxisCategoryType
        #ExFor:AxisCrosses
        #ExFor:Chart.axis_x
        #ExFor:Chart.axis_y
        #ExFor:Chart.axis_z
        #ExSummary:Shows how to insert a chart and modify the appearance of its axes.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.COLUMN, 500, 300)
        chart = shape.chart

        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()

        # Insert a chart series with categories for the X-axis and respective numeric values for the Y-axis.
        chart.series.add("Aspose Test Series",
            ["Word", "PDF", "Excel", "GoogleDocs", "Note"],
            [640, 320, 280, 120, 150])

        # Chart axes have various options that can change their appearance,
        # such as their direction, major/minor unit ticks, and tick marks.
        x_axis = chart.axis_x
        x_axis.category_type = aw.drawing.charts.AxisCategoryType.CATEGORY
        x_axis.crosses = aw.drawing.charts.AxisCrosses.MINIMUM
        x_axis.reverse_order = False
        x_axis.major_tick_mark = aw.drawing.charts.AxisTickMark.INSIDE
        x_axis.minor_tick_mark = aw.drawing.charts.AxisTickMark.CROSS
        x_axis.major_unit = 10.0
        x_axis.minor_unit = 15.0
        x_axis.tick_label_offset = 50
        x_axis.tick_label_position = aw.drawing.charts.AxisTickLabelPosition.LOW
        x_axis.tick_label_spacing_is_auto = False
        x_axis.tick_mark_spacing = 1

        y_axis = chart.axis_y
        y_axis.category_type = aw.drawing.charts.AxisCategoryType.AUTOMATIC
        y_axis.crosses = aw.drawing.charts.AxisCrosses.MAXIMUM
        y_axis.reverse_order = True
        y_axis.major_tick_mark = aw.drawing.charts.AxisTickMark.INSIDE
        y_axis.minor_tick_mark = aw.drawing.charts.AxisTickMark.CROSS
        y_axis.major_unit = 100.0
        y_axis.minor_unit = 20.0
        y_axis.tick_label_position = aw.drawing.charts.AxisTickLabelPosition.NEXT_TO_AXIS

        # Column charts do not have a Z-axis.
        self.assertIsNone(chart.axis_z)

        doc.save(ARTIFACTS_DIR + "Charts.axis_properties.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Charts.axis_properties.docx")
        chart = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart

        self.assertEqual(aw.drawing.charts.AxisCategoryType.CATEGORY, chart.axis_x.category_type)
        self.assertEqual(aw.drawing.charts.AxisCrosses.MINIMUM, chart.axis_x.crosses)
        self.assertFalse(chart.axis_x.reverse_order)
        self.assertEqual(aw.drawing.charts.AxisTickMark.INSIDE, chart.axis_x.major_tick_mark)
        self.assertEqual(aw.drawing.charts.AxisTickMark.CROSS, chart.axis_x.minor_tick_mark)
        self.assertEqual(1.0, chart.axis_x.major_unit)
        self.assertEqual(0.5, chart.axis_x.minor_unit)
        self.assertEqual(50, chart.axis_x.tick_label_offset)
        self.assertEqual(aw.drawing.charts.AxisTickLabelPosition.LOW, chart.axis_x.tick_label_position)
        self.assertFalse(chart.axis_x.tick_label_spacing_is_auto)
        self.assertEqual(1, chart.axis_x.tick_mark_spacing)

        self.assertEqual(aw.drawing.charts.AxisCategoryType.CATEGORY, chart.axis_y.category_type)
        self.assertEqual(aw.drawing.charts.AxisCrosses.MAXIMUM, chart.axis_y.crosses)
        self.assertTrue(chart.axis_y.reverse_order)
        self.assertEqual(aw.drawing.charts.AxisTickMark.INSIDE, chart.axis_y.major_tick_mark)
        self.assertEqual(aw.drawing.charts.AxisTickMark.CROSS, chart.axis_y.minor_tick_mark)
        self.assertEqual(100.0, chart.axis_y.major_unit)
        self.assertEqual(20.0, chart.axis_y.minor_unit)
        self.assertEqual(aw.drawing.charts.AxisTickLabelPosition.NEXT_TO_AXIS, chart.axis_y.tick_label_position)

    def test_date_time_values(self):

        #ExStart
        #ExFor:AxisBound
        #ExFor:AxisBound.__init__(float)
        #ExFor:AxisBound.__init__(datetime)
        #ExFor:AxisScaling.minimum
        #ExFor:AxisScaling.maximum
        #ExFor:ChartAxis.scaling
        #ExFor:AxisTickMark
        #ExFor:AxisTickLabelPosition
        #ExFor:AxisTimeUnit
        #ExFor:ChartAxis.base_time_unit
        #ExFor:ChartAxis.has_major_gridlines
        #ExFor:ChartAxis.has_minor_gridlines
        #ExSummary:Shows how to insert chart with date/time values.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 500, 300)
        chart = shape.chart

        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()

        # Add a custom series containing date/time values for the X-axis, and respective decimal values for the Y-axis.
        dates = [
            date(2017, 11, 6),
            date(2017, 11, 9),
            date(2017, 11, 15),
            date(2017, 11, 21),
            date(2017, 11, 25),
            date(2017, 11, 29)
            ]
        chart.series.add("Aspose Test Series", dates=dates, values=[1.2, 0.3, 2.1, 2.9, 4.2, 5.3])

        # Set lower and upper bounds for the X-axis.
        x_axis = chart.axis_x
        x_axis.scaling.minimum = aw.drawing.charts.AxisBound(date(2017, 11, 5))
        x_axis.scaling.maximum = aw.drawing.charts.AxisBound(date(2017, 12, 3))

        # Set the major units of the X-axis to a week, and the minor units to a day.
        x_axis.base_time_unit = aw.drawing.charts.AxisTimeUnit.DAYS
        x_axis.major_unit = 7.0
        x_axis.major_tick_mark = aw.drawing.charts.AxisTickMark.CROSS
        x_axis.minor_unit = 1.0
        x_axis.minor_tick_mark = aw.drawing.charts.AxisTickMark.OUTSIDE
        x_axis.has_major_gridlines = True
        x_axis.has_minor_gridlines = True

        # Define Y-axis properties for decimal values.
        y_axis = chart.axis_y
        y_axis.tick_label_position = aw.drawing.charts.AxisTickLabelPosition.HIGH
        y_axis.major_unit = 100.0
        y_axis.minor_unit = 50.0
        y_axis.display_unit.unit = aw.drawing.charts.AxisBuiltInUnit.HUNDREDS
        y_axis.scaling.minimum = aw.drawing.charts.AxisBound(100)
        y_axis.scaling.maximum = aw.drawing.charts.AxisBound(700)
        y_axis.has_major_gridlines = True
        y_axis.has_minor_gridlines = True

        doc.save(ARTIFACTS_DIR + "Charts.date_time_values.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Charts.date_time_values.docx")
        chart = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart

        self.assertEqual(aw.drawing.charts.AxisBound(date(2017, 11, 5)), chart.axis_x.scaling.minimum)
        self.assertEqual(aw.drawing.charts.AxisBound(date(2017, 12, 3)), chart.axis_x.scaling.maximum)
        self.assertEqual(aw.drawing.charts.AxisTimeUnit.DAYS, chart.axis_x.base_time_unit)
        self.assertEqual(7.0, chart.axis_x.major_unit)
        self.assertEqual(1.0, chart.axis_x.minor_unit)
        self.assertEqual(aw.drawing.charts.AxisTickMark.CROSS, chart.axis_x.major_tick_mark)
        self.assertEqual(aw.drawing.charts.AxisTickMark.OUTSIDE, chart.axis_x.minor_tick_mark)
        self.assertEqual(True, chart.axis_x.has_major_gridlines)
        self.assertEqual(True, chart.axis_x.has_minor_gridlines)

        self.assertEqual(aw.drawing.charts.AxisTickLabelPosition.HIGH, chart.axis_y.tick_label_position)
        self.assertEqual(100.0, chart.axis_y.major_unit)
        self.assertEqual(50.0, chart.axis_y.minor_unit)
        self.assertEqual(aw.drawing.charts.AxisBuiltInUnit.HUNDREDS, chart.axis_y.display_unit.unit)
        self.assertEqual(aw.drawing.charts.AxisBound(100), chart.axis_y.scaling.minimum)
        self.assertEqual(aw.drawing.charts.AxisBound(700), chart.axis_y.scaling.maximum)
        self.assertEqual(True, chart.axis_y.has_major_gridlines)
        self.assertEqual(True, chart.axis_y.has_minor_gridlines)

    def test_hide_chart_axis(self):

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

        doc.save(ARTIFACTS_DIR + "Charts.hide_chart_axis.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Charts.hide_chart_axis.docx")
        chart = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart

        self.assertTrue(chart.axis_x.hidden)
        self.assertTrue(chart.axis_y.hidden)

    def test_set_number_format_to_chart_axis(self):

        #ExStart
        #ExFor:ChartAxis.number_format
        #ExFor:ChartNumberFormat
        #ExFor:ChartNumberFormat.format_code
        #ExFor:ChartNumberFormat.is_linked_to_source
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

        doc.save(ARTIFACTS_DIR + "Charts.set_number_format_to_chart_axis.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Charts.set_number_format_to_chart_axis.docx")
        chart = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart

        self.assertEqual("#,##0", chart.axis_y.number_format.format_code)

    def test_display_charts_with_conversion(self):

        for chart_type in (aw.drawing.charts.ChartType.COLUMN,
                           aw.drawing.charts.ChartType.LINE,
                           aw.drawing.charts.ChartType.PIE,
                           aw.drawing.charts.ChartType.BAR,
                           aw.drawing.charts.ChartType.AREA):
            with self.subTest(chart_type=chart_type):
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                shape = builder.insert_chart(chart_type, 500, 300)
                chart = shape.chart
                chart.series.clear()

                chart.series.add("Aspose Test Series",
                    ["Word", "PDF", "Excel", "GoogleDocs", "Note"],
                    [1900000, 850000, 2100000, 600000, 1500000])

                doc.save(ARTIFACTS_DIR + "Charts.display_charts_with_conversion.docx")
                doc.save(ARTIFACTS_DIR + "Charts.display_charts_with_conversion.pdf")

    def test_surface3d_chart(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.SURFACE_3D, 500, 300)
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

        doc.save(ARTIFACTS_DIR + "Charts.surface3d_chart.docx")
        doc.save(ARTIFACTS_DIR + "Charts.surface3d_chart.pdf")

    def test_data_labels_bubble_chart(self):

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
        data_labels = series.data_labels
        data_labels.show_bubble_size = True
        data_labels.show_category_name = True
        data_labels.show_series_name = True
        data_labels.separator = " & "

        doc.save(ARTIFACTS_DIR + "Charts.data_labels_bubble_chart.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Charts.data_labels_bubble_chart.docx")
        data_labels = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart.series[0].data_labels

        self.assertTrue(data_labels.show_bubble_size)
        self.assertTrue(data_labels.show_category_name)
        self.assertTrue(data_labels.show_series_name)
        self.assertEqual(" & ", data_labels.separator)

    def test_data_labels_pie_chart(self):

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
            ["Word", "PDF", "Excel"],
            [2.7, 3.2, 0.8])

        # Enable data labels that will display both percentage and frequency of each sector, and modify their appearance.
        series.has_data_labels = True
        data_labels = series.data_labels
        data_labels.show_leader_lines = True
        data_labels.show_legend_key = True
        data_labels.show_percentage = True
        data_labels.show_value = True
        data_labels.separator = "; "

        doc.save(ARTIFACTS_DIR + "Charts.data_labels_pie_chart.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Charts.data_labels_pie_chart.docx")
        data_labels = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart.series[0].data_labels

        self.assertTrue(data_labels.show_leader_lines)
        self.assertTrue(data_labels.show_legend_key)
        self.assertTrue(data_labels.show_percentage)
        self.assertTrue(data_labels.show_value)
        self.assertEqual("; ", data_labels.separator)

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
    #ExFor:ChartDataLabelCollection.clear_format
    #ExFor:ChartDataLabelCollection.count
    #ExFor:ChartDataLabelCollection.__iter__
    #ExFor:ChartDataLabelCollection.__getitem__(int)
    #ExSummary:Shows how to apply labels to data points in a line chart.
    def test_data_labels(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        chart_shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 400, 300)
        chart = chart_shape.chart

        self.assertEqual(3, chart.series.count)
        self.assertEqual("Series 1", chart.series[0].name)
        self.assertEqual("Series 2", chart.series[1].name)
        self.assertEqual("Series 3", chart.series[2].name)

        # Apply data labels to every series in the chart.
        # These labels will appear next to each data point in the graph and display its value.
        for series in chart.series:
            self.apply_data_labels(series, 4, "000.0", ", ")
            self.assertEqual(4, series.data_labels.count)

        # Change the separator string for every data label in a series.
        for label in chart.series[0].data_labels:
            self.assertEqual(", ", label.separator)
            label.separator = " & "

        # For a cleaner looking graph, we can remove data labels individually.
        chart.series[1].data_labels[2].clear_format()

        # We can also strip an entire series of its data labels at once.
        chart.series[2].data_labels.clear_format()

        doc.save(ARTIFACTS_DIR + "Charts.data_labels.docx")

    def apply_data_labels(self, series: aw.drawing.charts.ChartSeries, labels_count: int, number_format: str, separator: str):
        """Apply data labels with custom number format and separator to several data points in a series."""

        for i in range(labels_count):
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

            series.data_labels[i].number_format.format_code = number_format
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
    #ExFor:ChartDataPointCollection.__iter__
    #ExFor:ChartDataPointCollection.__getitem__(int)
    #ExFor:ChartMarker
    #ExFor:ChartMarker.size
    #ExFor:ChartMarker.symbol
    #ExFor:IChartDataPoint
    #ExFor:IChartDataPoint.invert_if_negative
    #ExFor:IChartDataPoint.marker
    #ExFor:MarkerSymbol
    #ExSummary:Shows how to work with data points on a line chart.
    def test_chart_data_point(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 500, 350)
        chart = shape.chart

        self.assertEqual(3, chart.series.count)
        self.assertEqual("Series 1", chart.series[0].name)
        self.assertEqual("Series 2", chart.series[1].name)
        self.assertEqual("Series 3", chart.series[2].name)

        # Emphasize the chart's data points by making them appear as diamond shapes.
        for series in chart.series:
            self.apply_data_points(series, 4, aw.drawing.charts.MarkerSymbol.DIAMOND, 15)

        # Smooth out the line that represents the first data series.
        chart.series[0].smooth = True

        # Verify that data points for the first series will not invert their colors if the value is negative.
        for data_point in chart.series[0].data_points:
            self.assertFalse(data_point.invert_if_negative)

        # For a cleaner looking graph, we can clear format individually.
        chart.series[1].data_points[2].clear_format()

        # We can also strip an entire series of data points at once.
        chart.series[2].data_points.clear_format()

        doc.save(ARTIFACTS_DIR + "Charts.chart_data_point.docx")

    def apply_data_points(self, series: aw.drawing.charts.ChartSeries, data_points_count: int, marker_symbol: aw.drawing.charts.MarkerSymbol, data_point_size: int):
        """Applies a number of data points to a series."""

        for i in range(data_points_count):
            point = series.data_points[i]
            point.marker.symbol = marker_symbol
            point.marker.size = data_point_size

            self.assertEqual(i, point.index)

    #ExEnd

    def test_pie_chart_explosion(self):

        #ExStart
        #ExFor:IChartDataPoint.explosion
        #ExSummary:Shows how to move the slices of a pie chart away from the center.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.PIE, 500, 350)
        chart = shape.chart

        self.assertEqual(1, chart.series.count)
        self.assertEqual("Sales", chart.series[0].name)

        # "Slices" of a pie chart may be moved away from the center by a distance via the respective data point's Explosion attribute.
        # Add a data point to the first portion of the pie chart and move it away from the center by 10 points.
        # Aspose.Words create data points automatically if them does not exist.
        data_point = chart.series[0].data_points[0]
        data_point.explosion = 10

        # Displace the second portion by a greater distance.
        data_point = chart.series[0].data_points[1]
        data_point.explosion = 40

        doc.save(ARTIFACTS_DIR + "Charts.pie_chart_explosion.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Charts.pie_chart_explosion.docx")
        series = (doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()).chart.series[0]

        self.assertEqual(10, series.data_points[0].explosion)
        self.assertEqual(40, series.data_points[1].explosion)

    def test_bubble_3d(self):

        #ExStart
        #ExFor:ChartDataLabel.show_bubble_size
        #ExFor:Charts.ChartDataLabel.font
        #ExFor:IChartDataPoint.bubble_3d
        #ExSummary:Shows how to use 3D effects with bubble charts.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.BUBBLE_3D, 500, 350)
        chart = shape.chart

        self.assertEqual(1, chart.series.count)
        self.assertEqual("Y-Values", chart.series[0].name)
        self.assertTrue(chart.series[0].bubble_3d)

        # Apply a data label to each bubble that displays its diameter.
        for i in range(3):
            chart.series[0].has_data_labels = True
            chart.series[0].data_labels[i].show_bubble_size = True
            chart.series[0].data_labels[i].font.size = 12

        doc.save(ARTIFACTS_DIR + "Charts.bubble_3d.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Charts.bubble_3d.docx")
        series = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart.series[0]

        for i in range(3):
            self.assertTrue(series.data_labels[i].show_bubble_size)

    #ExStart
    #ExFor:ChartAxis.type
    #ExFor:ChartAxisType
    #ExFor:ChartType
    #ExFor:Chart.series
    #ExFor:ChartSeriesCollection.add(str,List[datetime],List[float])
    #ExFor:ChartSeriesCollection.add(str,List[float],List[float])
    #ExFor:ChartSeriesCollection.add(str,List[float],List[float],List[float])
    #ExFor:ChartSeriesCollection.add(str,List[str],List[float])
    #ExSummary:Shows how to create an appropriate type of chart series for a graph type.
    def test_chart_series_collection(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # There are several ways of populating a chart's series collection.
        # Different series schemas are intended for different chart types.
        # 1 -  Column chart with columns grouped and banded along the X-axis by category:
        chart = self.append_chart(builder, aw.drawing.charts.ChartType.COLUMN, 500, 300)

        categories = ["Category 1", "Category 2", "Category 3"]

        # Insert two series of decimal values containing a value for each respective category.
        # This column chart will have three groups, each with two columns.
        chart.series.add("Series 1", categories, [76.6, 82.1, 91.6])
        chart.series.add("Series 2", categories, [64.2, 79.5, 94.0])

        # Categories are distributed along the X-axis, and values are distributed along the Y-axis.
        self.assertEqual(aw.drawing.charts.ChartAxisType.CATEGORY, chart.axis_x.type)
        self.assertEqual(aw.drawing.charts.ChartAxisType.VALUE, chart.axis_y.type)

        # 2 -  Area chart with dates distributed along the X-axis:
        chart = self.append_chart(builder, aw.drawing.charts.ChartType.AREA, 500, 300)

        dates = [
            date(2014, 3, 31),
            date(2017, 1, 23),
            date(2017, 6, 18),
            date(2019, 11, 22),
            date(2020, 9, 7)
            ]

        # Insert a series with a decimal value for each respective date.
        # The dates will be distributed along a linear X-axis,
        # and the values added to this series will create data points.
        chart.series.add("Series 1", dates=dates, values=[15.8, 21.5, 22.9, 28.7, 33.1])

        self.assertEqual(aw.drawing.charts.ChartAxisType.CATEGORY, chart.axis_x.type)
        self.assertEqual(aw.drawing.charts.ChartAxisType.VALUE, chart.axis_y.type)

        # 3 -  2D scatter plot:
        chart = self.append_chart(builder, aw.drawing.charts.ChartType.SCATTER, 500, 300)

        # Each series will need two decimal arrays of equal length.
        # The first array contains X-values, and the second contains corresponding Y-values
        # of data points on the chart's graph.
        chart.series.add("Series 1",
            x_values=[3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6],
            y_values=[3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9])
        chart.series.add("Series 2",
            x_values=[2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3],
            y_values=[7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6])

        self.assertEqual(aw.drawing.charts.ChartAxisType.VALUE, chart.axis_x.type)
        self.assertEqual(aw.drawing.charts.ChartAxisType.VALUE, chart.axis_y.type)

        # 4 -  Bubble chart:
        chart = self.append_chart(builder, aw.drawing.charts.ChartType.BUBBLE, 500, 300)

        # Each series will need three decimal arrays of equal length.
        # The first array contains X-values, the second contains corresponding Y-values,
        # and the third contains diameters for each of the graph's data points.
        chart.series.add("Series 1",
            [1.1, 5.0, 9.8],
            [1.2, 4.9, 9.9],
            [2.0, 4.0, 8.0])

        doc.save(ARTIFACTS_DIR + "Charts.chart_series_collection.docx")

    def append_chart(self, builder: aw.DocumentBuilder,
                     chart_type: aw.drawing.charts.ChartType,
                     width: float, height: float) -> aw.drawing.charts.Chart:
        """Insert a chart using a document builder of a specified ChartType, width and height, and remove its demo data."""

        chart_shape = builder.insert_chart(chart_type, width, height)
        chart = chart_shape.chart
        chart.series.clear()
        self.assertEqual(0, chart.series.count) #ExSkip

        return chart

    #ExEnd

    def test_chart_series_collection_modify(self):

        #ExStart
        #ExFor:ChartSeriesCollection
        #ExFor:ChartSeriesCollection.clear
        #ExFor:ChartSeriesCollection.count
        #ExFor:ChartSeriesCollection.__iter__
        #ExFor:ChartSeriesCollection.__getitem__(int)
        #ExFor:ChartSeriesCollection.remove_at(int)
        #ExSummary:Shows how to add and remove series data in a chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a column chart that will contain three series of demo data by default.
        chart_shape = builder.insert_chart(aw.drawing.charts.ChartType.COLUMN, 400, 300)
        chart = chart_shape.chart

        # Each series has four decimal values: one for each of the four categories.
        # Four clusters of three columns will represent this data.
        chart_data = chart.series

        self.assertEqual(3, chart_data.count)

        # Print the name of every series in the chart.
        for series in chart.series:
            print(series.name)

        # These are the names of the categories in the chart.
        categories = ["Category 1", "Category 2", "Category 3", "Category 4"]

        # We can add a series with new values for existing categories.
        # This chart will now contain four clusters of four columns.
        chart.series.add("Series 4", categories, [4.4, 7.0, 3.5, 2.1])
        self.assertEqual(4, chart_data.count) #ExSkip
        self.assertEqual("Series 4", chart_data[3].name) #ExSkip

        # A chart series can also be removed by index, like this.
        # This will remove one of the three demo series that came with the chart.
        chart_data.remove_at(2)

        self.assertFalse(any(s for s in chart_data if s.name == "Series 3"))
        self.assertEqual(3, chart_data.count) #ExSkip
        self.assertEqual("Series 4", chart_data[2].name) #ExSkip

        # We can also clear all the chart's data at once with this method.
        # When creating a new chart, this is the way to wipe all the demo data
        # before we can begin working on a blank chart.
        chart_data.clear()
        self.assertEqual(0, chart_data.count) #ExSkip
        #ExEnd

    def test_axis_scaling(self):

        #ExStart
        #ExFor:AxisScaleType
        #ExFor:AxisScaling
        #ExFor:AxisScaling.log_base
        #ExFor:AxisScaling.type
        #ExSummary:Shows how to apply logarithmic scaling to a chart axis.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        chart_shape = builder.insert_chart(aw.drawing.charts.ChartType.SCATTER, 450, 300)
        chart = chart_shape.chart

        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()

        # Insert a series with X/Y coordinates for five points.
        chart.series.add("Series 1",
            x_values=[1.0, 2.0, 3.0, 4.0, 5.0],
            y_values=[1.0, 20.0, 400.0, 8000.0, 160000.0])

        # The scaling of the X-axis is linear by default,
        # displaying evenly incrementing values that cover our X-value range (0, 1, 2, 3...).
        # A linear axis is not ideal for our Y-values
        # since the points with the smaller Y-values will be harder to read.
        # A logarithmic scaling with a base of 20 (1, 20, 400, 8000...)
        # will spread the plotted points, allowing us to read their values on the chart more easily.
        chart.axis_y.scaling.type = aw.drawing.charts.AxisScaleType.LOGARITHMIC
        chart.axis_y.scaling.log_base = 20

        doc.save(ARTIFACTS_DIR + "Charts.axis_scaling.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Charts.axis_scaling.docx")
        chart = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart

        self.assertEqual(aw.drawing.charts.AxisScaleType.LINEAR, chart.axis_x.scaling.type)
        self.assertEqual(aw.drawing.charts.AxisScaleType.LOGARITHMIC, chart.axis_y.scaling.type)
        self.assertEqual(20.0, chart.axis_y.scaling.log_base)

    def test_axis_bound(self):

        #ExStart
        #ExFor:AxisBound.__init__(float)
        #ExFor:AxisBound.__init__(datetime)
        #ExFor:AxisBound.is_auto
        #ExFor:AxisBound.value
        #ExFor:AxisBound.value_as_date
        #ExSummary:Shows how to set custom axis bounds.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        chart_shape = builder.insert_chart(aw.drawing.charts.ChartType.SCATTER, 450, 300)
        chart = chart_shape.chart

        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()

        # Add a series with two decimal arrays. The first array contains the X-values,
        # and the second contains corresponding Y-values for points in the scatter chart.
        chart.series.add("Series 1",
            x_values = [1.1, 5.4, 7.9, 3.5, 2.1, 9.7],
            y_values = [2.1, 0.3, 0.6, 3.3, 1.4, 1.9])

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
        chart_shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 450, 300)
        chart = chart_shape.chart
        chart.series.clear()

        dates = [
            date(1973, 5, 11),
            date(1981, 2, 4),
            date(1985, 9, 23),
            date(1989, 6, 28),
            date(1994, 12, 15)
            ]

        chart.series.add("Series 1", dates=dates, values=[3.0, 4.7, 5.9, 7.1, 8.9])

        # We can set axis bounds in the form of dates as well, limiting the chart to a period.
        # Setting the range to 1980-1990 will omit the two of the series values
        # that are outside of the range from the graph.
        chart.axis_x.scaling.minimum = aw.drawing.charts.AxisBound(date(1980, 1, 1))
        chart.axis_x.scaling.maximum = aw.drawing.charts.AxisBound(date(1990, 1, 1))

        doc.save(ARTIFACTS_DIR + "Charts.axis_bound.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Charts.axis_bound.docx")
        chart = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart

        self.assertFalse(chart.axis_x.scaling.minimum.is_auto)
        self.assertEqual(0.0, chart.axis_x.scaling.minimum.value)
        self.assertEqual(10.0, chart.axis_x.scaling.maximum.value)

        self.assertFalse(chart.axis_y.scaling.minimum.is_auto)
        self.assertEqual(0.0, chart.axis_y.scaling.minimum.value)
        self.assertEqual(10.0, chart.axis_y.scaling.maximum.value)

        chart = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape().chart

        self.assertFalse(chart.axis_x.scaling.minimum.is_auto)
        self.assertEqual(aw.drawing.charts.AxisBound(date(1980, 1, 1)), chart.axis_x.scaling.minimum)
        self.assertEqual(aw.drawing.charts.AxisBound(date(1990, 1, 1)), chart.axis_x.scaling.maximum)

        self.assertTrue(chart.axis_y.scaling.minimum.is_auto)

    def test_chart_legend(self):

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

        doc.save(ARTIFACTS_DIR + "Charts.chart_legend.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Charts.chart_legend.docx")

        legend = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart.legend

        self.assertTrue(legend.overlay)
        self.assertEqual(aw.drawing.charts.LegendPosition.TOP_RIGHT, legend.position)

    def test_axis_cross(self):

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

        doc.save(ARTIFACTS_DIR + "Charts.axis_cross.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Charts.axis_cross.docx")
        axis = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart.axis_x

        self.assertTrue(axis.axis_between_categories)
        self.assertEqual(aw.drawing.charts.AxisCrosses.CUSTOM, axis.crosses)
        self.assertEqual(3.0, axis.crosses_at)

    def test_axis_display_unit(self):

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

        doc.save(ARTIFACTS_DIR + "Charts.axis_display_unit.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Charts.axis_display_unit.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.assertEqual(450.0, shape.width)
        self.assertEqual(250.0, shape.height)

        axis = shape.chart.axis_x

        self.assertEqual(aw.drawing.charts.AxisTickMark.INSIDE, axis.major_tick_mark)
        self.assertEqual(aw.drawing.charts.AxisTickMark.INSIDE, axis.minor_tick_mark)
        self.assertEqual(10.0, axis.major_unit)
        self.assertEqual(-10.0, axis.scaling.minimum.value)
        self.assertEqual(30.0, axis.scaling.maximum.value)
        self.assertEqual(1, axis.tick_label_spacing)
        self.assertEqual(aw.ParagraphAlignment.RIGHT, axis.tick_label_alignment)
        self.assertEqual(aw.drawing.charts.AxisBuiltInUnit.CUSTOM, axis.display_unit.unit)
        self.assertEqual(1000000.0, axis.display_unit.custom_unit)

        axis = shape.chart.axis_y

        self.assertEqual(aw.drawing.charts.AxisTickMark.CROSS, axis.major_tick_mark)
        self.assertEqual(aw.drawing.charts.AxisTickMark.OUTSIDE, axis.minor_tick_mark)
        self.assertEqual(10.0, axis.major_unit)
        self.assertEqual(1.0, axis.minor_unit)
        self.assertEqual(-10.0, axis.scaling.minimum.value)
        self.assertEqual(20.0, axis.scaling.maximum.value)

    def test_marker_formatting(self):

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
        series = chart.series.add("AW Series 1", x_values=[0.7, 1.8, 2.6, 3.9], y_values=[2.7, 3.2, 0.8, 1.7])

        # Set marker formatting.
        series.marker.size = 40
        series.marker.symbol = aw.drawing.charts.MarkerSymbol.SQUARE
        data_points = series.data_points
        data_points[0].marker.format.fill.preset_textured(aw.drawing.PresetTexture.DENIM)
        data_points[0].marker.format.stroke.fore_color = drawing.Color.yellow
        data_points[0].marker.format.stroke.back_color = drawing.Color.red
        data_points[1].marker.format.fill.preset_textured(aw.drawing.PresetTexture.WATER_DROPLETS)
        data_points[1].marker.format.stroke.fore_color = drawing.Color.yellow
        data_points[1].marker.format.stroke.visible = False
        data_points[2].marker.format.fill.preset_textured(aw.drawing.PresetTexture.GREEN_MARBLE)
        data_points[2].marker.format.stroke.fore_color = drawing.Color.yellow
        data_points[3].marker.format.fill.preset_textured(aw.drawing.PresetTexture.OAK)
        data_points[3].marker.format.stroke.fore_color = drawing.Color.yellow
        data_points[3].marker.format.stroke.transparency = 0.5

        doc.save(ARTIFACTS_DIR + "Charts.marker_formatting.docx")
        #ExEnd

    def test_series_color(self):

        #ExStart
        #ExFor:ChartSeries.format
        #ExSummary:Sows how to set series color.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(aw.drawing.charts.ChartType.COLUMN, 432, 252)

        chart = shape.chart
        series_coll = chart.series

        # Delete default generated series.
        series_coll.clear()

        # Create category names array.
        categories = ["Category 1", "Category 2"]

        # Adding new series. Value and category arrays must be the same size.
        series1 = series_coll.add("Series 1", categories, [1, 2])
        series2 = series_coll.add("Series 2", categories, [3, 4])
        series3 = series_coll.add("Series 3", categories, [5, 6])

        # Set series color.
        series1.format.fill.fore_color = drawing.Color.red
        series2.format.fill.fore_color = drawing.Color.yellow
        series3.format.fill.fore_color = drawing.Color.blue

        doc.save(ARTIFACTS_DIR + "Charts.series_color.docx")
        #ExEnd

    def test_data_points_formatting(self):

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
        data_points = series.data_points
        data_points[0].format.fill.preset_textured(aw.drawing.PresetTexture.DENIM)
        data_points[1].format.fill.fore_color = drawing.Color.red
        data_points[2].format.fill.fore_color = drawing.Color.yellow
        data_points[3].format.fill.fore_color = drawing.Color.blue

        doc.save(ARTIFACTS_DIR + "Charts.data_points_formatting.docx")
        #ExEnd

    def test_legend_entries(self):
        #ExStart
        #ExFor:ChartLegendEntryCollection
        #ExFor:ChartLegend.legend_entries
        #ExFor:ChartLegendEntry.is_hidden
        #ExFor:ChartLegendEntry.font
        #ExSummary:Shows how to work with a legend entry for chart series.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(awdc.ChartType.COLUMN, 432, 252)

        chart = shape.chart
        series = chart.series
        series.clear()

        categories = ["AW Category 1", "AW Category 2"]
        series1 = series.add("Series 1", categories, [1, 2])
        series.add("Series 2", categories, [3, 4])
        series.add("Series 3", categories, [5, 6])
        series.add("Series 4", categories, [0, 0])

        legend_entries = chart.legend.legend_entries
        legend_entries[3].is_hidden = True

        for legend_entry in legend_entries:
            legend_entry.font.size = 12

        series1.legend_entry.font.italic = True

        doc.save(ARTIFACTS_DIR + "Charts.LegendEntries.docx")
        #ExEnd

    def test_axis_collection(self):
        #ExStart
        #ExFor:ChartAxisCollection
        #ExFor:Chart.axes
        #ExSummary:Shows how to work with axes collection.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_chart(awdc.ChartType.COLUMN, 500, 300)
        chart = shape.chart

        # Hide the major grid lines on the primary and secondary Y axes.
        for axis in chart.axes:
            if axis.type == awdc.ChartAxisType.VALUE:
                axis.has_major_gridlines = False

        doc.save(ARTIFACTS_DIR + "Charts.AxisCollection.docx")
        #ExEnd

    def test_format_data_lables(self):
        #ExStart
        #ExFor:ChartDataLabelCollection.format
        #ExFor:ChartFormat.shape_type
        #ExFor:ChartShapeType
        #ExSummary:Shows how to set fill, stroke and callout formatting for chart data labels.

        doc = Document()
        builder = DocumentBuilder(doc)

        shape = builder.insert_chart(ChartType.COLUMN, 432, 252)
        chart = shape.chart

        # Delete default generated series.
        chart.series.clear()

        # Add new series.
        series = chart.series.add("AW Series 1", ["AW Category 1", "AW Category 2", "AW Category 3", "AW Category 4"],
                                  [100, 200, 300, 400])

        # Show data labels.
        series.has_data_labels = True
        series.data_labels.show_value = True

        # Format data labels as callouts.
        format = series.data_labels.format
        format.shape_type = ChartShapeType.WEDGE_RECT_CALLOUT
        format.stroke.color = Color.dark_green
        format.fill.solid(Color.green)
        series.data_labels.font.color = Color.yellow

        # Change fill and stroke of an individual data label.
        labelFormat = series.data_labels[0].format
        labelFormat.stroke.color = Color.dark_blue
        labelFormat.fill.solid(Color.blue)

        doc.save(ARTIFACTS_DIR + "Charts.FormatDataLabels.docx")
        #ExEnd

    def test_chart_axis_title(self):
        #ExStart
        #ExFor:ChartAxisTitle
        #ExFor:ChartAxisTitle.text
        #ExFor:ChartAxisTitle.show
        #ExFor:ChartAxisTitle.overlay
        #ExFor: ChartAxisTitle.font
        #ExSummary: Shows how to set chart axis title.

        doc = Document()

        builder = DocumentBuilder(doc)
        shape = builder.insert_chart(ChartType.COLUMN, 432, 252)

        chart = shape.chart

        series_coll = chart.series
        # Delete default generated series.
        series_coll.clear()

        series_coll.add("AW Series 1", ["AW Category 1", "AW Category 2"], [1, 2])

        # Set axis title.
        chart.axis_x.title.text = "Categories"
        chart.axis_x.title.show = True
        chart.axis_y.title.text = "Values"
        chart.axis_y.title.show = True
        chart.axis_y.title.overlay = True
        chart.axis_y.title.font.size = 12
        chart.axis_y.title.font.color = drawing.Color.blue

        doc.save(ARTIFACTS_DIR + "Charts.ChartAxisTitle.docx")
        #ExEnd


    def test_copy_data_point_format(self):
        #ExStart
        #ExFor:ChartSeries.copy_format_from(int)
        #ExFor:ChartDataPointCollection.has_default_format(int)
        #ExFor:ChartDataPointCollection.copy_format(int, int)
        #ExSummary:Shows how to copy data point format.

        doc = aw.Document(MY_DIR + "DataPoint format.docx")

        # Get the chart and series to update format.
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        series = shape.chart.series[0]

        data_points = series.data_points

        self.assertTrue(data_points.has_default_format(0))
        self.assertFalse(data_points.has_default_format(1))

        # Copy format of the data point with index 1 to the data point with index 2
        # so that the data point 2 looks the same as the data point 1.
        data_points.copy_format(0, 1)

        self.assertTrue(data_points.has_default_format(0))
        self.assertTrue(data_points.has_default_format(1))

        # Copy format of the data point with index 0 to the series defaults so that all data points
        # in the series that have the default format look the same as the data point 0.
        series.copy_format_from(1)

        self.assertTrue(data_points.has_default_format(0))
        self.assertTrue(data_points.has_default_format(1))

        doc.save(ARTIFACTS_DIR + "Charts.CopyDataPointFormat.docx")
        #ExEnd

    def test_reset_data_point_fill(self):
        #ExStart
        #ExFor:ChartFormat.is_defined
        #ExFor:ChartFormat.set_default_fill
        #ExSummary: Shows how to reset the fill to the default value defined in the series.
        doc = Document(MY_DIR + "DataPoint format.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        series = shape.chart.series[0]
        data_point = series.data_points[1]

        self.assertTrue(data_point.format.is_defined)

        data_point.format.set_default_fill()

        doc.save(ARTIFACTS_DIR + "Charts.ResetDataPointFill.docx")
        #ExEnd
