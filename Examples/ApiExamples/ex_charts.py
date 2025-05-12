# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import locale
from aspose.words import Document, DocumentBuilder, NodeType
from aspose.pydrawing import Color
from aspose.words.drawing.charts import ChartXValue, ChartYValue, ChartSeriesType, ChartType, ChartShapeType
from datetime import date
from math import nan
import aspose.pydrawing
import aspose.words as aw
import aspose.words.drawing
import aspose.words.drawing.charts
import datetime
import math
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExCharts(ApiExampleBase):

    def test_chart_title(self):
        #ExStart:ChartTitle
        #ExFor:Chart
        #ExFor:Chart.title
        #ExFor:ChartTitle
        #ExFor:ChartTitle.overlay
        #ExFor:ChartTitle.show
        #ExFor:ChartTitle.text
        #ExFor:ChartTitle.font
        #ExSummary:Shows how to insert a chart and set a title.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a chart shape with a document builder and get its chart.
        chart_shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.BAR, width=400, height=300)
        chart = chart_shape.chart
        # Use the "Title" property to give our chart a title, which appears at the top center of the chart area.
        title = chart.title
        title.text = 'My Chart'
        title.font.size = 15
        title.font.color = aspose.pydrawing.Color.blue
        # Set the "Show" property to "true" to make the title visible.
        title.show = True
        # Set the "Overlay" property to "true" Give other chart elements more room by allowing them to overlap the title
        title.overlay = True
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.ChartTitle.docx')
        #ExEnd:ChartTitle
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.ChartTitle.docx')
        chart_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual(aw.drawing.ShapeType.NON_PRIMITIVE, chart_shape.shape_type)
        self.assertTrue(chart_shape.has_chart)
        title = chart_shape.chart.title
        self.assertEqual('My Chart', title.text)
        self.assertTrue(title.overlay)
        self.assertTrue(title.show)

    def test_data_label_number_format(self):
        #ExStart
        #ExFor:ChartDataLabelCollection.number_format
        #ExFor:ChartDataLabelCollection.font
        #ExFor:ChartNumberFormat.format_code
        #ExFor:ChartSeries.has_data_labels
        #ExSummary:Shows how to enable and configure data labels for a chart series.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Add a line chart, then clear its demo data series to start with a clean chart,
        # and then set a title.
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.LINE, width=500, height=300)
        chart = shape.chart
        chart.series.clear()
        chart.title.text = 'Monthly sales report'
        # Insert a custom chart series with months as categories for the X-axis,
        # and respective decimal amounts for the Y-axis.
        series = chart.series.add(series_name='Revenue', categories=['January', 'February', 'March'], values=[25.611, 21.439, 33.75])
        # Enable data labels, and then apply a custom number format for values displayed in the data labels.
        # This format will treat displayed decimal values as millions of US Dollars.
        series.has_data_labels = True
        data_labels = series.data_labels
        data_labels.show_value = True
        data_labels.number_format.format_code = '"US$" #,##0.000"M"'
        data_labels.font.size = 12
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.DataLabelNumberFormat.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.DataLabelNumberFormat.docx')
        series = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart.series[0]
        self.assertTrue(series.has_data_labels)
        self.assertTrue(series.data_labels.show_value)
        self.assertEqual('"US$" #,##0.000"M"', series.data_labels.number_format.format_code)

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
        #ExFor:ChartAxis.document
        #ExFor:ChartAxis.tick_labels
        #ExFor:ChartAxis.format
        #ExFor:AxisTickLabels
        #ExFor:AxisTickLabels.offset
        #ExFor:AxisTickLabels.position
        #ExFor:AxisTickLabels.is_auto_spacing
        #ExFor:AxisTickLabels.alignment
        #ExFor:AxisTickLabels.font
        #ExFor:AxisTickLabels.spacing
        #ExFor:ChartAxis.tick_mark_spacing
        #ExFor:AxisCategoryType
        #ExFor:AxisCrosses
        #ExFor:Chart.axis_x
        #ExFor:Chart.axis_y
        #ExFor:Chart.axis_z
        #ExSummary:Shows how to insert a chart and modify the appearance of its axes.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=500, height=300)
        chart = shape.chart
        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()
        # Insert a chart series with categories for the X-axis and respective numeric values for the Y-axis.
        chart.series.add(series_name='Aspose Test Series', categories=['Word', 'PDF', 'Excel', 'GoogleDocs', 'Note'], values=[640, 320, 280, 120, 150])
        # Chart axes have various options that can change their appearance,
        # such as their direction, major/minor unit ticks, and tick marks.
        x_axis = chart.axis_x
        x_axis.category_type = aw.drawing.charts.AxisCategoryType.CATEGORY
        x_axis.crosses = aw.drawing.charts.AxisCrosses.MINIMUM
        x_axis.reverse_order = False
        x_axis.major_tick_mark = aw.drawing.charts.AxisTickMark.INSIDE
        x_axis.minor_tick_mark = aw.drawing.charts.AxisTickMark.CROSS
        x_axis.major_unit = 10
        x_axis.minor_unit = 15
        x_axis.tick_labels.offset = 50
        x_axis.tick_labels.position = aw.drawing.charts.AxisTickLabelPosition.LOW
        x_axis.tick_labels.is_auto_spacing = False
        x_axis.tick_mark_spacing = 1
        self.assertEqual(doc, x_axis.document)
        y_axis = chart.axis_y
        y_axis.category_type = aw.drawing.charts.AxisCategoryType.AUTOMATIC
        y_axis.crosses = aw.drawing.charts.AxisCrosses.MAXIMUM
        y_axis.reverse_order = True
        y_axis.major_tick_mark = aw.drawing.charts.AxisTickMark.INSIDE
        y_axis.minor_tick_mark = aw.drawing.charts.AxisTickMark.CROSS
        y_axis.major_unit = 100
        y_axis.minor_unit = 20
        y_axis.tick_labels.position = aw.drawing.charts.AxisTickLabelPosition.NEXT_TO_AXIS
        y_axis.tick_labels.alignment = aw.ParagraphAlignment.CENTER
        y_axis.tick_labels.font.color = aspose.pydrawing.Color.red
        y_axis.tick_labels.spacing = 1
        # Column charts do not have a Z-axis.
        self.assertIsNone(chart.axis_z)
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.AxisProperties.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.AxisProperties.docx')
        chart = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart
        self.assertEqual(aw.drawing.charts.AxisCategoryType.CATEGORY, chart.axis_x.category_type)
        self.assertEqual(aw.drawing.charts.AxisCrosses.MINIMUM, chart.axis_x.crosses)
        self.assertFalse(chart.axis_x.reverse_order)
        self.assertEqual(aw.drawing.charts.AxisTickMark.INSIDE, chart.axis_x.major_tick_mark)
        self.assertEqual(aw.drawing.charts.AxisTickMark.CROSS, chart.axis_x.minor_tick_mark)
        self.assertEqual(1, chart.axis_x.major_unit)
        self.assertEqual(0.5, chart.axis_x.minor_unit)
        self.assertEqual(50, chart.axis_x.tick_labels.offset)
        self.assertEqual(aw.drawing.charts.AxisTickLabelPosition.LOW, chart.axis_x.tick_labels.position)
        self.assertFalse(chart.axis_x.tick_labels.is_auto_spacing)
        self.assertEqual(1, chart.axis_x.tick_mark_spacing)
        self.assertTrue(chart.axis_x.format.is_defined)
        self.assertEqual(aw.drawing.charts.AxisCategoryType.CATEGORY, chart.axis_y.category_type)
        self.assertEqual(aw.drawing.charts.AxisCrosses.MAXIMUM, chart.axis_y.crosses)
        self.assertTrue(chart.axis_y.reverse_order)
        self.assertEqual(aw.drawing.charts.AxisTickMark.INSIDE, chart.axis_y.major_tick_mark)
        self.assertEqual(aw.drawing.charts.AxisTickMark.CROSS, chart.axis_y.minor_tick_mark)
        self.assertEqual(100, chart.axis_y.major_unit)
        self.assertEqual(20, chart.axis_y.minor_unit)
        self.assertEqual(aw.drawing.charts.AxisTickLabelPosition.NEXT_TO_AXIS, chart.axis_y.tick_labels.position)
        self.assertEqual(aw.ParagraphAlignment.CENTER, chart.axis_y.tick_labels.alignment)
        self.assertEqual(aspose.pydrawing.Color.red.to_argb(), chart.axis_y.tick_labels.font.color.to_argb())
        self.assertEqual(1, chart.axis_y.tick_labels.spacing)
        self.assertTrue(chart.axis_y.format.is_defined)

    def test_axis_collection(self):
        #ExStart
        #ExFor:ChartAxisCollection
        #ExFor:Chart.axes
        #ExSummary:Shows how to work with axes collection.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=500, height=300)
        chart = shape.chart
        # Hide the major grid lines on the primary and secondary Y axes.
        for axis in chart.axes:
            if axis.type == aw.drawing.charts.ChartAxisType.VALUE:
                axis.has_major_gridlines = False
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.AxisCollection.docx')
        #ExEnd

    def test_hide_chart_axis(self):
        #ExStart
        #ExFor:ChartAxis.hidden
        #ExSummary:Shows how to hide chart axes.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.LINE, width=500, height=300)
        chart = shape.chart
        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()
        # Add a custom series with categories for the X-axis, and respective decimal values for the Y-axis.
        chart.series.add(series_name='AW Series 1', categories=['Item 1', 'Item 2', 'Item 3', 'Item 4', 'Item 5'], values=[1.2, 0.3, 2.1, 2.9, 4.2])
        # Hide the chart axes to simplify the appearance of the chart.
        chart.axis_x.hidden = True
        chart.axis_y.hidden = True
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.HideChartAxis.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.HideChartAxis.docx')
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
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=500, height=300)
        chart = shape.chart
        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()
        # Add a custom series to the chart with categories for the X-axis,
        # and large respective numeric values for the Y-axis.
        chart.series.add(series_name='Aspose Test Series', categories=['Word', 'PDF', 'Excel', 'GoogleDocs', 'Note'], values=[1900000, 850000, 2100000, 600000, 1500000])
        # Set the number format of the Y-axis tick labels to not group digits with commas.
        chart.axis_y.number_format.format_code = '#,##0'
        # This flag can override the above value and draw the number format from the source cell.
        self.assertFalse(chart.axis_y.number_format.is_linked_to_source)
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.SetNumberFormatToChartAxis.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.SetNumberFormatToChartAxis.docx')
        chart = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart
        self.assertEqual('#,##0', chart.axis_y.number_format.format_code)

    def test_display_charts_with_conversion(self):
        for chart_type in [aw.drawing.charts.ChartType.COLUMN, aw.drawing.charts.ChartType.LINE, aw.drawing.charts.ChartType.PIE, aw.drawing.charts.ChartType.BAR, aw.drawing.charts.ChartType.AREA]:
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            shape = builder.insert_chart(chart_type=chart_type, width=500, height=300)
            chart = shape.chart
            chart.series.clear()
            chart.series.add(series_name='Aspose Test Series', categories=['Word', 'PDF', 'Excel', 'GoogleDocs', 'Note'], values=[1900000, 850000, 2100000, 600000, 1500000])
            doc.save(file_name=ARTIFACTS_DIR + 'Charts.TestDisplayChartsWithConversion.docx')
            doc.save(file_name=ARTIFACTS_DIR + 'Charts.TestDisplayChartsWithConversion.pdf')

    def test_surface_3d_chart(self):
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.SURFACE_3D, width=500, height=300)
        chart = shape.chart
        chart.series.clear()
        chart.series.add(series_name='Aspose Test Series 1', categories=['Word', 'PDF', 'Excel', 'GoogleDocs', 'Note'], values=[1900000, 850000, 2100000, 600000, 1500000])
        chart.series.add(series_name='Aspose Test Series 2', categories=['Word', 'PDF', 'Excel', 'GoogleDocs', 'Note'], values=[900000, 50000, 1100000, 400000, 2500000])
        chart.series.add(series_name='Aspose Test Series 3', categories=['Word', 'PDF', 'Excel', 'GoogleDocs', 'Note'], values=[500000, 820000, 1500000, 400000, 100000])
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.SurfaceChart.docx')
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.SurfaceChart.pdf')

    def test_data_labels_bubble_chart(self):
        #ExStart
        #ExFor:ChartDataLabelCollection.separator
        #ExFor:ChartDataLabelCollection.show_bubble_size
        #ExFor:ChartDataLabelCollection.show_category_name
        #ExFor:ChartDataLabelCollection.show_series_name
        #ExSummary:Shows how to work with data labels of a bubble chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        chart = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.BUBBLE, width=500, height=300).chart
        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()
        # Add a custom series with X/Y coordinates and diameter of each of the bubbles.
        series = chart.series.add(series_name='Aspose Test Series', x_values=[2.9, 3.5, 1.1, 4, 4], y_values=[1.9, 8.5, 2.1, 6, 1.5], bubble_sizes=[9, 4.5, 2.5, 8, 5])
        # Enable data labels, and then modify their appearance.
        series.has_data_labels = True
        data_labels = series.data_labels
        data_labels.show_bubble_size = True
        data_labels.show_category_name = True
        data_labels.show_series_name = True
        data_labels.separator = ' & '
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.DataLabelsBubbleChart.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.DataLabelsBubbleChart.docx')
        data_labels = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart.series[0].data_labels
        self.assertTrue(data_labels.show_bubble_size)
        self.assertTrue(data_labels.show_category_name)
        self.assertTrue(data_labels.show_series_name)
        self.assertEqual(' & ', data_labels.separator)

    def test_data_labels_pie_chart(self):
        #ExStart
        #ExFor:ChartDataLabelCollection.separator
        #ExFor:ChartDataLabelCollection.show_leader_lines
        #ExFor:ChartDataLabelCollection.show_legend_key
        #ExFor:ChartDataLabelCollection.show_percentage
        #ExFor:ChartDataLabelCollection.show_value
        #ExSummary:Shows how to work with data labels of a pie chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        chart = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.PIE, width=500, height=300).chart
        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()
        # Insert a custom chart series with a category name for each of the sectors, and their frequency table.
        series = chart.series.add(series_name='Aspose Test Series', categories=['Word', 'PDF', 'Excel'], values=[2.7, 3.2, 0.8])
        # Enable data labels that will display both percentage and frequency of each sector, and modify their appearance.
        series.has_data_labels = True
        data_labels = series.data_labels
        data_labels.show_leader_lines = True
        data_labels.show_legend_key = True
        data_labels.show_percentage = True
        data_labels.show_value = True
        data_labels.separator = '; '
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.DataLabelsPieChart.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.DataLabelsPieChart.docx')
        data_labels = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart.series[0].data_labels
        self.assertTrue(data_labels.show_leader_lines)
        self.assertTrue(data_labels.show_legend_key)
        self.assertTrue(data_labels.show_percentage)
        self.assertTrue(data_labels.show_value)
        self.assertEqual('; ', data_labels.separator)

    def test_pie_chart_explosion(self):
        #ExStart
        #ExFor:IChartDataPoint.explosion
        #ExFor:ChartDataPoint.explosion
        #ExSummary:Shows how to move the slices of a pie chart away from the center.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.PIE, width=500, height=350)
        chart = shape.chart
        self.assertEqual(1, chart.series.count)
        self.assertEqual('Sales', chart.series[0].name)
        # "Slices" of a pie chart may be moved away from the center by a distance via the respective data point's Explosion attribute.
        # Add a data point to the first portion of the pie chart and move it away from the center by 10 points.
        # Aspose.Words create data points automatically if them does not exist.
        data_point = chart.series[0].data_points[0]
        data_point.explosion = 10
        # Displace the second portion by a greater distance.
        data_point = chart.series[0].data_points[1]
        data_point.explosion = 40
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.PieChartExplosion.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.PieChartExplosion.docx')
        series = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart.series[0]
        self.assertEqual(10, series.data_points[0].explosion)
        self.assertEqual(40, series.data_points[1].explosion)

    def test_bubble_3d(self):
        #ExStart
        #ExFor:ChartDataLabel.show_bubble_size
        #ExFor:ChartDataLabel.font
        #ExFor:IChartDataPoint.bubble_3d
        #ExFor:ChartSeries.bubble_3d
        #ExSummary:Shows how to use 3D effects with bubble charts.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.BUBBLE_3D, width=500, height=350)
        chart = shape.chart
        self.assertEqual(1, chart.series.count)
        self.assertEqual('Y-Values', chart.series[0].name)
        self.assertTrue(chart.series[0].bubble_3d)
        # Apply a data label to each bubble that displays its diameter.
        i = 0
        while i < 3:
            chart.series[0].has_data_labels = True
            chart.series[0].data_labels[i].show_bubble_size = True
            chart.series[0].data_labels[i].font.size = 12
            i += 1
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.Bubble3D.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.Bubble3D.docx')
        series = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart.series[0]
        i = 0
        while i < 3:
            self.assertTrue(series.data_labels[i].show_bubble_size)
            i += 1

    def test_axis_scaling(self):
        #ExStart
        #ExFor:AxisScaleType
        #ExFor:AxisScaling
        #ExFor:AxisScaling.log_base
        #ExFor:AxisScaling.type
        #ExSummary:Shows how to apply logarithmic scaling to a chart axis.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        chart_shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.SCATTER, width=450, height=300)
        chart = chart_shape.chart
        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()
        # Insert a series with X/Y coordinates for five points.
        chart.series.add(series_name='Series 1', x_values=[1, 2, 3, 4, 5], y_values=[1, 20, 400, 8000, 160000])
        # The scaling of the X-axis is linear by default,
        # displaying evenly incrementing values that cover our X-value range (0, 1, 2, 3...).
        # A linear axis is not ideal for our Y-values
        # since the points with the smaller Y-values will be harder to read.
        # A logarithmic scaling with a base of 20 (1, 20, 400, 8000...)
        # will spread the plotted points, allowing us to read their values on the chart more easily.
        chart.axis_y.scaling.type = aw.drawing.charts.AxisScaleType.LOGARITHMIC
        chart.axis_y.scaling.log_base = 20
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.AxisScaling.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.AxisScaling.docx')
        chart = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart
        self.assertEqual(aw.drawing.charts.AxisScaleType.LINEAR, chart.axis_x.scaling.type)
        self.assertEqual(aw.drawing.charts.AxisScaleType.LOGARITHMIC, chart.axis_y.scaling.type)
        self.assertEqual(20, chart.axis_y.scaling.log_base)

    def test_axis_bound(self):
        #ExStart
        #ExFor:AxisBound.__init__
        #ExFor:AxisBound.is_auto
        #ExFor:AxisBound.value
        #ExFor:AxisBound.value_as_date
        #ExSummary:Shows how to set custom axis bounds.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        chart_shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.SCATTER, width=450, height=300)
        chart = chart_shape.chart
        # Clear the chart's demo data series to start with a clean chart.
        chart.series.clear()
        # Add a series with two decimal arrays. The first array contains the X-values,
        # and the second contains corresponding Y-values for points in the scatter chart.
        chart.series.add(series_name='Series 1', x_values=[1.1, 5.4, 7.9, 3.5, 2.1, 9.7], y_values=[2.1, 0.3, 0.6, 3.3, 1.4, 1.9])
        # By default, default scaling is applied to the graph's X and Y-axes,
        # so that both their ranges are big enough to encompass every X and Y-value of every series.
        self.assertTrue(chart.axis_x.scaling.minimum.is_auto)
        # We can define our own axis bounds.
        # In this case, we will make both the X and Y-axis rulers show a range of 0 to 10.
        chart.axis_x.scaling.minimum = aw.drawing.charts.AxisBound(value=0)
        chart.axis_x.scaling.maximum = aw.drawing.charts.AxisBound(value=10)
        chart.axis_y.scaling.minimum = aw.drawing.charts.AxisBound(value=0)
        chart.axis_y.scaling.maximum = aw.drawing.charts.AxisBound(value=10)
        self.assertFalse(chart.axis_x.scaling.minimum.is_auto)
        self.assertFalse(chart.axis_y.scaling.minimum.is_auto)
        # Create a line chart with a series requiring a range of dates on the X-axis, and decimal values for the Y-axis.
        chart_shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.LINE, width=450, height=300)
        chart = chart_shape.chart
        chart.series.clear()
        dates = [datetime.datetime(1973, 5, 11), datetime.datetime(1981, 2, 4), datetime.datetime(1985, 9, 23), datetime.datetime(1989, 6, 28), datetime.datetime(1994, 12, 15)]
        chart.series.add(series_name='Series 1', dates=dates, values=[3, 4.7, 5.9, 7.1, 8.9])
        # We can set axis bounds in the form of dates as well, limiting the chart to a period.
        # Setting the range to 1980-1990 will omit the two of the series values
        # that are outside of the range from the graph.
        chart.axis_x.scaling.minimum = aw.drawing.charts.AxisBound(datetime=datetime.datetime(1980, 1, 1))
        chart.axis_x.scaling.maximum = aw.drawing.charts.AxisBound(datetime=datetime.datetime(1990, 1, 1))
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.AxisBound.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.AxisBound.docx')
        chart = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart
        self.assertFalse(chart.axis_x.scaling.minimum.is_auto)
        self.assertEqual(0, chart.axis_x.scaling.minimum.value)
        self.assertEqual(10, chart.axis_x.scaling.maximum.value)
        self.assertFalse(chart.axis_y.scaling.minimum.is_auto)
        self.assertEqual(0, chart.axis_y.scaling.minimum.value)
        self.assertEqual(10, chart.axis_y.scaling.maximum.value)
        chart = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape().chart
        self.assertFalse(chart.axis_x.scaling.minimum.is_auto)
        self.assertEqual(aw.drawing.charts.AxisBound(datetime=datetime.datetime(1980, 1, 1)), chart.axis_x.scaling.minimum)
        self.assertEqual(aw.drawing.charts.AxisBound(datetime=datetime.datetime(1990, 1, 1)), chart.axis_x.scaling.maximum)
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
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.LINE, width=450, height=300)
        chart = shape.chart
        self.assertEqual(3, chart.series.count)
        self.assertEqual('Series 1', chart.series[0].name)
        self.assertEqual('Series 2', chart.series[1].name)
        self.assertEqual('Series 3', chart.series[2].name)
        # Move the chart's legend to the top right corner.
        legend = chart.legend
        legend.position = aw.drawing.charts.LegendPosition.TOP_RIGHT
        # Give other chart elements, such as the graph, more room by allowing them to overlap the legend.
        legend.overlay = True
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.ChartLegend.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.ChartLegend.docx')
        legend = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart.legend
        self.assertTrue(legend.overlay)
        self.assertEqual(aw.drawing.charts.LegendPosition.TOP_RIGHT, legend.position)

    def test_axis_cross(self):
        #ExStart
        #ExFor:ChartAxis.axis_between_categories
        #ExFor:ChartAxis.crosses_at
        #ExSummary:Shows how to get a graph axis to cross at a custom location.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=450, height=250)
        chart = shape.chart
        self.assertEqual(3, chart.series.count)
        self.assertEqual('Series 1', chart.series[0].name)
        self.assertEqual('Series 2', chart.series[1].name)
        self.assertEqual('Series 3', chart.series[2].name)
        # For column charts, the Y-axis crosses at zero by default,
        # which means that columns for all values below zero point down to represent negative values.
        # We can set a different value for the Y-axis crossing. In this case, we will set it to 3.
        axis = chart.axis_x
        axis.crosses = aw.drawing.charts.AxisCrosses.CUSTOM
        axis.crosses_at = 3
        axis.axis_between_categories = True
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.AxisCross.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.AxisCross.docx')
        axis = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart.axis_x
        self.assertTrue(axis.axis_between_categories)
        self.assertEqual(aw.drawing.charts.AxisCrosses.CUSTOM, axis.crosses)
        self.assertEqual(3, axis.crosses_at)

    def test_axis_display_unit(self):
        #ExStart
        #ExFor:AxisBuiltInUnit
        #ExFor:ChartAxis.display_unit
        #ExFor:ChartAxis.major_unit_is_auto
        #ExFor:ChartAxis.major_unit_scale
        #ExFor:ChartAxis.minor_unit_is_auto
        #ExFor:ChartAxis.minor_unit_scale
        #ExFor:AxisDisplayUnit
        #ExFor:AxisDisplayUnit.custom_unit
        #ExFor:AxisDisplayUnit.unit
        #ExFor:AxisDisplayUnit.document
        #ExSummary:Shows how to manipulate the tick marks and displayed values of a chart axis.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.SCATTER, width=450, height=250)
        chart = shape.chart
        self.assertEqual(1, chart.series.count)
        self.assertEqual('Y-Values', chart.series[0].name)
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
        axis.scaling.minimum = aw.drawing.charts.AxisBound(value=-10)
        axis.scaling.maximum = aw.drawing.charts.AxisBound(value=20)
        # For the X-axis, set the major tick marks at every 10 units,
        # every minor tick mark at 2.5 units.
        axis = chart.axis_x
        axis.major_unit = 10
        axis.minor_unit = 2.5
        # Configure both types of tick marks to appear inside the graph plot area.
        axis.major_tick_mark = aw.drawing.charts.AxisTickMark.INSIDE
        axis.minor_tick_mark = aw.drawing.charts.AxisTickMark.INSIDE
        # Set the X-axis bounds so that the X-axis spans 5 major tick marks and 12 minor tick marks.
        axis.scaling.minimum = aw.drawing.charts.AxisBound(value=-10)
        axis.scaling.maximum = aw.drawing.charts.AxisBound(value=30)
        axis.tick_labels.alignment = aw.ParagraphAlignment.RIGHT
        self.assertEqual(1, axis.tick_labels.spacing)
        self.assertEqual(doc, axis.display_unit.document)
        # Set the tick labels to display their value in millions.
        axis.display_unit.unit = aw.drawing.charts.AxisBuiltInUnit.MILLIONS
        # We can set a more specific value by which tick labels will display their values.
        # This statement is equivalent to the one above.
        axis.display_unit.custom_unit = 1000000
        self.assertEqual(aw.drawing.charts.AxisBuiltInUnit.CUSTOM, axis.display_unit.unit)  #ExSkip
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.AxisDisplayUnit.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.AxisDisplayUnit.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual(450, shape.width)
        self.assertEqual(250, shape.height)
        axis = shape.chart.axis_x
        self.assertEqual(aw.drawing.charts.AxisTickMark.INSIDE, axis.major_tick_mark)
        self.assertEqual(aw.drawing.charts.AxisTickMark.INSIDE, axis.minor_tick_mark)
        self.assertEqual(10, axis.major_unit)
        self.assertEqual(-10, axis.scaling.minimum.value)
        self.assertEqual(30, axis.scaling.maximum.value)
        self.assertEqual(1, axis.tick_labels.spacing)
        self.assertEqual(aw.ParagraphAlignment.RIGHT, axis.tick_labels.alignment)
        self.assertEqual(aw.drawing.charts.AxisBuiltInUnit.CUSTOM, axis.display_unit.unit)
        self.assertEqual(1000000, axis.display_unit.custom_unit)
        axis = shape.chart.axis_y
        self.assertEqual(aw.drawing.charts.AxisTickMark.CROSS, axis.major_tick_mark)
        self.assertEqual(aw.drawing.charts.AxisTickMark.OUTSIDE, axis.minor_tick_mark)
        self.assertEqual(10, axis.major_unit)
        self.assertEqual(1, axis.minor_unit)
        self.assertEqual(-10, axis.scaling.minimum.value)
        self.assertEqual(20, axis.scaling.maximum.value)

    def test_marker_formatting(self):
        #ExStart
        #ExFor:ChartDataPoint.marker
        #ExFor:ChartMarker.format
        #ExFor:ChartFormat.fill
        #ExFor:ChartSeries.marker
        #ExFor:ChartFormat.stroke
        #ExFor:Stroke.fore_color
        #ExFor:Stroke.back_color
        #ExFor:Stroke.visible
        #ExFor:Stroke.transparency
        #ExFor:PresetTexture
        #ExFor:Fill.preset_textured(PresetTexture)
        #ExSummary:Show how to set marker formatting.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.SCATTER, width=432, height=252)
        chart = shape.chart
        # Delete default generated series.
        chart.series.clear()
        series = chart.series.add(series_name='AW Series 1', x_values=[0.7, 1.8, 2.6, 3.9], y_values=[2.7, 3.2, 0.8, 1.7])
        # Set marker formatting.
        series.marker.size = 40
        series.marker.symbol = aw.drawing.charts.MarkerSymbol.SQUARE
        data_points = series.data_points
        data_points[0].marker.format.fill.preset_textured(aw.drawing.PresetTexture.DENIM)
        data_points[0].marker.format.stroke.fore_color = aspose.pydrawing.Color.yellow
        data_points[0].marker.format.stroke.back_color = aspose.pydrawing.Color.red
        data_points[1].marker.format.fill.preset_textured(aw.drawing.PresetTexture.WATER_DROPLETS)
        data_points[1].marker.format.stroke.fore_color = aspose.pydrawing.Color.yellow
        data_points[1].marker.format.stroke.visible = False
        data_points[2].marker.format.fill.preset_textured(aw.drawing.PresetTexture.GREEN_MARBLE)
        data_points[2].marker.format.stroke.fore_color = aspose.pydrawing.Color.yellow
        data_points[3].marker.format.fill.preset_textured(aw.drawing.PresetTexture.OAK)
        data_points[3].marker.format.stroke.fore_color = aspose.pydrawing.Color.yellow
        data_points[3].marker.format.stroke.transparency = 0.5
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.MarkerFormatting.docx')
        #ExEnd

    def test_series_color(self):
        #ExStart
        #ExFor:ChartSeries.format
        #ExSummary:Sows how to set series color.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=432, height=252)
        chart = shape.chart
        series_coll = chart.series
        # Delete default generated series.
        series_coll.clear()
        # Create category names array.
        categories = ['Category 1', 'Category 2']
        # Adding new series. Value and category arrays must be the same size.
        series1 = series_coll.add(series_name='Series 1', categories=categories, values=[1, 2])
        series2 = series_coll.add(series_name='Series 2', categories=categories, values=[3, 4])
        series3 = series_coll.add(series_name='Series 3', categories=categories, values=[5, 6])
        # Set series color.
        series1.format.fill.fore_color = aspose.pydrawing.Color.red
        series2.format.fill.fore_color = aspose.pydrawing.Color.yellow
        series3.format.fill.fore_color = aspose.pydrawing.Color.blue
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.SeriesColor.docx')
        #ExEnd

    def test_data_points_formatting(self):
        #ExStart
        #ExFor:ChartDataPoint.format
        #ExSummary:Shows how to set individual formatting for categories of a column chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=432, height=252)
        chart = shape.chart
        # Delete default generated series.
        chart.series.clear()
        # Adding new series.
        series = chart.series.add(series_name='Series 1', categories=['Category 1', 'Category 2', 'Category 3', 'Category 4'], values=[1, 2, 3, 4])
        # Set column formatting.
        data_points = series.data_points
        data_points[0].format.fill.preset_textured(aw.drawing.PresetTexture.DENIM)
        data_points[1].format.fill.fore_color = aspose.pydrawing.Color.red
        data_points[2].format.fill.fore_color = aspose.pydrawing.Color.yellow
        data_points[3].format.fill.fore_color = aspose.pydrawing.Color.blue
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.DataPointsFormatting.docx')
        #ExEnd

    def test_legend_entries(self):
        #ExStart
        #ExFor:ChartLegendEntryCollection
        #ExFor:ChartLegend.legend_entries
        #ExFor:ChartLegendEntry.is_hidden
        #ExSummary:Shows how to work with a legend entry for chart series.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=432, height=252)
        chart = shape.chart
        series = chart.series
        series.clear()
        categories = ['AW Category 1', 'AW Category 2']
        series1 = series.add(series_name='Series 1', categories=categories, values=[1, 2])
        series.add(series_name='Series 2', categories=categories, values=[3, 4])
        series.add(series_name='Series 3', categories=categories, values=[5, 6])
        series.add(series_name='Series 4', categories=categories, values=[0, 0])
        legend_entries = chart.legend.legend_entries
        legend_entries[3].is_hidden = True
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.LegendEntries.docx')
        #ExEnd

    def test_legend_font(self):
        #ExStart:LegendFont
        #ExFor:ChartLegendEntry
        #ExFor:ChartLegendEntry.font
        #ExFor:ChartLegend.font
        #ExFor:ChartSeries.legend_entry
        #ExSummary:Shows how to work with a legend font.
        doc = aw.Document(file_name=MY_DIR + 'Reporting engine template - Chart series.docx')
        chart = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart
        chart_legend = chart.legend
        # Set default font size all legend entries.
        chart_legend.font.size = 14
        # Change font for specific legend entry.
        chart_legend.legend_entries[1].font.italic = True
        chart_legend.legend_entries[1].font.size = 12
        # Get legend entry for chart series.
        legend_entry = chart.series[0].legend_entry
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.LegendFont.docx')
        #ExEnd:LegendFont

    def test_remove_specific_chart_series(self):
        #ExStart
        #ExFor:ChartSeries.series_type
        #ExFor:ChartSeriesType
        #ExSummary:Shows how to remove specific chart serie.
        doc = aw.Document(file_name=MY_DIR + 'Reporting engine template - Chart series.docx')
        chart = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().chart
        # Remove all series of the Column type.
        i = chart.series.count - 1
        while i >= 0:
            if chart.series[i].series_type == aw.drawing.charts.ChartSeriesType.COLUMN:
                chart.series.remove_at(i)
            i -= 1
        chart.series.add(series_name='Aspose Series', categories=['Category 1', 'Category 2', 'Category 3', 'Category 4'], values=[5.6, 7.1, 2.9, 8.9])
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.RemoveSpecificChartSeries.docx')
        #ExEnd

    def test_populate_chart_with_data(self):
        #ExStart
        #ExFor:ChartXValue
        #ExFor:ChartXValue.from_double(float)
        #ExFor:ChartYValue.from_double(float)
        #ExFor:ChartSeries.add(ChartXValue)
        #ExFor:ChartSeries.add(ChartXValue,ChartYValue)
        #ExFor:ChartSeries.add(ChartXValue,ChartYValue,float)
        #ExFor:ChartSeries.clear_values
        #ExFor:ChartSeries.clear
        #ExSummary:Shows how to populate chart series with data.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=432, height=252)
        chart = shape.chart
        series1 = chart.series[0]
        # Clear X and Y values of the first series.
        series1.clear_values()
        # Populate the series with data.
        series1.add(x_value=aw.drawing.charts.ChartXValue.from_double(3), y_value=aw.drawing.charts.ChartYValue.from_double(10), bubble_size=10)
        series1.add(x_value=aw.drawing.charts.ChartXValue.from_double(5), y_value=aw.drawing.charts.ChartYValue.from_double(5))
        series1.add(x_value=aw.drawing.charts.ChartXValue.from_double(7), y_value=aw.drawing.charts.ChartYValue.from_double(11))
        series1.add(x_value=aw.drawing.charts.ChartXValue.from_double(9))
        series2 = chart.series[1]
        # Clear X and Y values of the second series.
        series2.clear()
        # Populate the series with data.
        series2.add(x_value=aw.drawing.charts.ChartXValue.from_double(2), y_value=aw.drawing.charts.ChartYValue.from_double(4))
        series2.add(x_value=aw.drawing.charts.ChartXValue.from_double(4), y_value=aw.drawing.charts.ChartYValue.from_double(7))
        series2.add(x_value=aw.drawing.charts.ChartXValue.from_double(6), y_value=aw.drawing.charts.ChartYValue.from_double(14))
        series2.add(x_value=aw.drawing.charts.ChartXValue.from_double(8), y_value=aw.drawing.charts.ChartYValue.from_double(7))
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.PopulateChartWithData.docx')
        #ExEnd

    def test_get_chart_series_data(self):
        #ExStart
        #ExFor:ChartXValueCollection
        #ExFor:ChartYValueCollection
        #ExSummary:Shows how to get chart series data.
        doc = aw.Document()
        builder = aw.DocumentBuilder()
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=432, height=252)
        chart = shape.chart
        series = chart.series[0]
        min_value = 1.7976931348623157e+308
        min_value_index = 0
        max_value = -1.7976931348623157e+308
        max_value_index = 0
        i = 0
        while i < series.y_values.count:
            # Clear individual format of all data points.
            # Data points and data values are one-to-one in column charts.
            series.data_points[i].clear_format()
            # Get Y value.
            y_value = series.y_values[i].double_value
            if y_value < min_value:
                min_value = y_value
                min_value_index = i
            if y_value > max_value:
                max_value = y_value
                max_value_index = i
            i += 1
        # Change colors of the max and min values.
        series.data_points[min_value_index].format.fill.fore_color = aspose.pydrawing.Color.red
        series.data_points[max_value_index].format.fill.fore_color = aspose.pydrawing.Color.green
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.GetChartSeriesData.docx')
        #ExEnd

    def test_chart_data_values(self):
        #ExStart
        #ExFor:ChartXValue.from_string(str)
        #ExFor:ChartSeries.remove(int)
        #ExFor:ChartSeries.add(ChartXValue,ChartYValue)
        #ExSummary:Shows how to add/remove chart data values.
        doc = aw.Document()
        builder = aw.DocumentBuilder()
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=432, height=252)
        chart = shape.chart
        department_1_series = chart.series[0]
        department_2_series = chart.series[1]
        # Remove the first value in the both series.
        department_1_series.remove(0)
        department_2_series.remove(0)
        # Add new values to the both series.
        new_x_category = aw.drawing.charts.ChartXValue.from_string('Q1, 2023')
        department_1_series.add(x_value=new_x_category, y_value=aw.drawing.charts.ChartYValue.from_double(10.3))
        department_2_series.add(x_value=new_x_category, y_value=aw.drawing.charts.ChartYValue.from_double(5.7))
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.ChartDataValues.docx')
        #ExEnd

    def test_format_data_lables(self):
        #ExStart
        #ExFor:ChartDataLabelCollection.format
        #ExFor:ChartFormat.shape_type
        #ExFor:ChartShapeType
        #ExSummary:Shows how to set fill, stroke and callout formatting for chart data labels.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=432, height=252)
        chart = shape.chart
        # Delete default generated series.
        chart.series.clear()
        # Add new series.
        series = chart.series.add(series_name='AW Series 1', categories=['AW Category 1', 'AW Category 2', 'AW Category 3', 'AW Category 4'], values=[100, 200, 300, 400])
        # Show data labels.
        series.has_data_labels = True
        series.data_labels.show_value = True
        # Format data labels as callouts.
        format = series.data_labels.format
        format.shape_type = aw.drawing.charts.ChartShapeType.WEDGE_RECT_CALLOUT
        format.stroke.color = aspose.pydrawing.Color.dark_green
        format.fill.solid(aspose.pydrawing.Color.green)
        series.data_labels.font.color = aspose.pydrawing.Color.yellow
        # Change fill and stroke of an individual data label.
        label_format = series.data_labels[0].format
        label_format.stroke.color = aspose.pydrawing.Color.dark_blue
        label_format.fill.solid(aspose.pydrawing.Color.blue)
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.FormatDataLables.docx')
        #ExEnd

    def test_chart_axis_title(self):
        #ExStart:ChartAxisTitle
        #ExFor:ChartAxis.title
        #ExFor:ChartAxisTitle
        #ExFor:ChartAxisTitle.text
        #ExFor:ChartAxisTitle.show
        #ExFor:ChartAxisTitle.overlay
        #ExFor:ChartAxisTitle.font
        #ExSummary:Shows how to set chart axis title.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=432, height=252)
        chart = shape.chart
        series_coll = chart.series
        # Delete default generated series.
        series_coll.clear()
        series_coll.add(series_name='AW Series 1', categories=['AW Category 1', 'AW Category 2'], values=[1, 2])
        chart_axis_x_title = chart.axis_x.title
        chart_axis_x_title.text = 'Categories'
        chart_axis_x_title.show = True
        chart_axis_y_title = chart.axis_y.title
        chart_axis_y_title.text = 'Values'
        chart_axis_y_title.show = True
        chart_axis_y_title.overlay = True
        chart_axis_y_title.font.size = 12
        chart_axis_y_title.font.color = aspose.pydrawing.Color.blue
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.ChartAxisTitle.docx')
        #ExEnd:ChartAxisTitle

    def test_copy_data_point_format(self):
        #ExStart:CopyDataPointFormat
        #ExFor:ChartSeries.copy_format_from(int)
        #ExFor:ChartDataPointCollection.has_default_format(int)
        #ExFor:ChartDataPointCollection.copy_format(int,int)
        #ExSummary:Shows how to copy data point format.
        doc = aw.Document(file_name=MY_DIR + 'DataPoint format.docx')
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
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.CopyDataPointFormat.docx')
        #ExEnd:CopyDataPointFormat

    def test_reset_data_point_fill(self):
        #ExStart:ResetDataPointFill
        #ExFor:ChartFormat.is_defined
        #ExFor:ChartFormat.set_default_fill
        #ExSummary:Shows how to reset the fill to the default value defined in the series.
        doc = aw.Document(file_name=MY_DIR + 'DataPoint format.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        series = shape.chart.series[0]
        data_point = series.data_points[1]
        self.assertTrue(data_point.format.is_defined)
        data_point.format.set_default_fill()
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.ResetDataPointFill.docx')
        #ExEnd:ResetDataPointFill

    def test_data_table(self):
        #ExStart:DataTable
        #ExFor:Chart.data_table
        #ExFor:ChartDataTable
        #ExFor:ChartDataTable.show
        #ExFor:ChartDataTable.format
        #ExFor:ChartDataTable.font
        #ExFor:ChartDataTable.has_legend_keys
        #ExFor:ChartDataTable.has_horizontal_border
        #ExFor:ChartDataTable.has_vertical_border
        #ExFor:ChartDataTable.has_outline_border
        #ExSummary:Shows how to show data table with chart series data.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=432, height=252)
        chart = shape.chart
        series = chart.series
        series.clear()
        x_values = [2020, 2021, 2022, 2023]
        series.add(series_name='Series1', x_values=x_values, y_values=[5, 11, 2, 7])
        series.add(series_name='Series2', x_values=x_values, y_values=[6, 5.5, 7, 7.8])
        series.add(series_name='Series3', x_values=x_values, y_values=[10, 8, 7, 9])
        data_table = chart.data_table
        data_table.show = True
        data_table.has_legend_keys = False
        data_table.has_horizontal_border = False
        data_table.has_vertical_border = False
        data_table.has_outline_border = False
        data_table.font.italic = True
        data_table.format.stroke.weight = 1
        data_table.format.stroke.dash_style = aw.drawing.DashStyle.SHORT_DOT
        data_table.format.stroke.color = aspose.pydrawing.Color.dark_blue
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.DataTable.docx')
        #ExEnd:DataTable

    def test_chart_format(self):
        #ExStart:ChartFormat
        #ExFor:ChartFormat
        #ExFor:Chart.format
        #ExFor:ChartTitle.format
        #ExFor:ChartAxisTitle.format
        #ExFor:ChartLegend.format
        #ExFor:Fill.solid(Color)
        #ExSummary:Shows how to use chart formating.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=432, height=252)
        chart = shape.chart
        # Delete series generated by default.
        series = chart.series
        series.clear()
        categories = ['Category 1', 'Category 2']
        series.add(series_name='Series 1', categories=categories, values=[1, 2])
        series.add(series_name='Series 2', categories=categories, values=[3, 4])
        # Format chart background.
        chart.format.fill.solid(aspose.pydrawing.Color.dark_slate_gray)
        # Hide axis tick labels.
        chart.axis_x.tick_labels.position = aw.drawing.charts.AxisTickLabelPosition.NONE
        chart.axis_y.tick_labels.position = aw.drawing.charts.AxisTickLabelPosition.NONE
        # Format chart title.
        chart.title.format.fill.solid(aspose.pydrawing.Color.light_goldenrod_yellow)
        # Format axis title.
        chart.axis_x.title.show = True
        chart.axis_x.title.format.fill.solid(aspose.pydrawing.Color.light_goldenrod_yellow)
        # Format legend.
        chart.legend.format.fill.solid(aspose.pydrawing.Color.light_goldenrod_yellow)
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.ChartFormat.docx')
        #ExEnd:ChartFormat
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.ChartFormat.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        chart = shape.chart
        self.assertEqual(aspose.pydrawing.Color.dark_slate_gray.to_argb(), chart.format.fill.color.to_argb())
        self.assertEqual(aspose.pydrawing.Color.light_goldenrod_yellow.to_argb(), chart.title.format.fill.color.to_argb())
        self.assertEqual(aspose.pydrawing.Color.light_goldenrod_yellow.to_argb(), chart.axis_x.title.format.fill.color.to_argb())
        self.assertEqual(aspose.pydrawing.Color.light_goldenrod_yellow.to_argb(), chart.legend.format.fill.color.to_argb())

    def test_secondary_axis(self):
        #ExStart:SecondaryAxis
        #ExFor:ChartSeriesGroup
        #ExFor:ChartSeriesGroup.series_type
        #ExFor:ChartSeriesGroup.axis_group
        #ExFor:ChartSeriesGroup.axis_x
        #ExFor:ChartSeriesGroup.axis_y
        #ExFor:ChartSeriesGroup.series
        #ExFor:ChartSeriesGroupCollection
        #ExFor:ChartSeriesGroupCollection.add(ChartSeriesType)
        #ExFor:AxisGroup
        #ExSummary:Shows how to work with the secondary axis of chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.LINE, width=450, height=250)
        chart = shape.chart
        series = chart.series
        # Delete default generated series.
        series.clear()
        categories = ['Category 1', 'Category 2', 'Category 3']
        series.add(series_name='Series 1 of primary series group', categories=categories, values=[2, 3, 4])
        series.add(series_name='Series 2 of primary series group', categories=categories, values=[5, 2, 3])
        # Create an additional series group, also of the line type.
        new_series_group = chart.series_groups.add(aw.drawing.charts.ChartSeriesType.LINE)
        # Specify the use of secondary axes for the new series group.
        new_series_group.axis_group = aw.drawing.charts.AxisGroup.SECONDARY
        # Hide the secondary X axis.
        new_series_group.axis_x.hidden = True
        # Define title of the secondary Y axis.
        new_series_group.axis_y.title.show = True
        new_series_group.axis_y.title.text = 'Secondary Y axis'
        self.assertEqual(aw.drawing.charts.ChartSeriesType.LINE, new_series_group.series_type)
        # Add a series to the new series group.
        series3 = new_series_group.series.add(series_name='Series of secondary series group', categories=categories, values=[13, 11, 16])
        series3.format.stroke.weight = 3.5
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.SecondaryAxis.docx')
        #ExEnd:SecondaryAxis

    def test_configure_gap_overlap(self):
        #ExStart:ConfigureGapOverlap
        #ExFor:Chart.series_groups
        #ExFor:ChartSeriesGroup.gap_width
        #ExFor:ChartSeriesGroup.overlap
        #ExSummary:Show how to configure gap width and overlap.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=450, height=250)
        series_group = shape.chart.series_groups[0]
        # Set column gap width and overlap.
        series_group.gap_width = 450
        series_group.overlap = -75
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.ConfigureGapOverlap.docx')
        #ExEnd:ConfigureGapOverlap

    def test_bubble_scale(self):
        #ExStart:BubbleScale
        #ExFor:ChartSeriesGroup.bubble_scale
        #ExSummary:Show how to set size of the bubbles.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a bubble 3D chart.
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.BUBBLE_3D, width=450, height=250)
        series_group = shape.chart.series_groups[0]
        # Set bubble scale to 200%.
        series_group.bubble_scale = 200
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.BubbleScale.docx')
        #ExEnd:BubbleScale

    def test_remove_secondary_axis(self):
        #ExStart:RemoveSecondaryAxis
        #ExFor:ChartSeriesGroupCollection.count
        #ExFor:ChartSeriesGroupCollection.__getitem__(int)
        #ExFor:ChartSeriesGroupCollection.remove_at(int)
        #ExSummary:Show how to remove secondary axis.
        doc = aw.Document(file_name=MY_DIR + 'Combo chart.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        chart = shape.chart
        series_groups = chart.series_groups
        # Find secondary axis and remove from the collection.
        i = 0
        while i < series_groups.count:
            if series_groups[i].axis_group == aw.drawing.charts.AxisGroup.SECONDARY:
                series_groups.remove_at(i)
            i += 1
        #ExEnd:RemoveSecondaryAxis

    def test_histogram_chart(self):
        #ExStart:HistogramChart
        #ExFor:ChartSeriesCollection.add(str,List[float])
        #ExSummary:Shows how to create histogram chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a Histogram chart.
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.HISTOGRAM, width=450, height=450)
        chart = shape.chart
        chart.title.text = 'Avg Temperature since 1991'
        # Delete default generated series.
        chart.series.clear()
        # Add a series.
        chart.series.add(series_name='Avg Temperature', x_values=[51.8, 53.6, 50.3, 54.7, 53.9, 54.3, 53.4, 52.9, 53.3, 53.7, 53.8, 52, 55, 52.1, 53.4, 53.8, 53.8, 51.9, 52.1, 52.7, 51.8, 56.6, 53.3, 55.6, 56.3, 56.2, 56.1, 56.2, 53.6, 55.7, 56.3, 55.9, 55.6])
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.Histogram.docx')
        #ExEnd:HistogramChart

    def test_pareto_chart(self):
        #ExStart:ParetoChart
        #ExFor:ChartSeriesCollection.add(str,List[str],List[float])
        #ExSummary:Shows how to create pareto chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a Pareto chart.
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.PARETO, width=450, height=450)
        chart = shape.chart
        chart.title.text = 'Best-Selling Car'
        # Delete default generated series.
        chart.series.clear()
        # Add a series.
        chart.series.add(series_name='Best-Selling Car', categories=['Tesla Model Y', 'Toyota Corolla', 'Toyota RAV4', 'Ford F-Series', 'Honda CR-V'], values=[1.43, 0.91, 1.17, 0.98, 0.85])
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.Pareto.docx')
        #ExEnd:ParetoChart

    def test_box_and_whisker_chart(self):
        #ExStart:BoxAndWhiskerChart
        #ExFor:ChartSeriesCollection.add(str,List[str],List[float])
        #ExSummary:Shows how to create box and whisker chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a Box & Whisker chart.
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.BOX_AND_WHISKER, width=450, height=450)
        chart = shape.chart
        chart.title.text = 'Points by Years'
        # Delete default generated series.
        chart.series.clear()
        # Add a series.
        series = chart.series.add(series_name='Points by Years', categories=['WC', 'WC', 'WC', 'WC', 'WC', 'WC', 'WC', 'WC', 'WC', 'WC', 'NR', 'NR', 'NR', 'NR', 'NR', 'NR', 'NR', 'NR', 'NR', 'NR', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA'], values=[91, 80, 100, 77, 90, 104, 105, 118, 120, 101, 114, 107, 110, 60, 79, 78, 77, 102, 101, 113, 94, 93, 84, 71, 80, 103, 80, 94, 100, 101])
        # Show data labels.
        series.has_data_labels = True
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.BoxAndWhisker.docx')
        #ExEnd:BoxAndWhiskerChart

    def test_waterfall_chart(self):
        #ExStart:WaterfallChart
        #ExFor:ChartSeriesCollection.add(str,List[str],List[float],List[bool])
        #ExSummary:Shows how to create waterfall chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a Waterfall chart.
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.WATERFALL, width=450, height=450)
        chart = shape.chart
        chart.title.text = 'New Zealand GDP'
        # Delete default generated series.
        chart.series.clear()
        # Add a series.
        series = chart.series.add(series_name='New Zealand GDP', categories=['2018', '2019 growth', '2020 growth', '2020', '2021 growth', '2022 growth', '2022'], values=[100, 0.57, -0.25, 100.32, 20.22, -2.92, 117.62], is_subtotal=[True, False, False, True, False, False, True])
        # Show data labels.
        series.has_data_labels = True
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.Waterfall.docx')
        #ExEnd:WaterfallChart

    def test_label_orientation_rotation(self):
        #ExStart:LabelOrientationRotation
        #ExFor:ChartDataLabelCollection.orientation
        #ExFor:ChartDataLabelCollection.rotation
        #ExFor:ChartDataLabel.rotation
        #ExFor:ChartDataLabel.orientation
        #ExFor:ShapeTextOrientation
        #ExSummary:Shows how to change orientation and rotation for data labels.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=432, height=252)
        series = shape.chart.series[0]
        data_labels = series.data_labels
        # Show data labels.
        series.has_data_labels = True
        data_labels.show_value = True
        data_labels.show_category_name = True
        # Define data label shape.
        data_labels.format.shape_type = aw.drawing.charts.ChartShapeType.UP_ARROW
        data_labels.format.stroke.fill.solid(aspose.pydrawing.Color.dark_blue)
        # Set data label orientation and rotation for the entire series.
        data_labels.orientation = aw.drawing.ShapeTextOrientation.VERTICAL_FAR_EAST
        data_labels.rotation = -45
        # Change orientation and rotation of the first data label.
        data_labels[0].orientation = aw.drawing.ShapeTextOrientation.HORIZONTAL
        data_labels[0].rotation = 45
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.LabelOrientationRotation.docx')
        #ExEnd:LabelOrientationRotation

    def test_tick_labels_orientation_rotation(self):
        #ExStart:TickLabelsOrientationRotation
        #ExFor:AxisTickLabels.rotation
        #ExFor:AxisTickLabels.orientation
        #ExSummary:Shows how to change orientation and rotation for axis tick labels.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a column chart.
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=432, height=252)
        x_tick_labels = shape.chart.axis_x.tick_labels
        y_tick_labels = shape.chart.axis_y.tick_labels
        # Set axis tick label orientation and rotation.
        x_tick_labels.orientation = aw.drawing.ShapeTextOrientation.VERTICAL_FAR_EAST
        x_tick_labels.rotation = -30
        y_tick_labels.orientation = aw.drawing.ShapeTextOrientation.HORIZONTAL
        y_tick_labels.rotation = 45
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.TickLabelsOrientationRotation.docx')
        #ExEnd:TickLabelsOrientationRotation

    def test_doughnut_chart(self):
        #ExStart:DoughnutChart
        #ExFor:ChartSeriesGroup.doughnut_hole_size
        #ExFor:ChartSeriesGroup.first_slice_angle
        #ExSummary:Shows how to create and format Doughnut chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.DOUGHNUT, width=400, height=400)
        chart = shape.chart
        # Delete the default generated series.
        chart.series.clear()
        categories = ['Category 1', 'Category 2', 'Category 3']
        chart.series.add(series_name='Series 1', categories=categories, values=[4, 2, 5])
        # Format the Doughnut chart.
        series_group = chart.series_groups[0]
        series_group.doughnut_hole_size = 10
        series_group.first_slice_angle = 270
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.DoughnutChart.docx')
        #ExEnd:DoughnutChart

    def test_pie_of_pie_chart(self):
        #ExStart:PieOfPieChart
        #ExFor:ChartSeriesGroup.second_section_size
        #ExSummary:Shows how to create and format pie of Pie chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.PIE_OF_PIE, width=440, height=300)
        chart = shape.chart
        # Delete the default generated series.
        chart.series.clear()
        categories = ['Category 1', 'Category 2', 'Category 3', 'Category 4']
        chart.series.add(series_name='Series 1', categories=categories, values=[11, 8, 4, 3])
        # Format the Pie of Pie chart.
        series_group = chart.series_groups[0]
        series_group.gap_width = 10
        series_group.second_section_size = 77
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.PieOfPieChart.docx')
        #ExEnd:PieOfPieChart

    def test_format_code(self):
        #ExStart:FormatCode
        #ExFor:ChartXValueCollection.format_code
        #ExFor:ChartYValueCollection.format_code
        #ExFor:BubbleSizeCollection.format_code
        #ExFor:ChartSeries.bubble_sizes
        #ExFor:ChartSeries.x_values
        #ExFor:ChartSeries.y_values
        #ExSummary:Shows how to work with the format code of the chart data.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a Bubble chart.
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.BUBBLE, width=432, height=252)
        chart = shape.chart
        # Delete default generated series.
        chart.series.clear()
        series = chart.series.add(series_name='Series1', x_values=[1, 1.9, 2.45, 3], y_values=[1, -0.9, 1.82, 0], bubble_sizes=[2, 1.1, 2.95, 2])
        # Show data labels.
        series.has_data_labels = True
        series.data_labels.show_category_name = True
        series.data_labels.show_value = True
        series.data_labels.show_bubble_size = True
        # Set data format codes.
        series.x_values.format_code = '#,##0.0#'
        series.y_values.format_code = '#,##0.0#;[Red]\\-#,##0.0#'
        series.bubble_sizes.format_code = '#,##0.0#'
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.FormatCode.docx')
        #ExEnd:FormatCode
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.FormatCode.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        chart = shape.chart
        series_collection = chart.series
        for series_properties in series_collection:
            self.assertEqual('#,##0.0#', series_properties.x_values.format_code)
            self.assertEqual('#,##0.0#;[Red]\\-#,##0.0#', series_properties.y_values.format_code)
            self.assertEqual('#,##0.0#', series_properties.bubble_sizes.format_code)

    def test_data_lable_position(self):
        #ExStart:DataLablePosition
        #ExFor:ChartDataLabelCollection.position
        #ExFor:ChartDataLabel.position
        #ExFor:ChartDataLabelPosition
        #ExSummary:Shows how to set the position of the data label.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert column chart.
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=432, height=252)
        chart = shape.chart
        series_coll = chart.series
        # Delete default generated series.
        series_coll.clear()
        # Add series.
        series = series_coll.add(series_name='Series 1', categories=['Category 1', 'Category 2', 'Category 3'], values=[4, 5, 6])
        # Show data labels and set font color.
        series.has_data_labels = True
        data_labels = series.data_labels
        data_labels.show_value = True
        data_labels.font.color = aspose.pydrawing.Color.white
        # Set data label position.
        data_labels.position = aw.drawing.charts.ChartDataLabelPosition.INSIDE_BASE
        data_labels[0].position = aw.drawing.charts.ChartDataLabelPosition.OUTSIDE_END
        data_labels[0].font.color = aspose.pydrawing.Color.dark_red
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.LabelPosition.docx')
        #ExEnd:DataLablePosition

    def test_doughnut_chart_label_position(self):
        #ExStart:DoughnutChartLabelPosition
        #ExFor:ChartDataLabel.left
        #ExFor:ChartDataLabel.left_mode
        #ExFor:ChartDataLabel.top
        #ExFor:ChartDataLabel.top_mode
        #ExFor:ChartDataLabelLocationMode
        #ExSummary:Shows how to place data labels of doughnut chart outside doughnut.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        chart_width = 432
        chart_height = 252
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.DOUGHNUT, width=chart_width, height=chart_height)
        chart = shape.chart
        series_coll = chart.series
        # Delete default generated series.
        series_coll.clear()
        # Hide the legend.
        chart.legend.position = aw.drawing.charts.LegendPosition.NONE
        # Generate data.
        data_length = 20
        total_value = 0
        categories = [None for i in range(0, data_length)]
        values = [None for i in range(0, data_length)]
        i = 0
        while i < data_length:
            categories[i] = f'Category {i}'
            values[i] = data_length - i
            total_value = total_value + values[i]
            i += 1
        series = series_coll.add(series_name='Series 1', categories=categories, values=values)
        series.has_data_labels = True
        data_labels = series.data_labels
        data_labels.show_value = True
        data_labels.show_leader_lines = True
        # The Position property cannot be used for doughnut charts. Let's place data labels using the Left and Top
        # properties around a circle outside of the chart doughnut.
        # The origin is in the upper left corner of the chart.
        title_area_height = 25.5  # This can be calculated using title text and font.
        doughnut_center_y = title_area_height + (chart_height - title_area_height) / 2
        doughnut_center_x = chart_width / 2
        label_height = 16.5  # This can be calculated using label font.
        one_char_label_width = 12.75  # This can be calculated for each label using its text and font.
        two_char_label_width = 17.25  # This can be calculated for each label using its text and font.
        y_margin = 0.75
        label_margin = 1.5
        label_circle_radius = chart_height - doughnut_center_y - y_margin - label_height / 2
        # Because the data points start at the top, the X coordinates used in the Left and Top properties of
        # the data labels point to the right and the Y coordinates point down, the starting angle is -PI/2.
        total_angle = -math.pi / 2
        previous_label = None
        i = 0
        while i < series.y_values.count:
            data_label = data_labels[i]
            value = series.y_values[i].double_value
            label_width = None
            if value < 10:
                label_width = one_char_label_width
            else:
                label_width = two_char_label_width
            label_segment_angle = value / total_value * 2 * math.pi
            label_angle = label_segment_angle / 2 + total_angle
            label_center_x = label_circle_radius * math.cos(label_angle) + doughnut_center_x
            label_center_y = label_circle_radius * math.sin(label_angle) + doughnut_center_y
            label_left = label_center_x - label_width / 2
            label_top = label_center_y - label_height / 2
            # If the current data label overlaps other labels, move it horizontally.
            if previous_label != None and math.fabs(previous_label.top - label_top) < label_height and (math.fabs(previous_label.left - label_left) < label_width):
                # Move right on the top, left on the bottom.
                is_on_top = total_angle < 0 or total_angle >= math.pi
                factor = None
                if is_on_top:
                    factor = 1
                else:
                    factor = -1
                label_left = previous_label.left + label_width * factor + label_margin
            data_label.left = label_left
            data_label.left_mode = aw.drawing.charts.ChartDataLabelLocationMode.ABSOLUTE
            data_label.top = label_top
            data_label.top_mode = aw.drawing.charts.ChartDataLabelLocationMode.ABSOLUTE
            total_angle = total_angle + label_segment_angle
            previous_label = data_label
            i += 1
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.DoughnutChartLabelPosition.docx')
        #ExEnd:DoughnutChartLabelPosition

    def test_insert_chart_series(self):
        #ExStart
        #ExFor:ChartSeries.insert(int,ChartXValue)
        #ExFor:ChartSeries.insert(int,ChartXValue,ChartYValue)
        #ExFor:ChartSeries.insert(int,ChartXValue,ChartYValue,float)
        #ExSummary:Shows how to insert data into a chart series.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.LINE, width=432, height=252)
        chart = shape.chart
        series1 = chart.series[0]
        # Clear X and Y values of the first series.
        series1.clear_values()
        # Populate the series with data.
        series1.insert(index=0, x_value=aw.drawing.charts.ChartXValue.from_double(3))
        series1.insert(index=1, x_value=aw.drawing.charts.ChartXValue.from_double(3), y_value=aw.drawing.charts.ChartYValue.from_double(10))
        series1.insert(index=2, x_value=aw.drawing.charts.ChartXValue.from_double(3), y_value=aw.drawing.charts.ChartYValue.from_double(10))
        series1.insert(index=3, x_value=aw.drawing.charts.ChartXValue.from_double(3), y_value=aw.drawing.charts.ChartYValue.from_double(10), bubble_size=10)
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.PopulateChartWithData.docx')
        #ExEnd

    def test_set_chart_style(self):
        #ExStart
        #ExFor:ChartStyle
        #ExSummary:Shows how to set and get chart style.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a chart in the Black style.
        builder.insert_chart(chart_type=aw.drawing.charts.ChartType.COLUMN, width=400, height=250, chart_style=aw.drawing.charts.ChartStyle.BLACK)
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.SetChartStyle.docx')
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Charts.SetChartStyle.docx')
        # Get a chart to update.
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        chart = shape.chart
        # Get the chart style.
        self.assertEqual(aw.drawing.charts.ChartStyle.BLACK, chart.style)
        #ExEnd

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
        dates = [date(2017, 11, 6), date(2017, 11, 9), date(2017, 11, 15), date(2017, 11, 21), date(2017, 11, 25), date(2017, 11, 29)]
        chart.series.add('Aspose Test Series', dates=dates, values=[1.2, 0.3, 2.1, 2.9, 4.2, 5.3])
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
        y_axis.tick_labels.position = aw.drawing.charts.AxisTickLabelPosition.HIGH
        y_axis.major_unit = 100.0
        y_axis.minor_unit = 50.0
        y_axis.display_unit.unit = aw.drawing.charts.AxisBuiltInUnit.HUNDREDS
        y_axis.scaling.minimum = aw.drawing.charts.AxisBound(100)
        y_axis.scaling.maximum = aw.drawing.charts.AxisBound(700)
        y_axis.has_major_gridlines = True
        y_axis.has_minor_gridlines = True
        doc.save(ARTIFACTS_DIR + 'Charts.date_time_values.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Charts.date_time_values.docx')
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
        self.assertEqual(aw.drawing.charts.AxisTickLabelPosition.HIGH, chart.axis_y.tick_labels.position)
        self.assertEqual(100.0, chart.axis_y.major_unit)
        self.assertEqual(50.0, chart.axis_y.minor_unit)
        self.assertEqual(aw.drawing.charts.AxisBuiltInUnit.HUNDREDS, chart.axis_y.display_unit.unit)
        self.assertEqual(aw.drawing.charts.AxisBound(100), chart.axis_y.scaling.minimum)
        self.assertEqual(aw.drawing.charts.AxisBound(700), chart.axis_y.scaling.maximum)
        self.assertEqual(True, chart.axis_y.has_major_gridlines)
        self.assertEqual(True, chart.axis_y.has_minor_gridlines)

    def test_data_labels(self):
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

        def data_labels():
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc)
            chart_shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 400, 300)
            chart = chart_shape.chart
            self.assertEqual(3, chart.series.count)
            self.assertEqual('Series 1', chart.series[0].name)
            self.assertEqual('Series 2', chart.series[1].name)
            self.assertEqual('Series 3', chart.series[2].name)
            # Apply data labels to every series in the chart.
            # These labels will appear next to each data point in the graph and display its value.
            for series in chart.series:
                apply_data_labels(series, 4, '000.0', ', ')
                self.assertEqual(4, series.data_labels.count)
            # Change the separator string for every data label in a series.
            for label in chart.series[0].data_labels:
                self.assertEqual(', ', label.separator)
                label.separator = ' & '
            # For a cleaner looking graph, we can remove data labels individually.
            chart.series[1].data_labels[2].clear_format()
            # We can also strip an entire series of its data labels at once.
            chart.series[2].data_labels.clear_format()
            doc.save(ARTIFACTS_DIR + 'Charts.data_labels.docx')

        def apply_data_labels(series: aw.drawing.charts.ChartSeries, labels_count: int, number_format: str, separator: str):
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
        data_labels()

    def test_chart_data_point(self):
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

        def chart_data_point():
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc)
            shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 500, 350)
            chart = shape.chart
            self.assertEqual(3, chart.series.count)
            self.assertEqual('Series 1', chart.series[0].name)
            self.assertEqual('Series 2', chart.series[1].name)
            self.assertEqual('Series 3', chart.series[2].name)
            # Emphasize the chart's data points by making them appear as diamond shapes.
            for series in chart.series:
                apply_data_points(series, 4, aw.drawing.charts.MarkerSymbol.DIAMOND, 15)
            # Smooth out the line that represents the first data series.
            chart.series[0].smooth = True
            # Verify that data points for the first series will not invert their colors if the value is negative.
            for data_point in chart.series[0].data_points:
                self.assertFalse(data_point.invert_if_negative)
            # For a cleaner looking graph, we can clear format individually.
            chart.series[1].data_points[2].clear_format()
            # We can also strip an entire series of data points at once.
            chart.series[2].data_points.clear_format()
            doc.save(ARTIFACTS_DIR + 'Charts.chart_data_point.docx')

        def apply_data_points(series: aw.drawing.charts.ChartSeries, data_points_count: int, marker_symbol: aw.drawing.charts.MarkerSymbol, data_point_size: int):
            """Applies a number of data points to a series."""
            for i in range(data_points_count):
                point = series.data_points[i]
                point.marker.symbol = marker_symbol
                point.marker.size = data_point_size
                self.assertEqual(i, point.index)
        #ExEnd
        chart_data_point()

    def test_chart_series_collection(self):
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

        def chart_series_collection():
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc)
            # There are several ways of populating a chart's series collection.
            # Different series schemas are intended for different chart types.
            # 1 -  Column chart with columns grouped and banded along the X-axis by category:
            chart = append_chart(builder, aw.drawing.charts.ChartType.COLUMN, 500, 300)
            categories = ['Category 1', 'Category 2', 'Category 3']
            # Insert two series of decimal values containing a value for each respective category.
            # This column chart will have three groups, each with two columns.
            chart.series.add('Series 1', categories, [76.6, 82.1, 91.6])
            chart.series.add('Series 2', categories, [64.2, 79.5, 94.0])
            # Categories are distributed along the X-axis, and values are distributed along the Y-axis.
            self.assertEqual(aw.drawing.charts.ChartAxisType.CATEGORY, chart.axis_x.type)
            self.assertEqual(aw.drawing.charts.ChartAxisType.VALUE, chart.axis_y.type)
            # 2 -  Area chart with dates distributed along the X-axis:
            chart = append_chart(builder, aw.drawing.charts.ChartType.AREA, 500, 300)
            dates = [date(2014, 3, 31), date(2017, 1, 23), date(2017, 6, 18), date(2019, 11, 22), date(2020, 9, 7)]
            # Insert a series with a decimal value for each respective date.
            # The dates will be distributed along a linear X-axis,
            # and the values added to this series will create data points.
            chart.series.add('Series 1', dates=dates, values=[15.8, 21.5, 22.9, 28.7, 33.1])
            self.assertEqual(aw.drawing.charts.ChartAxisType.CATEGORY, chart.axis_x.type)
            self.assertEqual(aw.drawing.charts.ChartAxisType.VALUE, chart.axis_y.type)
            # 3 -  2D scatter plot:
            chart = append_chart(builder, aw.drawing.charts.ChartType.SCATTER, 500, 300)
            # Each series will need two decimal arrays of equal length.
            # The first array contains X-values, and the second contains corresponding Y-values
            # of data points on the chart's graph.
            chart.series.add('Series 1', x_values=[3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6], y_values=[3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9])
            chart.series.add('Series 2', x_values=[2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3], y_values=[7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6])
            self.assertEqual(aw.drawing.charts.ChartAxisType.VALUE, chart.axis_x.type)
            self.assertEqual(aw.drawing.charts.ChartAxisType.VALUE, chart.axis_y.type)
            # 4 -  Bubble chart:
            chart = append_chart(builder, aw.drawing.charts.ChartType.BUBBLE, 500, 300)
            # Each series will need three decimal arrays of equal length.
            # The first array contains X-values, the second contains corresponding Y-values,
            # and the third contains diameters for each of the graph's data points.
            chart.series.add('Series 1', [1.1, 5.0, 9.8], [1.2, 4.9, 9.9], [2.0, 4.0, 8.0])
            doc.save(ARTIFACTS_DIR + 'Charts.chart_series_collection.docx')

        def append_chart(builder: aw.DocumentBuilder, chart_type: aw.drawing.charts.ChartType, width: float, height: float) -> aw.drawing.charts.Chart:
            """Insert a chart using a document builder of a specified ChartType, width and height, and remove its demo data."""
            chart_shape = builder.insert_chart(chart_type, width, height)
            chart = chart_shape.chart
            chart.series.clear()
            self.assertEqual(0, chart.series.count)  #ExSkip
            return chart
        #ExEnd
        chart_series_collection()

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
        categories = ['Category 1', 'Category 2', 'Category 3', 'Category 4']
        # We can add a series with new values for existing categories.
        # This chart will now contain four clusters of four columns.
        chart.series.add('Series 4', categories, [4.4, 7.0, 3.5, 2.1])
        self.assertEqual(4, chart_data.count)  #ExSkip
        self.assertEqual('Series 4', chart_data[3].name)  #ExSkip
        # A chart series can also be removed by index, like this.
        # This will remove one of the three demo series that came with the chart.
        chart_data.remove_at(2)
        self.assertFalse(any((s for s in chart_data if s.name == 'Series 3')))
        self.assertEqual(3, chart_data.count)  #ExSkip
        self.assertEqual('Series 4', chart_data[2].name)  #ExSkip
        # We can also clear all the chart's data at once with this method.
        # When creating a new chart, this is the way to wipe all the demo data
        # before we can begin working on a blank chart.
        chart_data.clear()
        self.assertEqual(0, chart_data.count)  #ExSkip
        #ExEnd

    def test_data_arrays_wrong_size(self):
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        shape = builder.insert_chart(aw.drawing.charts.ChartType.LINE, 500, 300)
        chart = shape.chart
        series_coll = chart.series
        series_coll.clear()
        categories = ['Cat1', None, 'Cat3', 'Cat4', 'Cat5', None]
        series_coll.add('AW Series 1', categories, [1, 2, nan, 4, 5, 6])
        series_coll.add('AW Series 2', categories, [2, 3, nan, 5, 6, 7])
        with self.assertRaises(Exception):
            series_coll.add('AW Series 3', categories, [nan, 4, 5, nan, nan])
        with self.assertRaises(Exception):
            series_coll.add('AW Series 4', categories, [nan, nan, nan, nan, nan])

    def test_treemap_chart(self):
        #ExStart:TreemapChart
        #ExFor:ChartSeriesCollection.add_multilevel_value(str,List[ChartMultilevelValue],List[float])
        #ExFor:ChartMultilevelValue.__init__(str,str)
        #ExSummary:Shows how to create treemap chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Insert a Treemap chart.
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.TREEMAP, width=450, height=280)
        chart = shape.chart
        chart.title.text = 'World Population'
        # Delete default generated series.
        chart.series.clear()
        # Add a series.
        series = chart.series.add_multilevel_value(series_name='Population by Region', categories=[aw.drawing.charts.ChartMultilevelValue(level1='Asia', level2='China'), aw.drawing.charts.ChartMultilevelValue(level1='Asia', level2='India'), aw.drawing.charts.ChartMultilevelValue(level1='Asia', level2='Indonesia'), aw.drawing.charts.ChartMultilevelValue(level1='Asia', level2='Pakistan'), aw.drawing.charts.ChartMultilevelValue(level1='Asia', level2='Bangladesh'), aw.drawing.charts.ChartMultilevelValue(level1='Asia', level2='Japan'), aw.drawing.charts.ChartMultilevelValue(level1='Asia', level2='Philippines'), aw.drawing.charts.ChartMultilevelValue(level1='Asia', level2='Other'), aw.drawing.charts.ChartMultilevelValue(level1='Africa', level2='Nigeria'), aw.drawing.charts.ChartMultilevelValue(level1='Africa', level2='Ethiopia'), aw.drawing.charts.ChartMultilevelValue(level1='Africa', level2='Egypt'), aw.drawing.charts.ChartMultilevelValue(level1='Africa', level2='Other'), aw.drawing.charts.ChartMultilevelValue(level1='Europe', level2='Russia'), aw.drawing.charts.ChartMultilevelValue(level1='Europe', level2='Germany'), aw.drawing.charts.ChartMultilevelValue(level1='Europe', level2='Other'), aw.drawing.charts.ChartMultilevelValue(level1='Latin America', level2='Brazil'), aw.drawing.charts.ChartMultilevelValue(level1='Latin America', level2='Mexico'), aw.drawing.charts.ChartMultilevelValue(level1='Latin America', level2='Other'), aw.drawing.charts.ChartMultilevelValue(level1='Northern America', level2='United States'), aw.drawing.charts.ChartMultilevelValue(level1='Northern America', level2='Other'), aw.drawing.charts.ChartMultilevelValue(level1='Oceania')], values=[1409670000, 1400744000, 279118866, 241499431, 169828911, 123930000, 112892781, 764000000, 223800000, 107334000, 105914499, 903000000, 146150789, 84607016, 516000000, 203080756, 129713690, 310000000, 335893238, 35000000, 42000000])
        # Show data labels.
        series.has_data_labels = True
        series.data_labels.show_value = True
        series.data_labels.show_category_name = True
        decimal_separator = locale.localeconv()['decimal_point']
        series.data_labels.number_format.format_code = f'0{decimal_separator}0%'
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.Treemap.docx')
        #ExEnd:TreemapChart

    def test_sunburst_chart(self):
        #ExStart:SunburstChart
        #ExFor:ChartSeriesCollection.add(str,List[ChartMultilevelValue],List[float])
        #ExSummary:Shows how to create sunburst chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Insert a Sunburst chart.
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.SUNBURST, width=450, height=450)
        chart = shape.chart
        chart.title.text = 'Sales'
        # Delete default generated series.
        chart.series.clear()
        # Add a series.
        series = chart.series.add_multilevel_value(series_name='Sales', categories=[aw.drawing.charts.ChartMultilevelValue(level1='Sales - Europe', level2='UK', level3='London Dep.'), aw.drawing.charts.ChartMultilevelValue(level1='Sales - Europe', level2='UK', level3='Liverpool Dep.'), aw.drawing.charts.ChartMultilevelValue(level1='Sales - Europe', level2='UK', level3='Manchester Dep.'), aw.drawing.charts.ChartMultilevelValue(level1='Sales - Europe', level2='France', level3='Paris Dep.'), aw.drawing.charts.ChartMultilevelValue(level1='Sales - Europe', level2='France', level3='Lyon Dep.'), aw.drawing.charts.ChartMultilevelValue(level1='Sales - NA', level2='USA', level3='Denver Dep.'), aw.drawing.charts.ChartMultilevelValue(level1='Sales - NA', level2='USA', level3='Seattle Dep.'), aw.drawing.charts.ChartMultilevelValue(level1='Sales - NA', level2='USA', level3='Detroit Dep.'), aw.drawing.charts.ChartMultilevelValue(level1='Sales - NA', level2='USA', level3='Houston Dep.'), aw.drawing.charts.ChartMultilevelValue(level1='Sales - NA', level2='Canada', level3='Toronto Dep.'), aw.drawing.charts.ChartMultilevelValue(level1='Sales - NA', level2='Canada', level3='Montreal Dep.'), aw.drawing.charts.ChartMultilevelValue(level1='Sales - Oceania', level2='Australia', level3='Sydney Dep.'), aw.drawing.charts.ChartMultilevelValue(level1='Sales - Oceania', level2='New Zealand', level3='Auckland Dep.')], values=[1236, 851, 536, 468, 179, 527, 799, 1148, 921, 457, 482, 761, 694])
        # Show data labels.
        series.has_data_labels = True
        series.data_labels.show_value = False
        series.data_labels.show_category_name = True
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.Sunburst.docx')
        #ExEnd:SunburstChart

    def test_funnel_chart(self):
        #ExStart:FunnelChart
        #ExFor:ChartSeriesCollection.add(str,List[str],List[float])
        #ExSummary:Shows how to create funnel chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Insert a Funnel chart.
        shape = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.FUNNEL, width=450, height=450)
        chart = shape.chart
        chart.title.text = 'Population by Age Group'
        # Delete default generated series.
        chart.series.clear()
        # Add a series.
        series = chart.series.add(series_name='Population by Age Group', categories=['0-9', '10-19', '20-29', '30-39', '40-49', '50-59', '60-69', '70-79', '80-89', '90-'], values=[0.121, 0.128, 0.132, 0.146, 0.124, 0.124, 0.111, 0.075, 0.032, 0.007])
        # Show data labels.
        series.has_data_labels = True
        decimal_separator = locale.localeconv()['decimal_point']
        series.data_labels.number_format.format_code = f'0{decimal_separator}0%'
        doc.save(file_name=ARTIFACTS_DIR + 'Charts.Funnel.docx')
        #ExEnd:FunnelChart