# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import io
import json
import unittest
from datetime import datetime
from typing import List, Optional

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, IMAGE_DIR, GOLDS_DIR
from document_helper import DocumentHelper
from testdata import Common, ClientTestClass, ColorItemTestClass, DocumentTestClass, ImageTestClass, MessageTestClass, NumericTestClass

class ExReportingEngine(ApiExampleBase):

    def __init__(self, *args, **kwargs):
        ApiExampleBase.__init__(self, *args, **kwargs)

        self.image = IMAGE_DIR + "Logo.jpg"
        self.document = MY_DIR + "Reporting engine template - Data table.docx"

    def test_simple_case(self):

        doc = DocumentHelper.create_simple_document("<<[s.Name]>> says: <<[s.Message]>>")

        sender = MessageTestClass("LINQ Reporting Engine", "Hello World")
        self.build_report(doc, self.create_json_data_source(sender), "s", options=aw.reporting.ReportBuildOptions.INLINE_ERROR_MESSAGES)

        doc = DocumentHelper.save_open(doc)

        self.assertEqual("LINQ Reporting Engine says: Hello World\f", doc.get_text())

    def test_string_format(self):

        doc = DocumentHelper.create_simple_document(
            "<<[s.Name]:lower>> says: <<[s.Message]:upper>>, <<[s.Message]:caps>>, <<[s.Message]:firstCap>>")

        sender = MessageTestClass("LINQ Reporting Engine", "hello world")
        self.build_report(doc, self.create_json_data_source(sender), "s")

        doc = DocumentHelper.save_open(doc)

        self.assertEqual("linq reporting engine says: HELLO WORLD, Hello World, Hello world\f", doc.get_text())

    def test_number_format(self):

        doc = DocumentHelper.create_simple_document(
            "<<[s.Value1]:alphabetic>> : <<[s.Value2]:roman:lower>>, <<[s.Value3]:ordinal>>, <<[s.Value1]:ordinalText:upper>>" +
            ", <<[s.Value2]:cardinal>>, <<[s.Value3]:hex>>, <<[s.Value3]:arabicDash>>")

        sender = NumericTestClass(1, 2.2, 200, None)
        self.build_report(doc, self.create_json_data_source(sender), "s")

        doc = DocumentHelper.save_open(doc)

        self.assertEqual("A : ii, 200th, FIRST, Two, C8, - 200 -\f", doc.get_text())

    def test_data_table(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Data table.docx")

        self.build_report(doc, self.create_json_data_source(Common.get_contracts()), "Contracts")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.test_data_table.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.data_table.docx", GOLDS_DIR + "ReportingEngine.TestDataTable Gold.docx"))

    def test_total(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Total.docx")

        self.build_report(doc, self.create_json_data_source(Common.get_contracts()), "Contracts")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.total.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.total.docx", GOLDS_DIR + "ReportingEngine.Total Gold.docx"))

    def test_nested_data_table(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Nested data table.docx")

        self.build_report(doc, self.create_json_data_source(Common.get_managers()), "Managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.nested_data_table.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.nested_data_table.docx", GOLDS_DIR + "ReportingEngine.TestNestedDataTable Gold.docx"))

    def test_restarting_list_numbering_dynamically(self):

        template = aw.Document(MY_DIR + "Reporting engine template - List numbering.docx")

        self.build_report(template, self.create_json_data_source(Common.get_managers()), "Managers", options=aw.reporting.ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS)

        template.save(ARTIFACTS_DIR + "ReportingEngine.restarting_list_numbering_dynamically.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.restarting_list_numbering_dynamically.docx", GOLDS_DIR + "ReportingEngine.RestartingListNumberingDynamically Gold.docx"))

    def test_restarting_list_numbering_dynamically_while_inserting_document_dynamically(self):

        template = DocumentHelper.create_simple_document("<<doc [src.Document] -build>>")

        doc = DocumentTestClass(doc=aw.Document(MY_DIR + "Reporting engine template - List numbering.docx"))

        self.build_report(template, [doc, Common.get_managers()], ["src", "Managers"], options=aw.reporting.ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS)

        template.save(ARTIFACTS_DIR + "ReportingEngine.restarting_list_numbering_dynamically_while_inserting_document_dynamically.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.restarting_list_numbering_dynamically_while_inserting_document_dynamically.docx", GOLDS_DIR + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically Gold.docx"))

    def test_restarting_list_numbering_dynamically_while_multiple_insertions_document_dynamically(self):

        main_template = DocumentHelper.create_simple_document("<<doc [src] -build>>")
        template1 = DocumentHelper.create_simple_document("<<doc [src1] -build>>")
        template2 = DocumentHelper.create_simple_document("<<doc [src2.Document] -build>>")

        doc = DocumentTestClass(doc=aw.Document(MY_DIR + "Reporting engine template - List numbering.docx"))

        self.build_report(main_template,
                          [template1, template2, doc, Common.get_managers()],
                          ["src", "src1", "src2", "Managers"],
                          options=aw.reporting.ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS)

        main_template.save(ARTIFACTS_DIR + "ReportingEngine.restarting_list_numbering_dynamically_while_multiple_insertions_document_dynamically.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.restarting_list_numbering_dynamically_while_multiple_insertions_document_dynamically.docx", GOLDS_DIR + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically Gold.docx"))

    def test_chart(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Chart.docx")

        self.build_report(doc, self.create_json_data_source(Common.get_managers()), "managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.chart.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.chart.docx", GOLDS_DIR + "ReportingEngine.TestChart Gold.docx"))

    def test_bubble_chart(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Bubble chart.docx")

        self.build_report(doc, self.create_json_data_source(Common.get_managers()), "managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.bubble_chart.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.bubble_chart.docx", GOLDS_DIR + "ReportingEngine.TestBubbleChart Gold.docx"))

    def test_set_chart_series_colors_dynamically(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Chart series color.docx")

        self.build_report(doc, Common.get_managers(), "managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.set_chart_series_color_dynamically.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.set_chart_series_color_dynamically.docx", GOLDS_DIR + "ReportingEngine.SetChartSeriesColorDynamically Gold.docx"))

    def test_set_point_colors_dynamically(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Point color.docx")

        colors = [
            ColorItemTestClass("Black", drawing.Color.black.to_argb(), value1=1.0, value2=2.5, value3=3.5),
            ColorItemTestClass("Red", drawing.Color.red.to_argb(), value1=2.0, value2=4.0, value3=2.5),
            ColorItemTestClass("Green", drawing.Color.green.to_argb(), value1=0.5, value2=1.5, value3=2.5),
            ColorItemTestClass("Blue", drawing.Color.blue.to_argb(), value1=4.5, value2=3.5, value3=1.5),
            ColorItemTestClass("Yellow", drawing.Color.yellow.to_argb(), value1=5.0, value2=2.5, value3=1.5),
            ]

        self.build_report(doc, colors, "colorItems", [type(ColorItemTestClass)])

        doc.save(ARTIFACTS_DIR + "ReportingEngine.set_point_color_dynamically.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.set_point_color_dynamically.docx", GOLDS_DIR + "ReportingEngine.SetPointColorDynamically Gold.docx"))

    def test_conditional_expression_for_leave_chart_series(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Chart series.docx")

        condition = 3
        self.build_report(doc, [self.create_json_data_source(Common.get_managers()), condition], ["managers", "condition"])

        doc.save(ARTIFACTS_DIR + "ReportingEngine.conditional_expression_for_leave_chart_series.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.conditional_expression_for_leave_chart_series.docx", GOLDS_DIR + "ReportingEngine.TestLeaveChartSeries Gold.docx"))

    def test_conditional_expression_for_remove_chart_series(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Chart series.docx")

        condition = 2
        self.build_report(doc, [self.create_json_data_source(Common.get_managers()), condition], ["managers", "condition"])

        doc.save(ARTIFACTS_DIR + "ReportingEngine.conditional_expression_for_remove_chart_series.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.conditional_expression_for_remove_chart_series.docx", GOLDS_DIR + "ReportingEngine.TestRemoveChartSeries Gold.docx"))

    def test_index_of(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Index of.docx")

        self.build_report(doc, self.create_json_data_source(Common.get_managers()), "Managers")

        doc = DocumentHelper.save_open(doc)

        self.assertEqual("The names are: John Smith, Tony Anderson, July James\f", doc.get_text())

    def test_if_else(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - If-else.docx")

        self.build_report(doc, self.create_json_data_source(Common.get_managers()), "m")

        doc = DocumentHelper.save_open(doc)

        self.assertEqual("You have chosen 3 item(s).\f", doc.get_text())

    def test_if_else_without_data(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - If-else.docx")

        self.build_report(doc, Common.get_empty_managers(), "m")

        doc = DocumentHelper.save_open(doc)

        self.assertEqual("You have chosen no items.\f", doc.get_text())

    def test_extension_methods(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Extension methods.docx")

        self.build_report(doc, self.create_json_data_source(Common.get_managers()), "Managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.extension_methods.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.extension_methods.docx", GOLDS_DIR + "ReportingEngine.ExtensionMethods Gold.docx"))

    def test_operators(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Operators.docx")

        test_data = NumericTestClass(1, 2.0, 3, None, logical=True)

        report = aw.reporting.ReportingEngine()
        report.known_types.add(type(NumericTestClass))
        report.build_report(doc, test_data, "ds")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.operators.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.operators.docx", GOLDS_DIR + "ReportingEngine.Operators Gold.docx"))

    def test_header_variable(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Header variable.docx")

        self.build_report(doc, self.create_json_data_source({}), data_source_name="", options=aw.reporting.ReportBuildOptions.USE_LEGACY_HEADER_FOOTER_VISITING)

        doc.save(ARTIFACTS_DIR + "ReportingEngine.HeaderVariable.docx");

        self.assertEqual("Value of myHeaderVariable is: I am header variable", doc.first_section.body.first_paragraph.get_text().strip())

    def test_contextual_object_member_access(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Contextual object member access.docx")

        self.build_report(doc, self.create_json_data_source(Common.get_managers()), "Managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.contextual_object_member_access.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.contextual_object_member_access.docx", GOLDS_DIR + "ReportingEngine.ContextualObjectMemberAccess Gold.docx"))

    def test_insert_document_dynamically_with_additional_template_checking(self):

        template = DocumentHelper.create_simple_document("<<doc [src.Document] -build>>")

        doc = DocumentTestClass(doc=aw.Document(MY_DIR + "Reporting engine template - Data table.docx"))

        self.build_report(template, [doc, Common.get_contracts()], ["src", "Contracts"],
            options=aw.reporting.ReportBuildOptions.NONE)
        template.save(ARTIFACTS_DIR + "ReportingEngine.insert_document_dynamically_with_additional_template_checking.docx")

        self.assertTrue(
            DocumentHelper.compare_docs(
                ARTIFACTS_DIR + "ReportingEngine.insert_document_dynamically_with_additional_template_checking.docx",
                GOLDS_DIR + "ReportingEngine.InsertDocumentDynamicallyWithAdditionalTemplateChecking Gold.docx"),
            "Fail inserting document by document")

    def test_insert_document_dynamically_with_styles(self):

        template = DocumentHelper.create_simple_document("<<doc [src.Document] -sourceStyles>>")

        doc = DocumentTestClass(doc=aw.Document(MY_DIR + "Reporting engine template - Data table.docx"))

        self.build_report(template, doc, "src", options=aw.reporting.ReportBuildOptions.NONE)
        template.save(ARTIFACTS_DIR + "ReportingEngine.insert_document_dynamically_with_styles.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.insert_document_dynamically_with_styles.docx",
            GOLDS_DIR + "ReportingEngine.insert_document_dynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by document")

    def test_insert_document_dynamically_trim_last_paragraph(self):
        template = DocumentHelper.create_simple_document("<<doc [src.Document] -inline>>")

        doc = DocumentTestClass(doc=aw.Document(MY_DIR + "Reporting engine template - Data table.docx"))

        self.build_report(template, doc, "src", options=aw.reporting.ReportBuildOptions.NONE)
        template.save(ARTIFACTS_DIR + "ReportingEngine.InsertDocumentDynamically.docx")

        template = aw.Document(ARTIFACTS_DIR + "ReportingEngine.InsertDocumentDynamically.docx")
        self.assertEqual(1, template.first_section.body.paragraphs.count)

    def test_insert_document_dynamically_by_stream(self):

        template = DocumentHelper.create_simple_document("<<doc [src.DocumentStream]>>")

        with open(self.document, "rb") as stream:
            doc_stream = DocumentTestClass(doc_stream=stream)

        self.build_report(template, doc_stream, "src", options=aw.reporting.ReportBuildOptions.NONE)
        template.save(ARTIFACTS_DIR + "ReportingEngine.insert_document_dynamically_by_stream.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.insert_document_dynamically_by_stream.docx",
            GOLDS_DIR + "ReportingEngine.insert_document_dynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by stream")

    def test_insert_document_dynamically_by_bytes(self):

        template = DocumentHelper.create_simple_document("<<doc [src.DocumentBytes]>>")

        with open(MY_DIR + "Reporting engine template - Data table.docx", "rb") as stream:
            doc_bytes = DocumentTestClass(doc_bytes=stream.read())

        self.build_report(template, doc_bytes, "src", options=aw.reporting.ReportBuildOptions.NONE)
        template.save(ARTIFACTS_DIR + "ReportingEngine.insert_document_dynamically_by_bytes.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.insert_document_dynamically_by_bytes.docx",
            GOLDS_DIR + "ReportingEngine.insert_document_dynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes")

    def test_insert_document_dynamically_by_uri(self):

        template = DocumentHelper.create_simple_document("<<doc [src.DocumentString]>>")

        doc_uri = DocumentTestClass(doc_string="http://www.snee.com/xml/xslt/sample.doc")

        self.build_report(template, self.create_json_data_source(doc_uri), "src", options=aw.reporting.ReportBuildOptions.NONE)
        template.save(ARTIFACTS_DIR + "ReportingEngine.insert_document_dynamically_by_uri.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.insert_document_dynamically_by_uri.docx",
            GOLDS_DIR + "ReportingEngine.InsertDocumentDynamically(uri) Gold.docx"), "Fail inserting document by uri")

    def test_insert_document_dynamically_by_base64(self):

        template = DocumentHelper.create_simple_document("<<doc [src.DocumentString]>>")
        with open(MY_DIR + "Reporting engine template - Data table (base64).txt", "rb") as file:
            base64_template = file.read().decode('utf-8')

        doc_base64 = DocumentTestClass(doc_string=base64_template)

        self.build_report(template, self.create_json_data_source(doc_base64), "src", options=aw.reporting.ReportBuildOptions.NONE)
        template.save(ARTIFACTS_DIR + "ReportingEngine.insert_document_dynamically_by_base64.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.insert_document_dynamically_by_base64.docx",
            GOLDS_DIR + "ReportingEngine.InsertDocumentDynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by uri")

    @unittest.skip("drawing.Image type isn't supported yet")
    def test_insert_image_dynamically(self):

        template = DocumentHelper.create_template_document_with_draw_objects("<<image [src.Image]>>", aw.drawing.ShapeType.TEXT_BOX)

        image = ImageTestClass(image=drawing.Image.from_file(self.image, True))

        self.build_report(template, image, "src", options=aw.reporting.ReportBuildOptions.NONE)
        template.save(ARTIFACTS_DIR + "ReportingEngine.insert_image_dynamically.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.insert_image_dynamically.docx",
            GOLDS_DIR + "ReportingEngine.insert_image_dynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes")

    def test_insert_image_dynamically_by_stream(self):

        template = DocumentHelper.create_template_document_with_draw_objects("<<image [src.ImageStream]>>", aw.drawing.ShapeType.TEXT_BOX)
        with open(self.image, "rb") as stream:
            image_stream = ImageTestClass(image_stream=stream)

            self.build_report(template, image_stream, "src", options=aw.reporting.ReportBuildOptions.NONE)
            template.save(ARTIFACTS_DIR + "ReportingEngine.insert_image_dynamically_by_stream.docx")

            self.assertTrue(DocumentHelper.compare_docs(
                ARTIFACTS_DIR + "ReportingEngine.insert_image_dynamically_by_stream.docx",
                GOLDS_DIR + "ReportingEngine.insert_image_dynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes")

    def test_insert_image_dynamically_by_bytes(self):

        template = DocumentHelper.create_template_document_with_draw_objects("<<image [src.ImageBytes]>>", aw.drawing.ShapeType.TEXT_BOX)
        with open(self.image, "rb") as stream:
            image_bytes = ImageTestClass(image_bytes=stream.read())

        self.build_report(template, image_bytes, "src", options=aw.reporting.ReportBuildOptions.NONE)
        template.save(ARTIFACTS_DIR + "ReportingEngine.insert_image_dynamically_by_bytes.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.insert_image_dynamically_by_bytes.docx", GOLDS_DIR + "ReportingEngine.insert_image_dynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes")

    def test_insert_image_dynamically_by_uri(self):

        template = DocumentHelper.create_template_document_with_draw_objects("<<image [src.ImageString]>>", aw.drawing.ShapeType.TEXT_BOX)
        image_uri = ImageTestClass(image_string="http://joomla-aspose.dynabic.com/templates/aspose/App_Themes/V3/images/customers/americanexpress.png")

        self.build_report(template, self.create_json_data_source(image_uri), "src", options=aw.reporting.ReportBuildOptions.NONE)
        template.save(ARTIFACTS_DIR + "ReportingEngine.insert_image_dynamically_by_uri.docx")

        self.assertTrue(
            DocumentHelper.compare_docs(
                ARTIFACTS_DIR + "ReportingEngine.insert_image_dynamically_by_uri.docx",
                GOLDS_DIR + "ReportingEngine.InsertImageDynamically(uri) Gold.docx"),
            "Fail inserting document by bytes")

    def test_insert_image_dynamically_by_base64(self):

        template = DocumentHelper.create_template_document_with_draw_objects("<<image [src.ImageString]>>", aw.drawing.ShapeType.TEXT_BOX)
        with open(MY_DIR + "Reporting engine template - base64 image.txt", "rb") as file:
            base64_template = file.read().decode('utf-8')

        image_base64 = ImageTestClass(image_string=base64_template)

        self.build_report(template, self.create_json_data_source(image_base64), "src", options=aw.reporting.ReportBuildOptions.NONE)
        template.save(ARTIFACTS_DIR + "ReportingEngine.insert_image_dynamically_by_base64.docx")

        self.assertTrue(
            DocumentHelper.compare_docs(
                ARTIFACTS_DIR + "ReportingEngine.insert_image_dynamically_by_base64.docx",
                GOLDS_DIR + "ReportingEngine.InsertImageDynamically(stream,doc,bytes) Gold.docx"),
            "Fail inserting document by bytes")

    def test_image_exif_orientation(self):
        template = aw.Document(MY_DIR + "Reporting engine template - Image exif orientation.docx")
        with open(IMAGE_DIR + "RightF.jpg", "rb") as source1:
            image1_bytes = source1.read()
        with open(IMAGE_DIR + "WrongF.jpg", "rb") as source2:
            image2_bytes = source2.read()
        self.build_report(template, [image1_bytes, image2_bytes], ["image1", "image2"], options=aw.reporting.ReportBuildOptions.RESPECT_JPEG_EXIF_ORIENTATION)
        template.save(ARTIFACTS_DIR + "ReportingEngine.ImageExifOrientation.docx")

    @unittest.skip("drawing.Image type isn't supported yet")
    def test_dynamic_stretching_image_within_text_box(self):

        template = aw.Document(MY_DIR + "Reporting engine template - Dynamic stretching.docx")

        image = ImageTestClass(image=drawing.Image.from_file(self.image, True))

        self.build_report(template, image, "src", options=aw.reporting.ReportBuildOptions.NONE)
        template.save(ARTIFACTS_DIR + "ReportingEngine.dynamic_stretching_image_within_text_box.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.dynamic_stretching_image_within_text_box.docx",
            GOLDS_DIR + "ReportingEngine.DynamicStretchingImageWithinTextBox Gold.docx"))

    def test_insert_hyperlinks_dynamically(self):
        links = [
            "https://auckland.dynabic.com/wiki/display/org/Supported+dynamic+insertion+of+hyperlinks+for+LINQ+Reporting+Engine",
            "Bookmark"
            ]
        for link in links:
            with self.subTest(link=link):
                template = aw.Document(MY_DIR + "Reporting engine template - Inserting hyperlinks.docx")
                self.build_report(template,
                    [
                        link, # Use URI or the name of a bookmark within the same document for a hyperlink
                        "Aspose"
                    ],
                    [
                        "uri_or_bookmark_expression",
                        "display_text_expression"
                    ])

                template.save(ARTIFACTS_DIR + "ReportingEngine.insert_hyperlinks_dynamically.docx")

    def test_insert_bookmarks_dynamically(self):

        doc = DocumentHelper.create_simple_document(
            "<<bookmark [bookmark_expression]>><<foreach [m in Contracts]>><<[m.client.Name]>><</foreach>><</bookmark>>")

        self.build_report(doc, ["BookmarkOne", self.create_json_data_source(Common.get_contracts())],
            ["bookmark_expression", "Contracts"])

        doc.save(ARTIFACTS_DIR + "ReportingEngine.insert_bookmarks_dynamically.docx")

    def test_without_known_type(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("<<[new DateTime()]:”dd.m_m.yyyy”>>")

        engine = aw.reporting.ReportingEngine()
        with self.assertRaises(Exception):
            engine.build_report(doc, "")

    def test_work_with_known_types(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("<<[new DateTime(2016, 1, 20)]:”dd.m_m.yyyy”>>")
        builder.writeln("<<[new DateTime(2016, 1, 20)]:”dd”>>")
        builder.writeln("<<[new DateTime(2016, 1, 20)]:”MM”>>")
        builder.writeln("<<[new DateTime(2016, 1, 20)]:”yyyy”>>")
        builder.writeln("<<[new DateTime(2016, 1, 20).month]>>")

        self.build_report(doc, "", known_types=[type(datetime)])

        doc.save(ARTIFACTS_DIR + "ReportingEngine.work_with_known_types.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "ReportingEngine.work_with_known_types.docx", GOLDS_DIR + "ReportingEngine.KnownTypes Gold.docx"))

    def test_work_with_content_controls(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - CheckBox Content Control.docx")
        self.build_report(doc, self.create_json_data_source(Common.get_managers()), "Managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.work_with_content_controls.docx")

    def test_work_with_single_column_table_row(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Table row.docx")
        self.build_report(doc, self.create_json_data_source(Common.get_managers()), "Managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.single_column_table_row.docx")

    def test_work_with_single_column_table_row_greedy(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Table row greedy.docx")
        self.build_report(doc, self.create_json_data_source(Common.get_managers()), "Managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.single_column_table_row_greedy.docx")

    def test_table_row_conditional_blocks(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Table row conditional blocks.docx")

        clients = [
            ClientTestClass(
                name="John Monrou",
                country="France",
                local_address="27 RUE PASTEUR"),
            ClientTestClass(
                name="James White",
                country="England",
                local_address="14 Tottenham Court Road"),
            ClientTestClass(
                name="Kate Otts",
                country="New Zealand",
                local_address="Wellington 6004"),
            ]
        
        self.build_report(doc, self.create_json_data_source(clients), "clients")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.table_row_conditional_blocks.docx")

    def test_if_greedy(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - If greedy.docx")

        obj = ExReportingEngine.AsposeData(["abc"])

        self.build_report(doc, obj)

        doc.save(ARTIFACTS_DIR + "ReportingEngine.if_greedy.docx")

    class AsposeData:

        def __init__(self, list: List[str]):
            self.List = list

    def test_stretch_imagefit_height(self):

        doc = DocumentHelper.create_template_document_with_draw_objects(
            "<<image [src.ImageStream] -fitHeight>>", aw.drawing.ShapeType.TEXT_BOX)

        with open(self.image, "rb") as stream:
            image_stream = ImageTestClass(image_stream=stream)
            self.build_report(doc, image_stream, "src", options=aw.reporting.ReportBuildOptions.NONE)

            doc = DocumentHelper.save_open(doc)

            shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)

            for shape in shapes:
                shape = shape.as_shape()

                # Assert that the image is really insert in textbox.
                self.assertIsNotNone(shape.fill.image_bytes)

                # Assert that the width is preserved, and the height is changed.
                self.assertNotEqual(346.35, shape.height)
                self.assertEqual(431.5, shape.width)

    def test_stretch_imagefit_width(self):

        doc = DocumentHelper.create_template_document_with_draw_objects(
            "<<image [src.ImageStream] -fitWidth>>", aw.drawing.ShapeType.TEXT_BOX)

        with open(self.image, "rb") as stream:
            image_stream = ImageTestClass(image_stream=stream)
            self.build_report(doc, image_stream, "src", options=aw.reporting.ReportBuildOptions.NONE)

            doc = DocumentHelper.save_open(doc)

            shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)

            for shape in shapes:
                shape = shape.as_shape()

                self.assertIsNotNone(shape.fill.image_bytes)

                # Assert that the height is preserved, and the width is changed
                self.assertNotEqual(431.5, shape.width)
                self.assertEqual(346.35, shape.height)

    def test_stretch_imagefit_size(self):

        doc = DocumentHelper.create_template_document_with_draw_objects(
            "<<image [src.ImageStream] -fitSize>>", aw.drawing.ShapeType.TEXT_BOX)

        with open(self.image, "rb") as stream:
            image_stream = ImageTestClass(image_stream=stream)
            self.build_report(doc, image_stream, "src", options=aw.reporting.ReportBuildOptions.NONE)

            doc = DocumentHelper.save_open(doc)

            shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)

            for shape in shapes:
                shape = shape.as_shape()

                self.assertNotNone(shape.fill.image_bytes)

                # Assert that the height and the width are changed.
                self.assertNotEqual(346.35, shape.height)
                self.assertNotEqual(431.5, shape.width)

    def test_stretch_imagefit_size_lim(self):

        doc = DocumentHelper.create_template_document_with_draw_objects(
            "<<image [src.ImageStream] -fitSizeLim>>", aw.drawing.ShapeType.TEXT_BOX)

        with open(self.image, "rb") as stream:
            image_stream = ImageTestClass(image_stream=stream)
            self.build_report(doc, image_stream, "src", options=aw.reporting.ReportBuildOptions.NONE)

            doc = DocumentHelper.save_open(doc)

            shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)

            for shape in shapes:
                shape = shape.as_shape()

                self.assertNotNone(shape.fill.image_bytes)

                # Assert that textbox size are equal image size.
                self.assertEqual(300.0, shape.height)
                self.assertEqual(300.0, shape.width)

    def test_without_missing_members(self):

        builder = aw.DocumentBuilder()

        # Add templete to the document for reporting engine
        DocumentHelper.insert_builder_text(builder,
            ["<<[missingObject.first().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>"])

        # Assert that build report failed without "ReportBuildOptions.ALLOW_MISSING_MEMBERS"
        with self.assertRaises(Exception):
            self.build_report(builder.document, DataSet(), "", options=aw.reporting.ReportBuildOptions.NONE)

    def test_with_missing_members(self):

        builder = aw.DocumentBuilder()

        # Add templete to the document for reporting engine
        DocumentHelper.insert_builder_text(builder,
            ["<<[missingObject.first().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>"])

        self.build_report(builder.document, self.create_json_data_source({}), "", options=aw.reporting.ReportBuildOptions.ALLOW_MISSING_MEMBERS)

        # Assert that build report success with "ReportBuildOptions.ALLOW_MISSING_MEMBERS"
        self.assertEqual(
            aw.ControlChar.PARAGRAPH_BREAK + aw.ControlChar.PARAGRAPH_BREAK + aw.ControlChar.SECTION_BREAK,
            builder.document.get_text())

    def test_inline_error_messages(self):

        parameters = [
            ("<<[missingObject.First().id]>>", "<<[missingObject.First( Error! Can not get the value of member 'missingObject' on type 'System.Collections.ArrayList'. ).id]>>", "Can not get the value of member"),
            ("<<[new DateTime()]:\"dd.m_m.yyyy\">>", "<<[new DateTime( Error! A type identifier is expected. )]:\"dd.m_m.yyyy\">>", "A type identifier is expected"),
            ("<<]>>", "<<] Error! Character ']' is unexpected. >>", "Character is unexpected"),
            ("<<[>>", "<<[>> Error! An expression is expected.", "An expression is expected"),
            ("<<>>", "<<>> Error! Tag end is unexpected.", "Tag end is unexpected"),
            ]

        for template_text, result, test_name in parameters:
            with self.subTest(test_name=test_name):
                builder = aw.DocumentBuilder()
                DocumentHelper.insert_builder_text(builder, [template_text])

                self.build_report(builder.document, self.create_json_data_source({}), "", options=aw.reporting.ReportBuildOptions.INLINE_ERROR_MESSAGES)

                self.assertEqual(
                    builder.document.first_section.body.paragraphs[0].get_text().rstrip(),
                    result)

    def test_set_background_color_dynamically(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Background color.docx")

        colors = [
            ColorItemTestClass("Black", drawing.Color.black),
            ColorItemTestClass("Red", drawing.Color.from_argb(255, 0, 0)),
            ColorItemTestClass("Empty", drawing.Color.empty()),
        ]

        self.build_report(doc, colors, "Colors")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.set_background_color_dynamically.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.set_background_color_dynamically.docx",
            GOLDS_DIR + "ReportingEngine.SetBackgroundColorDynamically Gold.docx"))
    
    def test_set_text_color_dynamically(self):
        doc = aw.Document(MY_DIR + "Reporting engine template - Text color.docx")

        colors = [
            ColorItemTestClass("Black", drawing.Color.black),
            ColorItemTestClass("Red", drawing.Color.from_argb(255, 0, 0)),
            ColorItemTestClass("Empty", drawing.Color.empty()),
        ]

        self.build_report(doc, colors, "Colors")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.set_text_color_dynamically.docx");

        self.assertTrue(DocumentHelper.CompareDocs(ARTIFACTS_DIR + "ReportingEngine.set_text_color_dynamically.docx",
            GOLDS_DIR + "ReportingEngine.SetTextColorDynamically Gold.docx"));

    def test_do_not_remove_empty_paragraphs(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Remove empty paragraphs.docx")

        self.build_report(doc, self.create_json_data_source(Common.get_managers()), "Managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.do_not_remove_empty_paragraphs.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.do_not_remove_empty_paragraphs.docx",
            GOLDS_DIR + "ReportingEngine.DoNotRemoveEmptyParagraphs Gold.docx"))

    def test_remove_empty_paragraphs(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Remove empty paragraphs.docx")

        self.build_report(doc, self.create_json_data_source(Common.get_managers()), "Managers", options=aw.reporting.ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS)

        doc.save(ARTIFACTS_DIR + "ReportingEngine.remove_empty_paragraphs.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.remove_empty_paragraphs.docx",
            GOLDS_DIR + "ReportingEngine.RemoveEmptyParagraphs Gold.docx"))

    def test_merging_table_cells_dynamically(self):

        parameters = [
            ("Hello", "Hello", "ReportingEngine.MergingTableCellsDynamically.Merged", "Cells in the first two tables must be merged"),
            ("Hello", "Name", "ReportingEngine.MergingTableCellsDynamically.NotMerged", "Only last table cells must be merge"),
            ]

        for value1, value2, result_document_name, test_name in parameters:
            with self.subTest(test_name=test_name):
                artifact_path = ARTIFACTS_DIR + result_document_name + aw.FileFormatUtil.save_format_to_extension(aw.SaveFormat.DOCX)
                gold_path = GOLDS_DIR + result_document_name + " Gold" + aw.FileFormatUtil.save_format_to_extension(aw.SaveFormat.DOCX)

                doc = aw.Document(MY_DIR + "Reporting engine template - Merging table cells dynamically.docx")

                clients = [
                    ClientTestClass(
                        name="John Monrou",
                        country="France",
                        local_address="27 RUE PASTEUR"),
                    ClientTestClass(
                        name="James White",
                        country="New Zealand",
                        local_address="14 Tottenham Court Road"),
                    ClientTestClass(
                        name="Kate Otts",
                        country="New Zealand",
                        local_address="Wellington 6004"),
                    ]

                self.build_report(doc, [value1, value2, self.create_json_data_source(clients)], ["value1", "value2", "clients"])
                doc.save(artifact_path)

                self.assertTrue(DocumentHelper.compare_docs(artifact_path, gold_path))

    def test_xml_data_string_without_schema(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - XML data destination.docx")

        data_source = aw.reporting.XmlDataSource(MY_DIR + "List of people.xml")
        self.build_report(doc, data_source, "persons")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.xml_data_string_without_schema.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.xml_data_string_without_schema.docx",
            GOLDS_DIR + "ReportingEngine.DataSource Gold.docx"))

    def test_xml_data_stream_without_schema(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - XML data destination.docx")

        with open(MY_DIR + "List of people.xml", "rb") as stream:
            data_source = aw.reporting.XmlDataSource(stream)
            self.build_report(doc, data_source, "persons")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.xml_data_stream_without_schema.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.xml_data_stream_without_schema.docx",
            GOLDS_DIR + "ReportingEngine.DataSource Gold.docx"))

    def test_xml_data_with_nested_elements(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Data destination with nested elements.docx")

        data_source = aw.reporting.XmlDataSource(MY_DIR + "Nested elements.xml")
        self.build_report(doc, data_source, "managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.xml_data_with_nested_elements.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.xml_data_with_nested_elements.docx",
            GOLDS_DIR + "ReportingEngine.DataSourceWithNestedElements Gold.docx"))

    def test_json_data_string(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - JSON data destination.docx")

        options = aw.reporting.JsonDataLoadOptions()
        options.exact_date_time_parse_formats = ["MM/dd/yyyy", "MM.d.yy", "MM d yy"]

        data_source = aw.reporting.JsonDataSource(MY_DIR + "List of people.json", options)
        self.build_report(doc, data_source, "persons")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.json_data_string.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.json_data_string.docx",
            GOLDS_DIR + "ReportingEngine.JsonDataString Gold.docx"))

    def test_json_data_string_exception(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - JSON data destination.docx")

        options = aw.reporting.JsonDataLoadOptions()
        options.simple_value_parse_mode = aw.reporting.JsonSimpleValueParseMode.STRICT

        data_source = aw.reporting.JsonDataSource(MY_DIR + "List of people.json", options)
        with self.assertRaises(Exception):
            self.build_report(doc, data_source, "persons")

    def test_json_data_stream(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - JSON data destination.docx")

        options = aw.reporting.JsonDataLoadOptions()
        options.exact_date_time_parse_formats = ["MM/dd/yyyy", "MM.d.yy", "MM d yy"]

        with open(MY_DIR + "List of people.json", "rb") as stream:
            data_source = aw.reporting.JsonDataSource(stream, options)
            self.build_report(doc, data_source, "persons")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.json_data_stream.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.json_data_stream.docx",
            GOLDS_DIR + "ReportingEngine.JsonDataString Gold.docx"))

    def test_json_data_with_nested_elements(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - Data destination with nested elements.docx")

        data_source = aw.reporting.JsonDataSource(MY_DIR + "Nested elements.json")
        self.build_report(doc, data_source, "managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.json_data_with_nested_elements.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.json_data_with_nested_elements.docx",
            GOLDS_DIR + "ReportingEngine.DataSourceWithNestedElements Gold.docx"))

    def test_csv_data_string(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - CSV data destination.docx")

        load_options = aw.reporting.CsvDataLoadOptions(True)
        load_options.delimiter = ';'
        load_options.comment_char = '$'

        data_source = aw.reporting.CsvDataSource(MY_DIR + "List of people.csv", load_options)
        self.build_report(doc, data_source, "persons")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.csv_data_string.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.csv_data_string.docx",
            GOLDS_DIR + "ReportingEngine.CsvData Gold.docx"))

    def test_csv_data_stream(self):

        doc = aw.Document(MY_DIR + "Reporting engine template - CSV data destination.docx")

        load_options = aw.reporting.CsvDataLoadOptions(True)
        load_options.delimiter = ';'
        load_options.comment_char = '$'

        with open(MY_DIR + "List of people.csv", "rb") as stream:
            data_source = aw.reporting.CsvDataSource(stream, load_options)
            self.build_report(doc, data_source, "persons")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.csv_data_stream.docx")

        self.assertTrue(DocumentHelper.compare_docs(
            ARTIFACTS_DIR + "ReportingEngine.csv_data_stream.docx",
            GOLDS_DIR + "ReportingEngine.CsvData Gold.docx"))

    def test_insert_combobox_dropdown_list_items_dynamically(self):

        for sdt_type in (aw.markup.SdtType.COMBO_BOX, aw.markup.SdtType.DROP_DOWN_LIST):
            with self.subTest(sdt_type=sdt_type):
                template = "<<item[\"three\"] [\"3\"]>><<if [false]>><<item [\"four\"] [null]>><</if>><<item[\"five\"] [\"5\"]>>"

                static_items = [
                    aw.markup.SdtListItem("1", "one"),
                    aw.markup.SdtListItem("2", "two"),
                    ]

                doc = aw.Document()

                sdt = aw.markup.StructuredDocumentTag(doc, sdt_type, aw.markup.MarkupLevel.BLOCK)
                sdt.title = template

                for item in static_items:
                    sdt.list_items.add(item)

                doc.first_section.body.append_child(sdt)

                self.build_report(doc, self.create_json_data_source({}), "")

                doc.save(ARTIFACTS_DIR + "ReportingEngine.insert_combobox_dropdown_list_items_dynamically_{}.docx".format(sdt_type))

                doc = aw.Document(ARTIFACTS_DIR + "ReportingEngine.insert_combobox_dropdown_list_items_dynamically_{}.docx".format(sdt_type))

                sdt = doc.get_child(aw.NodeType.STRUCTURED_DOCUMENT_TAG, 0, True).as_structured_document_tag()

                expected_items = [
                    aw.markup.SdtListItem("1", "one"),
                    aw.markup.SdtListItem("2", "two"),
                    aw.markup.SdtListItem("3", "three"),
                    aw.markup.SdtListItem("5", "five"),
                    ]

                self.assertEqual(len(expected_items), sdt.list_items.count)

                for i, expected_item in enumerate(expected_items):
                    self.assertEqual(expected_item.value, sdt.list_items[i].value)
                    self.assertEqual(expected_item.display_text, sdt.list_items[i].display_text)

    def build_report(self, document: aw.Document, data_source, data_source_name = None,
                     known_types = None, options: Optional[aw.reporting.ReportBuildOptions] = None):

        engine = aw.reporting.ReportingEngine()

        if options is not None:
            engine.options = options

        if known_types is not None:
            for known_type in known_types:
                engine.known_types.add(known_type)

        if data_source_name is not None:
            engine.build_report(document, data_source, data_source_name)
        else:
            engine.build_report(document, data_source)

    def create_json_data_source(self, obj: object) -> aw.reporting.JsonDataSource:
        def default(item):
            if isinstance(item, datetime):
                return item.isoformat()
            return vars(item)

        json_data = json.dumps(obj, default=default)
        with io.BytesIO(json_data.encode('utf-8')) as stream:
            return aw.reporting.JsonDataSource(stream)
