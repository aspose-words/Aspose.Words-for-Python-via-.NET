# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from document_helper import DocumentHelper
from datetime import timedelta, timezone
from enum import Enum
import glob
import aspose.pydrawing
import aspose.words as aw
import aspose.words.digitalsignatures
import aspose.words.drawing
import aspose.words.drawing.charts
import aspose.words.fields
import aspose.words.lists
import aspose.words.notes
import aspose.words.tables
import datetime
import document_helper
import io
import system_helper
import test_util
import unittest
import uuid
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, GOLDS_DIR, IMAGE_DIR, IMAGE_URL, MY_DIR

class ExDocumentBuilder(ApiExampleBase):

    def test_write_and_font(self):
        #ExStart
        #ExFor:Font.size
        #ExFor:Font.bold
        #ExFor:Font.name
        #ExFor:Font.color
        #ExFor:Font.underline
        #ExFor:DocumentBuilder.__init__
        #ExSummary:Shows how to insert formatted text using DocumentBuilder.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Specify font formatting, then add text.
        font = builder.font
        font.size = 16
        font.bold = True
        font.color = aspose.pydrawing.Color.blue
        font.name = 'Courier New'
        font.underline = aw.Underline.DASH
        builder.write('Hello world!')
        #ExEnd
        doc = document_helper.DocumentHelper.save_open(builder.document)
        first_run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('Hello world!', first_run.get_text().strip())
        self.assertEqual(16, first_run.font.size)
        self.assertTrue(first_run.font.bold)
        self.assertEqual('Courier New', first_run.font.name)
        self.assertEqual(aspose.pydrawing.Color.blue.to_argb(), first_run.font.color.to_argb())
        self.assertEqual(aw.Underline.DASH, first_run.font.underline)

    def test_headers_and_footers(self):
        #ExStart
        #ExFor:DocumentBuilder
        #ExFor:DocumentBuilder.__init__(Document)
        #ExFor:DocumentBuilder.move_to_header_footer
        #ExFor:DocumentBuilder.move_to_section
        #ExFor:DocumentBuilder.insert_break
        #ExFor:DocumentBuilder.writeln
        #ExFor:HeaderFooterType
        #ExFor:PageSetup.different_first_page_header_footer
        #ExFor:PageSetup.odd_and_even_pages_header_footer
        #ExFor:BreakType
        #ExSummary:Shows how to create headers and footers in a document using DocumentBuilder.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Specify that we want different headers and footers for first, even and odd pages.
        builder.page_setup.different_first_page_header_footer = True
        builder.page_setup.odd_and_even_pages_header_footer = True
        # Create the headers, then add three pages to the document to display each header type.
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_FIRST)
        builder.write('Header for the first page')
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_EVEN)
        builder.write('Header for even pages')
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.write('Header for all other pages')
        builder.move_to_section(0)
        builder.writeln('Page1')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln('Page2')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln('Page3')
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.HeadersAndFooters.docx')
        #ExEnd
        headers_footers = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.HeadersAndFooters.docx').first_section.headers_footers
        self.assertEqual(3, headers_footers.count)
        self.assertEqual('Header for the first page', headers_footers.get_by_header_footer_type(aw.HeaderFooterType.HEADER_FIRST).get_text().strip())
        self.assertEqual('Header for even pages', headers_footers.get_by_header_footer_type(aw.HeaderFooterType.HEADER_EVEN).get_text().strip())
        self.assertEqual('Header for all other pages', headers_footers.get_by_header_footer_type(aw.HeaderFooterType.HEADER_PRIMARY).get_text().strip())

    def test_merge_fields(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_field(str)
        #ExFor:DocumentBuilder.move_to_merge_field(str,bool,bool)
        #ExSummary:Shows how to insert fields, and move the document builder's cursor to them.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.insert_field(field_code='MERGEFIELD MyMergeField1 \\* MERGEFORMAT')
        builder.insert_field(field_code='MERGEFIELD MyMergeField2 \\* MERGEFORMAT')
        # Move the cursor to the first MERGEFIELD.
        builder.move_to_merge_field(field_name='MyMergeField1', is_after=True, is_delete_field=False)
        # Note that the cursor is placed immediately after the first MERGEFIELD, and before the second.
        self.assertEqual(doc.range.fields[1].start, builder.current_node)
        self.assertEqual(doc.range.fields[0].end, builder.current_node.previous_sibling)
        # If we wish to edit the field's field code or contents using the builder,
        # its cursor would need to be inside a field.
        # To place it inside a field, we would need to call the document builder's MoveTo method
        # and pass the field's start or separator node as an argument.
        builder.write(' Text between our merge fields. ')
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.MergeFields.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.MergeFields.docx')
        self.assertEqual('\x13MERGEFIELD MyMergeField1 \\* MERGEFORMAT\x14«MyMergeField1»\x15' + ' Text between our merge fields. ' + '\x13MERGEFIELD MyMergeField2 \\* MERGEFORMAT\x14«MyMergeField2»\x15', doc.get_text().strip())
        self.assertEqual(2, doc.range.fields.count)
        test_util.TestUtil.verify_field(expected_type=aw.fields.FieldType.FIELD_MERGE_FIELD, expected_field_code='MERGEFIELD MyMergeField1 \\* MERGEFORMAT', expected_result='«MyMergeField1»', field=doc.range.fields[0])
        test_util.TestUtil.verify_field(expected_type=aw.fields.FieldType.FIELD_MERGE_FIELD, expected_field_code='MERGEFIELD MyMergeField2 \\* MERGEFORMAT', expected_result='«MyMergeField2»', field=doc.range.fields[1])

    def test_insert_horizontal_rule(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_horizontal_rule
        #ExFor:ShapeBase.is_horizontal_rule
        #ExFor:Shape.horizontal_rule_format
        #ExFor:HorizontalRuleAlignment
        #ExFor:HorizontalRuleFormat
        #ExFor:HorizontalRuleFormat.alignment
        #ExFor:HorizontalRuleFormat.width_percent
        #ExFor:HorizontalRuleFormat.height
        #ExFor:HorizontalRuleFormat.color
        #ExFor:HorizontalRuleFormat.no_shade
        #ExSummary:Shows how to insert a horizontal rule shape, and customize its formatting.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shape = builder.insert_horizontal_rule()
        horizontal_rule_format = shape.horizontal_rule_format
        horizontal_rule_format.alignment = aw.drawing.HorizontalRuleAlignment.CENTER
        horizontal_rule_format.width_percent = 70
        horizontal_rule_format.height = 3
        horizontal_rule_format.color = aspose.pydrawing.Color.blue
        horizontal_rule_format.no_shade = True
        self.assertTrue(shape.is_horizontal_rule)
        self.assertTrue(shape.horizontal_rule_format.no_shade)
        #ExEnd
        doc = document_helper.DocumentHelper.save_open(doc)
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual(aw.drawing.HorizontalRuleAlignment.CENTER, shape.horizontal_rule_format.alignment)
        self.assertEqual(70, shape.horizontal_rule_format.width_percent)
        self.assertEqual(3, shape.horizontal_rule_format.height)
        self.assertEqual(aspose.pydrawing.Color.blue.to_argb(), shape.horizontal_rule_format.color.to_argb())

    def test_insert_hyperlink(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_hyperlink
        #ExFor:Font.clear_formatting
        #ExFor:Font.color
        #ExFor:Font.underline
        #ExFor:Underline
        #ExSummary:Shows how to insert a hyperlink field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.write('For more information, please visit the ')
        # Insert a hyperlink and emphasize it with custom formatting.
        # The hyperlink will be a clickable piece of text which will take us to the location specified in the URL.
        builder.font.color = aspose.pydrawing.Color.blue
        builder.font.underline = aw.Underline.SINGLE
        builder.insert_hyperlink('Google website', 'https://www.google.com', False)
        builder.font.clear_formatting()
        builder.writeln('.')
        # Ctrl + left clicking the link in the text in Microsoft Word will take us to the URL via a new web browser window.
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertHyperlink.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertHyperlink.docx')
        hyperlink = doc.range.fields[0].as_field_hyperlink()
        self.assertEqual('https://www.google.com', hyperlink.address)
        field_contents = hyperlink.start.next_sibling.as_run()
        self.assertEqual(aspose.pydrawing.Color.blue.to_argb(), field_contents.font.color.to_argb())
        self.assertEqual(aw.Underline.SINGLE, field_contents.font.underline)
        self.assertEqual('HYPERLINK "https://www.google.com"', field_contents.get_text().strip())

    def test_push_pop_font(self):
        #ExStart
        #ExFor:DocumentBuilder.push_font
        #ExFor:DocumentBuilder.pop_font
        #ExFor:DocumentBuilder.insert_hyperlink
        #ExSummary:Shows how to use a document builder's formatting stack.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Set up font formatting, then write the text that goes before the hyperlink.
        builder.font.name = 'Arial'
        builder.font.size = 24
        builder.write('To visit Google, hold Ctrl and click ')
        # Preserve our current formatting configuration on the stack.
        builder.push_font()
        # Alter the builder's current formatting by applying a new style.
        builder.font.style_identifier = aw.StyleIdentifier.HYPERLINK
        builder.insert_hyperlink('here', 'http://www.google.com', False)
        self.assertEqual(aspose.pydrawing.Color.blue.to_argb(), builder.font.color.to_argb())
        self.assertEqual(aw.Underline.SINGLE, builder.font.underline)
        # Restore the font formatting that we saved earlier and remove the element from the stack.
        builder.pop_font()
        self.assertEqual(aspose.pydrawing.Color.empty().to_argb(), builder.font.color.to_argb())
        self.assertEqual(aw.Underline.NONE, builder.font.underline)
        builder.write('. We hope you enjoyed the example.')
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.PushPopFont.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.PushPopFont.docx')
        runs = doc.first_section.body.first_paragraph.runs
        self.assertEqual(4, runs.count)
        self.assertEqual('To visit Google, hold Ctrl and click', runs[0].get_text().strip())
        self.assertEqual('. We hope you enjoyed the example.', runs[3].get_text().strip())
        self.assertEqual(runs[0].font.color, runs[3].font.color)
        self.assertEqual(runs[0].font.underline, runs[3].font.underline)
        self.assertEqual('here', runs[2].get_text().strip())
        self.assertEqual(aspose.pydrawing.Color.blue.to_argb(), runs[2].font.color.to_argb())
        self.assertEqual(aw.Underline.SINGLE, runs[2].font.underline)
        self.assertNotEqual(runs[0].font.color, runs[2].font.color)
        self.assertNotEqual(runs[0].font.underline, runs[2].font.underline)
        self.assertEqual('http://www.google.com', doc.range.fields[0].as_field_hyperlink().address)

    @unittest.skip("drawing.Image type isn't supported yet")
    def test_insert_watermark(self):
        #ExStart
        #ExFor:DocumentBuilder.move_to_header_footer
        #ExFor:PageSetup.page_width
        #ExFor:PageSetup.page_height
        #ExFor:WrapType
        #ExFor:RelativeHorizontalPosition
        #ExFor:RelativeVerticalPosition
        #ExSummary:Shows how to insert an image, and use it as a watermark.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert the image into the header so that it will be visible on every page.
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        shape = builder.insert_image(file_name=IMAGE_DIR + 'Transparent background logo.png')
        shape.wrap_type = aw.drawing.WrapType.NONE
        shape.behind_text = True
        # Place the image at the center of the page.
        shape.relative_horizontal_position = aw.drawing.RelativeHorizontalPosition.PAGE
        shape.relative_vertical_position = aw.drawing.RelativeVerticalPosition.PAGE
        shape.left = (builder.page_setup.page_width - shape.width) / 2
        shape.top = (builder.page_setup.page_height - shape.height) / 2
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertWatermark.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertWatermark.docx')
        shape = doc.first_section.headers_footers.get_by_header_footer_type(aw.HeaderFooterType.HEADER_PRIMARY).get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.PNG, shape)
        self.assertEqual(aw.drawing.WrapType.NONE, shape.wrap_type)
        self.assertTrue(shape.behind_text)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.PAGE, shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PAGE, shape.relative_vertical_position)
        self.assertEqual((doc.first_section.page_setup.page_width - shape.width) / 2, shape.left)
        self.assertEqual((doc.first_section.page_setup.page_height - shape.height) / 2, shape.top)

    def test_insert_ole_object(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_ole_object(str,bool,bool,BytesIO)
        #ExFor:DocumentBuilder.insert_ole_object(str,str,bool,bool,BytesIO)
        #ExFor:DocumentBuilder.insert_ole_object_as_icon(str,bool,str,str)
        #ExSummary:Shows how to insert an OLE object into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # OLE objects are links to files in our local file system that can be opened by other installed applications.
        # Double clicking these shapes will launch the application, and then use it to open the linked object.
        # There are three ways of using the InsertOleObject method to insert these shapes and configure their appearance.
        # 1 -  Image taken from the local file system:
        with system_helper.io.FileStream(IMAGE_DIR + 'Logo.jpg', system_helper.io.FileMode.OPEN) as image_stream:
            # If 'presentation' is omitted and 'asIcon' is set, this overloaded method selects
            # the icon according to the file extension and uses the filename for the icon caption.
            builder.insert_ole_object(file_name=MY_DIR + 'Spreadsheet.xlsx', is_linked=False, as_icon=False, presentation=image_stream)
        # If 'presentation' is omitted and 'asIcon' is set, this overloaded method selects
        # the icon according to 'progId' and uses the filename for the icon caption.
        # 2 -  Icon based on the application that will open the object:
        builder.insert_ole_object(file_name=MY_DIR + 'Spreadsheet.xlsx', prog_id='Excel.Sheet', is_linked=False, as_icon=True, presentation=None)
        # If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
        # the icon according to 'progId' and uses the predefined icon caption.
        # 3 -  Image icon that's 32 x 32 pixels or smaller from the local file system, with a custom caption:
        builder.insert_ole_object_as_icon(file_name=MY_DIR + 'Presentation.pptx', is_linked=False, icon_file=IMAGE_DIR + 'Logo icon.ico', icon_caption='Double click to view presentation!')
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertOleObject.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertOleObject.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual(aw.drawing.ShapeType.OLE_OBJECT, shape.shape_type)
        self.assertEqual('Excel.Sheet.12', shape.ole_format.prog_id)
        self.assertEqual('.xlsx', shape.ole_format.suggested_extension)
        shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()
        self.assertEqual(aw.drawing.ShapeType.OLE_OBJECT, shape.shape_type)
        self.assertEqual('Package', shape.ole_format.prog_id)
        self.assertEqual('.xlsx', shape.ole_format.suggested_extension)
        shape = doc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()
        self.assertEqual(aw.drawing.ShapeType.OLE_OBJECT, shape.shape_type)
        self.assertEqual('PowerPoint.Show.12', shape.ole_format.prog_id)
        self.assertEqual('.pptx', shape.ole_format.suggested_extension)

    def test_insert_html(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_html(str)
        #ExSummary:Shows how to use a document builder to insert html content into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        html = "<p align='right'>Paragraph right</p>" + '<b>Implicit paragraph left</b>' + "<div align='center'>Div center</div>" + "<h1 align='left'>Heading 1 left.</h1>"
        builder.insert_html(html=html)
        # Inserting HTML code parses the formatting of each element into equivalent document text formatting.
        paragraphs = doc.first_section.body.paragraphs
        self.assertEqual('Paragraph right', paragraphs[0].get_text().strip())
        self.assertEqual(aw.ParagraphAlignment.RIGHT, paragraphs[0].paragraph_format.alignment)
        self.assertEqual('Implicit paragraph left', paragraphs[1].get_text().strip())
        self.assertEqual(aw.ParagraphAlignment.LEFT, paragraphs[1].paragraph_format.alignment)
        self.assertTrue(paragraphs[1].runs[0].font.bold)
        self.assertEqual('Div center', paragraphs[2].get_text().strip())
        self.assertEqual(aw.ParagraphAlignment.CENTER, paragraphs[2].paragraph_format.alignment)
        self.assertEqual('Heading 1 left.', paragraphs[3].get_text().strip())
        self.assertEqual('Heading 1', paragraphs[3].paragraph_format.style.name)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertHtml.docx')
        #ExEnd

    def test_math_ml(self):
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        math_ml = '<math xmlns="http://www.w3.org/1998/Math/MathML"><mrow><msub><mi>a</mi><mrow><mn>1</mn></mrow></msub><mo>+</mo><msub><mi>b</mi><mrow><mn>1</mn></mrow></msub></mrow></math>'
        builder.insert_html(html=math_ml)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.MathML.docx')
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.MathML.pdf')
        self.assertTrue(document_helper.DocumentHelper.compare_docs(GOLDS_DIR + 'DocumentBuilder.MathML Gold.docx', ARTIFACTS_DIR + 'DocumentBuilder.MathML.docx'))

    def test_insert_text_and_bookmark(self):
        #ExStart
        #ExFor:DocumentBuilder.start_bookmark
        #ExFor:DocumentBuilder.end_bookmark
        #ExSummary:Shows how create a bookmark.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # A valid bookmark needs to have document body text enclosed by
        # BookmarkStart and BookmarkEnd nodes created with a matching bookmark name.
        builder.start_bookmark('MyBookmark')
        builder.writeln('Hello world!')
        builder.end_bookmark('MyBookmark')
        self.assertEqual(1, doc.range.bookmarks.count)
        self.assertEqual('MyBookmark', doc.range.bookmarks[0].name)
        self.assertEqual('Hello world!', doc.range.bookmarks[0].text.strip())
        #ExEnd

    def test_create_form(self):
        #ExStart
        #ExFor:TextFormFieldType
        #ExFor:DocumentBuilder.insert_text_input
        #ExFor:DocumentBuilder.insert_combo_box
        #ExSummary:Shows how to create form fields.
        builder = aw.DocumentBuilder()
        # Form fields are objects in the document that the user can interact with by being prompted to enter values.
        # We can create them using a document builder, and below are two ways of doing so.
        # 1 -  Basic text input:
        builder.insert_text_input('My text input', aw.fields.TextFormFieldType.REGULAR, '', 'Enter your name here', 30)
        # 2 -  Combo box with prompt text, and a range of possible values:
        items = ['-- Select your favorite footwear --', 'Sneakers', 'Oxfords', 'Flip-flops', 'Other']
        builder.insert_paragraph()
        builder.insert_combo_box('My combo box', items, 0)
        builder.document.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.CreateForm.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.CreateForm.docx')
        form_field = doc.range.form_fields[0]
        self.assertEqual('My text input', form_field.name)
        self.assertEqual(aw.fields.TextFormFieldType.REGULAR, form_field.text_input_type)
        self.assertEqual('Enter your name here', form_field.result)
        form_field = doc.range.form_fields[1]
        self.assertEqual('My combo box', form_field.name)
        self.assertEqual(aw.fields.TextFormFieldType.REGULAR, form_field.text_input_type)
        self.assertEqual('-- Select your favorite footwear --', form_field.result)
        self.assertEqual(0, form_field.drop_down_selected_index)
        self.assertSequenceEqual(['-- Select your favorite footwear --', 'Sneakers', 'Oxfords', 'Flip-flops', 'Other'], list(form_field.drop_down_items))

    def test_insert_check_box(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_check_box(str,bool,bool,int)
        #ExFor:DocumentBuilder.insert_check_box(str,bool,int)
        #ExSummary:Shows how to insert checkboxes into the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert checkboxes of varying sizes and default checked statuses.
        builder.write('Unchecked check box of a default size: ')
        builder.insert_check_box(name='', default_value=False, checked_value=False, size=0)
        builder.insert_paragraph()
        builder.write('Large checked check box: ')
        builder.insert_check_box(name='CheckBox_Default', default_value=True, checked_value=True, size=50)
        builder.insert_paragraph()
        # Form fields have a name length limit of 20 characters.
        builder.write('Very large checked check box: ')
        builder.insert_check_box(name='CheckBox_OnlyCheckedValue', checked_value=True, size=100)
        self.assertEqual('CheckBox_OnlyChecked', doc.range.form_fields[2].name)
        # We can interact with these check boxes in Microsoft Word by double clicking them.
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertCheckBox.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertCheckBox.docx')
        form_fields = doc.range.form_fields
        self.assertEqual('', form_fields[0].name)
        self.assertEqual(False, form_fields[0].checked)
        self.assertEqual(False, form_fields[0].default)
        self.assertEqual(10, form_fields[0].check_box_size)
        self.assertEqual('CheckBox_Default', form_fields[1].name)
        self.assertEqual(True, form_fields[1].checked)
        self.assertEqual(True, form_fields[1].default)
        self.assertEqual(50, form_fields[1].check_box_size)
        self.assertEqual('CheckBox_OnlyChecked', form_fields[2].name)
        self.assertEqual(True, form_fields[2].checked)
        self.assertEqual(True, form_fields[2].default)
        self.assertEqual(100, form_fields[2].check_box_size)

    def test_insert_check_box_empty_name(self):
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Checking that the checkbox insertion with an empty name working correctly
        builder.insert_check_box(name='', default_value=True, checked_value=False, size=1)
        builder.insert_check_box(name='', checked_value=False, size=1)

    def test_working_with_nodes(self):
        #ExStart
        #ExFor:DocumentBuilder.move_to(Node)
        #ExFor:DocumentBuilder.move_to_bookmark(str)
        #ExFor:DocumentBuilder.current_paragraph
        #ExFor:DocumentBuilder.current_node
        #ExFor:DocumentBuilder.move_to_document_start
        #ExFor:DocumentBuilder.move_to_document_end
        #ExFor:DocumentBuilder.is_at_end_of_paragraph
        #ExFor:DocumentBuilder.is_at_start_of_paragraph
        #ExSummary:Shows how to move a document builder's cursor to different nodes in a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Create a valid bookmark, an entity that consists of nodes enclosed by a bookmark start node,
        # and a bookmark end node.
        builder.start_bookmark('MyBookmark')
        builder.write('Bookmark contents.')
        builder.end_bookmark('MyBookmark')
        first_paragraph_nodes = doc.first_section.body.first_paragraph.get_child_nodes(aw.NodeType.ANY, False)
        self.assertEqual(aw.NodeType.BOOKMARK_START, first_paragraph_nodes[0].node_type)
        self.assertEqual(aw.NodeType.RUN, first_paragraph_nodes[1].node_type)
        self.assertEqual('Bookmark contents.', first_paragraph_nodes[1].get_text().strip())
        self.assertEqual(aw.NodeType.BOOKMARK_END, first_paragraph_nodes[2].node_type)
        # The document builder's cursor is always ahead of the node that we last added with it.
        # If the builder's cursor is at the end of the document, its current node will be null.
        # The previous node is the bookmark end node that we last added.
        # Adding new nodes with the builder will append them to the last node.
        self.assertIsNone(builder.current_node)
        # If we wish to edit a different part of the document with the builder,
        # we will need to bring its cursor to the node we wish to edit.
        builder.move_to_bookmark(bookmark_name='MyBookmark')
        # Moving it to a bookmark will move it to the first node within the bookmark start and end nodes, the enclosed run.
        self.assertEqual(first_paragraph_nodes[1], builder.current_node)
        # We can also move the cursor to an individual node like this.
        builder.move_to(doc.first_section.body.first_paragraph.get_child_nodes(aw.NodeType.ANY, False)[0])
        self.assertEqual(aw.NodeType.BOOKMARK_START, builder.current_node.node_type)
        self.assertEqual(doc.first_section.body.first_paragraph, builder.current_paragraph)
        self.assertTrue(builder.is_at_start_of_paragraph)
        # We can use specific methods to move to the start/end of a document.
        builder.move_to_document_end()
        self.assertTrue(builder.is_at_end_of_paragraph)
        builder.move_to_document_start()
        self.assertTrue(builder.is_at_start_of_paragraph)
        #ExEnd

    def test_fill_merge_fields(self):
        #ExStart
        #ExFor:DocumentBuilder.move_to_merge_field(str)
        #ExFor:DocumentBuilder.bold
        #ExFor:DocumentBuilder.italic
        #ExSummary:Shows how to fill MERGEFIELDs with data with a document builder instead of a mail merge.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert some MERGEFIELDS, which accept data from columns of the same name in a data source during a mail merge,
        # and then fill them manually.
        builder.insert_field(field_code=' MERGEFIELD Chairman ')
        builder.insert_field(field_code=' MERGEFIELD ChiefFinancialOfficer ')
        builder.insert_field(field_code=' MERGEFIELD ChiefTechnologyOfficer ')
        builder.move_to_merge_field(field_name='Chairman')
        builder.bold = True
        builder.writeln('John Doe')
        builder.move_to_merge_field(field_name='ChiefFinancialOfficer')
        builder.italic = True
        builder.writeln('Jane Doe')
        builder.move_to_merge_field(field_name='ChiefTechnologyOfficer')
        builder.italic = True
        builder.writeln('John Bloggs')
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.FillMergeFields.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.FillMergeFields.docx')
        paragraphs = doc.first_section.body.paragraphs
        self.assertTrue(paragraphs[0].runs[0].font.bold)
        self.assertEqual('John Doe', paragraphs[0].runs[0].get_text().strip())
        self.assertTrue(paragraphs[1].runs[0].font.italic)
        self.assertEqual('Jane Doe', paragraphs[1].runs[0].get_text().strip())
        self.assertTrue(paragraphs[2].runs[0].font.italic)
        self.assertEqual('John Bloggs', paragraphs[2].runs[0].get_text().strip())

    def test_insert_toc(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_table_of_contents
        #ExFor:Document.update_fields
        #ExFor:DocumentBuilder.__init__(Document)
        #ExFor:ParagraphFormat.style_identifier
        #ExFor:DocumentBuilder.insert_break
        #ExFor:BreakType
        #ExSummary:Shows how to insert a Table of contents (TOC) into a document using heading styles as entries.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a table of contents for the first page of the document.
        # Configure the table to pick up paragraphs with headings of levels 1 to 3.
        # Also, set its entries to be hyperlinks that will take us
        # to the location of the heading when left-clicked in Microsoft Word.
        builder.insert_table_of_contents('\\o "1-3" \\h \\z \\u')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        # Populate the table of contents by adding paragraphs with heading styles.
        # Each such heading with a level between 1 and 3 will create an entry in the table.
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
        builder.writeln('Heading 1')
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING2
        builder.writeln('Heading 1.1')
        builder.writeln('Heading 1.2')
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
        builder.writeln('Heading 2')
        builder.writeln('Heading 3')
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING2
        builder.writeln('Heading 3.1')
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING3
        builder.writeln('Heading 3.1.1')
        builder.writeln('Heading 3.1.2')
        builder.writeln('Heading 3.1.3')
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING4
        builder.writeln('Heading 3.1.3.1')
        builder.writeln('Heading 3.1.3.2')
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING2
        builder.writeln('Heading 3.2')
        builder.writeln('Heading 3.3')
        # A table of contents is a field of a type that needs to be updated to show an up-to-date result.
        doc.update_fields()
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertToc.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertToc.docx')
        table_of_contents = doc.range.fields[0].as_field_toc()
        self.assertEqual('1-3', table_of_contents.heading_level_range)
        self.assertTrue(table_of_contents.insert_hyperlinks)
        self.assertTrue(table_of_contents.hide_in_web_layout)
        self.assertTrue(table_of_contents.use_paragraph_outline_level)

    def test_insert_table(self):
        #ExStart
        #ExFor:DocumentBuilder
        #ExFor:DocumentBuilder.write
        #ExFor:DocumentBuilder.start_table
        #ExFor:DocumentBuilder.insert_cell
        #ExFor:DocumentBuilder.end_row
        #ExFor:DocumentBuilder.end_table
        #ExFor:DocumentBuilder.cell_format
        #ExFor:DocumentBuilder.row_format
        #ExFor:CellFormat
        #ExFor:CellFormat.fit_text
        #ExFor:CellFormat.width
        #ExFor:CellFormat.vertical_alignment
        #ExFor:CellFormat.shading
        #ExFor:CellFormat.orientation
        #ExFor:CellFormat.wrap_text
        #ExFor:RowFormat
        #ExFor:RowFormat.borders
        #ExFor:RowFormat.clear_formatting
        #ExFor:Shading.clear_formatting
        #ExSummary:Shows how to build a table with custom borders.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.start_table()
        # Setting table formatting options for a document builder
        # will apply them to every row and cell that we add with it.
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        builder.cell_format.clear_formatting()
        builder.cell_format.width = 150
        builder.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.CENTER
        builder.cell_format.shading.background_pattern_color = aspose.pydrawing.Color.green_yellow
        builder.cell_format.wrap_text = False
        builder.cell_format.fit_text = True
        builder.row_format.clear_formatting()
        builder.row_format.height_rule = aw.HeightRule.EXACTLY
        builder.row_format.height = 50
        builder.row_format.borders.line_style = aw.LineStyle.ENGRAVE_3D
        builder.row_format.borders.color = aspose.pydrawing.Color.orange
        builder.insert_cell()
        builder.write('Row 1, Col 1')
        builder.insert_cell()
        builder.write('Row 1, Col 2')
        builder.end_row()
        # Changing the formatting will apply it to the current cell,
        # and any new cells that we create with the builder afterward.
        # This will not affect the cells that we have added previously.
        builder.cell_format.shading.clear_formatting()
        builder.insert_cell()
        builder.write('Row 2, Col 1')
        builder.insert_cell()
        builder.write('Row 2, Col 2')
        builder.end_row()
        # Increase row height to fit the vertical text.
        builder.insert_cell()
        builder.row_format.height = 150
        builder.cell_format.orientation = aw.TextOrientation.UPWARD
        builder.write('Row 3, Col 1')
        builder.insert_cell()
        builder.cell_format.orientation = aw.TextOrientation.DOWNWARD
        builder.write('Row 3, Col 2')
        builder.end_row()
        builder.end_table()
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertTable.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertTable.docx')
        table = doc.first_section.body.tables[0]
        self.assertEqual('Row 1, Col 1\x07', table.rows[0].cells[0].get_text().strip())
        self.assertEqual('Row 1, Col 2\x07', table.rows[0].cells[1].get_text().strip())
        self.assertEqual(aw.HeightRule.EXACTLY, table.rows[0].row_format.height_rule)
        self.assertEqual(50, table.rows[0].row_format.height)
        self.assertEqual(aw.LineStyle.ENGRAVE_3D, table.rows[0].row_format.borders.line_style)
        self.assertEqual(aspose.pydrawing.Color.orange.to_argb(), table.rows[0].row_format.borders.color.to_argb())
        for c in table.rows[0].cells:
            c = c.as_cell()
            self.assertEqual(150, c.cell_format.width)
            self.assertEqual(aw.tables.CellVerticalAlignment.CENTER, c.cell_format.vertical_alignment)
            self.assertEqual(aspose.pydrawing.Color.green_yellow.to_argb(), c.cell_format.shading.background_pattern_color.to_argb())
            self.assertFalse(c.cell_format.wrap_text)
            self.assertTrue(c.cell_format.fit_text)
            self.assertEqual(aw.ParagraphAlignment.CENTER, c.first_paragraph.paragraph_format.alignment)
        self.assertEqual('Row 2, Col 1\x07', table.rows[1].cells[0].get_text().strip())
        self.assertEqual('Row 2, Col 2\x07', table.rows[1].cells[1].get_text().strip())
        for c in table.rows[1].cells:
            c = c.as_cell()
            self.assertEqual(150, c.cell_format.width)
            self.assertEqual(aw.tables.CellVerticalAlignment.CENTER, c.cell_format.vertical_alignment)
            self.assertEqual(aspose.pydrawing.Color.empty().to_argb(), c.cell_format.shading.background_pattern_color.to_argb())
            self.assertFalse(c.cell_format.wrap_text)
            self.assertTrue(c.cell_format.fit_text)
            self.assertEqual(aw.ParagraphAlignment.CENTER, c.first_paragraph.paragraph_format.alignment)
        self.assertEqual(150, table.rows[2].row_format.height)
        self.assertEqual('Row 3, Col 1\x07', table.rows[2].cells[0].get_text().strip())
        self.assertEqual(aw.TextOrientation.UPWARD, table.rows[2].cells[0].cell_format.orientation)
        self.assertEqual(aw.ParagraphAlignment.CENTER, table.rows[2].cells[0].first_paragraph.paragraph_format.alignment)
        self.assertEqual('Row 3, Col 2\x07', table.rows[2].cells[1].get_text().strip())
        self.assertEqual(aw.TextOrientation.DOWNWARD, table.rows[2].cells[1].cell_format.orientation)
        self.assertEqual(aw.ParagraphAlignment.CENTER, table.rows[2].cells[1].first_paragraph.paragraph_format.alignment)

    def test_insert_table_with_style(self):
        #ExStart
        #ExFor:Table.style_identifier
        #ExFor:Table.style_options
        #ExFor:TableStyleOptions
        #ExFor:Table.auto_fit
        #ExFor:AutoFitBehavior
        #ExSummary:Shows how to build a new table while applying a style.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        table = builder.start_table()
        # We must insert at least one row before setting any table formatting.
        builder.insert_cell()
        # Set the table style used based on the style identifier.
        # Note that not all table styles are available when saving to .doc format.
        table.style_identifier = aw.StyleIdentifier.MEDIUM_SHADING1_ACCENT1
        # Partially apply the style to features of the table based on predicates, then build the table.
        table.style_options = aw.tables.TableStyleOptions.FIRST_COLUMN | aw.tables.TableStyleOptions.ROW_BANDS | aw.tables.TableStyleOptions.FIRST_ROW
        table.auto_fit(aw.tables.AutoFitBehavior.AUTO_FIT_TO_CONTENTS)
        builder.writeln('Item')
        builder.cell_format.right_padding = 40
        builder.insert_cell()
        builder.writeln('Quantity (kg)')
        builder.end_row()
        builder.insert_cell()
        builder.writeln('Apples')
        builder.insert_cell()
        builder.writeln('20')
        builder.end_row()
        builder.insert_cell()
        builder.writeln('Bananas')
        builder.insert_cell()
        builder.writeln('40')
        builder.end_row()
        builder.insert_cell()
        builder.writeln('Carrots')
        builder.insert_cell()
        builder.writeln('50')
        builder.end_row()
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertTableWithStyle.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertTableWithStyle.docx')
        doc.expand_table_styles_to_direct_formatting()
        self.assertEqual('Medium Shading 1 Accent 1', table.style.name)
        self.assertEqual(aw.tables.TableStyleOptions.FIRST_COLUMN | aw.tables.TableStyleOptions.ROW_BANDS | aw.tables.TableStyleOptions.FIRST_ROW, table.style_options)
        self.assertEqual(189, table.first_row.first_cell.cell_format.shading.background_pattern_color.b)
        self.assertEqual(aspose.pydrawing.Color.white.to_argb(), table.first_row.first_cell.first_paragraph.runs[0].font.color.to_argb())
        self.assertNotEqual(aspose.pydrawing.Color.light_blue.to_argb(), table.last_row.first_cell.cell_format.shading.background_pattern_color.b)
        self.assertEqual(aspose.pydrawing.Color.empty().to_argb(), table.last_row.first_cell.first_paragraph.runs[0].font.color.to_argb())

    def test_insert_table_set_heading_row(self):
        #ExStart
        #ExFor:RowFormat.heading_format
        #ExSummary:Shows how to build a table with rows that repeat on every page.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        table = builder.start_table()
        # Any rows inserted while the "HeadingFormat" flag is set to "true"
        # will show up at the top of the table on every page that it spans.
        builder.row_format.heading_format = True
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        builder.cell_format.width = 100
        builder.insert_cell()
        builder.write('Heading row 1')
        builder.end_row()
        builder.insert_cell()
        builder.write('Heading row 2')
        builder.end_row()
        builder.cell_format.width = 50
        builder.paragraph_format.clear_formatting()
        builder.row_format.heading_format = False
        # Add enough rows for the table to span two pages.
        i = 0
        while i < 50:
            builder.insert_cell()
            builder.write(f'Row {table.rows.count}, column 1.')
            builder.insert_cell()
            builder.write(f'Row {table.rows.count}, column 2.')
            builder.end_row()
            i += 1
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertTableSetHeadingRow.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertTableSetHeadingRow.docx')
        table = doc.first_section.body.tables[0]
        i = 0
        while i < table.rows.count:
            self.assertEqual(i < 2, table.rows[i].row_format.heading_format)
            i += 1

    def test_insert_table_with_preferred_width(self):
        #ExStart
        #ExFor:Table.preferred_width
        #ExFor:PreferredWidth.from_percent
        #ExFor:PreferredWidth
        #ExSummary:Shows how to set a table to auto fit to 50% of the width of the page.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        table = builder.start_table()
        builder.insert_cell()
        builder.write('Cell #1')
        builder.insert_cell()
        builder.write('Cell #2')
        builder.insert_cell()
        builder.write('Cell #3')
        table.preferred_width = aw.tables.PreferredWidth.from_percent(50)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertTableWithPreferredWidth.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertTableWithPreferredWidth.docx')
        table = doc.first_section.body.tables[0]
        self.assertEqual(aw.tables.PreferredWidthType.PERCENT, table.preferred_width.type)
        self.assertEqual(50, table.preferred_width.value)

    def test_insert_cells_with_preferred_widths(self):
        #ExStart
        #ExFor:CellFormat.preferred_width
        #ExFor:PreferredWidth
        #ExFor:PreferredWidth.auto
        #ExFor:PreferredWidth.__eq__(PreferredWidth)
        #ExFor:PreferredWidth.__eq__(object)
        #ExFor:PreferredWidth.from_points
        #ExFor:PreferredWidth.from_percent
        #ExFor:PreferredWidth.__hash__
        #ExFor:PreferredWidth.__str__
        #ExSummary:Shows how to set a preferred width for table cells.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        table = builder.start_table()
        # There are two ways of applying the "PreferredWidth" class to table cells.
        # 1 -  Set an absolute preferred width based on points:
        builder.insert_cell()
        builder.cell_format.preferred_width = aw.tables.PreferredWidth.from_points(40)
        builder.cell_format.shading.background_pattern_color = aspose.pydrawing.Color.light_yellow
        builder.writeln(f'Cell with a width of {builder.cell_format.preferred_width}.')
        # 2 -  Set a relative preferred width based on percent of the table's width:
        builder.insert_cell()
        builder.cell_format.preferred_width = aw.tables.PreferredWidth.from_percent(20)
        builder.cell_format.shading.background_pattern_color = aspose.pydrawing.Color.light_blue
        builder.writeln(f'Cell with a width of {builder.cell_format.preferred_width}.')
        builder.insert_cell()
        # A cell with no preferred width specified will take up the rest of the available space.
        builder.cell_format.preferred_width = aw.tables.PreferredWidth.AUTO
        # Each configuration of the "PreferredWidth" property creates a new object.
        self.assertNotEqual(hash(table.first_row.cells[1].cell_format.preferred_width), hash(builder.cell_format.preferred_width))
        builder.cell_format.shading.background_pattern_color = aspose.pydrawing.Color.light_green
        builder.writeln('Automatically sized cell.')
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertCellsWithPreferredWidths.docx')
        #ExEnd
        self.assertEqual(100, aw.tables.PreferredWidth.from_percent(100).value)
        self.assertEqual(100, aw.tables.PreferredWidth.from_points(100).value)
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertCellsWithPreferredWidths.docx')
        table = doc.first_section.body.tables[0]
        self.assertEqual(aw.tables.PreferredWidthType.POINTS, table.first_row.cells[0].cell_format.preferred_width.type)
        self.assertEqual(40, table.first_row.cells[0].cell_format.preferred_width.value)
        self.assertEqual('Cell with a width of 800.\r\x07', table.first_row.cells[0].get_text().strip())
        self.assertEqual(aw.tables.PreferredWidthType.PERCENT, table.first_row.cells[1].cell_format.preferred_width.type)
        self.assertEqual(20, table.first_row.cells[1].cell_format.preferred_width.value)
        self.assertEqual('Cell with a width of 20%.\r\x07', table.first_row.cells[1].get_text().strip())
        self.assertEqual(aw.tables.PreferredWidthType.AUTO, table.first_row.cells[2].cell_format.preferred_width.type)
        self.assertEqual(0, table.first_row.cells[2].cell_format.preferred_width.value)
        self.assertEqual('Automatically sized cell.\r\x07', table.first_row.cells[2].get_text().strip())

    def test_insert_table_from_html(self):
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert the table from HTML. Note that AutoFitSettings does not apply to tables
        # inserted from HTML.
        builder.insert_html(html='<table>' + '<tr>' + '<td>Row 1, Cell 1</td>' + '<td>Row 1, Cell 2</td>' + '</tr>' + '<tr>' + '<td>Row 2, Cell 2</td>' + '<td>Row 2, Cell 2</td>' + '</tr>' + '</table>')
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertTableFromHtml.docx')
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertTableFromHtml.docx')
        self.assertEqual(1, doc.get_child_nodes(aw.NodeType.TABLE, True).count)
        self.assertEqual(2, doc.get_child_nodes(aw.NodeType.ROW, True).count)
        self.assertEqual(4, doc.get_child_nodes(aw.NodeType.CELL, True).count)

    def test_insert_nested_table(self):
        #ExStart
        #ExFor:Cell.first_paragraph
        #ExSummary:Shows how to create a nested table using a document builder.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Build the outer table.
        cell = builder.insert_cell()
        builder.writeln('Outer Table Cell 1')
        builder.insert_cell()
        builder.writeln('Outer Table Cell 2')
        builder.end_table()
        # Move to the first cell of the outer table, the build another table inside the cell.
        builder.move_to(cell.first_paragraph)
        builder.insert_cell()
        builder.writeln('Inner Table Cell 1')
        builder.insert_cell()
        builder.writeln('Inner Table Cell 2')
        builder.end_table()
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertNestedTable.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertNestedTable.docx')
        self.assertEqual(2, doc.get_child_nodes(aw.NodeType.TABLE, True).count)
        self.assertEqual(4, doc.get_child_nodes(aw.NodeType.CELL, True).count)
        self.assertEqual(1, cell.tables[0].count)
        self.assertEqual(2, cell.tables[0].first_row.cells.count)

    def test_create_table(self):
        #ExStart
        #ExFor:DocumentBuilder
        #ExFor:DocumentBuilder.write
        #ExFor:DocumentBuilder.insert_cell
        #ExSummary:Shows how to use a document builder to create a table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Start the table, then populate the first row with two cells.
        builder.start_table()
        builder.insert_cell()
        builder.write('Row 1, Cell 1.')
        builder.insert_cell()
        builder.write('Row 1, Cell 2.')
        # Call the builder's "EndRow" method to start a new row.
        builder.end_row()
        builder.insert_cell()
        builder.write('Row 2, Cell 1.')
        builder.insert_cell()
        builder.write('Row 2, Cell 2.')
        builder.end_table()
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.CreateTable.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.CreateTable.docx')
        table = doc.first_section.body.tables[0]
        self.assertEqual(4, table.get_child_nodes(aw.NodeType.CELL, True).count)
        self.assertEqual('Row 1, Cell 1.\x07', table.rows[0].cells[0].get_text().strip())
        self.assertEqual('Row 1, Cell 2.\x07', table.rows[0].cells[1].get_text().strip())
        self.assertEqual('Row 2, Cell 1.\x07', table.rows[1].cells[0].get_text().strip())
        self.assertEqual('Row 2, Cell 2.\x07', table.rows[1].cells[1].get_text().strip())

    def test_build_formatted_table(self):
        #ExStart
        #ExFor:RowFormat.height
        #ExFor:RowFormat.height_rule
        #ExFor:Table.left_indent
        #ExFor:DocumentBuilder.paragraph_format
        #ExFor:DocumentBuilder.font
        #ExSummary:Shows how to create a formatted table using DocumentBuilder.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        table = builder.start_table()
        builder.insert_cell()
        table.left_indent = 20
        # Set some formatting options for text and table appearance.
        builder.row_format.height = 40
        builder.row_format.height_rule = aw.HeightRule.AT_LEAST
        builder.cell_format.shading.background_pattern_color = aspose.pydrawing.Color.from_argb(198, 217, 241)
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        builder.font.size = 16
        builder.font.name = 'Arial'
        builder.font.bold = True
        # Configuring the formatting options in a document builder will apply them
        # to the current cell/row its cursor is in,
        # as well as any new cells and rows created using that builder.
        builder.write('Header Row,\n Cell 1')
        builder.insert_cell()
        builder.write('Header Row,\n Cell 2')
        builder.insert_cell()
        builder.write('Header Row,\n Cell 3')
        builder.end_row()
        # Reconfigure the builder's formatting objects for new rows and cells that we are about to make.
        # The builder will not apply these to the first row already created so that it will stand out as a header row.
        builder.cell_format.shading.background_pattern_color = aspose.pydrawing.Color.white
        builder.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.CENTER
        builder.row_format.height = 30
        builder.row_format.height_rule = aw.HeightRule.AUTO
        builder.insert_cell()
        builder.font.size = 12
        builder.font.bold = False
        builder.write('Row 1, Cell 1.')
        builder.insert_cell()
        builder.write('Row 1, Cell 2.')
        builder.insert_cell()
        builder.write('Row 1, Cell 3.')
        builder.end_row()
        builder.insert_cell()
        builder.write('Row 2, Cell 1.')
        builder.insert_cell()
        builder.write('Row 2, Cell 2.')
        builder.insert_cell()
        builder.write('Row 2, Cell 3.')
        builder.end_row()
        builder.end_table()
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.CreateFormattedTable.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.CreateFormattedTable.docx')
        table = doc.first_section.body.tables[0]
        self.assertEqual(20, table.left_indent)
        self.assertEqual(aw.HeightRule.AT_LEAST, table.rows[0].row_format.height_rule)
        self.assertEqual(40, table.rows[0].row_format.height)
        for c in doc.get_child_nodes(aw.NodeType.CELL, True):
            c = c.as_cell()
            self.assertEqual(aw.ParagraphAlignment.CENTER, c.first_paragraph.paragraph_format.alignment)
            for r in c.first_paragraph.runs:
                r = r.as_run()
                self.assertEqual('Arial', r.font.name)
                if c.parent_row == table.first_row:
                    self.assertEqual(16, r.font.size)
                    self.assertTrue(r.font.bold)
                else:
                    self.assertEqual(12, r.font.size)
                    self.assertFalse(r.font.bold)

    def test_table_borders_and_shading(self):
        #ExStart
        #ExFor:Shading
        #ExFor:Table.set_borders
        #ExFor:BorderCollection.left
        #ExFor:BorderCollection.right
        #ExFor:BorderCollection.top
        #ExFor:BorderCollection.bottom
        #ExSummary:Shows how to apply border and shading color while building a table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Start a table and set a default color/thickness for its borders.
        table = builder.start_table()
        table.set_borders(aw.LineStyle.SINGLE, 2, aspose.pydrawing.Color.black)
        # Create a row with two cells with different background colors.
        builder.insert_cell()
        builder.cell_format.shading.background_pattern_color = aspose.pydrawing.Color.light_sky_blue
        builder.writeln('Row 1, Cell 1.')
        builder.insert_cell()
        builder.cell_format.shading.background_pattern_color = aspose.pydrawing.Color.orange
        builder.writeln('Row 1, Cell 2.')
        builder.end_row()
        # Reset cell formatting to disable the background colors
        # set a custom border thickness for all new cells created by the builder,
        # then build a second row.
        builder.cell_format.clear_formatting()
        builder.cell_format.borders.left.line_width = 4
        builder.cell_format.borders.right.line_width = 4
        builder.cell_format.borders.top.line_width = 4
        builder.cell_format.borders.bottom.line_width = 4
        builder.insert_cell()
        builder.writeln('Row 2, Cell 1.')
        builder.insert_cell()
        builder.writeln('Row 2, Cell 2.')
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.TableBordersAndShading.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.TableBordersAndShading.docx')
        table = doc.first_section.body.tables[0]
        for c in table.first_row:
            c = c.as_cell()
            self.assertEqual(0.5, c.cell_format.borders.top.line_width)
            self.assertEqual(0.5, c.cell_format.borders.bottom.line_width)
            self.assertEqual(0.5, c.cell_format.borders.left.line_width)
            self.assertEqual(0.5, c.cell_format.borders.right.line_width)
            self.assertEqual(aspose.pydrawing.Color.empty().to_argb(), c.cell_format.borders.left.color.to_argb())
            self.assertEqual(aw.LineStyle.SINGLE, c.cell_format.borders.left.line_style)
        self.assertEqual(aspose.pydrawing.Color.light_sky_blue.to_argb(), table.first_row.first_cell.cell_format.shading.background_pattern_color.to_argb())
        self.assertEqual(aspose.pydrawing.Color.orange.to_argb(), table.first_row.cells[1].cell_format.shading.background_pattern_color.to_argb())
        for c in table.last_row:
            c = c.as_cell()
            self.assertEqual(4, c.cell_format.borders.top.line_width)
            self.assertEqual(4, c.cell_format.borders.bottom.line_width)
            self.assertEqual(4, c.cell_format.borders.left.line_width)
            self.assertEqual(4, c.cell_format.borders.right.line_width)
            self.assertEqual(aspose.pydrawing.Color.empty().to_argb(), c.cell_format.borders.left.color.to_argb())
            self.assertEqual(aw.LineStyle.SINGLE, c.cell_format.borders.left.line_style)
            self.assertEqual(aspose.pydrawing.Color.empty().to_argb(), c.cell_format.shading.background_pattern_color.to_argb())

    def test_set_preferred_type_convert_util(self):
        #ExStart
        #ExFor:PreferredWidth.from_points
        #ExSummary:Shows how to use unit conversion tools while specifying a preferred width for a cell.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        table = builder.start_table()
        builder.cell_format.preferred_width = aw.tables.PreferredWidth.from_points(aw.ConvertUtil.inch_to_point(3))
        builder.insert_cell()
        self.assertEqual(216, table.first_row.first_cell.cell_format.preferred_width.value)
        #ExEnd

    def test_insert_hyperlink_to_local_bookmark(self):
        #ExStart
        #ExFor:DocumentBuilder.start_bookmark
        #ExFor:DocumentBuilder.end_bookmark
        #ExFor:DocumentBuilder.insert_hyperlink
        #ExSummary:Shows how to insert a hyperlink which references a local bookmark.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.start_bookmark('Bookmark1')
        builder.write('Bookmarked text. ')
        builder.end_bookmark('Bookmark1')
        builder.writeln('Text outside of the bookmark.')
        # Insert a HYPERLINK field that links to the bookmark. We can pass field switches
        # to the "InsertHyperlink" method as part of the argument containing the referenced bookmark's name.
        builder.font.color = aspose.pydrawing.Color.blue
        builder.font.underline = aw.Underline.SINGLE
        hyperlink = builder.insert_hyperlink('Link to Bookmark1', 'Bookmark1', True).as_field_hyperlink()
        hyperlink.screen_tip = 'Hyperlink Tip'
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertHyperlinkToLocalBookmark.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertHyperlinkToLocalBookmark.docx')
        hyperlink = doc.range.fields[0].as_field_hyperlink()
        test_util.TestUtil.verify_field(expected_type=aw.fields.FieldType.FIELD_HYPERLINK, expected_field_code=' HYPERLINK \\l "Bookmark1" \\o "Hyperlink Tip" ', expected_result='Link to Bookmark1', field=hyperlink)
        self.assertEqual('Bookmark1', hyperlink.sub_address)
        self.assertEqual('Hyperlink Tip', hyperlink.screen_tip)
        self.assertTrue(any([b.name == 'Bookmark1' for b in doc.range.bookmarks]))

    def test_cursor_position(self):
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.write('Hello world!')
        # If the builder's cursor is at the end of the document,
        # there will be no nodes in front of it so that the current node will be null.
        self.assertIsNone(builder.current_node)
        self.assertEqual('Hello world!', builder.current_paragraph.get_text().strip())
        # Move to the beginning of the document and place the cursor at an existing node.
        builder.move_to_document_start()
        self.assertEqual(aw.NodeType.RUN, builder.current_node.node_type)

    def test_move_to(self):
        #ExStart
        #ExFor:Story.last_paragraph
        #ExFor:DocumentBuilder.move_to(Node)
        #ExSummary:Shows how to move a DocumentBuilder's cursor position to a specified node.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Run 1. ')
        # The document builder has a cursor, which acts as the part of the document
        # where the builder appends new nodes when we use its document construction methods.
        # This cursor functions in the same way as Microsoft Word's blinking cursor,
        # and it also always ends up immediately after any node that the builder just inserted.
        # To append content to a different part of the document,
        # we can move the cursor to a different node with the "MoveTo" method.
        self.assertEqual(doc.first_section.body.last_paragraph, builder.current_paragraph)  #ExSkip
        builder.move_to(doc.first_section.body.first_paragraph.runs[0])
        self.assertEqual(doc.first_section.body.first_paragraph, builder.current_paragraph)  #ExSkip
        # The cursor is now in front of the node that we moved it to.
        # Adding a second run will insert it in front of the first run.
        builder.writeln('Run 2. ')
        self.assertEqual('Run 2. \rRun 1.', doc.get_text().strip())
        # Move the cursor to the end of the document to continue appending text to the end as before.
        builder.move_to(doc.last_section.body.last_paragraph)
        builder.writeln('Run 3. ')
        self.assertEqual('Run 2. \rRun 1. \rRun 3.', doc.get_text().strip())
        self.assertEqual(doc.first_section.body.last_paragraph, builder.current_paragraph)  #ExSkip
        #ExEnd

    def test_move_to_paragraph(self):
        #ExStart
        #ExFor:DocumentBuilder.move_to_paragraph
        #ExSummary:Shows how to move a builder's cursor position to a specified paragraph.
        doc = aw.Document(file_name=MY_DIR + 'Paragraphs.docx')
        paragraphs = doc.first_section.body.paragraphs
        self.assertEqual(22, paragraphs.count)
        # Create document builder to edit the document. The builder's cursor,
        # which is the point where it will insert new nodes when we call its document construction methods,
        # is currently at the beginning of the document.
        builder = aw.DocumentBuilder(doc=doc)
        self.assertEqual(0, paragraphs.index_of(builder.current_paragraph))
        # Move that cursor to a different paragraph will place that cursor in front of that paragraph.
        builder.move_to_paragraph(2, 0)
        self.assertEqual(2, paragraphs.index_of(builder.current_paragraph))  #ExSkip
        # Any new content that we add will be inserted at that point.
        builder.writeln('This is a new third paragraph. ')
        #ExEnd
        self.assertEqual(3, paragraphs.index_of(builder.current_paragraph))
        doc = document_helper.DocumentHelper.save_open(doc)
        self.assertEqual('This is a new third paragraph.', doc.first_section.body.paragraphs[2].get_text().strip())

    def test_move_to_cell(self):
        #ExStart
        #ExFor:DocumentBuilder.move_to_cell
        #ExSummary:Shows how to move a document builder's cursor to a cell in a table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Create an empty 2x2 table.
        builder.start_table()
        builder.insert_cell()
        builder.insert_cell()
        builder.end_row()
        builder.insert_cell()
        builder.insert_cell()
        builder.end_table()
        # Because we have ended the table with the EndTable method,
        # the document builder's cursor is currently outside the table.
        # This cursor has the same function as Microsoft Word's blinking text cursor.
        # It can also be moved to a different location in the document using the builder's MoveTo methods.
        # We can move the cursor back inside the table to a specific cell.
        builder.move_to_cell(0, 1, 1, 0)
        builder.write('Column 2, cell 2.')
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.MoveToCell.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.MoveToCell.docx')
        table = doc.first_section.body.tables[0]
        self.assertEqual('Column 2, cell 2.\x07', table.rows[1].cells[1].get_text().strip())

    def test_move_to_bookmark(self):
        #ExStart
        #ExFor:DocumentBuilder.move_to_bookmark(str,bool,bool)
        #ExSummary:Shows how to move a document builder's node insertion point cursor to a bookmark.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # A valid bookmark consists of a BookmarkStart node, a BookmarkEnd node with a
        # matching bookmark name somewhere afterward, and contents enclosed by those nodes.
        builder.start_bookmark('MyBookmark')
        builder.write('Hello world! ')
        builder.end_bookmark('MyBookmark')
        # There are 4 ways of moving a document builder's cursor to a bookmark.
        # If we are between the BookmarkStart and BookmarkEnd nodes, the cursor will be inside the bookmark.
        # This means that any text added by the builder will become a part of the bookmark.
        # 1 -  Outside of the bookmark, in front of the BookmarkStart node:
        self.assertTrue(builder.move_to_bookmark(bookmark_name='MyBookmark', is_start=True, is_after=False))
        builder.write('1. ')
        self.assertEqual('Hello world! ', doc.range.bookmarks.get_by_name('MyBookmark').text)
        self.assertEqual('1. Hello world!', doc.get_text().strip())
        # 2 -  Inside the bookmark, right after the BookmarkStart node:
        self.assertTrue(builder.move_to_bookmark(bookmark_name='MyBookmark', is_start=True, is_after=True))
        builder.write('2. ')
        self.assertEqual('2. Hello world! ', doc.range.bookmarks.get_by_name('MyBookmark').text)
        self.assertEqual('1. 2. Hello world!', doc.get_text().strip())
        # 2 -  Inside the bookmark, right in front of the BookmarkEnd node:
        self.assertTrue(builder.move_to_bookmark(bookmark_name='MyBookmark', is_start=False, is_after=False))
        builder.write('3. ')
        self.assertEqual('2. Hello world! 3. ', doc.range.bookmarks.get_by_name('MyBookmark').text)
        self.assertEqual('1. 2. Hello world! 3.', doc.get_text().strip())
        # 4 -  Outside of the bookmark, after the BookmarkEnd node:
        self.assertTrue(builder.move_to_bookmark(bookmark_name='MyBookmark', is_start=False, is_after=True))
        builder.write('4.')
        self.assertEqual('2. Hello world! 3. ', doc.range.bookmarks.get_by_name('MyBookmark').text)
        self.assertEqual('1. 2. Hello world! 3. 4.', doc.get_text().strip())
        #ExEnd

    def test_build_table(self):
        #ExStart
        #ExFor:Table
        #ExFor:DocumentBuilder.start_table
        #ExFor:DocumentBuilder.end_row
        #ExFor:DocumentBuilder.end_table
        #ExFor:DocumentBuilder.cell_format
        #ExFor:DocumentBuilder.row_format
        #ExFor:DocumentBuilder.write(str)
        #ExFor:DocumentBuilder.writeln(str)
        #ExFor:CellVerticalAlignment
        #ExFor:CellFormat.orientation
        #ExFor:TextOrientation
        #ExFor:AutoFitBehavior
        #ExSummary:Shows how to build a formatted 2x2 table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        table = builder.start_table()
        builder.insert_cell()
        builder.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.CENTER
        builder.write('Row 1, cell 1.')
        builder.insert_cell()
        builder.write('Row 1, cell 2.')
        builder.end_row()
        # While building the table, the document builder will apply its current RowFormat/CellFormat property values
        # to the current row/cell that its cursor is in and any new rows/cells as it creates them.
        self.assertEqual(aw.tables.CellVerticalAlignment.CENTER, table.rows[0].cells[0].cell_format.vertical_alignment)
        self.assertEqual(aw.tables.CellVerticalAlignment.CENTER, table.rows[0].cells[1].cell_format.vertical_alignment)
        builder.insert_cell()
        builder.row_format.height = 100
        builder.row_format.height_rule = aw.HeightRule.EXACTLY
        builder.cell_format.orientation = aw.TextOrientation.UPWARD
        builder.write('Row 2, cell 1.')
        builder.insert_cell()
        builder.cell_format.orientation = aw.TextOrientation.DOWNWARD
        builder.write('Row 2, cell 2.')
        builder.end_row()
        builder.end_table()
        # Previously added rows and cells are not retroactively affected by changes to the builder's formatting.
        self.assertEqual(0, table.rows[0].row_format.height)
        self.assertEqual(aw.HeightRule.AUTO, table.rows[0].row_format.height_rule)
        self.assertEqual(100, table.rows[1].row_format.height)
        self.assertEqual(aw.HeightRule.EXACTLY, table.rows[1].row_format.height_rule)
        self.assertEqual(aw.TextOrientation.UPWARD, table.rows[1].cells[0].cell_format.orientation)
        self.assertEqual(aw.TextOrientation.DOWNWARD, table.rows[1].cells[1].cell_format.orientation)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.BuildTable.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.BuildTable.docx')
        table = doc.first_section.body.tables[0]
        self.assertEqual(2, table.rows.count)
        self.assertEqual(2, table.rows[0].cells.count)
        self.assertEqual(2, table.rows[1].cells.count)
        self.assertEqual(0, table.rows[0].row_format.height)
        self.assertEqual(aw.HeightRule.AUTO, table.rows[0].row_format.height_rule)
        self.assertEqual(100, table.rows[1].row_format.height)
        self.assertEqual(aw.HeightRule.EXACTLY, table.rows[1].row_format.height_rule)
        self.assertEqual('Row 1, cell 1.\x07', table.rows[0].cells[0].get_text().strip())
        self.assertEqual(aw.tables.CellVerticalAlignment.CENTER, table.rows[0].cells[0].cell_format.vertical_alignment)
        self.assertEqual('Row 1, cell 2.\x07', table.rows[0].cells[1].get_text().strip())
        self.assertEqual('Row 2, cell 1.\x07', table.rows[1].cells[0].get_text().strip())
        self.assertEqual(aw.TextOrientation.UPWARD, table.rows[1].cells[0].cell_format.orientation)
        self.assertEqual('Row 2, cell 2.\x07', table.rows[1].cells[1].get_text().strip())
        self.assertEqual(aw.TextOrientation.DOWNWARD, table.rows[1].cells[1].cell_format.orientation)

    def test_table_cell_vertical_rotated_far_east_text_orientation(self):
        doc = aw.Document(file_name=MY_DIR + 'Rotated cell text.docx')
        table = doc.first_section.body.tables[0]
        cell = table.first_row.first_cell
        self.assertEqual(aw.TextOrientation.VERTICAL_ROTATED_FAR_EAST, cell.cell_format.orientation)
        doc = document_helper.DocumentHelper.save_open(doc)
        table = doc.first_section.body.tables[0]
        cell = table.first_row.first_cell
        self.assertEqual(aw.TextOrientation.VERTICAL_ROTATED_FAR_EAST, cell.cell_format.orientation)

    def test_insert_floating_image(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_image(str,RelativeHorizontalPosition,float,RelativeVerticalPosition,float,float,float,WrapType)
        #ExSummary:Shows how to insert an image.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # There are two ways of using a document builder to source an image and then insert it as a floating shape.
        # 1 -  From a file in the local file system:
        builder.insert_image(file_name=IMAGE_DIR + 'Transparent background logo.png', horz_pos=aw.drawing.RelativeHorizontalPosition.MARGIN, left=100, vert_pos=aw.drawing.RelativeVerticalPosition.MARGIN, top=0, width=200, height=200, wrap_type=aw.drawing.WrapType.SQUARE)
        # 2 -  From a URL:
        builder.insert_image(file_name=IMAGE_URL, horz_pos=aw.drawing.RelativeHorizontalPosition.MARGIN, left=100, vert_pos=aw.drawing.RelativeVerticalPosition.MARGIN, top=250, width=200, height=200, wrap_type=aw.drawing.WrapType.SQUARE)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertFloatingImage.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertFloatingImage.docx')
        image = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.PNG, image)
        self.assertEqual(100, image.left)
        self.assertEqual(0, image.top)
        self.assertEqual(200, image.width)
        self.assertEqual(200, image.height)
        self.assertEqual(aw.drawing.WrapType.SQUARE, image.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, image.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, image.relative_vertical_position)
        image = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()
        test_util.TestUtil.verify_image_in_shape(272, 92, aw.drawing.ImageType.PNG, image)
        self.assertEqual(100, image.left)
        self.assertEqual(250, image.top)
        self.assertEqual(200, image.width)
        self.assertEqual(200, image.height)
        self.assertEqual(aw.drawing.WrapType.SQUARE, image.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, image.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, image.relative_vertical_position)

    def test_insert_image_original_size(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_image(str,RelativeHorizontalPosition,float,RelativeVerticalPosition,float,float,float,WrapType)
        #ExSummary:Shows how to insert an image from the local file system into a document while preserving its dimensions.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # The InsertImage method creates a floating shape with the passed image in its image data.
        # We can specify the dimensions of the shape can be passing them to this method.
        image_shape = builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg', horz_pos=aw.drawing.RelativeHorizontalPosition.MARGIN, left=0, vert_pos=aw.drawing.RelativeVerticalPosition.MARGIN, top=0, width=-1, height=-1, wrap_type=aw.drawing.WrapType.SQUARE)
        # Passing negative values as the intended dimensions will automatically define
        # the shape's dimensions based on the dimensions of its image.
        self.assertEqual(300, image_shape.width)
        self.assertEqual(300, image_shape.height)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertImageOriginalSize.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertImageOriginalSize.docx')
        image_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(0, image_shape.left)
        self.assertEqual(0, image_shape.top)
        self.assertEqual(300, image_shape.width)
        self.assertEqual(300, image_shape.height)
        self.assertEqual(aw.drawing.WrapType.SQUARE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, image_shape.relative_vertical_position)

    def test_insert_text_input(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_text_input
        #ExSummary:Shows how to insert a text input form field into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a form that prompts the user to enter text.
        builder.insert_text_input('TextInput', aw.fields.TextFormFieldType.REGULAR, '', 'Enter your text here', 0)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertTextInput.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertTextInput.docx')
        form_field = doc.range.form_fields[0]
        self.assertTrue(form_field.enabled)
        self.assertEqual('TextInput', form_field.name)
        self.assertEqual(0, form_field.max_length)
        self.assertEqual('Enter your text here', form_field.result)
        self.assertEqual(aw.fields.FieldType.FIELD_FORM_TEXT_INPUT, form_field.type)
        self.assertEqual('', form_field.text_input_format)
        self.assertEqual(aw.fields.TextFormFieldType.REGULAR, form_field.text_input_type)

    def test_insert_combo_box(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_combo_box
        #ExSummary:Shows how to insert a combo box form field into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a form that prompts the user to pick one of the items from the menu.
        builder.write('Pick a fruit: ')
        items = ['Apple', 'Banana', 'Cherry']
        builder.insert_combo_box('DropDown', items, 0)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertComboBox.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertComboBox.docx')
        form_field = doc.range.form_fields[0]
        self.assertTrue(form_field.enabled)
        self.assertEqual('DropDown', form_field.name)
        self.assertEqual(0, form_field.drop_down_selected_index)
        self.assertSequenceEqual(items, list(form_field.drop_down_items))
        self.assertEqual(aw.fields.FieldType.FIELD_FORM_DROP_DOWN, form_field.type)

    def test_signature_line_inline(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_signature_line(SignatureLineOptions,RelativeHorizontalPosition,float,RelativeVerticalPosition,float,WrapType)
        #ExSummary:Shows how to insert an inline signature line into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        options = aw.SignatureLineOptions()
        options.signer = 'John Doe'
        options.signer_title = 'Manager'
        options.email = 'johndoe@aspose.com'
        options.show_date = True
        options.default_instructions = False
        options.instructions = 'Please sign here.'
        options.allow_comments = True
        builder.insert_signature_line(signature_line_options=options, horz_pos=aw.drawing.RelativeHorizontalPosition.RIGHT_MARGIN, left=2, vert_pos=aw.drawing.RelativeVerticalPosition.PAGE, top=3, wrap_type=aw.drawing.WrapType.INLINE)
        # The signature line can be signed in Microsoft Word by double clicking it.
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.SignatureLineInline.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.SignatureLineInline.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        signature_line = shape.signature_line
        self.assertEqual('John Doe', signature_line.signer)
        self.assertEqual('Manager', signature_line.signer_title)
        self.assertEqual('johndoe@aspose.com', signature_line.email)
        self.assertTrue(signature_line.show_date)
        self.assertFalse(signature_line.default_instructions)
        self.assertEqual('Please sign here.', signature_line.instructions)
        self.assertTrue(signature_line.allow_comments)
        self.assertFalse(signature_line.is_signed)
        self.assertFalse(signature_line.is_valid)

    def test_set_paragraph_formatting(self):
        #ExStart
        #ExFor:ParagraphFormat.right_indent
        #ExFor:ParagraphFormat.left_indent
        #ExSummary:Shows how to configure paragraph formatting to create off-center text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Center all text that the document builder writes, and set up indents.
        # The indent configuration below will create a body of text that will sit asymmetrically on the page.
        # The "center" that we align the text to will be the middle of the body of text, not the middle of the page.
        paragraph_format = builder.paragraph_format
        paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        paragraph_format.left_indent = 100
        paragraph_format.right_indent = 50
        paragraph_format.space_after = 25
        builder.writeln('This paragraph demonstrates how left and right indentation affects word wrapping.')
        builder.writeln("The space between the above paragraph and this one depends on the DocumentBuilder's paragraph format.")
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.SetParagraphFormatting.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.SetParagraphFormatting.docx')
        for paragraph in doc.first_section.body.paragraphs:
            paragraph = paragraph.as_paragraph()
            self.assertEqual(aw.ParagraphAlignment.CENTER, paragraph.paragraph_format.alignment)
            self.assertEqual(100, paragraph.paragraph_format.left_indent)
            self.assertEqual(50, paragraph.paragraph_format.right_indent)
            self.assertEqual(25, paragraph.paragraph_format.space_after)

    @unittest.skip('Calculation problems')
    def test_set_cell_formatting(self):
        #ExStart
        #ExFor:DocumentBuilder.cell_format
        #ExFor:CellFormat.width
        #ExFor:CellFormat.left_padding
        #ExFor:CellFormat.right_padding
        #ExFor:CellFormat.top_padding
        #ExFor:CellFormat.bottom_padding
        #ExFor:DocumentBuilder.start_table
        #ExFor:DocumentBuilder.end_table
        #ExSummary:Shows how to format cells with a document builder.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        table = builder.start_table()
        builder.insert_cell()
        builder.write('Row 1, cell 1.')
        # Insert a second cell, and then configure cell text padding options.
        # The builder will apply these settings at its current cell, and any new cells creates afterwards.
        builder.insert_cell()
        cell_format = builder.cell_format
        cell_format.width = 250
        cell_format.left_padding = 30
        cell_format.right_padding = 30
        cell_format.top_padding = 30
        cell_format.bottom_padding = 30
        builder.write('Row 1, cell 2.')
        builder.end_row()
        builder.end_table()
        # The first cell was unaffected by the padding reconfiguration, and still holds the default values.
        self.assertEqual(0, table.first_row.cells[0].cell_format.width)
        self.assertEqual(5.4, table.first_row.cells[0].cell_format.left_padding)
        self.assertEqual(5.4, table.first_row.cells[0].cell_format.right_padding)
        self.assertEqual(0, table.first_row.cells[0].cell_format.top_padding)
        self.assertEqual(0, table.first_row.cells[0].cell_format.bottom_padding)
        self.assertEqual(250, table.first_row.cells[1].cell_format.width)
        self.assertEqual(30, table.first_row.cells[1].cell_format.left_padding)
        self.assertEqual(30, table.first_row.cells[1].cell_format.right_padding)
        self.assertEqual(30, table.first_row.cells[1].cell_format.top_padding)
        self.assertEqual(30, table.first_row.cells[1].cell_format.bottom_padding)
        # The first cell will still grow in the output document to match the size of its neighboring cell.
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.SetCellFormatting.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.SetCellFormatting.docx')
        table = doc.first_section.body.tables[0]
        self.assertEqual(159.3, table.first_row.cells[0].cell_format.width)
        self.assertEqual(5.4, table.first_row.cells[0].cell_format.left_padding)
        self.assertEqual(5.4, table.first_row.cells[0].cell_format.right_padding)
        self.assertEqual(0, table.first_row.cells[0].cell_format.top_padding)
        self.assertEqual(0, table.first_row.cells[0].cell_format.bottom_padding)
        self.assertEqual(310, table.first_row.cells[1].cell_format.width)
        self.assertEqual(30, table.first_row.cells[1].cell_format.left_padding)
        self.assertEqual(30, table.first_row.cells[1].cell_format.right_padding)
        self.assertEqual(30, table.first_row.cells[1].cell_format.top_padding)
        self.assertEqual(30, table.first_row.cells[1].cell_format.bottom_padding)

    def test_set_row_formatting(self):
        #ExStart
        #ExFor:DocumentBuilder.row_format
        #ExFor:HeightRule
        #ExFor:RowFormat.height
        #ExFor:RowFormat.height_rule
        #ExSummary:Shows how to format rows with a document builder.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        table = builder.start_table()
        builder.insert_cell()
        builder.write('Row 1, cell 1.')
        # Start a second row, and then configure its height. The builder will apply these settings to
        # its current row, as well as any new rows it creates afterwards.
        builder.end_row()
        row_format = builder.row_format
        row_format.height = 100
        row_format.height_rule = aw.HeightRule.EXACTLY
        builder.insert_cell()
        builder.write('Row 2, cell 1.')
        builder.end_table()
        # The first row was unaffected by the padding reconfiguration and still holds the default values.
        self.assertEqual(0, table.rows[0].row_format.height)
        self.assertEqual(aw.HeightRule.AUTO, table.rows[0].row_format.height_rule)
        self.assertEqual(100, table.rows[1].row_format.height)
        self.assertEqual(aw.HeightRule.EXACTLY, table.rows[1].row_format.height_rule)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.SetRowFormatting.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.SetRowFormatting.docx')
        table = doc.first_section.body.tables[0]
        self.assertEqual(0, table.rows[0].row_format.height)
        self.assertEqual(aw.HeightRule.AUTO, table.rows[0].row_format.height_rule)
        self.assertEqual(100, table.rows[1].row_format.height)
        self.assertEqual(aw.HeightRule.EXACTLY, table.rows[1].row_format.height_rule)

    def test_insert_footnote(self):
        #ExStart
        #ExFor:FootnoteType
        #ExFor:DocumentBuilder.insert_footnote(FootnoteType,str)
        #ExFor:DocumentBuilder.insert_footnote(FootnoteType,str,str)
        #ExSummary:Shows how to reference text with a footnote and an endnote.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert some text and mark it with a footnote with the IsAuto property set to "true" by default,
        # so the marker seen in the body text will be auto-numbered at "1",
        # and the footnote will appear at the bottom of the page.
        builder.write('This text will be referenced by a footnote.')
        builder.insert_footnote(footnote_type=aw.notes.FootnoteType.FOOTNOTE, footnote_text='Footnote comment regarding referenced text.')
        # Insert more text and mark it with an endnote with a custom reference mark,
        # which will be used in place of the number "2" and set "IsAuto" to false.
        builder.write('This text will be referenced by an endnote.')
        builder.insert_footnote(footnote_type=aw.notes.FootnoteType.ENDNOTE, footnote_text='Endnote comment regarding referenced text.', reference_mark='CustomMark')
        # Footnotes always appear at the bottom of their referenced text,
        # so this page break will not affect the footnote.
        # On the other hand, endnotes are always at the end of the document
        # so that this page break will push the endnote down to the next page.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertFootnote.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertFootnote.docx')
        test_util.TestUtil.verify_footnote(aw.notes.FootnoteType.FOOTNOTE, True, '', 'Footnote comment regarding referenced text.', doc.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote())
        test_util.TestUtil.verify_footnote(aw.notes.FootnoteType.ENDNOTE, False, 'CustomMark', 'CustomMark Endnote comment regarding referenced text.', doc.get_child(aw.NodeType.FOOTNOTE, 1, True).as_footnote())

    def test_apply_borders_and_shading(self):
        #ExStart
        #ExFor:BorderCollection.__getitem__(BorderType)
        #ExFor:Shading
        #ExFor:TextureIndex
        #ExFor:ParagraphFormat.shading
        #ExFor:Shading.texture
        #ExFor:Shading.background_pattern_color
        #ExFor:Shading.foreground_pattern_color
        #ExSummary:Shows how to decorate text with borders and shading.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        borders = builder.paragraph_format.borders
        borders.distance_from_text = 20
        borders.get_by_border_type(aw.BorderType.LEFT).line_style = aw.LineStyle.DOUBLE
        borders.get_by_border_type(aw.BorderType.RIGHT).line_style = aw.LineStyle.DOUBLE
        borders.get_by_border_type(aw.BorderType.TOP).line_style = aw.LineStyle.DOUBLE
        borders.get_by_border_type(aw.BorderType.BOTTOM).line_style = aw.LineStyle.DOUBLE
        shading = builder.paragraph_format.shading
        shading.texture = aw.TextureIndex.TEXTURE_DIAGONAL_CROSS
        shading.background_pattern_color = aspose.pydrawing.Color.light_coral
        shading.foreground_pattern_color = aspose.pydrawing.Color.light_salmon
        builder.write('This paragraph is formatted with a double border and shading.')
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.ApplyBordersAndShading.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.ApplyBordersAndShading.docx')
        borders = doc.first_section.body.first_paragraph.paragraph_format.borders
        self.assertEqual(20, borders.distance_from_text)
        self.assertEqual(aw.LineStyle.DOUBLE, borders.get_by_border_type(aw.BorderType.LEFT).line_style)
        self.assertEqual(aw.LineStyle.DOUBLE, borders.get_by_border_type(aw.BorderType.RIGHT).line_style)
        self.assertEqual(aw.LineStyle.DOUBLE, borders.get_by_border_type(aw.BorderType.TOP).line_style)
        self.assertEqual(aw.LineStyle.DOUBLE, borders.get_by_border_type(aw.BorderType.BOTTOM).line_style)
        self.assertEqual(aw.TextureIndex.TEXTURE_DIAGONAL_CROSS, shading.texture)
        self.assertEqual(aspose.pydrawing.Color.light_coral.to_argb(), shading.background_pattern_color.to_argb())
        self.assertEqual(aspose.pydrawing.Color.light_salmon.to_argb(), shading.foreground_pattern_color.to_argb())

    def test_delete_row(self):
        #ExStart
        #ExFor:DocumentBuilder.delete_row
        #ExSummary:Shows how to delete a row from a table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        table = builder.start_table()
        builder.insert_cell()
        builder.write('Row 1, cell 1.')
        builder.insert_cell()
        builder.write('Row 1, cell 2.')
        builder.end_row()
        builder.insert_cell()
        builder.write('Row 2, cell 1.')
        builder.insert_cell()
        builder.write('Row 2, cell 2.')
        builder.end_table()
        self.assertEqual(2, table.rows.count)
        # Delete the first row of the first table in the document.
        builder.delete_row(0, 0)
        self.assertEqual(1, table.rows.count)
        self.assertEqual('Row 2, cell 1.\x07Row 2, cell 2.\x07\x07', table.get_text().strip())
        #ExEnd

    def test_append_document_and_resolve_styles(self):
        for keep_source_numbering in [False, True]:
            #ExStart
            #ExFor:Document.append_document(Document,ImportFormatMode,ImportFormatOptions)
            #ExSummary:Shows how to manage list style clashes while appending a document.
            # Load a document with text in a custom style and clone it.
            src_doc = aw.Document(file_name=MY_DIR + 'Custom list numbering.docx')
            dst_doc = src_doc.clone()
            # We now have two documents, each with an identical style named "CustomStyle".
            # Change the text color for one of the styles to set it apart from the other.
            dst_doc.styles.get_by_name('CustomStyle').font.color = aspose.pydrawing.Color.dark_red
            # If there is a clash of list styles, apply the list format of the source document.
            # Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
            # Set the "KeepSourceNumbering" property to "true" import all clashing
            # list style numbering with the same appearance that it had in the source document.
            options = aw.ImportFormatOptions()
            options.keep_source_numbering = keep_source_numbering
            # Joining two documents that have different styles that share the same name causes a style clash.
            # We can specify an import format mode while appending documents to resolve this clash.
            dst_doc.append_document(src_doc=src_doc, import_format_mode=aw.ImportFormatMode.KEEP_DIFFERENT_STYLES, import_format_options=options)
            dst_doc.update_list_labels()
            dst_doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.AppendDocumentAndResolveStyles.docx')
            #ExEnd

    def test_insert_document_and_resolve_styles(self):
        for keep_source_numbering in [False, True]:
            #ExStart
            #ExFor:Document.append_document(Document,ImportFormatMode,ImportFormatOptions)
            #ExSummary:Shows how to manage list style clashes while inserting a document.
            dst_doc = aw.Document()
            builder = aw.DocumentBuilder(doc=dst_doc)
            builder.insert_break(aw.BreakType.PARAGRAPH_BREAK)
            dst_doc.lists.add(list_template=aw.lists.ListTemplate.NUMBER_DEFAULT)
            list = dst_doc.lists[0]
            builder.list_format.list = list
            i = 1
            while i <= 15:
                builder.write(f'List Item {i}\n')
                i += 1
            attach_doc = dst_doc.clone(True).as_document()
            # If there is a clash of list styles, apply the list format of the source document.
            # Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
            # Set the "KeepSourceNumbering" property to "true" import all clashing
            # list style numbering with the same appearance that it had in the source document.
            import_options = aw.ImportFormatOptions()
            import_options.keep_source_numbering = keep_source_numbering
            builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
            builder.insert_document(src_doc=attach_doc, import_format_mode=aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, import_format_options=import_options)
            dst_doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertDocumentAndResolveStyles.docx')
            #ExEnd

    def test_load_document_with_list_numbering(self):
        for keep_source_numbering in [False, True]:
            #ExStart
            #ExFor:Document.append_document(Document,ImportFormatMode,ImportFormatOptions)
            #ExSummary:Shows how to manage list style clashes while appending a clone of a document to itself.
            src_doc = aw.Document(file_name=MY_DIR + 'List item.docx')
            dst_doc = aw.Document(file_name=MY_DIR + 'List item.docx')
            # If there is a clash of list styles, apply the list format of the source document.
            # Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
            # Set the "KeepSourceNumbering" property to "true" import all clashing
            # list style numbering with the same appearance that it had in the source document.
            builder = aw.DocumentBuilder(doc=dst_doc)
            builder.move_to_document_end()
            builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
            options = aw.ImportFormatOptions()
            options.keep_source_numbering = keep_source_numbering
            builder.insert_document(src_doc=src_doc, import_format_mode=aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, import_format_options=options)
            dst_doc.update_list_labels()
            #ExEnd

    def test_ignore_text_boxes(self):
        for ignore_text_boxes in [True, False]:
            #ExStart
            #ExFor:ImportFormatOptions.ignore_text_boxes
            #ExSummary:Shows how to manage text box formatting while appending a document.
            # Create a document that will have nodes from another document inserted into it.
            dst_doc = aw.Document()
            builder = aw.DocumentBuilder(doc=dst_doc)
            builder.writeln('Hello world!')
            # Create another document with a text box, which we will import into the first document.
            src_doc = aw.Document()
            builder = aw.DocumentBuilder(doc=src_doc)
            text_box = builder.insert_shape(shape_type=aw.drawing.ShapeType.TEXT_BOX, width=300, height=100)
            builder.move_to(text_box.first_paragraph)
            builder.paragraph_format.style.font.name = 'Courier New'
            builder.paragraph_format.style.font.size = 24
            builder.write('Textbox contents')
            # Set a flag to specify whether to clear or preserve text box formatting
            # while importing them to other documents.
            import_format_options = aw.ImportFormatOptions()
            import_format_options.ignore_text_boxes = ignore_text_boxes
            # Import the text box from the source document into the destination document,
            # and then verify whether we have preserved the styling of its text contents.
            importer = aw.NodeImporter(src_doc=src_doc, dst_doc=dst_doc, import_format_mode=aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, import_format_options=import_format_options)
            imported_text_box = importer.import_node(text_box, True).as_shape()
            dst_doc.first_section.body.paragraphs[1].append_child(imported_text_box)
            if ignore_text_boxes:
                self.assertEqual(12, imported_text_box.first_paragraph.runs[0].font.size)
                self.assertEqual('Times New Roman', imported_text_box.first_paragraph.runs[0].font.name)
            else:
                self.assertEqual(24, imported_text_box.first_paragraph.runs[0].font.size)
                self.assertEqual('Courier New', imported_text_box.first_paragraph.runs[0].font.name)
            dst_doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.IgnoreTextBoxes.docx')
            #ExEnd

    def test_move_to_field(self):
        for move_cursor_to_after_the_field in [False, True]:
            #ExStart
            #ExFor:DocumentBuilder.move_to_field
            #ExSummary:Shows how to move a document builder's node insertion point cursor to a specific field.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            # Insert a field using the DocumentBuilder and add a run of text after it.
            field = builder.insert_field(field_code=' AUTHOR "John Doe" ')
            # The builder's cursor is currently at end of the document.
            self.assertIsNone(builder.current_node)
            # Move the cursor to the field while specifying whether to place that cursor before or after the field.
            builder.move_to_field(field, move_cursor_to_after_the_field)
            # Note that the cursor is outside of the field in both cases.
            # This means that we cannot edit the field using the builder like this.
            # To edit a field, we can use the builder's MoveTo method on a field's FieldStart
            # or FieldSeparator node to place the cursor inside.
            if move_cursor_to_after_the_field:
                self.assertIsNone(builder.current_node)
                builder.write(' Text immediately after the field.')
                self.assertEqual('\x13 AUTHOR "John Doe" \x14John Doe\x15 Text immediately after the field.', doc.get_text().strip())
            else:
                self.assertEqual(field.start, builder.current_node)
                builder.write('Text immediately before the field. ')
                self.assertEqual('Text immediately before the field. \x13 AUTHOR "John Doe" \x14John Doe\x15', doc.get_text().strip())
            #ExEnd

    def test_insert_pie_chart(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_chart(ChartType,float,float)
        #ExSummary:Shows how to insert a pie chart into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        chart = builder.insert_chart(chart_type=aw.drawing.charts.ChartType.PIE, width=aw.ConvertUtil.pixel_to_point(pixels=300), height=aw.ConvertUtil.pixel_to_point(pixels=300)).chart
        self.assertEqual(225, aw.ConvertUtil.pixel_to_point(pixels=300))  #ExSkip
        chart.series.clear()
        chart.series.add(series_name='My fruit', categories=['Apples', 'Bananas', 'Cherries'], values=[1.3, 2.2, 1.5])
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertPieChart.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertPieChart.docx')
        chart_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual('Chart Title', chart_shape.chart.title.text)
        self.assertEqual(225, chart_shape.width)
        self.assertEqual(225, chart_shape.height)

    def test_insert_chart_relative_position(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_chart(ChartType,RelativeHorizontalPosition,float,RelativeVerticalPosition,float,float,float,WrapType)
        #ExSummary:Shows how to specify position and wrapping while inserting a chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.insert_chart(chart_type=aw.drawing.charts.ChartType.PIE, horz_pos=aw.drawing.RelativeHorizontalPosition.MARGIN, left=100, vert_pos=aw.drawing.RelativeVerticalPosition.MARGIN, top=100, width=200, height=100, wrap_type=aw.drawing.WrapType.SQUARE)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertedChartRelativePosition.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertedChartRelativePosition.docx')
        chart_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual(100, chart_shape.top)
        self.assertEqual(100, chart_shape.left)
        self.assertEqual(200, chart_shape.width)
        self.assertEqual(100, chart_shape.height)
        self.assertEqual(aw.drawing.WrapType.SQUARE, chart_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, chart_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, chart_shape.relative_vertical_position)

    def test_insert_field_and_update(self):
        for update_inserted_fields_immediately in [False, True]:
            #ExStart
            #ExFor:DocumentBuilder.insert_field(FieldType,bool)
            #ExFor:Field.update
            #ExSummary:Shows how to insert a field into a document using FieldType.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            # Insert two fields while passing a flag which determines whether to update them as the builder inserts them.
            # In some cases, updating fields could be computationally expensive, and it may be a good idea to defer the update.
            doc.built_in_document_properties.author = 'John Doe'
            builder.write('This document was written by ')
            builder.insert_field(field_type=aw.fields.FieldType.FIELD_AUTHOR, update_field=update_inserted_fields_immediately)
            builder.insert_paragraph()
            builder.write('\nThis is page ')
            builder.insert_field(field_type=aw.fields.FieldType.FIELD_PAGE, update_field=update_inserted_fields_immediately)
            self.assertEqual(' AUTHOR ', doc.range.fields[0].get_field_code())
            self.assertEqual(' PAGE ', doc.range.fields[1].get_field_code())
            if update_inserted_fields_immediately:
                self.assertEqual('John Doe', doc.range.fields[0].result)
                self.assertEqual('1', doc.range.fields[1].result)
            else:
                self.assertEqual('', doc.range.fields[0].result)
                self.assertEqual('', doc.range.fields[1].result)
                # We will need to update these fields using the update methods manually.
                doc.range.fields[0].update()
                self.assertEqual('John Doe', doc.range.fields[0].result)
                doc.update_fields()
                self.assertEqual('1', doc.range.fields[1].result)
            #ExEnd
            doc = document_helper.DocumentHelper.save_open(doc)
            self.assertEqual('This document was written by \x13 AUTHOR \x14John Doe\x15' + '\r\rThis is page \x13 PAGE \x141\x15', doc.get_text().strip())
            test_util.TestUtil.verify_field(expected_type=aw.fields.FieldType.FIELD_AUTHOR, expected_field_code=' AUTHOR ', expected_result='John Doe', field=doc.range.fields[0])
            test_util.TestUtil.verify_field(expected_type=aw.fields.FieldType.FIELD_PAGE, expected_field_code=' PAGE ', expected_result='1', field=doc.range.fields[1])

    @unittest.skip('Failed')
    def test_insert_video_with_url(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_online_video(str,float,float)
        #ExSummary:Shows how to insert an online video into a document using a URL.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.insert_online_video(video_url='https://youtu.be/g1N9ke8Prmk', width=360, height=270)
        # We can watch the video from Microsoft Word by clicking on the shape.
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertVideoWithUrl.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertVideoWithUrl.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        test_util.TestUtil.verify_image_in_shape(480, 360, aw.drawing.ImageType.JPEG, shape)
        self.assertEqual('https://youtu.be/t_1LYZ102RA', shape.href)
        self.assertEqual(360, shape.width)
        self.assertEqual(270, shape.height)

    def test_insert_underline(self):
        #ExStart
        #ExFor:DocumentBuilder.underline
        #ExSummary:Shows how to format text inserted by a document builder.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.underline = aw.Underline.DASH
        builder.font.color = aspose.pydrawing.Color.blue
        builder.font.size = 32
        # The builder applies formatting to its current paragraph and any new text added by it afterward.
        builder.writeln('Large, blue, and underlined text.')
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertUnderline.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertUnderline.docx')
        first_run = doc.first_section.body.first_paragraph.runs[0]
        self.assertEqual('Large, blue, and underlined text.', first_run.get_text().strip())
        self.assertEqual(aw.Underline.DASH, first_run.font.underline)
        self.assertEqual(aspose.pydrawing.Color.blue.to_argb(), first_run.font.color.to_argb())
        self.assertEqual(32, first_run.font.size)

    def test_current_story(self):
        #ExStart
        #ExFor:DocumentBuilder.current_story
        #ExSummary:Shows how to work with a document builder's current story.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # A Story is a type of node that has child Paragraph nodes, such as a Body.
        self.assertEqual(builder.current_story, doc.first_section.body)
        self.assertEqual(builder.current_story, builder.current_paragraph.parent_node)
        self.assertEqual(aw.StoryType.MAIN_TEXT, builder.current_story.story_type)
        builder.current_story.append_paragraph('Text added to current Story.')
        # A Story can also contain tables.
        table = builder.start_table()
        builder.insert_cell()
        builder.write('Row 1, cell 1')
        builder.insert_cell()
        builder.write('Row 1, cell 2')
        builder.end_table()
        self.assertTrue(builder.current_story.tables.contains(table))
        #ExEnd
        doc = document_helper.DocumentHelper.save_open(doc)
        self.assertEqual(1, doc.first_section.body.tables.count)
        self.assertEqual('Row 1, cell 1\x07Row 1, cell 2\x07\x07\rText added to current Story.', doc.first_section.body.get_text().strip())

    def test_insert_ole_objects(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_ole_object(BytesIO,str,bool,BytesIO)
        #ExSummary:Shows how to use document builder to embed OLE objects in a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a Microsoft Excel spreadsheet from the local file system
        # into the document while keeping its default appearance.
        with system_helper.io.File.open(MY_DIR + 'Spreadsheet.xlsx', system_helper.io.FileMode.OPEN) as spreadsheet_stream:
            builder.writeln('Spreadsheet Ole object:')
            # If 'presentation' is omitted and 'asIcon' is set, this overloaded method selects
            # the icon according to 'progId' and uses the predefined icon caption.
            builder.insert_ole_object(stream=spreadsheet_stream, prog_id='OleObject.xlsx', as_icon=False, presentation=None)
        # Insert a Microsoft Powerpoint presentation as an OLE object.
        # This time, it will have an image downloaded from the web for an icon.
        with system_helper.io.File.open(MY_DIR + 'Presentation.pptx', system_helper.io.FileMode.OPEN) as powerpoint_stream:
            img_bytes = system_helper.io.File.read_all_bytes(IMAGE_DIR + 'Logo.jpg')
            with io.BytesIO(img_bytes) as image_stream:
                builder.insert_paragraph()
                builder.writeln('Powerpoint Ole object:')
                builder.insert_ole_object(stream=powerpoint_stream, prog_id='OleObject.pptx', as_icon=True, presentation=image_stream)
        # Double-click these objects in Microsoft Word to open
        # the linked files using their respective applications.
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertOleObjects.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertOleObjects.docx')
        self.assertEqual(2, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual('', shape.ole_format.icon_caption)
        self.assertFalse(shape.ole_format.ole_icon)
        shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()
        self.assertEqual('Unknown', shape.ole_format.icon_caption)
        self.assertTrue(shape.ole_format.ole_icon)

    def test_insert_style_separator(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_style_separator
        #ExSummary:Shows how to work with style separators.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Each paragraph can only have one style.
        # The InsertStyleSeparator method allows us to work around this limitation.
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
        builder.write('This text is in a Heading style. ')
        builder.insert_style_separator()
        para_style = builder.document.styles.add(aw.StyleType.PARAGRAPH, 'MyParaStyle')
        para_style.font.bold = False
        para_style.font.size = 8
        para_style.font.name = 'Arial'
        builder.paragraph_format.style_name = para_style.name
        builder.write('This text is in a custom style. ')
        # Calling the InsertStyleSeparator method creates another paragraph,
        # which can have a different style to the previous. There will be no break between paragraphs.
        # The text in the output document will look like one paragraph with two styles.
        self.assertEqual(2, doc.first_section.body.paragraphs.count)
        self.assertEqual('Heading 1', doc.first_section.body.paragraphs[0].paragraph_format.style.name)
        self.assertEqual('MyParaStyle', doc.first_section.body.paragraphs[1].paragraph_format.style.name)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertStyleSeparator.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertStyleSeparator.docx')
        self.assertEqual(2, doc.first_section.body.paragraphs.count)
        self.assertEqual('This text is in a Heading style. \r This text is in a custom style.', doc.get_text().strip())
        self.assertEqual('Heading 1', doc.first_section.body.paragraphs[0].paragraph_format.style.name)
        self.assertEqual('MyParaStyle', doc.first_section.body.paragraphs[1].paragraph_format.style.name)
        self.assertEqual(' ', doc.first_section.body.paragraphs[1].runs[0].get_text())
        test_util.TestUtil.doc_package_file_contains_string('w:rPr><w:vanish /><w:specVanish /></w:rPr>', ARTIFACTS_DIR + 'DocumentBuilder.InsertStyleSeparator.docx', 'document.xml')
        test_util.TestUtil.doc_package_file_contains_string('<w:t xml:space="preserve"> </w:t>', ARTIFACTS_DIR + 'DocumentBuilder.InsertStyleSeparator.docx', 'document.xml')

    @unittest.skip('Bug: does not insert headers and footers, all lists (bullets, numbering, multilevel) breaks')
    def test_insert_document(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_document(Document,ImportFormatMode)
        #ExFor:ImportFormatMode
        #ExSummary:Shows how to insert a document into another document.
        doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        builder = aw.DocumentBuilder(doc=doc)
        builder.move_to_document_end()
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        doc_to_insert = aw.Document(file_name=MY_DIR + 'Formatted elements.docx')
        builder.insert_document(src_doc=doc_to_insert, import_format_mode=aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
        builder.document.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertDocument.docx')
        #ExEnd
        self.assertEqual(29, doc.styles.count)
        self.assertTrue(document_helper.DocumentHelper.compare_docs(ARTIFACTS_DIR + 'DocumentBuilder.InsertDocument.docx', GOLDS_DIR + 'DocumentBuilder.InsertDocument Gold.docx'))

    def test_smart_style_behavior(self):
        #ExStart
        #ExFor:ImportFormatOptions
        #ExFor:ImportFormatOptions.smart_style_behavior
        #ExFor:DocumentBuilder.insert_document(Document,ImportFormatMode,ImportFormatOptions)
        #ExSummary:Shows how to resolve duplicate styles while inserting documents.
        dst_doc = aw.Document()
        builder = aw.DocumentBuilder(doc=dst_doc)
        my_style = builder.document.styles.add(aw.StyleType.PARAGRAPH, 'MyStyle')
        my_style.font.size = 14
        my_style.font.name = 'Courier New'
        my_style.font.color = aspose.pydrawing.Color.blue
        builder.paragraph_format.style_name = my_style.name
        builder.writeln('Hello world!')
        # Clone the document and edit the clone's "MyStyle" style, so it is a different color than that of the original.
        # If we insert the clone into the original document, the two styles with the same name will cause a clash.
        src_doc = dst_doc.clone()
        src_doc.styles.get_by_name('MyStyle').font.color = aspose.pydrawing.Color.red
        # When we enable SmartStyleBehavior and use the KeepSourceFormatting import format mode,
        # Aspose.Words will resolve style clashes by converting source document styles.
        # with the same names as destination styles into direct paragraph attributes.
        options = aw.ImportFormatOptions()
        options.smart_style_behavior = True
        builder.insert_document(src_doc=src_doc, import_format_mode=aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, import_format_options=options)
        dst_doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.SmartStyleBehavior.docx')
        #ExEnd
        dst_doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.SmartStyleBehavior.docx')
        self.assertEqual(aspose.pydrawing.Color.blue.to_argb(), dst_doc.styles.get_by_name('MyStyle').font.color.to_argb())
        self.assertEqual('MyStyle', dst_doc.first_section.body.paragraphs[0].paragraph_format.style.name)
        self.assertEqual('Normal', dst_doc.first_section.body.paragraphs[1].paragraph_format.style.name)
        self.assertEqual(14, dst_doc.first_section.body.paragraphs[1].runs[0].font.size)
        self.assertEqual('Courier New', dst_doc.first_section.body.paragraphs[1].runs[0].font.name)
        self.assertEqual(aspose.pydrawing.Color.red.to_argb(), dst_doc.first_section.body.paragraphs[1].runs[0].font.color.to_argb())

    def test_emphases_warning_source_markdown(self):
        #ExStart
        #ExFor:WarningInfo.source
        #ExFor:WarningSource
        #ExSummary:Shows how to work with the warning source.
        doc = aw.Document(file_name=MY_DIR + 'Emphases markdown warning.docx')
        warnings = aw.WarningInfoCollection()
        doc.warning_callback = warnings
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.EmphasesWarningSourceMarkdown.md')
        for warning_info in warnings:
            if warning_info.source == aw.WarningSource.MARKDOWN:
                self.assertEqual('The (*, 0:11) cannot be properly written into Markdown.', warning_info.description)
        #ExEnd

    def test_do_not_ignore_header_footer(self):
        #ExStart
        #ExFor:ImportFormatOptions.ignore_header_footer
        #ExSummary:Shows how to specifies ignoring or not source formatting of headers/footers content.
        dst_doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        src_doc = aw.Document(file_name=MY_DIR + 'Header and footer types.docx')
        # If 'IgnoreHeaderFooter' is false then the original formatting for header/footer content
        # from "Header and footer types.docx" will be used.
        # If 'IgnoreHeaderFooter' is true then the formatting for header/footer content
        # from "Document.docx" will be used.
        import_format_options = aw.ImportFormatOptions()
        import_format_options.ignore_header_footer = False
        dst_doc.append_document(src_doc=src_doc, import_format_mode=aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, import_format_options=import_format_options)
        dst_doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.DoNotIgnoreHeaderFooter.docx')
        #ExEnd

    def test_insert_online_video(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_online_video(str,RelativeHorizontalPosition,float,RelativeVerticalPosition,float,float,float,WrapType)
        #ExSummary:Shows how to insert an online video into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        video_url = 'https://vimeo.com/52477838'
        # Insert a shape that plays a video from the web when clicked in Microsoft Word.
        # This rectangular shape will contain an image based on the first frame of the linked video
        # and a "play button" visual prompt. The video has an aspect ratio of 16:9.
        # We will set the shape's size to that ratio, so the image does not appear stretched.
        builder.insert_online_video(video_url=video_url, horz_pos=aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN, left=0, vert_pos=aw.drawing.RelativeVerticalPosition.TOP_MARGIN, top=0, width=320, height=180, wrap_type=aw.drawing.WrapType.SQUARE)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertOnlineVideo.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertOnlineVideo.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        test_util.TestUtil.verify_image_in_shape(640, 360, aw.drawing.ImageType.JPEG, shape)
        self.assertEqual(320, shape.width)
        self.assertEqual(180, shape.height)
        self.assertEqual(0, shape.left)
        self.assertEqual(0, shape.top)
        self.assertEqual(aw.drawing.WrapType.SQUARE, shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.TOP_MARGIN, shape.relative_vertical_position)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN, shape.relative_horizontal_position)
        self.assertEqual('https://vimeo.com/52477838', shape.href)

    def test_insert_ole_object_as_icon(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_ole_object_as_icon(str,str,bool,str,str)
        #ExFor:DocumentBuilder.insert_ole_object_as_icon(BytesIO,str,str,str)
        #ExSummary:Shows how to insert an embedded or linked OLE object as icon into the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
        # the icon according to 'progId' and uses the filename for the icon caption.
        builder.insert_ole_object_as_icon(file_name=MY_DIR + 'Presentation.pptx', prog_id='Package', is_linked=False, icon_file=IMAGE_DIR + 'Logo icon.ico', icon_caption='My embedded file')
        builder.insert_break(aw.BreakType.LINE_BREAK)
        with system_helper.io.FileStream(MY_DIR + 'Presentation.pptx', system_helper.io.FileMode.OPEN) as stream:
            # If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
            # the icon according to the file extension and uses the filename for the icon caption.
            shape = builder.insert_ole_object_as_icon(stream=stream, prog_id='PowerPoint.Application', icon_file=IMAGE_DIR + 'Logo icon.ico', icon_caption='My embedded file stream')
            set_ole_package = shape.ole_format.ole_package
            set_ole_package.file_name = 'Presentation.pptx'
            set_ole_package.display_name = 'Presentation.pptx'
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.InsertOleObjectAsIcon.docx')
        #ExEnd

    def test_preserve_blocks(self):
        #ExStart
        #ExFor:HtmlInsertOptions
        #ExSummary:Shows how to allows better preserve borders and margins seen.
        html = "\n                <html>\n                    <div style='border:dotted'>\n                    <div style='border:solid'>\n                        <p>paragraph 1</p>\n                        <p>paragraph 2</p>\n                    </div>\n                    </div>\n                </html>"
        # Set the new mode of import HTML block-level elements.
        insert_options = aw.HtmlInsertOptions.PRESERVE_BLOCKS
        builder = aw.DocumentBuilder()
        builder.insert_html(html=html, options=insert_options)
        builder.document.save(file_name=ARTIFACTS_DIR + 'DocumentBuilder.PreserveBlocks.docx')
        #ExEnd

    def test_phonetic_guide(self):
        #ExStart
        #ExFor:Run.is_phonetic_guide
        #ExFor:Run.phonetic_guide
        #ExFor:PhoneticGuide
        #ExFor:PhoneticGuide.base_text
        #ExFor:PhoneticGuide.ruby_text
        #ExSummary:Shows how to get properties of the phonetic guide.
        doc = aw.Document(file_name=MY_DIR + 'Phonetic guide.docx')
        runs = doc.first_section.body.first_paragraph.runs
        # Use phonetic guide in the Asian text.
        self.assertEqual(True, runs[0].is_phonetic_guide)
        phonetic_guide = runs[0].phonetic_guide
        self.assertEqual('base', phonetic_guide.base_text)
        self.assertEqual('ruby', phonetic_guide.ruby_text)
        #ExEnd

    def test_horizontal_rule_format_exceptions(self):
        """Checking the boundary conditions of WidthPercent and Height properties."""
        builder = aw.DocumentBuilder()
        shape = builder.insert_horizontal_rule()
        horizontal_rule_format = shape.horizontal_rule_format
        horizontal_rule_format.width_percent = 1
        horizontal_rule_format.width_percent = 100
        with self.assertRaises(Exception):
            horizontal_rule_format.width_percent = 0
        with self.assertRaises(Exception):
            horizontal_rule_format.width_percent = 101
        horizontal_rule_format.height = 0
        horizontal_rule_format.height = 1584
        with self.assertRaises(Exception):
            horizontal_rule_format.height = -1
        with self.assertRaises(Exception):
            horizontal_rule_format.height = 1585

    def test_insert_html_with_formatting(self):
        for use_builder_formatting in (False, True):
            with self.subTest(use_builder_formatting=use_builder_formatting):
                #ExStart
                #ExFor:DocumentBuilder.insert_html(str,bool)
                #ExSummary:Shows how to apply a document builder's formatting while inserting HTML content.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                # Set a text alignment for the builder, insert an HTML paragraph with a specified alignment, and one without.
                builder.paragraph_format.alignment = aw.ParagraphAlignment.DISTRIBUTED
                builder.insert_html("<p align='right'>Paragraph 1.</p>" + '<p>Paragraph 2.</p>', use_builder_formatting)
                paragraphs = doc.first_section.body.paragraphs
                # The first paragraph has an alignment specified. When "insert_html" parses the HTML code,
                # the paragraph alignment value found in the HTML code always supersedes the document builder's value.
                self.assertEqual('Paragraph 1.', paragraphs[0].get_text().strip())
                self.assertEqual(aw.ParagraphAlignment.RIGHT, paragraphs[0].paragraph_format.alignment)
                # The second paragraph has no alignment specified. It can have its alignment value filled in
                # by the builder's value depending on the flag we passed to the "insert_html" method.
                self.assertEqual('Paragraph 2.', paragraphs[1].get_text().strip())
                self.assertEqual(aw.ParagraphAlignment.DISTRIBUTED if use_builder_formatting else aw.ParagraphAlignment.LEFT, paragraphs[1].paragraph_format.alignment)
                doc.save(ARTIFACTS_DIR + 'DocumentBuilder.insert_html_with_formatting.docx')
                #ExEnd

    def test_create_column_bookmark(self):
        #ExStart
        #ExFor:DocumentBuilder.start_column_bookmark
        #ExFor:DocumentBuilder.end_column_bookmark
        #ExSummary:Shows how to create a column bookmark.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.start_table()
        builder.insert_cell()
        # Cells 1,2,4,5 will be bookmarked.
        builder.start_column_bookmark('MyBookmark_1')
        # Badly formed bookmarks or bookmarks with duplicate names will be ignored when the document is saved.
        builder.start_column_bookmark('MyBookmark_1')
        builder.start_column_bookmark('BadStartBookmark')
        builder.write('Cell 1')
        builder.insert_cell()
        builder.write('Cell 2')
        builder.insert_cell()
        builder.write('Cell 3')
        builder.end_row()
        builder.insert_cell()
        builder.write('Cell 4')
        builder.insert_cell()
        builder.write('Cell 5')
        builder.end_column_bookmark('MyBookmark_1')
        builder.end_column_bookmark('MyBookmark_1')
        with self.assertRaises(Exception):  #ExSkip
            builder.end_column_bookmark('BadEndBookmark')  #ExSkip
        builder.insert_cell()
        builder.write('Cell 6')
        builder.end_row()
        builder.end_table()
        doc.save(ARTIFACTS_DIR + 'Bookmarks.create_column_bookmark.docx')
        #ExEnd

    def test_signature_line_provider_id(self):
        #ExStart
        #ExFor:SignatureLine.is_signed
        #ExFor:SignatureLine.is_valid
        #ExFor:SignatureLine.provider_id
        #ExFor:SignatureLineOptions.show_date
        #ExFor:SignatureLineOptions.email
        #ExFor:SignatureLineOptions.default_instructions
        #ExFor:SignatureLineOptions.instructions
        #ExFor:SignatureLineOptions.allow_comments
        #ExFor:DocumentBuilder.insert_signature_line(SignatureLineOptions)
        #ExFor:SignOptions.provider_id
        #ExSummary:Shows how to sign a document with a personal certificate and a signature line.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        signature_line_options = aw.SignatureLineOptions()
        signature_line_options.signer = 'vderyushev'
        signature_line_options.signer_title = 'QA'
        signature_line_options.email = 'vderyushev@aspose.com'
        signature_line_options.show_date = True
        signature_line_options.default_instructions = False
        signature_line_options.instructions = 'Please sign here.'
        signature_line_options.allow_comments = True
        signature_line = builder.insert_signature_line(signature_line_options).signature_line
        signature_line.provider_id = uuid.UUID('CF5A7BB4-8F3C-4756-9DF6-BEF7F13259A2')
        self.assertFalse(signature_line.is_signed)
        self.assertFalse(signature_line.is_valid)
        doc.save(ARTIFACTS_DIR + 'DocumentBuilder.signature_line_provider_id.docx')
        sign_options = aw.digitalsignatures.SignOptions()
        sign_options.signature_line_id = signature_line.id
        sign_options.provider_id = signature_line.provider_id
        sign_options.comments = 'Document was signed by vderyushev'
        sign_options.sign_time = datetime.datetime.utcnow()
        cert_holder = aw.digitalsignatures.CertificateHolder.create(MY_DIR + 'morzal.pfx', 'aw')
        aw.digitalsignatures.DigitalSignatureUtil.sign(ARTIFACTS_DIR + 'DocumentBuilder.signature_line_provider_id.docx', ARTIFACTS_DIR + 'DocumentBuilder.signature_line_provider_id.signed.docx', cert_holder, sign_options)
        # Re-open our saved document, and verify that the "is_signed" and "is_valid" properties both equal "True",
        # indicating that the signature line contains a signature.
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.signature_line_provider_id.signed.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        signature_line = shape.signature_line
        self.assertTrue(signature_line.is_signed)
        self.assertTrue(signature_line.is_valid)
        #ExEnd
        self.assertEqual('vderyushev', signature_line.signer)
        self.assertEqual('QA', signature_line.signer_title)
        self.assertEqual('vderyushev@aspose.com', signature_line.email)
        self.assertTrue(signature_line.show_date)
        self.assertFalse(signature_line.default_instructions)
        self.assertEqual('Please sign here.', signature_line.instructions)
        self.assertTrue(signature_line.allow_comments)
        self.assertTrue(signature_line.is_signed)
        self.assertTrue(signature_line.is_valid)
        signatures = aw.digitalsignatures.DigitalSignatureUtil.load_signatures(ARTIFACTS_DIR + 'DocumentBuilder.signature_line_provider_id.signed.docx')
        self.assertEqual(1, signatures.count)
        self.assertTrue(signatures[0].is_valid)
        self.assertEqual('Document was signed by vderyushev', signatures[0].comments)
        self.assertAlmostEqual(datetime.datetime.now(tz=timezone.utc), signatures[0].sign_time, delta=timedelta(seconds=5))
        self.assertEqual('CN=Morzal.Me', signatures[0].issuer_name)
        self.assertEqual(aw.digitalsignatures.DigitalSignatureType.XML_DSIG, signatures[0].signature_type)

    def test_insert_ole_object_exception(self):
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        with self.assertRaises(Exception):
            builder.insert_ole_object('', 'checkbox', False, True, None)

    def test_insert_field(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_field(str)
        #ExFor:Field
        #ExFor:Field.result
        #ExFor:Field.get_field_code()
        #ExFor:Field.type
        #ExFor:FieldType
        #ExSummary:Shows how to insert a field into a document using a field code.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        field = builder.insert_field('DATE \\@ "dddd, MMMM dd, yyyy"')
        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field.type)
        self.assertEqual('DATE \\@ "dddd, MMMM dd, yyyy"', field.get_field_code())
        # This overload of the "insert_field" method automatically updates inserted fields.
        self.assertAlmostEqual(datetime.datetime.strptime(field.result, '%A, %B %d, %Y'), datetime.datetime.now(), delta=timedelta(1))
        #ExEnd

    @unittest.skip('Calculation problems')
    def test_insert_online_video_custom_thumbnail(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_online_video(str,str,bytes,float,float)
        #ExFor:DocumentBuilder.insert_online_video(str,str,bytes,RelativeHorizontalPosition,float,RelativeVerticalPosition,float,float,float,WrapType)
        #ExSummary:Shows how to insert an online video into a document with a custom thumbnail.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        video_url = 'https://vimeo.com/52477838'
        video_embed_code = '<iframe src="https://player.vimeo.com/video/52477838" width="640" height="360" frameborder="0" ' + 'title="Aspose" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>'
        with open(IMAGE_DIR + 'Logo.jpg', 'rb') as file:
            thumbnail_image_bytes = file.read()
            image = aspose.pydrawing.Image.from_stream(io.BytesIO(thumbnail_image_bytes))
            # Below are two ways of creating a shape with a custom thumbnail, which links to an online video
            # that will play when we click on the shape in Microsoft Word.
            # 1 -  Insert an inline shape at the builder's node insertion cursor:
            builder.insert_online_video(video_url, video_embed_code, thumbnail_image_bytes, image.width, image.height)
            builder.insert_break(aw.BreakType.PAGE_BREAK)
            # 2 -  Insert a floating shape:
            left = builder.page_setup.right_margin - image.width
            top = builder.page_setup.bottom_margin - image.height
            builder.insert_online_video(video_url, video_embed_code, thumbnail_image_bytes, aw.drawing.RelativeHorizontalPosition.RIGHT_MARGIN, left, aw.drawing.RelativeVerticalPosition.BOTTOM_MARGIN, top, image.width, image.height, aw.drawing.WrapType.SQUARE)
        doc.save(ARTIFACTS_DIR + 'DocumentBuilder.insert_online_video_custom_thumbnail.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.insert_online_video_custom_thumbnail.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, shape)
        self.assertEqual(400.0, shape.width)
        self.assertEqual(400.0, shape.height)
        self.assertEqual(0.0, shape.left)
        self.assertEqual(0.0, shape.top)
        self.assertEqual(aw.drawing.WrapType.INLINE, shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, shape.relative_vertical_position)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, shape.relative_horizontal_position)
        self.assertEqual('https://vimeo.com/52477838', shape.href)
        shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()
        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, shape)
        self.assertEqual(400.0, shape.width)
        self.assertEqual(400.0, shape.height)
        self.assertEqual(-329.15, shape.left)
        self.assertEqual(-329.15, shape.top)
        self.assertEqual(aw.drawing.WrapType.SQUARE, shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.BOTTOM_MARGIN, shape.relative_vertical_position)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.RIGHT_MARGIN, shape.relative_horizontal_position)
        self.assertEqual('https://vimeo.com/52477838', shape.href)
        #ServicePointManager.security_protocol = SecurityProtocolType.TLS12
        self.verify_web_response_status_code(200, shape.href)

    def _test_markdown_document_emphases(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        builder = aw.DocumentBuilder()
        # Bold and Italic are represented as "Font.bold" and "Font.italic".
        builder.font.italic = True
        builder.writeln('This text will be italic')
        # Use clear formatting if we don't want to combine styles between paragraphs.
        builder.font.clear_formatting()
        builder.font.bold = True
        builder.writeln('This text will be bold')
        builder.font.clear_formatting()
        builder.font.italic = True
        builder.write('You ')
        builder.font.bold = True
        builder.write('can')
        builder.font.bold = False
        builder.writeln(' combine them')
        builder.font.clear_formatting()
        builder.font.strike_through = True
        builder.writeln('This text will be strikethrough')
        # Markdown treats asterisks (*), underscores (_) and tilde (~) as indicators of emphasis.
        builder.document.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _test_markdown_document_inline_code(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
        builder = aw.DocumentBuilder(doc)
        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.paragraph_format.clear_formatting()
        builder.writeln('\n')
        # Style with name that starts from word InlineCode, followed by optional dot (.) and number of backticks (`).
        # If number of backticks is missed, then one backtick will be used by default.
        inline_code1_back_ticks = doc.styles.add(aw.StyleType.CHARACTER, 'InlineCode')
        builder.font.style = inline_code1_back_ticks
        builder.writeln('Text with InlineCode style with one backtick')
        # Use optional dot (.) and number of backticks (`).
        # There will be 3 backticks.
        inline_code3_back_ticks = doc.styles.add(aw.StyleType.CHARACTER, 'InlineCode.3')
        builder.font.style = inline_code3_back_ticks
        builder.writeln('Text with InlineCode style with 3 backticks')
        builder.document.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _test_markdown_document_headings(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
        builder = aw.DocumentBuilder(doc)
        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.paragraph_format.clear_formatting()
        builder.writeln('\n')
        # By default, Heading styles in Word may have bold and italic formatting.
        # If we do not want text to be emphasized, set these properties explicitly to False.
        # Thus we can't use 'builder.font.clear_formatting()' because bold/italic will be set to True.
        builder.font.bold = False
        builder.font.italic = False
        # Create for one heading for each level.
        builder.paragraph_format.style_name = 'Heading 1'
        builder.font.italic = True
        builder.writeln('This is an italic H1 tag')
        # Reset our styles from the previous paragraph to not combine styles between paragraphs.
        builder.font.bold = False
        builder.font.italic = False
        # Structure-enhanced text heading can be added through style inheritance.
        setext_heading1 = doc.styles.add(aw.StyleType.PARAGRAPH, 'SetextHeading1')
        builder.paragraph_format.style = setext_heading1
        doc.styles.get_by_name('SetextHeading1').base_style_name = 'Heading 1'
        builder.writeln('SetextHeading 1')
        builder.paragraph_format.style_name = 'Heading 2'
        builder.writeln('This is an H2 tag')
        builder.font.bold = False
        builder.font.italic = False
        setext_heading2 = doc.styles.add(aw.StyleType.PARAGRAPH, 'SetextHeading2')
        builder.paragraph_format.style = setext_heading2
        doc.styles.get_by_name('SetextHeading2').base_style_name = 'Heading 2'
        builder.writeln('SetextHeading 2')
        builder.paragraph_format.style = doc.styles.get_by_name('Heading 3')
        builder.writeln('This is an H3 tag')
        builder.font.bold = False
        builder.font.italic = False
        builder.paragraph_format.style = doc.styles.get_by_name('Heading 4')
        builder.font.bold = True
        builder.writeln('This is an bold H4 tag')
        builder.font.bold = False
        builder.font.italic = False
        builder.paragraph_format.style = doc.styles.get_by_name('Heading 5')
        builder.font.italic = True
        builder.font.bold = True
        builder.writeln('This is an italic and bold H5 tag')
        builder.font.bold = False
        builder.font.italic = False
        builder.paragraph_format.style = doc.styles.get_by_name('Heading 6')
        builder.writeln('This is an H6 tag')
        doc.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _test_markdown_document_blockquotes(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
        builder = aw.DocumentBuilder(doc)
        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.paragraph_format.clear_formatting()
        builder.writeln('\n')
        # By default, the document stores blockquote style for the first level.
        builder.paragraph_format.style_name = 'Quote'
        builder.writeln('Blockquote')
        # Create styles for nested levels through style inheritance.
        quote_level2 = doc.styles.add(aw.StyleType.PARAGRAPH, 'Quote1')
        builder.paragraph_format.style = quote_level2
        doc.styles.get_by_name('Quote1').base_style_name = 'Quote'
        builder.writeln('1. Nested blockquote')
        quote_level3 = doc.styles.add(aw.StyleType.PARAGRAPH, 'Quote2')
        builder.paragraph_format.style = quote_level3
        doc.styles.get_by_name('Quote2').base_style_name = 'Quote1'
        builder.font.italic = True
        builder.writeln('2. Nested italic blockquote')
        quote_level4 = doc.styles.add(aw.StyleType.PARAGRAPH, 'Quote3')
        builder.paragraph_format.style = quote_level4
        doc.styles.get_by_name('Quote3').base_style_name = 'Quote2'
        builder.font.italic = False
        builder.font.bold = True
        builder.writeln('3. Nested bold blockquote')
        quote_level5 = doc.styles.add(aw.StyleType.PARAGRAPH, 'Quote4')
        builder.paragraph_format.style = quote_level5
        doc.styles.get_by_name('Quote4').base_style_name = 'Quote3'
        builder.font.bold = False
        builder.writeln('4. Nested blockquote')
        quote_level6 = doc.styles.add(aw.StyleType.PARAGRAPH, 'Quote5')
        builder.paragraph_format.style = quote_level6
        doc.styles.get_by_name('Quote5').base_style_name = 'Quote4'
        builder.writeln('5. Nested blockquote')
        quote_level7 = doc.styles.add(aw.StyleType.PARAGRAPH, 'Quote6')
        builder.paragraph_format.style = quote_level7
        doc.styles.get_by_name('Quote6').base_style_name = 'Quote5'
        builder.font.italic = True
        builder.font.bold = True
        builder.writeln('6. Nested italic bold blockquote')
        doc.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _test_markdown_document_indented_code(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
        builder = aw.DocumentBuilder(doc)
        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.writeln('\n')
        builder.paragraph_format.clear_formatting()
        builder.writeln('\n')
        indented_code = doc.styles.add(aw.StyleType.PARAGRAPH, 'IndentedCode')
        builder.paragraph_format.style = indented_code
        builder.writeln('This is an indented code')
        doc.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _test_markdown_document_fenced_code(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
        builder = aw.DocumentBuilder(doc)
        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.writeln('\n')
        builder.paragraph_format.clear_formatting()
        builder.writeln('\n')
        fenced_code = doc.styles.add(aw.StyleType.PARAGRAPH, 'FencedCode')
        builder.paragraph_format.style = fenced_code
        builder.writeln('This is a fenced code')
        fenced_code_with_info = doc.styles.add(aw.StyleType.PARAGRAPH, 'FencedCode.C#')
        builder.paragraph_format.style = fenced_code_with_info
        builder.writeln('This is a fenced code with info string')
        doc.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _test_markdown_document_horizontal_rule(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
        builder = aw.DocumentBuilder(doc)
        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.paragraph_format.clear_formatting()
        builder.writeln('\n')
        # Insert HorizontalRule that will be present in .md file as '-----'.
        builder.insert_horizontal_rule()
        builder.document.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _test_markdown_document_bulleted_list(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
        builder = aw.DocumentBuilder(doc)
        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.paragraph_format.clear_formatting()
        builder.writeln('\n')
        # Bulleted lists are represented using paragraph numbering.
        builder.list_format.apply_bullet_default()
        # There can be 3 types of bulleted lists.
        # The only diff in a numbering format of the very first level are ‘-’, ‘+’ or ‘*’ respectively.
        builder.list_format.list.list_levels[0].number_format = '-'
        builder.writeln('Item 1')
        builder.writeln('Item 2')
        builder.list_format.list_indent()
        builder.writeln('Item 2a')
        builder.writeln('Item 2b')
        builder.document.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _load_markdown_document_and_assert_content(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        parameters = [('Italic', 'Normal', True, False), ('Bold', 'Normal', False, True), ('ItalicBold', 'Normal', True, True), ('Text with InlineCode style with one backtick', 'InlineCode', False, False), ('Text with InlineCode style with 3 backticks', 'InlineCode.3', False, False), ('This is an italic H1 tag', 'Heading 1', True, False), ('SetextHeading 1', 'SetextHeading1', False, False), ('This is an H2 tag', 'Heading 2', False, False), ('SetextHeading 2', 'SetextHeading2', False, False), ('This is an H3 tag', 'Heading 3', False, False), ('This is an bold H4 tag', 'Heading 4', False, True), ('This is an italic and bold H5 tag', 'Heading 5', True, True), ('This is an H6 tag', 'Heading 6', False, False), ('Blockquote', 'Quote', False, False), ('1. Nested blockquote', 'Quote1', False, False), ('2. Nested italic blockquote', 'Quote2', True, False), ('3. Nested bold blockquote', 'Quote3', False, True), ('4. Nested blockquote', 'Quote4', False, False), ('5. Nested blockquote', 'Quote5', False, False), ('6. Nested italic bold blockquote', 'Quote6', True, True), ('This is an indented code', 'IndentedCode', False, False), ('This is a fenced code', 'FencedCode', False, False), ('This is a fenced code with info string', 'FencedCode.C#', False, False), ('Item 1', 'Normal', False, False)]
        for text, style_name, is_italic, is_bold in parameters:
            with self.subTest(text=text, style_name=style_name, is_italic=is_italic, is_bold=is_bold):
                # Load created document from previous tests.
                doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
                paragraphs = doc.first_section.body.paragraphs
                for paragraph in paragraphs:
                    paragraph = paragraph.as_paragraph()
                    if paragraph.runs.count != 0:
                        # Check that all document text has the necessary styles.
                        if paragraph.runs[0].text == text and 'InlineCode' not in text:
                            self.assertEqual(style_name, paragraph.paragraph_format.style.name)
                            self.assertEqual(is_italic, paragraph.runs[0].font.italic)
                            self.assertEqual(is_bold, paragraph.runs[0].font.bold)
                        elif paragraph.runs[0].text == text and 'InlineCode' not in text:
                            self.assertEqual(style_name, paragraph.runs[0].font.style_name)
                    # Check that document also has a HorizontalRule present as a shape.
                    shapes_collection = doc.first_section.body.get_child_nodes(aw.NodeType.SHAPE, True)
                    horizontal_rule_shape = shapes_collection[0].as_shape()
                    self.assertTrue(shapes_collection.count == 1)
                    self.assertTrue(horizontal_rule_shape.is_horizontal_rule)

    def _test_markdown_document_emphases(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        builder = aw.DocumentBuilder()
        # Bold and Italic are represented as "Font.bold" and "Font.italic".
        builder.font.italic = True
        builder.writeln('This text will be italic')
        # Use clear formatting if we don't want to combine styles between paragraphs.
        builder.font.clear_formatting()
        builder.font.bold = True
        builder.writeln('This text will be bold')
        builder.font.clear_formatting()
        builder.font.italic = True
        builder.write('You ')
        builder.font.bold = True
        builder.write('can')
        builder.font.bold = False
        builder.writeln(' combine them')
        builder.font.clear_formatting()
        builder.font.strike_through = True
        builder.writeln('This text will be strikethrough')
        # Markdown treats asterisks (*), underscores (_) and tilde (~) as indicators of emphasis.
        builder.document.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _test_markdown_document_inline_code(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
        builder = aw.DocumentBuilder(doc)
        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.paragraph_format.clear_formatting()
        builder.writeln('\n')
        # Style with name that starts from word InlineCode, followed by optional dot (.) and number of backticks (`).
        # If number of backticks is missed, then one backtick will be used by default.
        inline_code1_back_ticks = doc.styles.add(aw.StyleType.CHARACTER, 'InlineCode')
        builder.font.style = inline_code1_back_ticks
        builder.writeln('Text with InlineCode style with one backtick')
        # Use optional dot (.) and number of backticks (`).
        # There will be 3 backticks.
        inline_code3_back_ticks = doc.styles.add(aw.StyleType.CHARACTER, 'InlineCode.3')
        builder.font.style = inline_code3_back_ticks
        builder.writeln('Text with InlineCode style with 3 backticks')
        builder.document.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _test_markdown_document_headings(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
        builder = aw.DocumentBuilder(doc)
        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.paragraph_format.clear_formatting()
        builder.writeln('\n')
        # By default, Heading styles in Word may have bold and italic formatting.
        # If we do not want text to be emphasized, set these properties explicitly to False.
        # Thus we can't use 'builder.font.clear_formatting()' because bold/italic will be set to True.
        builder.font.bold = False
        builder.font.italic = False
        # Create for one heading for each level.
        builder.paragraph_format.style_name = 'Heading 1'
        builder.font.italic = True
        builder.writeln('This is an italic H1 tag')
        # Reset our styles from the previous paragraph to not combine styles between paragraphs.
        builder.font.bold = False
        builder.font.italic = False
        # Structure-enhanced text heading can be added through style inheritance.
        setext_heading1 = doc.styles.add(aw.StyleType.PARAGRAPH, 'SetextHeading1')
        builder.paragraph_format.style = setext_heading1
        doc.styles.get_by_name('SetextHeading1').base_style_name = 'Heading 1'
        builder.writeln('SetextHeading 1')
        builder.paragraph_format.style_name = 'Heading 2'
        builder.writeln('This is an H2 tag')
        builder.font.bold = False
        builder.font.italic = False
        setext_heading2 = doc.styles.add(aw.StyleType.PARAGRAPH, 'SetextHeading2')
        builder.paragraph_format.style = setext_heading2
        doc.styles.get_by_name('SetextHeading2').base_style_name = 'Heading 2'
        builder.writeln('SetextHeading 2')
        builder.paragraph_format.style = doc.styles.get_by_name('Heading 3')
        builder.writeln('This is an H3 tag')
        builder.font.bold = False
        builder.font.italic = False
        builder.paragraph_format.style = doc.styles.get_by_name('Heading 4')
        builder.font.bold = True
        builder.writeln('This is an bold H4 tag')
        builder.font.bold = False
        builder.font.italic = False
        builder.paragraph_format.style = doc.styles.get_by_name('Heading 5')
        builder.font.italic = True
        builder.font.bold = True
        builder.writeln('This is an italic and bold H5 tag')
        builder.font.bold = False
        builder.font.italic = False
        builder.paragraph_format.style = doc.styles.get_by_name('Heading 6')
        builder.writeln('This is an H6 tag')
        doc.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _test_markdown_document_blockquotes(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
        builder = aw.DocumentBuilder(doc)
        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.paragraph_format.clear_formatting()
        builder.writeln('\n')
        # By default, the document stores blockquote style for the first level.
        builder.paragraph_format.style_name = 'Quote'
        builder.writeln('Blockquote')
        # Create styles for nested levels through style inheritance.
        quote_level2 = doc.styles.add(aw.StyleType.PARAGRAPH, 'Quote1')
        builder.paragraph_format.style = quote_level2
        doc.styles.get_by_name('Quote1').base_style_name = 'Quote'
        builder.writeln('1. Nested blockquote')
        quote_level3 = doc.styles.add(aw.StyleType.PARAGRAPH, 'Quote2')
        builder.paragraph_format.style = quote_level3
        doc.styles.get_by_name('Quote2').base_style_name = 'Quote1'
        builder.font.italic = True
        builder.writeln('2. Nested italic blockquote')
        quote_level4 = doc.styles.add(aw.StyleType.PARAGRAPH, 'Quote3')
        builder.paragraph_format.style = quote_level4
        doc.styles.get_by_name('Quote3').base_style_name = 'Quote2'
        builder.font.italic = False
        builder.font.bold = True
        builder.writeln('3. Nested bold blockquote')
        quote_level5 = doc.styles.add(aw.StyleType.PARAGRAPH, 'Quote4')
        builder.paragraph_format.style = quote_level5
        doc.styles.get_by_name('Quote4').base_style_name = 'Quote3'
        builder.font.bold = False
        builder.writeln('4. Nested blockquote')
        quote_level6 = doc.styles.add(aw.StyleType.PARAGRAPH, 'Quote5')
        builder.paragraph_format.style = quote_level6
        doc.styles.get_by_name('Quote5').base_style_name = 'Quote4'
        builder.writeln('5. Nested blockquote')
        quote_level7 = doc.styles.add(aw.StyleType.PARAGRAPH, 'Quote6')
        builder.paragraph_format.style = quote_level7
        doc.styles.get_by_name('Quote6').base_style_name = 'Quote5'
        builder.font.italic = True
        builder.font.bold = True
        builder.writeln('6. Nested italic bold blockquote')
        doc.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _test_markdown_document_indented_code(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
        builder = aw.DocumentBuilder(doc)
        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.writeln('\n')
        builder.paragraph_format.clear_formatting()
        builder.writeln('\n')
        indented_code = doc.styles.add(aw.StyleType.PARAGRAPH, 'IndentedCode')
        builder.paragraph_format.style = indented_code
        builder.writeln('This is an indented code')
        doc.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _test_markdown_document_fenced_code(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
        builder = aw.DocumentBuilder(doc)
        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.writeln('\n')
        builder.paragraph_format.clear_formatting()
        builder.writeln('\n')
        fenced_code = doc.styles.add(aw.StyleType.PARAGRAPH, 'FencedCode')
        builder.paragraph_format.style = fenced_code
        builder.writeln('This is a fenced code')
        fenced_code_with_info = doc.styles.add(aw.StyleType.PARAGRAPH, 'FencedCode.C#')
        builder.paragraph_format.style = fenced_code_with_info
        builder.writeln('This is a fenced code with info string')
        doc.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _test_markdown_document_horizontal_rule(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
        builder = aw.DocumentBuilder(doc)
        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.paragraph_format.clear_formatting()
        builder.writeln('\n')
        # Insert HorizontalRule that will be present in .md file as '-----'.
        builder.insert_horizontal_rule()
        builder.document.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _test_markdown_document_bulleted_list(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
        builder = aw.DocumentBuilder(doc)
        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.paragraph_format.clear_formatting()
        builder.writeln('\n')
        # Bulleted lists are represented using paragraph numbering.
        builder.list_format.apply_bullet_default()
        # There can be 3 types of bulleted lists.
        # The only diff in a numbering format of the very first level are ‘-’, ‘+’ or ‘*’ respectively.
        builder.list_format.list.list_levels[0].number_format = '-'
        builder.writeln('Item 1')
        builder.writeln('Item 2')
        builder.list_format.list_indent()
        builder.writeln('Item 2a')
        builder.writeln('Item 2b')
        builder.document.save(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')

    def _load_markdown_document_and_assert_content(self):
        """All markdown tests work with the same file. That's why we need order for them."""
        parameters = [('Italic', 'Normal', True, False), ('Bold', 'Normal', False, True), ('ItalicBold', 'Normal', True, True), ('Text with InlineCode style with one backtick', 'InlineCode', False, False), ('Text with InlineCode style with 3 backticks', 'InlineCode.3', False, False), ('This is an italic H1 tag', 'Heading 1', True, False), ('SetextHeading 1', 'SetextHeading1', False, False), ('This is an H2 tag', 'Heading 2', False, False), ('SetextHeading 2', 'SetextHeading2', False, False), ('This is an H3 tag', 'Heading 3', False, False), ('This is an bold H4 tag', 'Heading 4', False, True), ('This is an italic and bold H5 tag', 'Heading 5', True, True), ('This is an H6 tag', 'Heading 6', False, False), ('Blockquote', 'Quote', False, False), ('1. Nested blockquote', 'Quote1', False, False), ('2. Nested italic blockquote', 'Quote2', True, False), ('3. Nested bold blockquote', 'Quote3', False, True), ('4. Nested blockquote', 'Quote4', False, False), ('5. Nested blockquote', 'Quote5', False, False), ('6. Nested italic bold blockquote', 'Quote6', True, True), ('This is an indented code', 'IndentedCode', False, False), ('This is a fenced code', 'FencedCode', False, False), ('This is a fenced code with info string', 'FencedCode.C#', False, False), ('Item 1', 'Normal', False, False)]
        for text, style_name, is_italic, is_bold in parameters:
            with self.subTest(text=text, style_name=style_name, is_italic=is_italic, is_bold=is_bold):
                # Load created document from previous tests.
                doc = aw.Document(ARTIFACTS_DIR + 'DocumentBuilder.markdown_document.md')
                paragraphs = doc.first_section.body.paragraphs
                for paragraph in paragraphs:
                    paragraph = paragraph.as_paragraph()
                    if paragraph.runs.count != 0:
                        # Check that all document text has the necessary styles.
                        if paragraph.runs[0].text == text and 'InlineCode' not in text:
                            self.assertEqual(style_name, paragraph.paragraph_format.style.name)
                            self.assertEqual(is_italic, paragraph.runs[0].font.italic)
                            self.assertEqual(is_bold, paragraph.runs[0].font.bold)
                        elif paragraph.runs[0].text == text and 'InlineCode' not in text:
                            self.assertEqual(style_name, paragraph.runs[0].font.style_name)
                    # Check that document also has a HorizontalRule present as a shape.
                    shapes_collection = doc.first_section.body.get_child_nodes(aw.NodeType.SHAPE, True)
                    horizontal_rule_shape = shapes_collection[0].as_shape()
                    self.assertTrue(shapes_collection.count == 1)
                    self.assertTrue(horizontal_rule_shape.is_horizontal_rule)