# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from datetime import date, timedelta
from document_helper import DocumentHelper
import sys
import aspose.pydrawing
import aspose.words as aw
import aspose.words.fields
import datetime
import document_helper
import test_util
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExParagraph(ApiExampleBase):

    def test_document_builder_insert_paragraph(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_paragraph
        #ExFor:ParagraphFormat.first_line_indent
        #ExFor:ParagraphFormat.alignment
        #ExFor:ParagraphFormat.keep_together
        #ExFor:ParagraphFormat.add_space_between_far_east_and_alpha
        #ExFor:ParagraphFormat.add_space_between_far_east_and_digit
        #ExFor:Paragraph.is_end_of_document
        #ExSummary:Shows how to insert a paragraph into the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        font = builder.font
        font.size = 16
        font.bold = True
        font.color = aspose.pydrawing.Color.blue
        font.name = 'Arial'
        font.underline = aw.Underline.DASH
        paragraph_format = builder.paragraph_format
        paragraph_format.first_line_indent = 8
        paragraph_format.alignment = aw.ParagraphAlignment.JUSTIFY
        paragraph_format.add_space_between_far_east_and_alpha = True
        paragraph_format.add_space_between_far_east_and_digit = True
        paragraph_format.keep_together = True
        # The "Writeln" method ends the paragraph after appending text
        # and then starts a new line, adding a new paragraph.
        builder.writeln('Hello world!')
        self.assertTrue(builder.current_paragraph.is_end_of_document)
        #ExEnd
        doc = document_helper.DocumentHelper.save_open(doc)
        paragraph = doc.first_section.body.first_paragraph
        self.assertEqual(8, paragraph.paragraph_format.first_line_indent)
        self.assertEqual(aw.ParagraphAlignment.JUSTIFY, paragraph.paragraph_format.alignment)
        self.assertTrue(paragraph.paragraph_format.add_space_between_far_east_and_alpha)
        self.assertTrue(paragraph.paragraph_format.add_space_between_far_east_and_digit)
        self.assertTrue(paragraph.paragraph_format.keep_together)
        self.assertEqual('Hello world!', paragraph.get_text().strip())
        run_font = paragraph.runs[0].font
        self.assertEqual(16, run_font.size)
        self.assertTrue(run_font.bold)
        self.assertEqual(aspose.pydrawing.Color.blue.to_argb(), run_font.color.to_argb())
        self.assertEqual('Arial', run_font.name)
        self.assertEqual(aw.Underline.DASH, run_font.underline)

    def test_composite_node_children(self):
        #ExStart
        #ExFor:CompositeNode.count
        #ExFor:CompositeNode.get_child_nodes(NodeType,bool)
        #ExFor:CompositeNode.insert_after
        #ExFor:CompositeNode.insert_before
        #ExFor:CompositeNode.prepend_child
        #ExFor:Paragraph.get_text
        #ExFor:Run
        #ExSummary:Shows how to add, update and delete child nodes in a CompositeNode's collection of children.
        doc = aw.Document()
        # An empty document, by default, has one paragraph.
        self.assertEqual(1, doc.first_section.body.paragraphs.count)
        # Composite nodes such as our paragraph can contain other composite and inline nodes as children.
        paragraph = doc.first_section.body.first_paragraph
        paragraph_text = aw.Run(doc=doc, text='Initial text. ')
        paragraph.append_child(paragraph_text)
        # Create three more run nodes.
        run1 = aw.Run(doc=doc, text='Run 1. ')
        run2 = aw.Run(doc=doc, text='Run 2. ')
        run3 = aw.Run(doc=doc, text='Run 3. ')
        # The document body will not display these runs until we insert them into a composite node
        # that itself is a part of the document's node tree, as we did with the first run.
        # We can determine where the text contents of nodes that we insert
        # appears in the document by specifying an insertion location relative to another node in the paragraph.
        self.assertEqual('Initial text.', paragraph.get_text().strip())
        # Insert the second run into the paragraph in front of the initial run.
        paragraph.insert_before(run2, paragraph_text)
        self.assertEqual('Run 2. Initial text.', paragraph.get_text().strip())
        # Insert the third run after the initial run.
        paragraph.insert_after(run3, paragraph_text)
        self.assertEqual('Run 2. Initial text. Run 3.', paragraph.get_text().strip())
        # Insert the first run to the start of the paragraph's child nodes collection.
        paragraph.prepend_child(run1)
        self.assertEqual('Run 1. Run 2. Initial text. Run 3.', paragraph.get_text().strip())
        self.assertEqual(4, paragraph.get_child_nodes(aw.NodeType.ANY, True).count)
        # We can modify the contents of the run by editing and deleting existing child nodes.
        paragraph.get_child_nodes(aw.NodeType.RUN, True)[1].as_run().text = 'Updated run 2. '
        paragraph.get_child_nodes(aw.NodeType.RUN, True).remove(paragraph_text)
        self.assertEqual('Run 1. Updated run 2. Run 3.', paragraph.get_text().strip())
        self.assertEqual(3, paragraph.get_child_nodes(aw.NodeType.ANY, True).count)
        #ExEnd

    def test_move_revisions(self):
        #ExStart
        #ExFor:Paragraph.is_move_from_revision
        #ExFor:Paragraph.is_move_to_revision
        #ExFor:ParagraphCollection
        #ExFor:ParagraphCollection.__getitem__(int)
        #ExFor:Story.paragraphs
        #ExSummary:Shows how to check whether a paragraph is a move revision.
        doc = aw.Document(file_name=MY_DIR + 'Revisions.docx')
        # This document contains "Move" revisions, which appear when we highlight text with the cursor,
        # and then drag it to move it to another location
        # while tracking revisions in Microsoft Word via "Review" -> "Track changes".
        self.assertEqual(6, len(list(filter(lambda r: r.revision_type == aw.RevisionType.MOVING, doc.revisions))))
        paragraphs = doc.first_section.body.paragraphs
        # Move revisions consist of pairs of "Move from", and "Move to" revisions.
        # These revisions are potential changes to the document that we can either accept or reject.
        # Before we accept/reject a move revision, the document
        # must keep track of both the departure and arrival destinations of the text.
        # The second and the fourth paragraph define one such revision, and thus both have the same contents.
        self.assertEqual(paragraphs[1].get_text(), paragraphs[3].get_text())
        # The "Move from" revision is the paragraph where we dragged the text from.
        # If we accept the revision, this paragraph will disappear,
        # and the other will remain and no longer be a revision.
        self.assertTrue(paragraphs[1].is_move_from_revision)
        # The "Move to" revision is the paragraph where we dragged the text to.
        # If we reject the revision, this paragraph instead will disappear, and the other will remain.
        self.assertTrue(paragraphs[3].is_move_to_revision)
        #ExEnd

    def test_range_revisions(self):
        #ExStart
        #ExFor:Range.revisions
        #ExSummary:Shows how to work with revisions in range.
        doc = aw.Document(file_name=MY_DIR + 'Revisions.docx')
        paragraph = doc.first_section.body.first_paragraph
        for revision in paragraph.range.revisions:
            if revision.revision_type == aw.RevisionType.DELETION:
                revision.accept()
        # Reject the first section revisions.
        doc.first_section.range.revisions.reject_all()
        #ExEnd

    def test_get_format_revision(self):
        #ExStart
        #ExFor:Paragraph.is_format_revision
        #ExSummary:Shows how to check whether a paragraph is a format revision.
        doc = aw.Document(file_name=MY_DIR + 'Format revision.docx')
        # This paragraph is a "Format" revision, which occurs when we change the formatting of existing text
        # while tracking revisions in Microsoft Word via "Review" -> "Track changes".
        self.assertTrue(doc.first_section.body.first_paragraph.is_format_revision)
        #ExEnd

    def test_is_revision(self):
        #ExStart
        #ExFor:Paragraph.is_delete_revision
        #ExFor:Paragraph.is_insert_revision
        #ExSummary:Shows how to work with revision paragraphs.
        doc = aw.Document()
        body = doc.first_section.body
        para = body.first_paragraph
        para.append_child(aw.Run(doc=doc, text='Paragraph 1. '))
        body.append_paragraph('Paragraph 2. ')
        body.append_paragraph('Paragraph 3. ')
        # The above paragraphs are not revisions.
        # Paragraphs that we add after starting revision tracking will register as "Insert" revisions.
        doc.start_track_revisions(author='John Doe', date_time=datetime.datetime.now())
        para = body.append_paragraph('Paragraph 4. ')
        self.assertTrue(para.is_insert_revision)
        # Paragraphs that we remove after starting revision tracking will register as "Delete" revisions.
        paragraphs = body.paragraphs
        self.assertEqual(4, paragraphs.count)
        para = paragraphs[2]
        para.remove()
        # Such paragraphs will remain until we either accept or reject the delete revision.
        # Accepting the revision will remove the paragraph for good,
        # and rejecting the revision will leave it in the document as if we never deleted it.
        self.assertEqual(4, paragraphs.count)
        self.assertTrue(para.is_delete_revision)
        # Accept the revision, and then verify that the paragraph is gone.
        doc.accept_all_revisions()
        self.assertEqual(3, paragraphs.count)
        self.assertEqual(0, para.count)
        self.assertEqual('Paragraph 1. \r' + 'Paragraph 2. \r' + 'Paragraph 4.', doc.get_text().strip())
        #ExEnd

    def test_break_is_style_separator(self):
        #ExStart
        #ExFor:Paragraph.break_is_style_separator
        #ExSummary:Shows how to write text to the same line as a TOC heading and have it not show up in the TOC.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.insert_table_of_contents('\\o \\h \\z \\u')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        # Insert a paragraph with a style that the TOC will pick up as an entry.
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
        # Both these strings are in the same paragraph and will therefore show up on the same TOC entry.
        builder.write('Heading 1. ')
        builder.write('Will appear in the TOC. ')
        # If we insert a style separator, we can write more text in the same paragraph
        # and use a different style without showing up in the TOC.
        # If we use a heading type style after the separator, we can draw multiple TOC entries from one document text line.
        builder.insert_style_separator()
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.QUOTE
        builder.write("Won't appear in the TOC. ")
        self.assertTrue(doc.first_section.body.first_paragraph.break_is_style_separator)
        doc.update_fields()
        doc.save(file_name=ARTIFACTS_DIR + 'Paragraph.BreakIsStyleSeparator.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Paragraph.BreakIsStyleSeparator.docx')
        test_util.TestUtil.verify_field(expected_type=aw.fields.FieldType.FIELD_TOC, expected_field_code='TOC \\o \\h \\z \\u', expected_result='\x13 HYPERLINK \\l "_Toc256000000" \x14Heading 1. Will appear in the TOC.\t\x13 PAGEREF _Toc256000000 \\h \x142\x15\x15\r', field=doc.range.fields[0])
        self.assertFalse(doc.first_section.body.first_paragraph.break_is_style_separator)

    def test_tab_stops(self):
        #ExStart
        #ExFor:TabLeader
        #ExFor:TabAlignment
        #ExFor:Paragraph.get_effective_tab_stops
        #ExSummary:Shows how to set custom tab stops for a paragraph.
        doc = aw.Document()
        para = doc.first_section.body.first_paragraph
        # If we are in a paragraph with no tab stops in this collection,
        # the cursor will jump 36 points each time we press the Tab key in Microsoft Word.
        self.assertEqual(0, len(doc.first_section.body.first_paragraph.get_effective_tab_stops()))
        # We can add custom tab stops in Microsoft Word if we enable the ruler via the "View" tab.
        # Each unit on this ruler is two default tab stops, which is 72 points.
        # We can add custom tab stops programmatically like this.
        tab_stops = doc.first_section.body.first_paragraph.paragraph_format.tab_stops
        tab_stops.add(position=72, alignment=aw.TabAlignment.LEFT, leader=aw.TabLeader.DOTS)
        tab_stops.add(position=216, alignment=aw.TabAlignment.CENTER, leader=aw.TabLeader.DASHES)
        tab_stops.add(position=360, alignment=aw.TabAlignment.RIGHT, leader=aw.TabLeader.LINE)
        # We can see these tab stops in Microsoft Word by enabling the ruler via "View" -> "Show" -> "Ruler".
        self.assertEqual(3, len(para.get_effective_tab_stops()))
        # Any tab characters we add will make use of the tab stops on the ruler and may,
        # depending on the tab leader's value, leave a line between the tab departure and arrival destinations.
        para.append_child(aw.Run(doc=doc, text='\tTab 1\tTab 2\tTab 3'))
        doc.save(file_name=ARTIFACTS_DIR + 'Paragraph.TabStops.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Paragraph.TabStops.docx')
        tab_stops = doc.first_section.body.first_paragraph.paragraph_format.tab_stops
        test_util.TestUtil.verify_tab_stop(72, aw.TabAlignment.LEFT, aw.TabLeader.DOTS, False, tab_stops[0])
        test_util.TestUtil.verify_tab_stop(216, aw.TabAlignment.CENTER, aw.TabLeader.DASHES, False, tab_stops[1])
        test_util.TestUtil.verify_tab_stop(360, aw.TabAlignment.RIGHT, aw.TabLeader.LINE, False, tab_stops[2])

    def test_join_runs(self):
        #ExStart
        #ExFor:Paragraph.join_runs_with_same_formatting
        #ExSummary:Shows how to simplify paragraphs by merging superfluous runs.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert four runs of text into the paragraph.
        builder.write('Run 1. ')
        builder.write('Run 2. ')
        builder.write('Run 3. ')
        builder.write('Run 4. ')
        # If we open this document in Microsoft Word, the paragraph will look like one seamless text body.
        # However, it will consist of four separate runs with the same formatting. Fragmented paragraphs like this
        # may occur when we manually edit parts of one paragraph many times in Microsoft Word.
        para = builder.current_paragraph
        self.assertEqual(4, para.runs.count)
        # Change the style of the last run to set it apart from the first three.
        para.runs[3].font.style_identifier = aw.StyleIdentifier.EMPHASIS
        # We can run the "JoinRunsWithSameFormatting" method to optimize the document's contents
        # by merging similar runs into one, reducing their overall count.
        # This method also returns the number of runs that this method merged.
        # These two merges occurred to combine Runs #1, #2, and #3,
        # while leaving out Run #4 because it has an incompatible style.
        self.assertEqual(2, para.join_runs_with_same_formatting())
        # The number of runs left will equal the original count
        # minus the number of run merges that the "JoinRunsWithSameFormatting" method carried out.
        self.assertEqual(2, para.runs.count)
        self.assertEqual('Run 1. Run 2. Run 3. ', para.runs[0].text)
        self.assertEqual('Run 4. ', para.runs[1].text)
        #ExEnd

    @unittest.skipUnless(sys.platform.startswith('win'), 'windows date time parameters')
    def test_append_field(self):
        #ExStart
        #ExFor:Paragraph.append_field(FieldType,bool)
        #ExFor:Paragraph.append_field(str)
        #ExFor:Paragraph.append_field(str,str)
        #ExSummary:Shows various ways of appending fields to a paragraph.
        doc = aw.Document()
        paragraph = doc.first_section.body.first_paragraph
        # Below are three ways of appending a field to the end of a paragraph.
        # 1 -  Append a DATE field using a field type, and then update it:
        paragraph.append_field(aw.fields.FieldType.FIELD_DATE, True)
        # 2 -  Append a TIME field using a field code:
        paragraph.append_field(' TIME  \\@ "HH:mm:ss" ')
        # 3 -  Append a QUOTE field using a field code, and get it to display a placeholder value:
        paragraph.append_field(' QUOTE "Real value"', 'Placeholder value')
        self.assertEqual('Placeholder value', doc.range.fields[2].result)
        # This field will display its placeholder value until we update it.
        doc.update_fields()
        self.assertEqual('Real value', doc.range.fields[2].result)
        doc.save(ARTIFACTS_DIR + 'Paragraph.append_field.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Paragraph.append_field.docx')
        self.verify_datetime_field(aw.fields.FieldType.FIELD_DATE, ' DATE ', datetime.datetime.now(), doc.range.fields[0], timedelta())
        self.verify_datetime_field(aw.fields.FieldType.FIELD_TIME, ' TIME  \\@ "HH:mm:ss" ', datetime.datetime.now(), doc.range.fields[1], timedelta(seconds=5))
        self.verify_field(aw.fields.FieldType.FIELD_QUOTE, ' QUOTE "Real value"', 'Real value', doc.range.fields[2])

    def test_insert_field(self):
        #ExStart
        #ExFor:Paragraph.insert_field(str,Node,bool)
        #ExFor:Paragraph.insert_field(FieldType,bool,Node,bool)
        #ExFor:Paragraph.insert_field(str,str,Node,bool)
        #ExSummary:Shows various ways of adding fields to a paragraph.
        doc = aw.Document()
        para = doc.first_section.body.first_paragraph
        # Below are three ways of inserting a field into a paragraph.
        # 1 -  Insert an AUTHOR field into a paragraph after one of the paragraph's child nodes:
        run = aw.Run(doc)
        run.text = 'This run was written by '
        para.append_child(run)
        doc.built_in_document_properties.get_by_name('Author').value = 'John Doe'
        para.insert_field(aw.fields.FieldType.FIELD_AUTHOR, True, run, True)
        # 2 -  Insert a QUOTE field after one of the paragraph's child nodes:
        run = aw.Run(doc)
        run.text = '.'
        para.append_child(run)
        field = para.insert_field(' QUOTE " Real value" ', run, True)
        # 3 -  Insert a QUOTE field before one of the paragraph's child nodes,
        # and get it to display a placeholder value:
        para.insert_field(' QUOTE " Real value."', ' Placeholder value.', field.start, False)
        self.assertEqual(' Placeholder value.', doc.range.fields[1].result)
        # This field will display its placeholder value until we update it.
        doc.update_fields()
        self.assertEqual(' Real value.', doc.range.fields[1].result)
        doc.save(ARTIFACTS_DIR + 'Paragraph.insert_field.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Paragraph.insert_field.docx')
        self.verify_field(aw.fields.FieldType.FIELD_AUTHOR, ' AUTHOR ', 'John Doe', doc.range.fields[0])
        self.verify_field(aw.fields.FieldType.FIELD_QUOTE, ' QUOTE " Real value."', ' Real value.', doc.range.fields[1])
        self.verify_field(aw.fields.FieldType.FIELD_QUOTE, ' QUOTE " Real value" ', ' Real value', doc.range.fields[2])

    def test_insert_field_before_text_in_paragraph(self):
        doc = DocumentHelper.create_document_fill_with_dummy_text()
        ExParagraph.insert_field_using_field_code(doc, ' AUTHOR ', None, False, 1)
        self.assertEqual('\x13 AUTHOR \x14Test Author\x15Hello World!\r', DocumentHelper.get_paragraph_text(doc, 1))

    @unittest.skipUnless(sys.platform.startswith('win'), 'windows date time parameters')
    def test_insert_field_after_text_in_paragraph(self):
        today = date.today().strftime('%d/%m/%Y').lstrip('0')
        doc = DocumentHelper.create_document_fill_with_dummy_text()
        ExParagraph.insert_field_using_field_code(doc, ' DATE ', None, True, 1)
        self.assertEqual(f'Hello World!\x13 DATE \x14{today}\x15\r', DocumentHelper.get_paragraph_text(doc, 1))

    def test_insert_field_before_text_in_paragraph_without_update_field(self):
        doc = DocumentHelper.create_document_fill_with_dummy_text()
        ExParagraph.insert_field_using_field_type(doc, aw.fields.FieldType.FIELD_AUTHOR, False, None, False, 1)
        self.assertEqual('\x13 AUTHOR \x14\x15Hello World!\r', DocumentHelper.get_paragraph_text(doc, 1))

    def test_insert_field_after_text_in_paragraph_without_update_field(self):
        doc = DocumentHelper.create_document_fill_with_dummy_text()
        ExParagraph.insert_field_using_field_type(doc, aw.fields.FieldType.FIELD_AUTHOR, False, None, True, 1)
        self.assertEqual('Hello World!\x13 AUTHOR \x14\x15\r', DocumentHelper.get_paragraph_text(doc, 1))

    def test_insert_field_without_separator(self):
        doc = DocumentHelper.create_document_fill_with_dummy_text()
        ExParagraph.insert_field_using_field_type(doc, aw.fields.FieldType.FIELD_LIST_NUM, True, None, False, 1)
        self.assertEqual('\x13 LISTNUM \x15Hello World!\r', DocumentHelper.get_paragraph_text(doc, 1))

    def test_insert_field_before_paragraph_without_document_author(self):
        doc = DocumentHelper.create_document_fill_with_dummy_text()
        doc.built_in_document_properties.author = ''
        ExParagraph.insert_field_using_field_code_field_string(doc, ' AUTHOR ', None, None, False, 1)
        self.assertEqual('\x13 AUTHOR \x14\x15Hello World!\r', DocumentHelper.get_paragraph_text(doc, 1))

    def test_insert_field_after_paragraph_without_changing_document_author(self):
        doc = DocumentHelper.create_document_fill_with_dummy_text()
        ExParagraph.insert_field_using_field_code_field_string(doc, ' AUTHOR ', None, None, True, 1)
        self.assertEqual('Hello World!\x13 AUTHOR \x14\x15\r', DocumentHelper.get_paragraph_text(doc, 1))

    def test_insert_field_before_run_text(self):
        doc = DocumentHelper.create_document_fill_with_dummy_text()
        #Add some text into the paragraph
        run = DocumentHelper.insert_new_run(doc, ' Hello World!', 1)
        ExParagraph.insert_field_using_field_code_field_string(doc, ' AUTHOR ', 'Test Field Value', run, False, 1)
        self.assertEqual('Hello World!\x13 AUTHOR \x14Test Field Value\x15 Hello World!\r', DocumentHelper.get_paragraph_text(doc, 1))

    def test_insert_field_after_run_text(self):
        doc = DocumentHelper.create_document_fill_with_dummy_text()
        # Add some text into the paragraph
        run = DocumentHelper.insert_new_run(doc, ' Hello World!', 1)
        ExParagraph.insert_field_using_field_code_field_string(doc, ' AUTHOR ', '', run, True, 1)
        self.assertEqual('Hello World! Hello World!\x13 AUTHOR \x14\x15\r', DocumentHelper.get_paragraph_text(doc, 1))

    def test_insert_field_empty_paragraph_without_update_field(self):
        doc = DocumentHelper.create_document_without_dummy_text()
        ExParagraph.insert_field_using_field_type(doc, aw.fields.FieldType.FIELD_AUTHOR, False, None, False, 1)
        self.assertEqual('\x13 AUTHOR \x14\x15\x0c', DocumentHelper.get_paragraph_text(doc, 1))

    def test_insert_field_empty_paragraph_with_update_field(self):
        doc = DocumentHelper.create_document_without_dummy_text()
        ExParagraph.insert_field_using_field_type(doc, aw.fields.FieldType.FIELD_AUTHOR, True, None, False, 0)
        self.assertEqual('\x13 AUTHOR \x14Test Author\x15\r', DocumentHelper.get_paragraph_text(doc, 0))

    def test_get_frame_properties(self):
        #ExStart
        #ExFor:Paragraph.frame_format
        #ExFor:FrameFormat
        #ExFor:FrameFormat.is_frame
        #ExFor:FrameFormat.width
        #ExFor:FrameFormat.height
        #ExFor:FrameFormat.height_rule
        #ExFor:FrameFormat.horizontal_alignment
        #ExFor:FrameFormat.vertical_alignment
        #ExFor:FrameFormat.horizontal_position
        #ExFor:FrameFormat.relative_horizontal_position
        #ExFor:FrameFormat.horizontal_distance_from_text
        #ExFor:FrameFormat.vertical_position
        #ExFor:FrameFormat.relative_vertical_position
        #ExFor:FrameFormat.vertical_distance_from_text
        #ExSummary:Shows how to get information about formatting properties of paragraphs that are frames.
        doc = aw.Document(MY_DIR + 'Paragraph frame.docx')
        for paragraph in doc.first_section.body.paragraphs:
            paragraph = paragraph.as_paragraph()
            if paragraph.frame_format.is_frame:
                paragraph_frame = paragraph
                break
        self.assertEqual(233.3, paragraph_frame.frame_format.width)
        self.assertEqual(138.8, paragraph_frame.frame_format.height)
        self.assertEqual(aw.HeightRule.AT_LEAST, paragraph_frame.frame_format.height_rule)
        self.assertEqual(aw.drawing.HorizontalAlignment.DEFAULT, paragraph_frame.frame_format.horizontal_alignment)
        self.assertEqual(aw.drawing.VerticalAlignment.DEFAULT, paragraph_frame.frame_format.vertical_alignment)
        self.assertEqual(34.05, paragraph_frame.frame_format.horizontal_position)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.PAGE, paragraph_frame.frame_format.relative_horizontal_position)
        self.assertEqual(9.0, paragraph_frame.frame_format.horizontal_distance_from_text)
        self.assertEqual(20.5, paragraph_frame.frame_format.vertical_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, paragraph_frame.frame_format.relative_vertical_position)
        self.assertEqual(0.0, paragraph_frame.frame_format.vertical_distance_from_text)
        #ExEnd

    @staticmethod
    def insert_field_using_field_type(doc: aw.Document, field_type: aw.fields.FieldType, update_field: bool, ref_node: aw.Node, is_after: bool, para_index: int):
        """Insert field into the first paragraph of the current document using field type."""
        para = DocumentHelper.get_paragraph(doc, para_index)
        para.insert_field(field_type, update_field, ref_node, is_after)

    @staticmethod
    def insert_field_using_field_code(doc: aw.Document, field_code: str, ref_node: aw.Node, is_after: bool, para_index: int):
        """Insert field into the first paragraph of the current document using field code."""
        para = DocumentHelper.get_paragraph(doc, para_index)
        para.insert_field(field_code, ref_node, is_after)

    @staticmethod
    def insert_field_using_field_code_field_string(doc: aw.Document, field_code: str, field_value: str, ref_node: aw.Node, is_after: bool, para_index: int):
        """Insert field into the first paragraph of the current document using field code and field String."""
        para = DocumentHelper.get_paragraph(doc, para_index)
        para.insert_field(field_code, field_value, ref_node, is_after)