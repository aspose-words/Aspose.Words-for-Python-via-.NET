from typing import List
import aspose.pydrawing
from aspose.words import Document, DocumentBuilder
from aspose.words.drawing import ShapeType
from aspose.words.replacing import FindReplaceOptions
import sys
# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import aspose.words.drawing
import aspose.words.notes
import aspose.words.replacing
import datetime
import system_helper
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExRange(ApiExampleBase):

    def test_replace(self):
        #ExStart
        #ExFor:Range.replace(str,str)
        #ExSummary:Shows how to perform a find-and-replace text operation on the contents of a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Greetings, _FullName_!')
        # Perform a find-and-replace operation on our document's contents and verify the number of replacements that took place.
        replacement_count = doc.range.replace(pattern='_FullName_', replacement='John Doe')
        self.assertEqual(1, replacement_count)
        self.assertEqual('Greetings, John Doe!', doc.get_text().strip())
        #ExEnd

    def test_replace_match_case(self):
        for match_case in [False, True]:
            #ExStart
            #ExFor:Range.replace(str,str,FindReplaceOptions)
            #ExFor:FindReplaceOptions
            #ExFor:FindReplaceOptions.match_case
            #ExSummary:Shows how to toggle case sensitivity when performing a find-and-replace operation.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.writeln('Ruby bought a ruby necklace.')
            # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            options = aw.replacing.FindReplaceOptions()
            # Set the "MatchCase" flag to "true" to apply case sensitivity while finding strings to replace.
            # Set the "MatchCase" flag to "false" to ignore character case while searching for text to replace.
            options.match_case = match_case
            doc.range.replace(pattern='Ruby', replacement='Jade', options=options)
            self.assertEqual('Jade bought a ruby necklace.' if match_case else 'Jade bought a Jade necklace.', doc.get_text().strip())
            #ExEnd

    def test_replace_find_whole_words_only(self):
        for find_whole_words_only in [False, True]:
            #ExStart
            #ExFor:Range.replace(str,str,FindReplaceOptions)
            #ExFor:FindReplaceOptions
            #ExFor:FindReplaceOptions.find_whole_words_only
            #ExSummary:Shows how to toggle standalone word-only find-and-replace operations.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.writeln('Jackson will meet you in Jacksonville.')
            # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            options = aw.replacing.FindReplaceOptions()
            # Set the "FindWholeWordsOnly" flag to "true" to replace the found text if it is not a part of another word.
            # Set the "FindWholeWordsOnly" flag to "false" to replace all text regardless of its surroundings.
            options.find_whole_words_only = find_whole_words_only
            doc.range.replace(pattern='Jackson', replacement='Louis', options=options)
            self.assertEqual('Louis will meet you in Jacksonville.' if find_whole_words_only else 'Louis will meet you in Louisville.', doc.get_text().strip())
            #ExEnd

    def test_ignore_deleted(self):
        for ignore_text_inside_delete_revisions in [True, False]:
            #ExStart
            #ExFor:FindReplaceOptions.ignore_deleted
            #ExSummary:Shows how to include or ignore text inside delete revisions during a find-and-replace operation.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.writeln('Hello world!')
            builder.writeln('Hello again!')
            # Start tracking revisions and remove the second paragraph, which will create a delete revision.
            # That paragraph will persist in the document until we accept the delete revision.
            doc.start_track_revisions(author='John Doe', date_time=datetime.datetime.now())
            doc.first_section.body.paragraphs[1].remove()
            doc.stop_track_revisions()
            self.assertTrue(doc.first_section.body.paragraphs[1].is_delete_revision)
            # We can use a "FindReplaceOptions" object to modify the find and replace process.
            options = aw.replacing.FindReplaceOptions()
            # Set the "IgnoreDeleted" flag to "true" to get the find-and-replace
            # operation to ignore paragraphs that are delete revisions.
            # Set the "IgnoreDeleted" flag to "false" to get the find-and-replace
            # operation to also search for text inside delete revisions.
            options.ignore_deleted = ignore_text_inside_delete_revisions
            doc.range.replace(pattern='Hello', replacement='Greetings', options=options)
            self.assertEqual('Greetings world!\rHello again!' if ignore_text_inside_delete_revisions else 'Greetings world!\rGreetings again!', doc.get_text().strip())
            #ExEnd

    def test_ignore_inserted(self):
        for ignore_text_inside_insert_revisions in [True, False]:
            #ExStart
            #ExFor:FindReplaceOptions.ignore_inserted
            #ExSummary:Shows how to include or ignore text inside insert revisions during a find-and-replace operation.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.writeln('Hello world!')
            # Start tracking revisions and insert a paragraph. That paragraph will be an insert revision.
            doc.start_track_revisions(author='John Doe', date_time=datetime.datetime.now())
            builder.writeln('Hello again!')
            doc.stop_track_revisions()
            self.assertTrue(doc.first_section.body.paragraphs[1].is_insert_revision)
            # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            options = aw.replacing.FindReplaceOptions()
            # Set the "IgnoreInserted" flag to "true" to get the find-and-replace
            # operation to ignore paragraphs that are insert revisions.
            # Set the "IgnoreInserted" flag to "false" to get the find-and-replace
            # operation to also search for text inside insert revisions.
            options.ignore_inserted = ignore_text_inside_insert_revisions
            doc.range.replace(pattern='Hello', replacement='Greetings', options=options)
            self.assertEqual('Greetings world!\rHello again!' if ignore_text_inside_insert_revisions else 'Greetings world!\rGreetings again!', doc.get_text().strip())
            #ExEnd

    def test_ignore_fields(self):
        for ignore_text_inside_fields in [True, False]:
            #ExStart
            #ExFor:FindReplaceOptions.ignore_fields
            #ExSummary:Shows how to ignore text inside fields.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.writeln('Hello world!')
            builder.insert_field(field_code='QUOTE', field_value='Hello again!')
            # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            options = aw.replacing.FindReplaceOptions()
            # Set the "IgnoreFields" flag to "true" to get the find-and-replace
            # operation to ignore text inside fields.
            # Set the "IgnoreFields" flag to "false" to get the find-and-replace
            # operation to also search for text inside fields.
            options.ignore_fields = ignore_text_inside_fields
            doc.range.replace(pattern='Hello', replacement='Greetings', options=options)
            self.assertEqual('Greetings world!\r\x13QUOTE\x14Hello again!\x15' if ignore_text_inside_fields else 'Greetings world!\r\x13QUOTE\x14Greetings again!\x15', doc.get_text().strip())
            #ExEnd

    def test_ignore_field_codes(self):
        for ignore_field_codes in [True, False]:
            #ExStart
            #ExFor:FindReplaceOptions.ignore_field_codes
            #ExSummary:Shows how to ignore text inside field codes.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.insert_field(field_code='INCLUDETEXT', field_value='Test IT!')
            options = aw.replacing.FindReplaceOptions()
            options.ignore_field_codes = ignore_field_codes
            # Replace 'T' in document ignoring text inside field code or not.
            doc.range.replace_regex(pattern='T', replacement='*', options=options)
            print(doc.get_text())
            self.assertEqual('\x13INCLUDETEXT\x14*est I*!\x15' if ignore_field_codes else '\x13INCLUDE*EX*\x14*est I*!\x15', doc.get_text().strip())
            #ExEnd

    def test_ignore_footnote(self):
        for is_ignore_footnotes in [True, False]:
            #ExStart
            #ExFor:FindReplaceOptions.ignore_footnotes
            #ExSummary:Shows how to ignore footnotes during a find-and-replace operation.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.write('Lorem ipsum dolor sit amet, consectetur adipiscing elit.')
            builder.insert_footnote(footnote_type=aw.notes.FootnoteType.FOOTNOTE, footnote_text='Lorem ipsum dolor sit amet, consectetur adipiscing elit.')
            builder.insert_paragraph()
            builder.write('Lorem ipsum dolor sit amet, consectetur adipiscing elit.')
            builder.insert_footnote(footnote_type=aw.notes.FootnoteType.ENDNOTE, footnote_text='Lorem ipsum dolor sit amet, consectetur adipiscing elit.')
            # Set the "IgnoreFootnotes" flag to "true" to get the find-and-replace
            # operation to ignore text inside footnotes.
            # Set the "IgnoreFootnotes" flag to "false" to get the find-and-replace
            # operation to also search for text inside footnotes.
            options = aw.replacing.FindReplaceOptions()
            options.ignore_footnotes = is_ignore_footnotes
            doc.range.replace(pattern='Lorem ipsum', replacement='Replaced Lorem ipsum', options=options)
            #ExEnd
            paragraphs = doc.first_section.body.paragraphs
            for para in paragraphs:
                para = para.as_paragraph()
                self.assertEqual('Replaced Lorem ipsum', para.runs[0].text)
            footnotes = list(map(lambda x: x.as_footnote(), list(doc.get_child_nodes(aw.NodeType.FOOTNOTE, True))))
            self.assertEqual('Lorem ipsum dolor sit amet, consectetur adipiscing elit.' if is_ignore_footnotes else 'Replaced Lorem ipsum dolor sit amet, consectetur adipiscing elit.', footnotes[0].to_string(save_format=aw.SaveFormat.TEXT).strip())
            self.assertEqual('Lorem ipsum dolor sit amet, consectetur adipiscing elit.' if is_ignore_footnotes else 'Replaced Lorem ipsum dolor sit amet, consectetur adipiscing elit.', footnotes[1].to_string(save_format=aw.SaveFormat.TEXT).strip())

    def test_ignore_shapes(self):
        #ExStart
        #ExFor:FindReplaceOptions.ignore_shapes
        #ExSummary:Shows how to ignore shapes while replacing text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.write('Lorem ipsum dolor sit amet, consectetur adipiscing elit.')
        builder.insert_shape(shape_type=aw.drawing.ShapeType.BALLOON, width=200, height=200)
        builder.write('Lorem ipsum dolor sit amet, consectetur adipiscing elit.')
        find_replace_options = aw.replacing.FindReplaceOptions()
        find_replace_options.ignore_shapes = True
        builder.document.range.replace(pattern='Lorem ipsum dolor sit amet, consectetur adipiscing elit.Lorem ipsum dolor sit amet, consectetur adipiscing elit.', replacement='Lorem ipsum dolor sit amet, consectetur adipiscing elit.', options=find_replace_options)
        self.assertEqual('Lorem ipsum dolor sit amet, consectetur adipiscing elit.', builder.document.get_text().strip())
        #ExEnd

    def test_update_fields_in_range(self):
        #ExStart
        #ExFor:Range.update_fields
        #ExSummary:Shows how to update all the fields in a range.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.insert_field(field_code=' DOCPROPERTY Category')
        builder.insert_break(aw.BreakType.SECTION_BREAK_EVEN_PAGE)
        builder.insert_field(field_code=' DOCPROPERTY Category')
        # The above DOCPROPERTY fields will display the value of this built-in document property.
        doc.built_in_document_properties.category = 'MyCategory'
        # If we update the value of a document property, we will need to update all the DOCPROPERTY fields to display it.
        self.assertEqual('', doc.range.fields[0].result)
        self.assertEqual('', doc.range.fields[1].result)
        # Update all the fields that are in the range of the first section.
        doc.first_section.range.update_fields()
        self.assertEqual('MyCategory', doc.range.fields[0].result)
        self.assertEqual('', doc.range.fields[1].result)
        #ExEnd

    def test_replace_with_string(self):
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('This one is sad.')
        builder.writeln('That one is mad.')
        options = aw.replacing.FindReplaceOptions()
        options.match_case = False
        options.find_whole_words_only = True
        doc.range.replace(pattern='sad', replacement='bad', options=options)
        doc.save(file_name=ARTIFACTS_DIR + 'Range.ReplaceWithString.docx')

    def test_replace_with_regex(self):
        #ExStart
        #ExFor:Range.replace_regex(Regex,str)
        #ExSummary:Shows how to replace all occurrences of a regular expression pattern with other text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('I decided to get the curtains in gray, ideal for the grey-accented room.')
        doc.range.replace_regex(pattern='gr(a|e)y', replacement='lavender')
        self.assertEqual('I decided to get the curtains in lavender, ideal for the lavender-accented room.', doc.get_text().strip())
        #ExEnd

    def test_apply_paragraph_format(self):
        #ExStart
        #ExFor:FindReplaceOptions.apply_paragraph_format
        #ExFor:Range.replace(str,str)
        #ExSummary:Shows how to add formatting to paragraphs in which a find-and-replace operation has found matches.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Every paragraph that ends with a full stop like this one will be right aligned.')
        builder.writeln('This one will not!')
        builder.write('This one also will.')
        paragraphs = doc.first_section.body.paragraphs
        self.assertEqual(aw.ParagraphAlignment.LEFT, paragraphs[0].paragraph_format.alignment)
        self.assertEqual(aw.ParagraphAlignment.LEFT, paragraphs[1].paragraph_format.alignment)
        self.assertEqual(aw.ParagraphAlignment.LEFT, paragraphs[2].paragraph_format.alignment)
        # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        options = aw.replacing.FindReplaceOptions()
        # Set the "Alignment" property to "ParagraphAlignment.Right" to right-align every paragraph
        # that contains a match that the find-and-replace operation finds.
        options.apply_paragraph_format.alignment = aw.ParagraphAlignment.RIGHT
        # Replace every full stop that is right before a paragraph break with an exclamation point.
        count = doc.range.replace(pattern='.&p', replacement='!&p', options=options)
        self.assertEqual(2, count)
        self.assertEqual(aw.ParagraphAlignment.RIGHT, paragraphs[0].paragraph_format.alignment)
        self.assertEqual(aw.ParagraphAlignment.LEFT, paragraphs[1].paragraph_format.alignment)
        self.assertEqual(aw.ParagraphAlignment.RIGHT, paragraphs[2].paragraph_format.alignment)
        self.assertEqual('Every paragraph that ends with a full stop like this one will be right aligned!\r' + 'This one will not!\r' + 'This one also will!', doc.get_text().strip())
        #ExEnd

    def test_delete_selection(self):
        #ExStart
        #ExFor:Node.range
        #ExFor:Range.delete
        #ExSummary:Shows how to delete all the nodes from a range.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Add text to the first section in the document, and then add another section.
        builder.write('Section 1. ')
        builder.insert_break(aw.BreakType.SECTION_BREAK_CONTINUOUS)
        builder.write('Section 2.')
        self.assertEqual('Section 1. \x0cSection 2.', doc.get_text().strip())
        # Remove the first section entirely by removing all the nodes
        # within its range, including the section itself.
        doc.sections[0].range.delete()
        self.assertEqual(1, doc.sections.count)
        self.assertEqual('Section 2.', doc.get_text().strip())
        #ExEnd

    def test_ranges_get_text(self):
        #ExStart
        #ExFor:Range
        #ExFor:Range.text
        #ExSummary:Shows how to get the text contents of all the nodes that a range covers.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.write('Hello world!')
        self.assertEqual('Hello world!', doc.range.text.strip())
        #ExEnd

    def test_use_substitutions(self):
        for use_substitutions in [False, True]:
            #ExStart
            #ExFor:FindReplaceOptions.use_substitutions
            #ExSummary:Shows how to replace the text with substitutions.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.writeln('John sold a car to Paul.')
            builder.writeln('Jane sold a house to Joe.')
            # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            options = aw.replacing.FindReplaceOptions()
            # Set the "UseSubstitutions" property to "true" to get
            # the find-and-replace operation to recognize substitution elements.
            # Set the "UseSubstitutions" property to "false" to ignore substitution elements.
            options.use_substitutions = use_substitutions
            regex = '([A-z]+) sold a ([A-z]+) to ([A-z]+)'
            doc.range.replace_regex(pattern=regex, replacement='$3 bought a $2 from $1', options=options)
            self.assertEqual('Paul bought a car from John.\rJoe bought a house from Jane.' if use_substitutions else '$3 bought a $2 from $1.\r$3 bought a $2 from $1.', doc.get_text().strip())
            #ExEnd

    @staticmethod
    def _insert_document(insertion_destination, doc_to_insert):
        if insertion_destination.node_type == aw.NodeType.PARAGRAPH or insertion_destination.node_type == aw.NodeType.TABLE:
            dst_story = insertion_destination.parent_node
            importer = aw.NodeImporter(src_doc=doc_to_insert, dst_doc=insertion_destination.document, import_format_mode=aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
            for src_section in filter(lambda a: a is not None, map(lambda b: system_helper.linq.Enumerable.of_type(lambda x: x.as_section(), b), list(doc_to_insert.sections))):
                for src_node in src_section.body:
                    # Skip the node if it is the last empty paragraph in a section.
                    if src_node.node_type == aw.NodeType.PARAGRAPH:
                        para = src_node.as_paragraph()
                        if para.is_end_of_section and (not para.has_child_nodes):
                            continue
                    new_node = importer.import_node(src_node, True)
                    dst_story.insert_after(new_node, insertion_destination)
                    insertion_destination = new_node
        else:
            raise Exception()
    #ExEnd

    @staticmethod
    def _test_insert_document_at_replace(doc):
        self.assertEqual('1) At text that can be identified by regex:\rHello World!\r' + '2) At a MERGEFIELD:\r\x13 MERGEFIELD  Document_1  \\* MERGEFORMAT \x14«Document_1»\x15\r' + '3) At a bookmark:', doc.first_section.body.get_text().strip())
    #ExStart:MatchEndNode
    #ExFor:ReplacingArgs.match_end_node
    #ExSummary:Shows how to get match end node.

    @unittest.skipIf(sys.platform.startswith('win'), 'Discrepancy in assertion between Python and .Net')
    def test_match_end_node(self):
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('1')
        builder.writeln('2')
        builder.writeln('3')
        replacing_callback = self.ReplacingCallback()
        options = aw.replacing.FindReplaceOptions()
        options.replacing_callback = replacing_callback
        doc.range.replace_regex(pattern='1[\\s\\S]*3', replacement='X', options=options)
        self.assertEqual('1', replacing_callback.start_node_text)
        self.assertEqual('3', replacing_callback.end_node_text)

    def test_ignore_office_math(self):
        for is_ignore_office_math in [True, False]:
            #ExStart:IgnoreOfficeMath
            #ExFor:FindReplaceOptions.ignore_office_math
            #ExSummary:Shows how to find and replace text within OfficeMath.
            doc = aw.Document(file_name=MY_DIR + 'Office math.docx')
            self.assertEqual('i+b-c≥iM+bM-cM', doc.first_section.body.first_paragraph.get_text().strip())
            options = aw.replacing.FindReplaceOptions()
            options.ignore_office_math = is_ignore_office_math
            doc.range.replace(pattern='b', replacement='x', options=options)
            if is_ignore_office_math:
                self.assertEqual('i+b-c≥iM+bM-cM', doc.first_section.body.first_paragraph.get_text().strip())
            else:
                self.assertEqual('i+x-c≥iM+xM-cM', doc.first_section.body.first_paragraph.get_text().strip())
            #ExEnd:IgnoreOfficeMath
    #ExStart
    #ExFor:FindReplaceOptions.replacing_callback
    #ExFor:Range.replace_regex(Regex,str,FindReplaceOptions)
    #ExFor:ReplacingArgs.replacement
    #ExFor:IReplacingCallback
    #ExFor:IReplacingCallback.replacing
    #ExFor:ReplacingArgs
    #ExSummary:Shows how to replace all occurrences of a regular expression pattern with another string, while tracking all such replacements (TextFindAndReplacementLogger).

    class TextFindAndReplacementLogger(aw.replacing.IReplacingCallback):

        def __init__(self):
            self.m_log = []

        def replacing(self, args):
            m_log.append(f'"{args.match.value}" converted to "{args.replacement}" {args.match_offset} characters into a {args.match_node.node_type} node.')
            args.replacement = f'(Old value:"{args.match.value}") {args.replacement}'
            return aw.Replacing.ReplaceAction.REPLACE

        def get_log(self):
            return str.join('', self.m_log)
    #ExEnd
    #ExStart
    #ExFor:FindReplaceOptions.apply_font
    #ExFor:FindReplaceOptions.replacing_callback
    #ExFor:ReplacingArgs.group_index
    #ExFor:ReplacingArgs.group_name
    #ExFor:ReplacingArgs.match
    #ExFor:ReplacingArgs.match_offset
    #ExSummary:Shows how to apply a different font to new content via FindReplaceOptions (NumberHexer).

    class NumberHexer(aw.replacing.IReplacingCallback):

        def __init__(self):
            self.m_current_replacement_number = None
            self.m_log = []

        def replacing(self, args):
            self.m_current_replacement_number += 1
            number = int(args.match.value)
            args.replacement = f'0x{number:X}'
            self.m_log.append(f'Match #{self.m_current_replacement_number}\n')
            self.m_log.append(f'\tOriginal value:\t{args.match.value}\n')
            self.m_log.append(f'\tReplacement:\t{args.replacement}\n')
            self.m_log.append(f'\tOffset in parent {args.match_node.node_type} node:\t{args.match_offset}\n')
            self.m_log.append(f'\tGroup index:\t{args.group_index}\n' if args.group_name is None or args.group_name == '' else f'\tGroup name:\t{args.group_name}\n')
            return aw.ReplaceAction.REPLACE

        def get_log(self):
            return str.join('', self.m_log)
    #ExEnd
    #ExStart
    #ExFor:FindReplaceOptions.use_legacy_order
    #ExSummary:Shows how to change the searching order of nodes when performing a find-and-replace text operation (TextReplacementTracker).

    class TextReplacementTracker(aw.replacing.IReplacingCallback):

        @property
        def matches(self):
            pass

        def replacing(self, e):
            matches.append(e.match.value)
            return aw.replacing.ReplaceAction.REPLACE
    #ExEnd
    #ExStart
    #ExFor:Range.replace_regex(Regex,str,FindReplaceOptions)
    #ExFor:IReplacingCallback
    #ExFor:ReplaceAction
    #ExFor:IReplacingCallback.replacing
    #ExFor:ReplacingArgs
    #ExFor:ReplacingArgs.match_node
    #ExSummary:Shows how to insert an entire document's contents as a replacement of a match in a find-and-replace operation (InsertDocumentAtReplaceHandler).

    class InsertDocumentAtReplaceHandler(aw.replacing.IReplacingCallback):

        def replacing(self, args):
            sub_doc = aw.Document(file_name=MY_DIR + 'Document.docx')
            # Insert a document after the paragraph containing the matched text.
            para = args.match_node.parent_node.as_paragraph()
            ExRange._insert_document(para, sub_doc)
            # Remove the paragraph with the matched text.
            para.remove()
            return aw.replacing.ReplaceAction.SKIP
    #ExStart
    #ExFor:FindReplaceOptions.direction
    #ExFor:FindReplaceDirection
    #ExSummary:Shows how to determine which direction a find-and-replace operation traverses the document in (TextReplacementRecorder).

    class TextReplacementRecorder(aw.replacing.IReplacingCallback):

        @property
        def matches(self):
            pass

        def replacing(self, e):
            matches.append(e.match.value)
            return aw.replacing.ReplaceAction.REPLACE
    #ExEnd

    class ReplacingCallback(aw.replacing.IReplacingCallback):

        @property
        def start_node_text(self):
            pass

        @start_node_text.setter
        def start_node_text(self, value):
            pass

        @property
        def end_node_text(self):
            pass

        @end_node_text.setter
        def end_node_text(self, value):
            pass

        def replacing(self, e):
            start_node_text = e.match_node.get_text().strip()
            end_node_text = e.match_end_node.get_text().strip()
            return aw.replacing.ReplaceAction.REPLACE
    #ExEnd:MatchEndNode

    def _test_insert_document_at_replace(self, doc: aw.Document):
        self.assertEqual('1) At text that can be identified by regex:\rHello World!\r' + '2) At a MERGEFIELD:\r\x13 MERGEFIELD  Document_1  \\* MERGEFORMAT \x14«Document_1»\x15\r' + '3) At a bookmark:', doc.first_section.body.get_text().strip())