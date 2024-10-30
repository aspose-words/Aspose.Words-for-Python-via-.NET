# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from aspose.words import Document, DocumentBuilder
from aspose.words.drawing import ShapeType
from aspose.words.replacing import FindReplaceOptions
from typing import List
import aspose.words as aw
import aspose.words.drawing
import aspose.words.notes
import aspose.words.replacing
import datetime
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

    def test_replace_match_case(self):
        for match_case in (False, True):
            with self.subTest(match_case=match_case):
                #ExStart
                #ExFor:Range.replace(str,str,FindReplaceOptions)
                #ExFor:FindReplaceOptions
                #ExFor:FindReplaceOptions.match_case
                #ExSummary:Shows how to toggle case sensitivity when performing a find-and-replace operation.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.writeln('Ruby bought a ruby necklace.')
                # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
                options = aw.replacing.FindReplaceOptions()
                # Set the "match_case" flag to "True" to apply case sensitivity while finding strings to replace.
                # Set the "match_case" flag to "False" to ignore character case while searching for text to replace.
                options.match_case = match_case
                doc.range.replace('Ruby', 'Jade', options)
                self.assertEqual('Jade bought a ruby necklace.' if match_case else 'Jade bought a Jade necklace.', doc.get_text().strip())
                #ExEnd

    def test_replace_find_whole_words_only(self):
        for find_whole_words_only in (False, True):
            with self.subTest(find_whole_words_on=find_whole_words_only):
                #ExStart
                #ExFor:Range.replace(str,str,FindReplaceOptions)
                #ExFor:FindReplaceOptions
                #ExFor:FindReplaceOptions.find_whole_words_only
                #ExSummary:Shows how to toggle standalone word-only find-and-replace operations.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.writeln('Jackson will meet you in Jacksonville.')
                # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
                options = aw.replacing.FindReplaceOptions()
                # Set the "find_whole_words_only" flag to "True" to replace the found text if it is not a part of another word.
                # Set the "find_whole_words_only" flag to "False" to replace all text regardless of its surroundings.
                options.find_whole_words_only = find_whole_words_only
                doc.range.replace('Jackson', 'Louis', options)
                self.assertEqual('Louis will meet you in Jacksonville.' if find_whole_words_only else 'Louis will meet you in Louisville.', doc.get_text().strip())
                #ExEnd

    def test_ignore_deleted(self):
        for ignore_text_inside_delete_revisions in (False, True):
            with self.subTest(ignore_text_inside_delete_revisions=ignore_text_inside_delete_revisions):
                #ExStart
                #ExFor:FindReplaceOptions.ignore_deleted
                #ExSummary:Shows how to include or ignore text inside delete revisions during a find-and-replace operation.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.writeln('Hello world!')
                builder.writeln('Hello again!')
                # Start tracking revisions and remove the second paragraph, which will create a delete revision.
                # That paragraph will persist in the document until we accept the delete revision.
                doc.start_track_revisions('John Doe', datetime.datetime.now())
                doc.first_section.body.paragraphs[1].remove()
                doc.stop_track_revisions()
                self.assertTrue(doc.first_section.body.paragraphs[1].is_delete_revision)
                # We can use a "FindReplaceOptions" object to modify the find and replace process.
                options = aw.replacing.FindReplaceOptions()
                # Set the "ignore_deleted" flag to "True" to get the find-and-replace
                # operation to ignore paragraphs that are delete revisions.
                # Set the "ignore_deleted" flag to "False" to get the find-and-replace
                # operation to also search for text inside delete revisions.
                options.ignore_deleted = ignore_text_inside_delete_revisions
                doc.range.replace('Hello', 'Greetings', options)
                self.assertEqual('Greetings world!\rHello again!' if ignore_text_inside_delete_revisions else 'Greetings world!\rGreetings again!', doc.get_text().strip())
                #ExEnd

    def test_ignore_inserted(self):
        for ignore_text_inside_insert_revisions in (True, False):
            with self.subTest(ignore_text_inside_insert_revisions=ignore_text_inside_insert_revisions):
                #ExStart
                #ExFor:FindReplaceOptions.ignore_inserted
                #ExSummary:Shows how to include or ignore text inside insert revisions during a find-and-replace operation.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.writeln('Hello world!')
                # Start tracking revisions and insert a paragraph. That paragraph will be an insert revision.
                doc.start_track_revisions('John Doe', datetime.datetime.now())
                builder.writeln('Hello again!')
                doc.stop_track_revisions()
                self.assertTrue(doc.first_section.body.paragraphs[1].is_insert_revision)
                # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
                options = aw.replacing.FindReplaceOptions()
                # Set the "ignore_inserted" flag to "True" to get the find-and-replace
                # operation to ignore paragraphs that are insert revisions.
                # Set the "ignore_inserted" flag to "False" to get the find-and-replace
                # operation to also search for text inside insert revisions.
                options.ignore_inserted = ignore_text_inside_insert_revisions
                doc.range.replace('Hello', 'Greetings', options)
                self.assertEqual('Greetings world!\rHello again!' if ignore_text_inside_insert_revisions else 'Greetings world!\rGreetings again!', doc.get_text().strip())
                #ExEnd

    def test_ignore_fields(self):
        for ignore_text_inside_fields in (True, False):
            with self.subTest(ignore_text_inside_fields=ignore_text_inside_fields):
                #ExStart
                #ExFor:FindReplaceOptions.ignore_fields
                #ExSummary:Shows how to ignore text inside fields.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.writeln('Hello world!')
                builder.insert_field('QUOTE', 'Hello again!')
                # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
                options = aw.replacing.FindReplaceOptions()
                # Set the "ignore_fields" flag to "True" to get the find-and-replace
                # operation to ignore text inside fields.
                # Set the "ignore_fields" flag to "False" to get the find-and-replace
                # operation to also search for text inside fields.
                options.ignore_fields = ignore_text_inside_fields
                doc.range.replace('Hello', 'Greetings', options)
                if ignore_text_inside_fields:
                    self.assertEqual('Greetings world!\r\x13QUOTE\x14Hello again!\x15', doc.get_text().strip())
                else:
                    self.assertEqual('Greetings world!\r\x13QUOTE\x14Greetings again!\x15', doc.get_text().strip())
                #ExEnd

    def test_ignore_field_codes(self):
        for ignore_field_codes in (True, False):
            with self.subTest(ignore_field_codes=ignore_field_codes):
                #ExStart
                #ExFor:FindReplaceOptions.ignore_field_codes
                #ExSummary:Shows how to ignore text inside field codes.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.insert_field('INCLUDETEXT', 'Test IT!')
                options = aw.replacing.FindReplaceOptions()
                options.ignore_field_codes = ignore_field_codes
                # Replace 'T' in document ignoring text inside field code or not.
                doc.range.replace_regex('T', '*', options)
                print(doc.get_text())
                if ignore_field_codes:
                    self.assertEqual('\x13INCLUDETEXT\x14*est I*!\x15', doc.get_text().strip())
                else:
                    self.assertEqual('\x13INCLUDE*EX*\x14*est I*!\x15', doc.get_text().strip())
                #ExEnd

    def test_ignore_footnote(self):
        for is_ignore_footnotes in (True, False):
            with self.subTest(is_ignore_footnotes=is_ignore_footnotes):
                #ExStart
                #ExFor:FindReplaceOptions.ignore_footnotes
                #ExSummary:Shows how to ignore footnotes during a find-and-replace operation.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.write('Lorem ipsum dolor sit amet, consectetur adipiscing elit.')
                builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.')
                builder.insert_paragraph()
                builder.write('Lorem ipsum dolor sit amet, consectetur adipiscing elit.')
                builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.')
                # Set the "ignore_footnotes" flag to "True" to get the find-and-replace
                # operation to ignore text inside footnotes.
                # Set the "ignore_footnotes" flag to "False" to get the find-and-replace
                # operation to also search for text inside footnotes.
                options = aw.replacing.FindReplaceOptions()
                options.ignore_footnotes = is_ignore_footnotes
                doc.range.replace('Lorem ipsum', 'Replaced Lorem ipsum', options)
                #ExEnd
                paragraphs = doc.first_section.body.paragraphs
                for para in paragraphs:
                    para = para.as_paragraph()
                    self.assertEqual('Replaced Lorem ipsum', para.runs[0].text)
                footnotes = [node.as_footnote() for node in doc.get_child_nodes(aw.NodeType.FOOTNOTE, True)]
                if is_ignore_footnotes:
                    expected_text = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.'
                else:
                    expected_text = 'Replaced Lorem ipsum dolor sit amet, consectetur adipiscing elit.'
                self.assertEqual(expected_text, footnotes[0].to_string(aw.SaveFormat.TEXT).strip())
                self.assertEqual(expected_text, footnotes[1].to_string(aw.SaveFormat.TEXT).strip())

    def test_replace_with_regex(self):
        #ExStart
        #ExFor:Range.replace(Regex,str)
        #ExSummary:Shows how to replace all occurrences of a regular expression pattern with other text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln('I decided to get the curtains in gray, ideal for the grey-accented room.')
        doc.range.replace_regex('gr(a|e)y', 'lavender')
        self.assertEqual('I decided to get the curtains in lavender, ideal for the lavender-accented room.', doc.get_text().strip())
        #ExEnd

    def test_use_substitutions(self):
        for use_substitutions in (False, True):
            with self.subTest(use_substitutions=use_substitutions):
                #ExStart
                #ExFor:FindReplaceOptions.use_substitutions
                #ExSummary:Shows how to replace the text with substitutions.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.writeln('John sold a car to Paul.')
                builder.writeln('Jane sold a house to Joe.')
                # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
                options = aw.replacing.FindReplaceOptions()
                # Set the "use_substitutions" property to "True" to get
                # the find-and-replace operation to recognize substitution elements.
                # Set the "use_substitutions" property to "False" to ignore substitution elements.
                options.use_substitutions = use_substitutions
                regex = '([A-z]+) sold a ([A-z]+) to ([A-z]+)'
                doc.range.replace_regex(regex, '$3 bought a $2 from $1', options)
                if use_substitutions:
                    self.assertEqual('Paul bought a car from John.\rJoe bought a house from Jane.', doc.get_text().strip())
                else:
                    self.assertEqual('$3 bought a $2 from $1.\r$3 bought a $2 from $1.', doc.get_text().strip())
                #ExEnd

    def _test_insert_document_at_replace(self, doc: aw.Document):
        self.assertEqual('1) At text that can be identified by regex:\rHello World!\r' + '2) At a MERGEFIELD:\r\x13 MERGEFIELD  Document_1  \\* MERGEFORMAT \x14«Document_1»\x15\r' + '3) At a bookmark:', doc.first_section.body.get_text().strip())