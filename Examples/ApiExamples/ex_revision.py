# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from datetime import date, timezone, timedelta
from document_helper import DocumentHelper
import aspose.words as aw
import aspose.words.comparing
import aspose.words.drawing
import aspose.words.layout
import aspose.words.notes
import datetime
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExRevision(ApiExampleBase):

    def test_get_info_about_revisions_in_revision_groups(self):
        #ExStart
        #ExFor:RevisionGroup
        #ExFor:RevisionGroup.author
        #ExFor:RevisionGroup.revision_type
        #ExFor:RevisionGroup.text
        #ExFor:RevisionGroupCollection
        #ExFor:RevisionGroupCollection.count
        #ExSummary:Shows how to print info about a group of revisions in a document.
        doc = aw.Document(file_name=MY_DIR + 'Revisions.docx')
        self.assertEqual(7, doc.revisions.groups.count)
        for group in doc.revisions.groups:
            print(f'Revision author: {group.author}; Revision type: {group.revision_type} \n\tRevision text: {group.text}')
        #ExEnd

    def test_get_specific_revision_group(self):
        #ExStart
        #ExFor:RevisionGroupCollection
        #ExFor:RevisionGroupCollection.__getitem__(int)
        #ExSummary:Shows how to get a group of revisions in a document.
        doc = aw.Document(file_name=MY_DIR + 'Revisions.docx')
        revision_group = doc.revisions.groups[0]
        #ExEnd
        self.assertEqual(aw.RevisionType.DELETION, revision_group.revision_type)
        self.assertEqual('Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ', revision_group.text)

    def test_show_revision_balloons(self):
        #ExStart
        #ExFor:RevisionOptions.show_in_balloons
        #ExSummary:Shows how to display revisions in balloons.
        doc = aw.Document(file_name=MY_DIR + 'Revisions.docx')
        # By default, text that is a revision has a different color to differentiate it from the other non-revision text.
        # Set a revision option to show more details about each revision in a balloon on the page's right margin.
        doc.layout_options.revision_options.show_in_balloons = aw.layout.ShowInBalloons.FORMAT_AND_DELETE
        doc.save(file_name=ARTIFACTS_DIR + 'Revision.ShowRevisionBalloons.pdf')
        #ExEnd

    def test_revision_options(self):
        #ExStart
        #ExFor:ShowInBalloons
        #ExFor:RevisionOptions.show_in_balloons
        #ExFor:RevisionOptions.comment_color
        #ExFor:RevisionOptions.deleted_text_color
        #ExFor:RevisionOptions.deleted_text_effect
        #ExFor:RevisionOptions.inserted_text_effect
        #ExFor:RevisionOptions.moved_from_text_color
        #ExFor:RevisionOptions.moved_from_text_effect
        #ExFor:RevisionOptions.moved_to_text_color
        #ExFor:RevisionOptions.moved_to_text_effect
        #ExFor:RevisionOptions.revised_properties_color
        #ExFor:RevisionOptions.revised_properties_effect
        #ExFor:RevisionOptions.revision_bars_color
        #ExFor:RevisionOptions.revision_bars_width
        #ExFor:RevisionOptions.show_original_revision
        #ExFor:RevisionOptions.show_revision_marks
        #ExFor:RevisionTextEffect
        #ExSummary:Shows how to modify the appearance of revisions.
        doc = aw.Document(file_name=MY_DIR + 'Revisions.docx')
        # Get the RevisionOptions object that controls the appearance of revisions.
        revision_options = doc.layout_options.revision_options
        # Render insertion revisions in green and italic.
        revision_options.inserted_text_color = aw.layout.RevisionColor.GREEN
        revision_options.inserted_text_effect = aw.layout.RevisionTextEffect.ITALIC
        # Render deletion revisions in red and bold.
        revision_options.deleted_text_color = aw.layout.RevisionColor.RED
        revision_options.deleted_text_effect = aw.layout.RevisionTextEffect.BOLD
        # The same text will appear twice in a movement revision:
        # once at the departure point and once at the arrival destination.
        # Render the text at the moved-from revision yellow with a double strike through
        # and double-underlined blue at the moved-to revision.
        revision_options.moved_from_text_color = aw.layout.RevisionColor.YELLOW
        revision_options.moved_from_text_effect = aw.layout.RevisionTextEffect.DOUBLE_STRIKE_THROUGH
        revision_options.moved_to_text_color = aw.layout.RevisionColor.CLASSIC_BLUE
        revision_options.moved_to_text_effect = aw.layout.RevisionTextEffect.DOUBLE_UNDERLINE
        # Render format revisions in dark red and bold.
        revision_options.revised_properties_color = aw.layout.RevisionColor.DARK_RED
        revision_options.revised_properties_effect = aw.layout.RevisionTextEffect.BOLD
        # Place a thick dark blue bar on the left side of the page next to lines affected by revisions.
        revision_options.revision_bars_color = aw.layout.RevisionColor.DARK_BLUE
        revision_options.revision_bars_width = 15
        # Show revision marks and original text.
        revision_options.show_original_revision = True
        revision_options.show_revision_marks = True
        # Get movement, deletion, formatting revisions, and comments to show up in green balloons
        # on the right side of the page.
        revision_options.show_in_balloons = aw.layout.ShowInBalloons.FORMAT
        revision_options.comment_color = aw.layout.RevisionColor.BRIGHT_GREEN
        # These features are only applicable to formats such as .pdf or .jpg.
        doc.save(file_name=ARTIFACTS_DIR + 'Revision.RevisionOptions.pdf')
        #ExEnd

    def test_accept_all_revisions(self):
        #ExStart
        #ExFor:Document.accept_all_revisions
        #ExSummary:Shows how to accept all tracking changes in the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Edit the document while tracking changes to create a few revisions.
        doc.start_track_revisions(author='John Doe')
        builder.write('Hello world! ')
        builder.write('Hello again! ')
        builder.write('This is another revision.')
        doc.stop_track_revisions()
        self.assertEqual(3, doc.revisions.count)
        # We can iterate through every revision and accept/reject it as a part of our document.
        # If we know we wish to accept every revision, we can do it more straightforwardly so by calling this method.
        doc.accept_all_revisions()
        self.assertEqual(0, doc.revisions.count)
        self.assertEqual('Hello world! Hello again! This is another revision.', doc.get_text().strip())
        #ExEnd

    def test_get_revised_properties_of_list(self):
        #ExStart
        #ExFor:RevisionsView
        #ExFor:Document.revisions_view
        #ExSummary:Shows how to switch between the revised and the original view of a document.
        doc = aw.Document(file_name=MY_DIR + 'Revisions at list levels.docx')
        doc.update_list_labels()
        paragraphs = doc.first_section.body.paragraphs
        self.assertEqual('1.', paragraphs[0].list_label.label_string)
        self.assertEqual('a.', paragraphs[1].list_label.label_string)
        self.assertEqual('', paragraphs[2].list_label.label_string)
        # View the document object as if all the revisions are accepted. Currently supports list labels.
        doc.revisions_view = aw.RevisionsView.FINAL
        self.assertEqual('', paragraphs[0].list_label.label_string)
        self.assertEqual('1.', paragraphs[1].list_label.label_string)
        self.assertEqual('a.', paragraphs[2].list_label.label_string)
        #ExEnd
        doc.revisions_view = aw.RevisionsView.ORIGINAL
        doc.accept_all_revisions()
        self.assertEqual('a.', paragraphs[0].list_label.label_string)
        self.assertEqual('', paragraphs[1].list_label.label_string)
        self.assertEqual('b.', paragraphs[2].list_label.label_string)

    def test_layout_options_revisions(self):
        #ExStart
        #ExFor:Document.layout_options
        #ExFor:LayoutOptions
        #ExFor:LayoutOptions.revision_options
        #ExFor:RevisionColor
        #ExFor:RevisionOptions
        #ExFor:RevisionOptions.inserted_text_color
        #ExFor:RevisionOptions.show_revision_bars
        #ExFor:RevisionOptions.revision_bars_position
        #ExSummary:Shows how to alter the appearance of revisions in a rendered output document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Insert a revision, then change the color of all revisions to green.
        builder.writeln('This is not a revision.')
        doc.start_track_revisions(author='John Doe', date_time=datetime.datetime.now())
        self.assertEqual(aw.layout.RevisionColor.BY_AUTHOR, doc.layout_options.revision_options.inserted_text_color)  #ExSkip
        self.assertTrue(doc.layout_options.revision_options.show_revision_bars)  #ExSkip
        builder.writeln('This is a revision.')
        doc.stop_track_revisions()
        builder.writeln('This is not a revision.')
        # Remove the bar that appears to the left of every revised line.
        doc.layout_options.revision_options.inserted_text_color = aw.layout.RevisionColor.BRIGHT_GREEN
        doc.layout_options.revision_options.show_revision_bars = False
        doc.layout_options.revision_options.revision_bars_position = aw.drawing.HorizontalAlignment.RIGHT
        doc.save(file_name=ARTIFACTS_DIR + 'Document.LayoutOptionsRevisions.pdf')
        #ExEnd

    def test_ignore_store_item_id(self):
        #ExStart:IgnoreStoreItemId
        #ExFor:AdvancedCompareOptions
        #ExFor:AdvancedCompareOptions.ignore_store_item_id
        #ExSummary:Shows how to compare SDT with same content but different store item id.
        doc_a = aw.Document(file_name=MY_DIR + 'Document with SDT 1.docx')
        doc_b = aw.Document(file_name=MY_DIR + 'Document with SDT 2.docx')
        # Configure options to compare SDT with same content but different store item id.
        compare_options = aw.comparing.CompareOptions()
        compare_options.advanced_options.ignore_store_item_id = False
        doc_a.compare(document=doc_b, author='user', date_time=datetime.datetime.now(), options=compare_options)
        self.assertEqual(8, doc_a.revisions.count)
        compare_options.advanced_options.ignore_store_item_id = True
        doc_a.revisions.reject_all()
        doc_a.compare(document=doc_b, author='user', date_time=datetime.datetime.now(), options=compare_options)
        self.assertEqual(0, doc_a.revisions.count)
        #ExEnd:IgnoreStoreItemId

    def test_revisions(self):
        #ExStart
        #ExFor:Revision
        #ExFor:Revision.accept
        #ExFor:Revision.author
        #ExFor:Revision.date_time
        #ExFor:Revision.group
        #ExFor:Revision.reject
        #ExFor:Revision.revision_type
        #ExFor:RevisionCollection
        #ExFor:RevisionCollection.__getitem__(int)
        #ExFor:RevisionCollection.count
        #ExFor:RevisionType
        #ExFor:Document.has_revisions
        #ExFor:Document.track_revisions
        #ExFor:Document.revisions
        #ExSummary:Shows how to work with revisions in a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Normal editing of the document does not count as a revision.
        builder.write('This does not count as a revision. ')
        self.assertFalse(doc.has_revisions)
        # To register our edits as revisions, we need to declare an author, and then start tracking them.
        doc.start_track_revisions('John Doe', datetime.datetime.now())
        builder.write('This is revision #1. ')
        self.assertTrue(doc.has_revisions)
        self.assertEqual(1, doc.revisions.count)
        # This flag corresponds to the "Review" -> "Tracking" -> "Track Changes" option in Microsoft Word.
        # The "start_track_revisions" method does not affect its value,
        # and the document is tracking revisions programmatically despite it having a value of "False".
        # If we open this document using Microsoft Word, it will not be tracking revisions.
        self.assertFalse(doc.track_revisions)
        # We have added text using the document builder, so the first revision is an insertion-type revision.
        revision = doc.revisions[0]
        self.assertEqual('John Doe', revision.author)
        self.assertEqual('This is revision #1. ', revision.parent_node.get_text())
        self.assertEqual(aw.RevisionType.INSERTION, revision.revision_type)
        self.assertEqual(revision.date_time.date(), date.today())
        self.assertEqual(doc.revisions.groups[0], revision.group)
        # Remove a run to create a deletion-type revision.
        doc.first_section.body.first_paragraph.runs[0].remove()
        # Adding a new revision places it at the beginning of the revision collection.
        self.assertEqual(aw.RevisionType.DELETION, doc.revisions[0].revision_type)
        self.assertEqual(2, doc.revisions.count)
        # Insert revisions show up in the document body even before we accept/reject the revision.
        # Rejecting the revision will remove its nodes from the body. Conversely, nodes that make up delete revisions
        # also linger in the document until we accept the revision.
        self.assertEqual('This does not count as a revision. This is revision #1.', doc.get_text().strip())
        # Accepting the delete revision will remove its parent node from the paragraph text
        # and then remove the collection's revision itself.
        doc.revisions[0].accept()
        self.assertEqual(1, doc.revisions.count)
        self.assertEqual('This is revision #1.', doc.get_text().strip())
        builder.writeln('')
        builder.write('This is revision #2.')
        # Now move the node to create a moving revision type.
        node = doc.first_section.body.paragraphs[1]
        end_node = doc.first_section.body.paragraphs[1].next_sibling
        reference_node = doc.first_section.body.paragraphs[0]
        while node != end_node:
            next_node = node.next_sibling
            doc.first_section.body.insert_before(node, reference_node)
            node = next_node
        self.assertEqual(aw.RevisionType.MOVING, doc.revisions[0].revision_type)
        self.assertEqual(8, doc.revisions.count)
        self.assertEqual('This is revision #2.\rThis is revision #1. \rThis is revision #2.', doc.get_text().strip())
        # The moving revision is now at index 1. Reject the revision to discard its contents.
        doc.revisions[1].reject()
        self.assertEqual(6, doc.revisions.count)
        self.assertEqual('This is revision #1. \rThis is revision #2.', doc.get_text().strip())
        #ExEnd

    def test_revision_collection(self):
        #ExStart
        #ExFor:Revision.parent_style
        #ExFor:RevisionCollection.__iter__
        #ExFor:RevisionCollection.groups
        #ExFor:RevisionCollection.reject_all
        #ExFor:RevisionGroupCollection.__iter__
        #ExSummary:Shows how to work with a document's collection of revisions.
        doc = aw.Document(MY_DIR + 'Revisions.docx')
        revisions = doc.revisions
        # This collection itself has a collection of revision groups.
        # Each group is a sequence of adjacent revisions.
        self.assertEqual(7, revisions.groups.count)  #ExSkip
        print(revisions.groups.count, 'revision groups:')
        # Iterate over the collection of groups and print the text that the revision concerns.
        for group in revisions.groups:
            print(f'\tGroup type "{group.revision_type}", ' + f'author: {group.author}, contents: [{group.text.strip()}]')
        # Each Run that a revision affects gets a corresponding Revision object.
        # The revisions' collection is considerably larger than the condensed form we printed above,
        # depending on how many Runs we have segmented the document into during Microsoft Word editing.
        self.assertEqual(11, revisions.count)  #ExSkip
        print(f'\n{revisions.count} revisions:')
        for revision in revisions:
            # A StyleDefinitionChange strictly affects styles and not document nodes. This means the "parent_style"
            # property will always be in use, while the "parent_node" will always be None.
            # Since all other changes affect nodes, "parent_node" will conversely be in use, and "parent_style" will be None.
            if revision.revision_type == aw.RevisionType.STYLE_DEFINITION_CHANGE:
                print(f'\tRevision type "{revision.revision_type}", ' + f'author: {revision.author}, style: [{revision.parent_style.name}]')
            else:
                print(f'\tRevision type "{revision.revision_type}", ' + f'author: {revision.author}, contents: [{revision.parent_node.get_text().strip()}]')
        # Reject all revisions via the collection, reverting the document to its original form.
        revisions.reject_all()
        self.assertEqual(0, revisions.count)
        #ExEnd

    def test_track_revisions(self):
        #ExStart
        #ExFor:Document.start_track_revisions(str)
        #ExFor:Document.start_track_revisions(str,datetime)
        #ExFor:Document.stop_track_revisions
        #ExSummary:Shows how to track revisions while editing a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Editing a document usually does not count as a revision until we begin tracking them.
        builder.write('Hello world! ')
        self.assertEqual(0, doc.revisions.count)
        self.assertFalse(doc.first_section.body.paragraphs[0].runs[0].is_insert_revision)
        doc.start_track_revisions('John Doe')
        builder.write('Hello again! ')
        self.assertEqual(1, doc.revisions.count)
        self.assertTrue(doc.first_section.body.paragraphs[0].runs[1].is_insert_revision)
        self.assertEqual('John Doe', doc.revisions[0].author)
        self.assertAlmostEqual(doc.revisions[0].date_time, datetime.datetime.now(tz=timezone.utc), delta=timedelta(seconds=1))
        # Stop tracking revisions to not count any future edits as revisions.
        doc.stop_track_revisions()
        builder.write('Hello again! ')
        self.assertEqual(1, doc.revisions.count)
        self.assertFalse(doc.first_section.body.paragraphs[0].runs[2].is_insert_revision)
        # Creating revisions gives them a date and time of the operation.
        # We can disable this by passing "datetime.min" when we start tracking revisions.
        doc.start_track_revisions('John Doe', datetime.datetime.min)
        builder.write('Hello again! ')
        self.assertEqual(2, doc.revisions.count)
        self.assertEqual('John Doe', doc.revisions[1].author)
        self.assertEqual(datetime.datetime.min, doc.revisions[1].date_time)
        # We can accept/reject these revisions programmatically
        # by calling methods such as "Document.accept_all_revisions", or each revision's "accept" method.
        # In Microsoft Word, we can process them manually via "Review" -> "Changes".
        doc.save(ARTIFACTS_DIR + 'Document.track_revisions.docx')
        #ExEnd

    def test_compare(self):
        #ExStart
        #ExFor:Document.compare(Document,str,datetime)
        #ExFor:RevisionCollection.accept_all
        #ExSummary:Shows how to compare documents.
        doc_original = aw.Document()
        builder = aw.DocumentBuilder(doc_original)
        builder.writeln('This is the original document.')
        doc_edited = aw.Document()
        builder = aw.DocumentBuilder(doc_edited)
        builder.writeln('This is the edited document.')
        # Comparing documents with revisions will throw an exception.
        if doc_original.revisions.count == 0 and doc_edited.revisions.count == 0:
            doc_original.compare(doc_edited, 'authorName', datetime.datetime.now())
        # After the comparison, the original document will gain a new revision
        # for every element that is different in the edited document.
        self.assertEqual(2, doc_original.revisions.count)  # ExSkip
        for revision in doc_original.revisions:
            print(f'Revision type: {revision.revision_type}, on a node of type "{revision.parent_node.node_type}"')
            print(f'\tChanged text: "{revision.parent_node.get_text()}"')
        # Accepting these revisions will transform the original document into the edited document.
        doc_original.revisions.accept_all()
        self.assertEqual(doc_original.get_text(), doc_edited.get_text())
        #ExEnd
        doc_original = DocumentHelper.save_open(doc_original)
        self.assertEqual(0, doc_original.revisions.count)

    def test_compare_document_with_revisions(self):
        doc1 = aw.Document()
        builder = aw.DocumentBuilder(doc1)
        builder.writeln('Hello world! This text is not a revision.')
        doc_with_revision = aw.Document()
        builder = aw.DocumentBuilder(doc_with_revision)
        doc_with_revision.start_track_revisions('John Doe')
        builder.writeln('This is a revision.')
        with self.assertRaises(Exception):
            doc_with_revision.compare(doc1, 'John Doe', datetime.datetime.now())

    def test_compare_options(self):
        #ExStart
        #ExFor:CompareOptions
        #ExFor:CompareOptions.ignore_formatting
        #ExFor:CompareOptions.ignore_case_changes
        #ExFor:CompareOptions.ignore_comments
        #ExFor:CompareOptions.ignore_tables
        #ExFor:CompareOptions.ignore_fields
        #ExFor:CompareOptions.ignore_footnotes
        #ExFor:CompareOptions.ignore_textboxes
        #ExFor:CompareOptions.ignore_headers_and_footers
        #ExFor:CompareOptions.target
        #ExFor:ComparisonTargetType
        #ExFor:Document.compare(Document,str,datetime,CompareOptions)
        #ExSummary:Shows how to filter specific types of document elements when making a comparison.
        # Create the original document and populate it with various kinds of elements.
        doc_original = aw.Document()
        builder = aw.DocumentBuilder(doc_original)
        # Paragraph text referenced with an endnote:
        builder.writeln('Hello world! This is the first paragraph.')
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, 'Original endnote text.')
        # Table:
        builder.start_table()
        builder.insert_cell()
        builder.write('Original cell 1 text')
        builder.insert_cell()
        builder.write('Original cell 2 text')
        builder.end_table()
        # Textbox:
        text_box = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 150, 20)
        builder.move_to(text_box.first_paragraph)
        builder.write('Original textbox contents')
        # DATE field:
        builder.move_to(doc_original.first_section.body.append_paragraph(''))
        builder.insert_field(' DATE ')
        # Comment:
        new_comment = aw.Comment(doc_original, 'John Doe', 'J.D.', datetime.datetime.now())
        new_comment.set_text('Original comment.')
        builder.current_paragraph.append_child(new_comment)
        # Header:
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.writeln('Original header contents.')
        # Create a clone of our document and perform a quick edit on each of the cloned document's elements.
        doc_edited = doc_original.clone(True).as_document()
        first_paragraph = doc_edited.first_section.body.first_paragraph
        first_paragraph.runs[0].text = 'hello world! this is the first paragraph, after editing.'
        first_paragraph.paragraph_format.style = doc_edited.styles.get_by_style_identifier(aw.StyleIdentifier.HEADING1)
        doc_edited.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote().first_paragraph.runs[1].text = 'Edited endnote text.'
        doc_edited.get_child(aw.NodeType.TABLE, 0, True).as_table().first_row.cells[1].first_paragraph.runs[0].text = 'Edited Cell 2 contents'
        doc_edited.get_child(aw.NodeType.SHAPE, 0, True).as_shape().first_paragraph.runs[0].text = 'Edited textbox contents'
        doc_edited.range.fields[0].as_field_date().use_lunar_calendar = True
        doc_edited.get_child(aw.NodeType.COMMENT, 0, True).as_comment().first_paragraph.runs[0].text = 'Edited comment.'
        doc_edited.first_section.headers_footers.header_primary.first_paragraph.runs[0].text = 'Edited header contents.'
        # Comparing documents creates a revision for every edit in the edited document.
        # A CompareOptions object has a series of flags that can suppress revisions
        # on each respective type of element, effectively ignoring their change.
        compare_options = aw.comparing.CompareOptions()
        compare_options.ignore_formatting = False
        compare_options.ignore_case_changes = False
        compare_options.ignore_comments = False
        compare_options.ignore_tables = False
        compare_options.ignore_fields = False
        compare_options.ignore_footnotes = False
        compare_options.ignore_textboxes = False
        compare_options.ignore_headers_and_footers = False
        compare_options.target = aw.comparing.ComparisonTargetType.NEW
        doc_original.compare(doc_edited, 'John Doe', datetime.datetime.now(), compare_options)
        doc_original.save(ARTIFACTS_DIR + 'Document.compare_options.docx')
        #ExEnd
        doc_original = aw.Document(ARTIFACTS_DIR + 'Document.compare_options.docx')
        self.verify_footnote(aw.notes.FootnoteType.ENDNOTE, True, '', 'OriginalEdited endnote text.', doc_original.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote())

    def test_ignore_dml_unique_id(self):
        for is_ignore_dml_unique_id in (False, True):
            with self.subTest(is_ignore_dml_unique_id=is_ignore_dml_unique_id):
                #ExStart
                #ExFor:CompareOptions.ignore_dml_unique_id
                #ExSummary:Shows how to compare documents ignoring DML unique ID.
                doc_a = aw.Document(MY_DIR + 'DML unique ID original.docx')
                doc_b = aw.Document(MY_DIR + 'DML unique ID compare.docx')
                # By default, Aspose.Words do not ignore DML's unique ID, and the revisions count was 2.
                # If we are ignoring DML's unique ID, and revisions count were 0.
                compare_options = aw.comparing.CompareOptions()
                compare_options.ignore_dml_unique_id = is_ignore_dml_unique_id
                doc_a.compare(doc_b, 'Aspose.Words', datetime.datetime.now(), compare_options)
                self.assertEqual(0 if is_ignore_dml_unique_id else 2, doc_a.revisions.count)
                #ExEnd

    def test_granularity_compare_option(self):
        for granularity in (aw.comparing.Granularity.CHAR_LEVEL, aw.comparing.Granularity.WORD_LEVEL):
            with self.subTest(granularity=granularity):
                #ExStart
                #ExFor:CompareOptions.granularity
                #ExFor:Granularity
                #ExSummary:Shows to specify a granularity while comparing documents.
                doc_a = aw.Document()
                builder_a = aw.DocumentBuilder(doc_a)
                builder_a.writeln('Alpha Lorem ipsum dolor sit amet, consectetur adipiscing elit')
                doc_b = aw.Document()
                builder_b = aw.DocumentBuilder(doc_b)
                builder_b.writeln('Lorems ipsum dolor sit amet consectetur - "adipiscing" elit')
                # Specify whether changes are tracking
                # by character ('Granularity.CHAR_LEVEL'), or by word ('Granularity.WORD_LEVEL').
                compare_options = aw.comparing.CompareOptions()
                compare_options.granularity = granularity
                doc_a.compare(doc_b, 'author', datetime.datetime.now(), compare_options)
                # The first document's collection of revision groups contains all the differences between documents.
                groups = doc_a.revisions.groups
                self.assertEqual(5, groups.count)
                #ExEnd
                if granularity == aw.comparing.Granularity.CHAR_LEVEL:
                    self.assertEqual(aw.RevisionType.DELETION, groups[0].revision_type)
                    self.assertEqual('Alpha ', groups[0].text)
                    self.assertEqual(aw.RevisionType.DELETION, groups[1].revision_type)
                    self.assertEqual(',', groups[1].text)
                    self.assertEqual(aw.RevisionType.INSERTION, groups[2].revision_type)
                    self.assertEqual('s', groups[2].text)
                    self.assertEqual(aw.RevisionType.INSERTION, groups[3].revision_type)
                    self.assertEqual('- "', groups[3].text)
                    self.assertEqual(aw.RevisionType.INSERTION, groups[4].revision_type)
                    self.assertEqual('"', groups[4].text)
                else:
                    self.assertEqual(aw.RevisionType.DELETION, groups[0].revision_type)
                    self.assertEqual('Alpha Lorem', groups[0].text)
                    self.assertEqual(aw.RevisionType.DELETION, groups[1].revision_type)
                    self.assertEqual(',', groups[1].text)
                    self.assertEqual(aw.RevisionType.INSERTION, groups[2].revision_type)
                    self.assertEqual('Lorems', groups[2].text)
                    self.assertEqual(aw.RevisionType.INSERTION, groups[3].revision_type)
                    self.assertEqual('- "', groups[3].text)
                    self.assertEqual(aw.RevisionType.INSERTION, groups[4].revision_type)
                    self.assertEqual('"', groups[4].text)
