# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import unittest
import io
from datetime import date, datetime

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR
from testutil import TestUtil

class ExInlineStory(ApiExampleBase):

    def test_position_footnote(self):

        for footnote_position in (aw.notes.FootnotePosition.BENEATH_TEXT,
                                  aw.notes.FootnotePosition.BOTTOM_OF_PAGE):
            with self.subTest(footnote_position=footnote_position):
                #ExStart
                #ExFor:Document.footnote_options
                #ExFor:FootnoteOptions
                #ExFor:FootnoteOptions.position
                #ExFor:FootnotePosition
                #ExSummary:Shows how to select a different place where the document collects and displays its footnotes.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # A footnote is a way to attach a reference or a side comment to text
                # that does not interfere with the main body text's flow.
                # Inserting a footnote adds a small superscript reference symbol
                # at the main body text where we insert the footnote.
                # Each footnote also creates an entry at the bottom of the page, consisting of a symbol
                # that matches the reference symbol in the main body text.
                # The reference text that we pass to the document builder's "insert_footnote" method.
                builder.write("Hello world!")
                builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Footnote contents.")

                # We can use the "position" property to determine where the document will place all its footnotes.
                # If we set the value of the "position" property to "FootnotePosition.BOTTOM_OF_PAGE",
                # every footnote will show up at the bottom of the page that contains its reference mark. This is the default value.
                # If we set the value of the "position" property to "FootnotePosition.BENEATH_TEXT",
                # every footnote will show up at the end of the page's text that contains its reference mark.
                doc.footnote_options.position = footnote_position

                doc.save(ARTIFACTS_DIR + "InlineStory.position_footnote.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "InlineStory.position_footnote.docx")

                self.assertEqual(footnote_position, doc.footnote_options.position)

                TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, True, "",
                    "Footnote contents.", doc.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote())

    def test_position_endnote(self):

        for endnote_position in (aw.notes.EndnotePosition.END_OF_DOCUMENT,
                                 aw.notes.EndnotePosition.END_OF_SECTION):
            with self.subTest(endnote_position=endnote_position):
                #ExStart
                #ExFor:Document.endnote_options
                #ExFor:EndnoteOptions
                #ExFor:EndnoteOptions.position
                #ExFor:EndnotePosition
                #ExSummary:Shows how to select a different place where the document collects and displays its endnotes.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # An endnote is a way to attach a reference or a side comment to text
                # that does not interfere with the main body text's flow.
                # Inserting an endnote adds a small superscript reference symbol
                # at the main body text where we insert the endnote.
                # Each endnote also creates an entry at the end of the document, consisting of a symbol
                # that matches the reference symbol in the main body text.
                # The reference text that we pass to the document builder's "insert_endnote" method.
                builder.write("Hello world!")
                builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Endnote contents.")
                builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
                builder.write("This is the second section.")

                # We can use the "position" property to determine where the document will place all its endnotes.
                # If we set the value of the "position" property to "EndnotePosition.END_OF_DOCUMENT",
                # every footnote will show up in a collection at the end of the document. This is the default value.
                # If we set the value of the "position" property to "EndnotePosition.END_OF_SECTION",
                # every footnote will show up in a collection at the end of the section whose text contains the endnote's reference mark.
                doc.endnote_options.position = endnote_position

                doc.save(ARTIFACTS_DIR + "InlineStory.position_endnote.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "InlineStory.position_endnote.docx")

                self.assertEqual(endnote_position, doc.endnote_options.position)

                TestUtil.verify_footnote(self, aw.notes.FootnoteType.ENDNOTE, True, "",
                    "Endnote contents.", doc.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote())

    def test_ref_mark_number_style(self):

        #ExStart
        #ExFor:Document.endnote_options
        #ExFor:EndnoteOptions
        #ExFor:EndnoteOptions.number_style
        #ExFor:Document.footnote_options
        #ExFor:FootnoteOptions
        #ExFor:FootnoteOptions.number_style
        #ExSummary:Shows how to change the number style of footnote/endnote reference marks.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Footnotes and endnotes are a way to attach a reference or a side comment to text
        # that does not interfere with the main body text's flow.
        # Inserting a footnote/endnote adds a small superscript reference symbol
        # at the main body text where we insert the footnote/endnote.
        # Each footnote/endnote also creates an entry, which consists of a symbol that matches the reference
        # symbol in the main body text. The reference text that we pass to the document builder's "insert_endnote" method.
        # Footnote entries, by default, show up at the bottom of each page that contains
        # their reference symbols, and endnotes show up at the end of the document.
        builder.write("Text 1. ")
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Footnote 1.")
        builder.write("Text 2. ")
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Footnote 2.")
        builder.write("Text 3. ")
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Footnote 3.", "Custom footnote reference mark")

        builder.insert_paragraph()

        builder.write("Text 1. ")
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Endnote 1.")
        builder.write("Text 2. ")
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Endnote 2.")
        builder.write("Text 3. ")
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Endnote 3.", "Custom endnote reference mark")

        # By default, the reference symbol for each footnote and endnote is its index
        # among all the document's footnotes/endnotes. Each document maintains separate counts
        # for footnotes and for endnotes. By default, footnotes display their numbers using Arabic numerals,
        # and endnotes display their numbers in lowercase Roman numerals.
        self.assertEqual(aw.NumberStyle.ARABIC, doc.footnote_options.number_style)
        self.assertEqual(aw.NumberStyle.LOWERCASE_ROMAN, doc.endnote_options.number_style)

        # We can use the "number_style" property to apply custom numbering styles to footnotes and endnotes.
        # This will not affect footnotes/endnotes with custom reference marks.
        doc.footnote_options.number_style = aw.NumberStyle.UPPERCASE_ROMAN
        doc.endnote_options.number_style = aw.NumberStyle.UPPERCASE_LETTER

        doc.save(ARTIFACTS_DIR + "InlineStory.ref_mark_number_style.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "InlineStory.ref_mark_number_style.docx")

        self.assertEqual(aw.NumberStyle.UPPERCASE_ROMAN, doc.footnote_options.number_style)
        self.assertEqual(aw.NumberStyle.UPPERCASE_LETTER, doc.endnote_options.number_style)

        TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, True, "",
            "Footnote 1.", doc.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, True, "",
            "Footnote 2.", doc.get_child(aw.NodeType.FOOTNOTE, 1, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, False, "Custom footnote reference mark",
            "Custom footnote reference mark Footnote 3.", doc.get_child(aw.NodeType.FOOTNOTE, 2, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.ENDNOTE, True, "",
            "Endnote 1.", doc.get_child(aw.NodeType.FOOTNOTE, 3, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.ENDNOTE, True, "",
            "Endnote 2.", doc.get_child(aw.NodeType.FOOTNOTE, 4, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.ENDNOTE, False, "Custom endnote reference mark",
            "Custom endnote reference mark Endnote 3.", doc.get_child(aw.NodeType.FOOTNOTE, 5, True).as_footnote())

    def test_numbering_rule(self):

        #ExStart
        #ExFor:Document.endnote_options
        #ExFor:EndnoteOptions
        #ExFor:EndnoteOptions.restart_rule
        #ExFor:FootnoteNumberingRule
        #ExFor:Document.footnote_options
        #ExFor:FootnoteOptions
        #ExFor:FootnoteOptions.restart_rule
        #ExSummary:Shows how to restart footnote/endnote numbering at certain places in the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Footnotes and endnotes are a way to attach a reference or a side comment to text
        # that does not interfere with the main body text's flow.
        # Inserting a footnote/endnote adds a small superscript reference symbol
        # at the main body text where we insert the footnote/endnote.
        # Each footnote/endnote also creates an entry, which consists of a symbol that matches the reference
        # symbol in the main body text. The reference text that we pass to the document builder's "insert_endnote" method.
        # Footnote entries, by default, show up at the bottom of each page that contains
        # their reference symbols, and endnotes show up at the end of the document.
        builder.write("Text 1. ")
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Footnote 1.")
        builder.write("Text 2. ")
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Footnote 2.")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.write("Text 3. ")
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Footnote 3.")
        builder.write("Text 4. ")
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Footnote 4.")

        builder.insert_break(aw.BreakType.PAGE_BREAK)

        builder.write("Text 1. ")
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Endnote 1.")
        builder.write("Text 2. ")
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Endnote 2.")
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
        builder.write("Text 3. ")
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Endnote 3.")
        builder.write("Text 4. ")
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Endnote 4.")

        # By default, the reference symbol for each footnote and endnote is its index
        # among all the document's footnotes/endnotes. Each document maintains separate counts
        # for footnotes and endnotes and does not restart these counts at any point.
        self.assertEqual(doc.footnote_options.restart_rule, aw.notes.FootnoteNumberingRule.DEFAULT)
        self.assertEqual(aw.notes.FootnoteNumberingRule.DEFAULT, aw.notes.FootnoteNumberingRule.CONTINUOUS)

        # We can use the "restart_rule" property to get the document to restart
        # the footnote/endnote counts at a new page or section.
        doc.footnote_options.restart_rule = aw.notes.FootnoteNumberingRule.RESTART_PAGE
        doc.endnote_options.restart_rule = aw.notes.FootnoteNumberingRule.RESTART_SECTION

        doc.save(ARTIFACTS_DIR + "InlineStory.numbering_rule.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "InlineStory.numbering_rule.docx")

        self.assertEqual(aw.notes.FootnoteNumberingRule.RESTART_PAGE, doc.footnote_options.restart_rule)
        self.assertEqual(aw.notes.FootnoteNumberingRule.RESTART_SECTION, doc.endnote_options.restart_rule)

        TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, True, "",
            "Footnote 1.", doc.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, True, "",
            "Footnote 2.", doc.get_child(aw.NodeType.FOOTNOTE, 1, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, True, "",
            "Footnote 3.", doc.get_child(aw.NodeType.FOOTNOTE, 2, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, True, "",
            "Footnote 4.", doc.get_child(aw.NodeType.FOOTNOTE, 3, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.ENDNOTE, True, "",
            "Endnote 1.", doc.get_child(aw.NodeType.FOOTNOTE, 4, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.ENDNOTE, True, "",
            "Endnote 2.", doc.get_child(aw.NodeType.FOOTNOTE, 5, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.ENDNOTE, True, "",
            "Endnote 3.", doc.get_child(aw.NodeType.FOOTNOTE, 6, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.ENDNOTE, True, "",
            "Endnote 4.", doc.get_child(aw.NodeType.FOOTNOTE, 7, True).as_footnote())

    def test_start_number(self):

        #ExStart
        #ExFor:Document.endnote_options
        #ExFor:EndnoteOptions
        #ExFor:EndnoteOptions.start_number
        #ExFor:Document.footnote_options
        #ExFor:FootnoteOptions
        #ExFor:FootnoteOptions.start_number
        #ExSummary:Shows how to set a number at which the document begins the footnote/endnote count.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Footnotes and endnotes are a way to attach a reference or a side comment to text
        # that does not interfere with the main body text's flow.
        # Inserting a footnote/endnote adds a small superscript reference symbol
        # at the main body text where we insert the footnote/endnote.
        # Each footnote/endnote also creates an entry, which consists of a symbol
        # that matches the reference symbol in the main body text.
        # The reference text that we pass to the document builder's "insert_endnote" method.
        # Footnote entries, by default, show up at the bottom of each page that contains
        # their reference symbols, and endnotes show up at the end of the document.
        builder.write("Text 1. ")
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Footnote 1.")
        builder.write("Text 2. ")
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Footnote 2.")
        builder.write("Text 3. ")
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Footnote 3.")

        builder.insert_paragraph()

        builder.write("Text 1. ")
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Endnote 1.")
        builder.write("Text 2. ")
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Endnote 2.")
        builder.write("Text 3. ")
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Endnote 3.")

        # By default, the reference symbol for each footnote and endnote is its index
        # among all the document's footnotes/endnotes. Each document maintains separate counts
        # for footnotes and for endnotes, which both begin at 1.
        self.assertEqual(1, doc.footnote_options.start_number)
        self.assertEqual(1, doc.endnote_options.start_number)

        # We can use the "start_number" property to get the document to
        # begin a footnote or endnote count at a different number.
        doc.endnote_options.number_style = aw.NumberStyle.ARABIC
        doc.endnote_options.start_number = 50

        doc.save(ARTIFACTS_DIR + "InlineStory.start_number.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "InlineStory.start_number.docx")

        self.assertEqual(1, doc.footnote_options.start_number)
        self.assertEqual(50, doc.endnote_options.start_number)
        self.assertEqual(aw.NumberStyle.ARABIC, doc.footnote_options.number_style)
        self.assertEqual(aw.NumberStyle.ARABIC, doc.endnote_options.number_style)

        TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, True, "",
            "Footnote 1.", doc.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, True, "",
            "Footnote 2.", doc.get_child(aw.NodeType.FOOTNOTE, 1, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, True, "",
            "Footnote 3.", doc.get_child(aw.NodeType.FOOTNOTE, 2, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.ENDNOTE, True, "",
            "Endnote 1.", doc.get_child(aw.NodeType.FOOTNOTE, 3, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.ENDNOTE, True, "",
            "Endnote 2.", doc.get_child(aw.NodeType.FOOTNOTE, 4, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.ENDNOTE, True, "",
            "Endnote 3.", doc.get_child(aw.NodeType.FOOTNOTE, 5, True).as_footnote())

    def test_add_footnote(self):

        #ExStart
        #ExFor:Footnote
        #ExFor:Footnote.is_auto
        #ExFor:Footnote.reference_mark
        #ExFor:InlineStory
        #ExFor:InlineStory.paragraphs
        #ExFor:InlineStory.first_paragraph
        #ExFor:FootnoteType
        #ExFor:Footnote.__init__
        #ExSummary:Shows how to insert and customize footnotes.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Add text, and reference it with a footnote. This footnote will place a small superscript reference
        # mark after the text that it references and create an entry below the main body text at the bottom of the page.
        # This entry will contain the footnote's reference mark and the reference text,
        # which we will pass to the document builder's "insert_footnote" method.
        builder.write("Main body text.")
        footnote = builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Footnote text.")

        # If this property is set to "True", then our footnote's reference mark
        # will be its index among all the section's footnotes.
        # This is the first footnote, so the reference mark will be "1".
        self.assertTrue(footnote.is_auto)

        # We can move the document builder inside the footnote to edit its reference text.
        builder.move_to(footnote.first_paragraph)
        builder.write(" More text added by a DocumentBuilder.")
        builder.move_to_document_end()

        self.assertEqual("\u0002 Footnote text. More text added by a DocumentBuilder.", footnote.get_text().strip())

        builder.write(" More main body text.")
        footnote = builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Footnote text.")

        # We can set a custom reference mark which the footnote will use instead of its index number.
        footnote.reference_mark = "RefMark"

        self.assertFalse(footnote.is_auto)

        # A bookmark with the "is_auto" flag set to true will still show its real index
        # even if previous bookmarks display custom reference marks, so this bookmark's reference mark will be a "3".
        builder.write(" More main body text.")
        footnote = builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Footnote text.")

        self.assertTrue(footnote.is_auto)

        doc.save(ARTIFACTS_DIR + "InlineStory.add_footnote.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "InlineStory.add_footnote.docx")

        TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, True, "",
            "Footnote text. More text added by a DocumentBuilder.", doc.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, False, "RefMark",
            "Footnote text.", doc.get_child(aw.NodeType.FOOTNOTE, 1, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, True, "",
            "Footnote text.", doc.get_child(aw.NodeType.FOOTNOTE, 2, True).as_footnote())

    def test_footnote_endnote(self):

        #ExStart
        #ExFor:Footnote.footnote_type
        #ExSummary:Shows the difference between footnotes and endnotes.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Below are two ways of attaching numbered references to the text. Both these references will add a
        # small superscript reference mark at the location that we insert them.
        # The reference mark, by default, is the index number of the reference among all the references in the document.
        # Each reference will also create an entry, which will have the same reference mark as in the body text
        # and reference text, which we will pass to the document builder's "insert_footnote" method.
        # 1 -  A footnote, whose entry will appear on the same page as the text that it references:
        builder.write("Footnote referenced main body text.")
        footnote = builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE,
            "Footnote text, will appear at the bottom of the page that contains the referenced text.")

        # 2 -  An endnote, whose entry will appear at the end of the document:
        builder.write("Endnote referenced main body text.")
        endnote = builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE,
            "Endnote text, will appear at the very end of the document.")

        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)

        self.assertEqual(aw.notes.FootnoteType.FOOTNOTE, footnote.footnote_type)
        self.assertEqual(aw.notes.FootnoteType.ENDNOTE, endnote.footnote_type)

        doc.save(ARTIFACTS_DIR + "InlineStory.footnote_endnote.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "InlineStory.footnote_endnote.docx")

        TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, True, "",
            "Footnote text, will appear at the bottom of the page that contains the referenced text.", doc.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote())
        TestUtil.verify_footnote(self, aw.notes.FootnoteType.ENDNOTE, True, "",
            "Endnote text, will appear at the very end of the document.", doc.get_child(aw.NodeType.FOOTNOTE, 1, True).as_footnote())

    def test_add_comment(self):

        #ExStart
        #ExFor:Comment
        #ExFor:InlineStory
        #ExFor:InlineStory.paragraphs
        #ExFor:InlineStory.first_paragraph
        #ExFor:Comment.__init__(DocumentBase,str,str,datetime)
        #ExSummary:Shows how to add a comment to a paragraph.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.write("Hello world!")

        comment = aw.Comment(doc, "John Doe", "JD", date.today())
        builder.current_paragraph.append_child(comment)
        builder.move_to(comment.append_child(aw.Paragraph(doc)))
        builder.write("Comment text.")

        self.assertEqual(date.today(), comment.date_time.date())

        # In Microsoft Word, we can right-click this comment in the document body to edit it, or reply to it.
        doc.save(ARTIFACTS_DIR + "InlineStory.add_comment.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "InlineStory.add_comment.docx")
        comment = doc.get_child(aw.NodeType.COMMENT, 0, True).as_comment()

        self.assertEqual("Comment text.\r", comment.get_text())
        self.assertEqual("John Doe", comment.author)
        self.assertEqual("JD", comment.initial)
        self.assertEqual(date.today(), comment.date_time.date())

    def test_inline_story_revisions(self):

        #ExStart
        #ExFor:InlineStory.is_delete_revision
        #ExFor:InlineStory.is_insert_revision
        #ExFor:InlineStory.is_move_from_revision
        #ExFor:InlineStory.is_move_to_revision
        #ExSummary:Shows how to view revision-related properties of InlineStory nodes.
        doc = aw.Document(MY_DIR + "Revision footnotes.docx")

        # When we edit the document while the "Track Changes" option, found in via Review -> Tracking,
        # is turned on in Microsoft Word, the changes we apply count as revisions.
        # When editing a document using Aspose.Words, we can begin tracking revisions by
        # invoking the document's "start_track_revisions" method and stop tracking by using the "stop_track_revisions" method.
        # We can either accept revisions to assimilate them into the document
        # or reject them to undo and discard the proposed change.
        self.assertTrue(doc.has_revisions)

        footnotes = [node.as_footnote() for node in doc.get_child_nodes(aw.NodeType.FOOTNOTE, True)]

        self.assertEqual(5, len(footnotes))

        # Below are five types of revisions that can flag an InlineStory node.
        # 1 -  An "insert" revision:
        # This revision occurs when we insert text while tracking changes.
        self.assertTrue(footnotes[2].is_insert_revision)

        # 2 -  A "move from" revision:
        # When we highlight text in Microsoft Word, and then drag it to a different place in the document
        # while tracking changes, two revisions appear.
        # The "move from" revision is a copy of the text originally before we moved it.
        self.assertTrue(footnotes[4].is_move_from_revision)

        # 3 -  A "move to" revision:
        # The "move to" revision is the text that we moved in its new position in the document.
        # "Move from" and "move to" revisions appear in pairs for every move revision we carry out.
        # Accepting a move revision deletes the "move from" revision and its text,
        # and keeps the text from the "move to" revision.
        # Rejecting a move revision conversely keeps the "move from" revision and deletes the "move to" revision.
        self.assertTrue(footnotes[1].is_move_to_revision)

        # 4 -  A "delete" revision:
        # This revision occurs when we delete text while tracking changes. When we delete text like this,
        # it will stay in the document as a revision until we either accept the revision,
        # which will delete the text for good, or reject the revision, which will keep the text we deleted where it was.
        self.assertTrue(footnotes[3].is_delete_revision)
        #ExEnd

    def test_insert_inline_story_nodes(self):

        #ExStart
        #ExFor:Comment.story_type
        #ExFor:Footnote.story_type
        #ExFor:InlineStory.ensure_minimum
        #ExFor:InlineStory.font
        #ExFor:InlineStory.last_paragraph
        #ExFor:InlineStory.parent_paragraph
        #ExFor:InlineStory.story_type
        #ExFor:InlineStory.tables
        #ExSummary:Shows how to insert InlineStory nodes.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        footnote = builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, None)

        # Table nodes have an "ensure_minimum()" method that makes sure the table has at least one cell.
        table = aw.tables.Table(doc)
        table.ensure_minimum()

        # We can place a table inside a footnote, which will make it appear at the referencing page's footer.
        self.assertEqual(0, footnote.tables.count)
        footnote.append_child(table)
        self.assertEqual(1, footnote.tables.count)
        self.assertEqual(aw.NodeType.TABLE, footnote.last_child.node_type)

        # An InlineStory has an "ensure_minimum()" method as well, but in this case,
        # it makes sure the last child of the node is a paragraph,
        # for us to be able to click and write text easily in Microsoft Word.
        footnote.ensure_minimum()
        self.assertEqual(aw.NodeType.PARAGRAPH, footnote.last_child.node_type)

        # Edit the appearance of the anchor, which is the small superscript number
        # in the main text that points to the footnote.
        footnote.font.name = "Arial"
        footnote.font.color = drawing.Color.green

        # All inline story nodes have their respective story types.
        self.assertEqual(aw.StoryType.FOOTNOTES, footnote.story_type)

        # A comment is another type of inline story.
        comment = builder.current_paragraph.append_child(aw.Comment(doc, "John Doe", "J. D.", datetime.now())).as_comment()

        # The parent paragraph of an inline story node will be the one from the main document body.
        self.assertEqual(doc.first_section.body.first_paragraph, comment.parent_paragraph)

        # However, the last paragraph is the one from the comment text contents,
        # which will be outside the main document body in a speech bubble.
        # A comment will not have any child nodes by default,
        # so we can apply the ensure_minimum() method to place a paragraph here as well.
        self.assertIsNone(comment.last_paragraph)
        comment.ensure_minimum()
        self.assertEqual(aw.NodeType.PARAGRAPH, comment.last_child.node_type)

        # Once we have a paragraph, we can move the builder to do it and write our comment.
        builder.move_to(comment.last_paragraph)
        builder.write("My comment.")

        self.assertEqual(aw.StoryType.COMMENTS, comment.story_type)

        doc.save(ARTIFACTS_DIR + "InlineStory.insert_inline_story_nodes.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "InlineStory.insert_inline_story_nodes.docx")

        footnote = doc.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote()

        TestUtil.verify_footnote(self, aw.notes.FootnoteType.FOOTNOTE, True, "", "",
            doc.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote())
        self.assertEqual("Arial", footnote.font.name)
        self.assertEqual(drawing.Color.green.to_argb(), footnote.font.color.to_argb())

        comment = doc.get_child(aw.NodeType.COMMENT, 0, True).as_comment()

        self.assertEqual("My comment.", comment.to_string(aw.SaveFormat.TEXT).strip())

    def test_delete_shapes(self):

        #ExStart
        #ExFor:Story
        #ExFor:Story.delete_shapes
        #ExFor:Story.story_type
        #ExFor:StoryType
        #ExSummary:Shows how to remove all shapes from a node.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Use a DocumentBuilder to insert a shape. This is an inline shape,
        # which has a parent Paragraph, which is a child node of the first section's Body.
        builder.insert_shape(aw.drawing.ShapeType.CUBE, 100.0, 100.0)

        self.assertEqual(1, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)

        # We can delete all shapes from the child paragraphs of this Body.
        self.assertEqual(aw.StoryType.MAIN_TEXT, doc.first_section.body.story_type)
        doc.first_section.body.delete_shapes()

        self.assertEqual(0, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)
        #ExEnd
