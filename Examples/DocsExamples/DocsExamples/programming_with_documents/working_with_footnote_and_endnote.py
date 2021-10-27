import unittest
import os
import sys

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithFootnotes(DocsExamplesBase):

    def test_set_foot_note_columns(self):

        #ExStart:SetFootNoteColumns
        doc = aw.Document(MY_DIR + "Document.docx")

        # Specify the number of columns with which the footnotes area is formatted.
        doc.footnote_options.columns = 3

        doc.save(ARTIFACTS_DIR + "WorkingWithFootnotes.set_foot_note_columns.docx")
        #ExEnd:SetFootNoteColumns

    def test_set_footnote_and_end_note_position(self):

        #ExStart:SetFootnoteAndEndNotePosition
        doc = aw.Document(MY_DIR + "Document.docx")

        doc.footnote_options.position = aw.notes.FootnotePosition.BENEATH_TEXT
        doc.endnote_options.position = aw.notes.EndnotePosition.END_OF_SECTION

        doc.save(ARTIFACTS_DIR + "WorkingWithFootnotes.set_footnote_and_end_note_position.docx")
        #ExEnd:SetFootnoteAndEndNotePosition

    def test_set_endnote_options(self):

        #ExStart:SetEndnoteOptions
        doc = aw.Document(MY_DIR + "Document.docx")
        builder = aw.DocumentBuilder(doc)

        builder.write("Some text")
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Footnote text.")

        option = doc.endnote_options
        option.restart_rule = aw.notes.FootnoteNumberingRule.RESTART_PAGE
        option.position = aw.notes.EndnotePosition.END_OF_SECTION

        doc.save(ARTIFACTS_DIR + "WorkingWithFootnotes.set_endnote_options.docx")
        #ExEnd:SetEndnoteOptions
