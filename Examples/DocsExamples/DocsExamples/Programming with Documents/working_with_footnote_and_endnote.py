import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithFootnotes(docs_base.DocsExamplesBase):

    def test_set_foot_note_columns(self) :

        #ExStart:SetFootNoteColumns
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        # Specify the number of columns with which the footnotes area is formatted.
        doc.footnote_options.columns = 3

        doc.save(docs_base.artifacts_dir + "WorkingWithFootnotes.set_foot_note_columns.docx")
        #ExEnd:SetFootNoteColumns


    def test_set_footnote_and_end_note_position(self) :

        #ExStart:SetFootnoteAndEndNotePosition
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        doc.footnote_options.position = aw.notes.FootnotePosition.BENEATH_TEXT
        doc.endnote_options.position = aw.notes.EndnotePosition.END_OF_SECTION

        doc.save(docs_base.artifacts_dir + "WorkingWithFootnotes.set_footnote_and_end_note_position.docx")
        #ExEnd:SetFootnoteAndEndNotePosition


    def test_set_endnote_options(self) :

        #ExStart:SetEndnoteOptions
        doc = aw.Document(docs_base.my_dir + "Document.docx")
        builder = aw.DocumentBuilder(doc)

        builder.write("Some text")
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Footnote text.")

        option = doc.endnote_options
        option.restart_rule = aw.notes.FootnoteNumberingRule.RESTART_PAGE
        option.position = aw.notes.EndnotePosition.END_OF_SECTION

        doc.save(docs_base.artifacts_dir + "WorkingWithFootnotes.set_endnote_options.docx")
        #ExEnd:SetEndnoteOptions


if __name__ == '__main__':
    unittest.main()
