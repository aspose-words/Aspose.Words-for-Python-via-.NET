import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithCleanupOptions(DocsExamplesBase):


    def test_cleanup_paragraphs_with_punctuation_marks(self):

        #ExStart:CleanupParagraphsWithPunctuationMarks
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        merge_field_option1 = builder.insert_field("MERGEFIELD", "Option_1").as_field_merge_field()
        merge_field_option1.field_name = "Option_1"

        # Here is the complete list of cleanable punctuation marks: ! , .:  ? ¡ ¿.
        builder.write(" ?  ")

        merge_field_option2 = builder.insert_field("MERGEFIELD", "Option_2").as_field_merge_field()
        merge_field_option2.field_name = "Option_2"

        doc.mail_merge.cleanup_options = aw.mailmerging.MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS
        # The option's default value is True, which means that the behavior was changed to mimic MS Word.
        # If you rely on the old behavior can revert it by setting the option to False.
        doc.mail_merge.cleanup_paragraphs_with_punctuation_marks = True

        doc.mail_merge.execute([ "Option_1", "Option_2" ], [ None, None ])

        doc.save(ARTIFACTS_DIR + "WorkingWithCleanupOptions.cleanup_paragraphs_with_punctuation_marks.docx")
        #ExEnd:CleanupParagraphsWithPunctuationMarks


    def test_remove_empty_paragraphs(self):

        #ExStart:RemoveEmptyParagraphs
        doc = aw.Document(MY_DIR + "Table with fields.docx")

        doc.mail_merge.cleanup_options = aw.mailmerging.MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS

        doc.mail_merge.execute([ "FullName", "Company", "Address", "Address2", "City" ],
            [ "James Bond", "MI5 Headquarters", "Milbank", "", "London" ])

        doc.save(ARTIFACTS_DIR + "WorkingWithCleanupOptions.remove_empty_paragraphs.docx")
        #ExEnd:RemoveEmptyParagraphs


    def test_remove_unused_fields(self):

        #ExStart:RemoveUnusedFields
        doc = aw.Document(MY_DIR + "Table with fields.docx")

        doc.mail_merge.cleanup_options = aw.mailmerging.MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS

        doc.mail_merge.execute([ "FullName", "Company", "Address", "Address2", "City" ],
            [ "James Bond", "MI5 Headquarters", "Milbank", "", "London" ])

        doc.save(ARTIFACTS_DIR + "WorkingWithCleanupOptions.remove_unused_fields.docx")
        #ExEnd:RemoveUnusedFields


    def test_remove_containing_fields(self):

        #ExStart:RemoveContainingFields
        doc = aw.Document(MY_DIR + "Table with fields.docx")

        doc.mail_merge.cleanup_options = aw.mailmerging.MailMergeCleanupOptions.REMOVE_CONTAINING_FIELDS

        doc.mail_merge.execute([ "FullName", "Company", "Address", "Address2", "City" ],
            [ "James Bond", "MI5 Headquarters", "Milbank", "", "London" ])

        doc.save(ARTIFACTS_DIR + "WorkingWithCleanupOptions.remove_containing_fields.docx")
        #ExEnd:RemoveContainingFields


    def test_remove_empty_table_rows(self):

        #ExStart:RemoveEmptyTableRows
        doc = aw.Document(MY_DIR + "Table with fields.docx")

        doc.mail_merge.cleanup_options = aw.mailmerging.MailMergeCleanupOptions.REMOVE_EMPTY_TABLE_ROWS

        doc.mail_merge.execute([ "FullName", "Company", "Address", "Address2", "City" ],
            [ "James Bond", "MI5 Headquarters", "Milbank", "", "London" ])

        doc.save(ARTIFACTS_DIR + "WorkingWithCleanupOptions.remove_empty_table_rows.docx")
        #ExEnd:RemoveEmptyTableRows


if __name__ == '__main__':
    unittest.main()
