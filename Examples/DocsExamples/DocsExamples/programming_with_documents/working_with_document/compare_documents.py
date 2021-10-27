import unittest
import os
import sys
from datetime import datetime

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class CompareDocument(DocsExamplesBase):

    def test_compare_for_equal(self):

        #ExStart:CompareForEqual
        doc_a = aw.Document(MY_DIR + "Document.docx")
        doc_b = doc_a.clone().as_document()

        # DocA now contains changes as revisions.
        doc_a.compare(doc_b, "user", datetime.today())

        print("Documents are equal" if doc_a.revisions.count == 0 else "Documents are not equal")
        #ExEnd:CompareForEqual

    def test_compare_options(self):

        #ExStart:CompareOptions
        doc_a = aw.Document(MY_DIR + "Document.docx")
        doc_b = doc_a.clone()

        options = aw.comparing.CompareOptions()

        options.ignore_formatting = True
        options.ignore_headers_and_footers = True
        options.ignore_case_changes = True
        options.ignore_tables = True
        options.ignore_fields = True
        options.ignore_comments = True
        options.ignore_textboxes = True
        options.ignore_footnotes = True

        doc_a.compare(doc_b, "user", datetime.now(), options)

        print("Documents are equal" if doc_a.revisions.count == 0 else "Documents are not equal")
        #ExEnd:CompareOptions

    def test_comparison_target(self):

        #ExStart:ComparisonTarget
        doc_a = aw.Document(MY_DIR + "Document.docx")
        doc_b = doc_a.clone()

        # Relates to Microsoft Word "Show changes in" option in "Compare Documents" dialog box.
        options = aw.comparing.CompareOptions()
        options.ignore_formatting = True
        options.target = aw.comparing.ComparisonTargetType.NEW

        doc_a.compare(doc_b, "user", datetime.now(), options)
        #ExEnd:ComparisonTarget

    def test_comparison_granularity(self):

        #ExStart:ComparisonGranularity
        builder_a = aw.DocumentBuilder(aw.Document())
        builder_b = aw.DocumentBuilder(aw.Document())

        builder_a.writeln("This is A simple word")
        builder_b.writeln("This is B simple words")

        compare_options = aw.comparing.CompareOptions()
        compare_options.granularity = aw.comparing.Granularity.CHAR_LEVEL

        builder_a.document.compare(builder_b.document, "author", datetime.now(), compare_options)
        #ExEnd:ComparisonGranularity

    def test_apply_compare_two_documents(self):

        #ExStart:ApplyCompareTwoDocuments
        # The source document doc1.
        doc1 = aw.Document()
        builder = aw.DocumentBuilder(doc1)
        builder.writeln("This is the original document.")

        # The target document doc2.
        doc2 = aw.Document()
        builder = aw.DocumentBuilder(doc2)
        builder.writeln("This is the edited document.")

        # If either document has a revision, an exception will be thrown.
        if doc1.revisions.count == 0 and doc2.revisions.count == 0:
            doc1.compare(doc2, "authorName", datetime.now())

        # If doc1 and doc2 are different, doc1 now has some revisions after the comparison, which can now be viewed and processed.
        self.assertEqual(2, doc1.revisions.count)

        for revision in doc1.revisions:
            print(f'Revision type: {revision.revision_type}, on a node of type "{revision.parent_node.node_type}"')
            print(f'\tChanged text: "{revision.parent_node.get_text()}"')

        # All the revisions in doc1 are differences between doc1 and doc2, so accepting them on doc1 transforms doc1 into doc2.
        doc1.revisions.accept_all()

        # doc1, when saved, now resembles doc2.
        doc1.save(ARTIFACTS_DIR + "Document.Compare.docx")
        doc1 = aw.Document(ARTIFACTS_DIR + "Document.Compare.docx")
        self.assertEqual(0, doc1.revisions.count)
        self.assertEqual(doc2.get_text().strip(), doc1.get_text().strip())
        #ExEnd:ApplyCompareTwoDocuments
