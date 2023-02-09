# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExAbsolutePositionTab(ApiExampleBase):

    ##ExStart
    ##ExFor:AbsolutePositionTab
    ##ExFor:AbsolutePositionTab.accept(DocumentVisitor)
    ##ExFor:DocumentVisitor.visit_absolute_position_tab
    ##ExSummary:Shows how to process absolute position tab characters with a document visitor.
    #def test_document_to_txt(self):

    #    doc = aw.Document(MY_DIR + "Absolute position tab.docx")

    #    # Extract the text contents of our document by accepting this custom document visitor.
    #    my_doc_text_extractor = ExAbsolutePositionTab.DocTextExtractor()
    #    doc.first_section.body.accept(my_doc_text_extractor)

    #    # The absolute position tab, which has no equivalent in string form, has been explicitly converted to a tab character.
    #    self.assertEqual("Before AbsolutePositionTab\tAfter AbsolutePositionTab", my_doc_text_extractor.get_text())

    #    # An AbsolutePositionTab can accept a DocumentVisitor by itself too.
    #    abs_position_tab = doc.first_section.body.first_paragraph.get_child(aw.NodeType.SPECIAL_CHAR, 0, True).as_absolute_position_tab()

    #    my_doc_text_extractor = ExAbsolutePositionTab.DocTextExtractor()
    #    abs_position_tab.accept(my_doc_text_extractor)

    #    self.assertEqual("\t", my_doc_text_extractor.get_text())

    #class DocTextExtractor(aw.DocumentVisitor):
    #    """Collects the text contents of all runs in the visited document. Replaces all absolute tab characters with ordinary tabs."""

    #    def __init__(self):

    #        self.builder = io.StringIO()

    #    def visit_run(self, run: aw.Run) -> aw.VisitorAction:
    #        """Called when a Run node is encountered in the document."""

    #        self.builder.write(run.text)
    #        return aw.VisitorAction.CONTINUE

    #    def visit_absolute_position_tab(self, tab: aw.AbsolutePositionTab) -> aw.VisitorAction:
    #        """Called when an AbsolutePositionTab node is encountered in the document."""

    #        self.builder.write("\t")
    #        return aw.VisitorAction.CONTINUE

    #    def get_text(self) -> str:
    #        """Plain text of the document that was accumulated by the visitor."""

    #        return self.builder.to_string()

    ##ExEnd
    pass
