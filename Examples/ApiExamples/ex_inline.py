# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import unittest
from api_example_base import ApiExampleBase, MY_DIR

class ExInline(ApiExampleBase):

    def test_inline_revisions(self):
        #ExStart
        #ExFor:Inline
        #ExFor:Inline.is_delete_revision
        #ExFor:Inline.is_format_revision
        #ExFor:Inline.is_insert_revision
        #ExFor:Inline.is_move_from_revision
        #ExFor:Inline.is_move_to_revision
        #ExFor:Inline.parent_paragraph
        #ExFor:Paragraph.runs
        #ExFor:Revision.parent_node
        #ExFor:RunCollection
        #ExFor:RunCollection.__getitem__(int)
        #ExFor:RunCollection.to_array
        #ExSummary:Shows how to determine the revision type of an inline node.
        doc = aw.Document(file_name=MY_DIR + 'Revision runs.docx')
        # When we edit the document while the "Track Changes" option, found in via Review -> Tracking,
        # is turned on in Microsoft Word, the changes we apply count as revisions.
        # When editing a document using Aspose.Words, we can begin tracking revisions by
        # invoking the document's "StartTrackRevisions" method and stop tracking by using the "StopTrackRevisions" method.
        # We can either accept revisions to assimilate them into the document
        # or reject them to change the proposed change effectively.
        self.assertEqual(6, doc.revisions.count)
        # The parent node of a revision is the run that the revision concerns. A Run is an Inline node.
        run = doc.revisions[0].parent_node.as_run()
        first_paragraph = run.parent_paragraph
        runs = first_paragraph.runs
        self.assertEqual(6, len(list(runs)))
        # Below are five types of revisions that can flag an Inline node.
        # 1 -  An "insert" revision:
        # This revision occurs when we insert text while tracking changes.
        self.assertTrue(runs[2].is_insert_revision)
        # 2 -  A "format" revision:
        # This revision occurs when we change the formatting of text while tracking changes.
        self.assertTrue(runs[2].is_format_revision)
        # 3 -  A "move from" revision:
        # When we highlight text in Microsoft Word, and then drag it to a different place in the document
        # while tracking changes, two revisions appear.
        # The "move from" revision is a copy of the text originally before we moved it.
        self.assertTrue(runs[4].is_move_from_revision)
        # 4 -  A "move to" revision:
        # The "move to" revision is the text that we moved in its new position in the document.
        # "Move from" and "move to" revisions appear in pairs for every move revision we carry out.
        # Accepting a move revision deletes the "move from" revision and its text,
        # and keeps the text from the "move to" revision.
        # Rejecting a move revision conversely keeps the "move from" revision and deletes the "move to" revision.
        self.assertTrue(runs[1].is_move_to_revision)
        # 5 -  A "delete" revision:
        # This revision occurs when we delete text while tracking changes. When we delete text like this,
        # it will stay in the document as a revision until we either accept the revision,
        # which will delete the text for good, or reject the revision, which will keep the text we deleted where it was.
        self.assertTrue(runs[5].is_delete_revision)
        #ExEnd