# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

from datetime import datetime, date

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExRevision(ApiExampleBase):

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
        builder.write("This does not count as a revision. ")

        self.assertFalse(doc.has_revisions)

        # To register our edits as revisions, we need to declare an author, and then start tracking them.
        doc.start_track_revisions("John Doe", datetime.now())

        builder.write("This is revision #1. ")

        self.assertTrue(doc.has_revisions)
        self.assertEqual(1, doc.revisions.count)

        # This flag corresponds to the "Review" -> "Tracking" -> "Track Changes" option in Microsoft Word.
        # The "start_track_revisions" method does not affect its value,
        # and the document is tracking revisions programmatically despite it having a value of "False".
        # If we open this document using Microsoft Word, it will not be tracking revisions.
        self.assertFalse(doc.track_revisions)

        # We have added text using the document builder, so the first revision is an insertion-type revision.
        revision = doc.revisions[0]
        self.assertEqual("John Doe", revision.author)
        self.assertEqual("This is revision #1. ", revision.parent_node.get_text())
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
        self.assertEqual("This does not count as a revision. This is revision #1.", doc.get_text().strip())

        # Accepting the delete revision will remove its parent node from the paragraph text
        # and then remove the collection's revision itself.
        doc.revisions[0].accept()

        self.assertEqual(1, doc.revisions.count)
        self.assertEqual("This is revision #1.", doc.get_text().strip())

        builder.writeln("")
        builder.write("This is revision #2.")

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
        self.assertEqual("This is revision #2.\rThis is revision #1. \rThis is revision #2.", doc.get_text().strip())

        # The moving revision is now at index 1. Reject the revision to discard its contents.
        doc.revisions[1].reject()

        self.assertEqual(6, doc.revisions.count)
        self.assertEqual("This is revision #1. \rThis is revision #2.", doc.get_text().strip())
        #ExEnd

    def test_revision_collection(self):

        #ExStart
        #ExFor:Revision.parent_style
        #ExFor:RevisionCollection.__iter__
        #ExFor:RevisionCollection.groups
        #ExFor:RevisionCollection.reject_all
        #ExFor:RevisionGroupCollection.__iter__
        #ExSummary:Shows how to work with a document's collection of revisions.
        doc = aw.Document(MY_DIR + "Revisions.docx")
        revisions = doc.revisions

        # This collection itself has a collection of revision groups.
        # Each group is a sequence of adjacent revisions.
        self.assertEqual(7, revisions.groups.count) #ExSkip
        print(revisions.groups.count, "revision groups:")

        # Iterate over the collection of groups and print the text that the revision concerns.
        for group in revisions.groups:
            print(f"\tGroup type \"{group.revision_type}\", " +
                  f"author: {group.author}, contents: [{group.text.strip()}]")

        # Each Run that a revision affects gets a corresponding Revision object.
        # The revisions' collection is considerably larger than the condensed form we printed above,
        # depending on how many Runs we have segmented the document into during Microsoft Word editing.
        self.assertEqual(11, revisions.count) #ExSkip
        print(f"\n{revisions.count} revisions:")

        for revision in revisions:
            # A StyleDefinitionChange strictly affects styles and not document nodes. This means the "parent_style"
            # property will always be in use, while the "parent_node" will always be None.
            # Since all other changes affect nodes, "parent_node" will conversely be in use, and "parent_style" will be None.
            if revision.revision_type == aw.RevisionType.STYLE_DEFINITION_CHANGE:
                print(f"\tRevision type \"{revision.revision_type}\", " +
                      f"author: {revision.author}, style: [{revision.parent_style.name}]")
            else:
                print(f"\tRevision type \"{revision.revision_type}\", " +
                      f"author: {revision.author}, contents: [{revision.parent_node.get_text().strip()}]")

        # Reject all revisions via the collection, reverting the document to its original form.
        revisions.reject_all()

        self.assertEqual(0, revisions.count)
        #ExEnd

    def test_get_info_about_revisions_in_revision_groups(self):

        #ExStart
        #ExFor:RevisionGroup
        #ExFor:RevisionGroup.author
        #ExFor:RevisionGroup.revision_type
        #ExFor:RevisionGroup.text
        #ExFor:RevisionGroupCollection
        #ExFor:RevisionGroupCollection.count
        #ExSummary:Shows how to print info about a group of revisions in a document.
        doc = aw.Document(MY_DIR + "Revisions.docx")

        self.assertEqual(7, doc.revisions.groups.count)

        for group in doc.revisions.groups:
            print(f"Revision author: {group.author}; Revision type: {group.revision_type} \n\tRevision text: {group.text}")

        #ExEnd

    def test_get_specific_revision_group(self):

        #ExStart
        #ExFor:RevisionGroupCollection
        #ExFor:RevisionGroupCollection.__getitem__(int)
        #ExSummary:Shows how to get a group of revisions in a document.
        doc = aw.Document(MY_DIR + "Revisions.docx")

        revision_group = doc.revisions.groups[0]
        #ExEnd

        self.assertEqual(aw.RevisionType.DELETION, revision_group.revision_type)
        self.assertEqual("Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ",
            revision_group.text)

    def test_show_revision_balloons(self):

        #ExStart
        #ExFor:RevisionOptions.show_in_balloons
        #ExSummary:Shows how to display revisions in balloons.
        doc = aw.Document(MY_DIR + "Revisions.docx")

        # By default, text that is a revision has a different color to differentiate it from the other non-revision text.
        # Set a revision option to show more details about each revision in a balloon on the page's right margin.
        doc.layout_options.revision_options.show_in_balloons = aw.layout.ShowInBalloons.FORMAT_AND_DELETE
        doc.save(ARTIFACTS_DIR + "Revision.show_revision_balloons.pdf")
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
        doc = aw.Document(MY_DIR + "Revisions.docx")

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
        revision_options.moved_from_text_effect = aw.layout.RevisionTextEffect.DOUBLE_UNDERLINE

        # Render format revisions in dark red and bold.
        revision_options.revised_properties_color = aw.layout.RevisionColor.DARK_RED
        revision_options.revised_properties_effect = aw.layout.RevisionTextEffect.BOLD

        # Place a thick dark blue bar on the left side of the page next to lines affected by revisions.
        revision_options.revision_bars_color = aw.layout.RevisionColor.DARK_BLUE
        revision_options.revision_bars_width = 15.0

        # Show revision marks and original text.
        revision_options.show_original_revision = True
        revision_options.show_revision_marks = True

        # Get movement, deletion, formatting revisions, and comments to show up in green balloons
        # on the right side of the page.
        revision_options.show_in_balloons = aw.layout.ShowInBalloons.FORMAT
        revision_options.comment_color = aw.layout.RevisionColor.BRIGHT_GREEN

        # These features are only applicable to formats such as .pdf or .jpg.
        doc.save(ARTIFACTS_DIR + "Revision.revision_options.pdf")
        #ExEnd
