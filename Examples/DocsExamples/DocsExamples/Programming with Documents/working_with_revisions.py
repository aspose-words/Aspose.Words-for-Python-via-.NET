import unittest
import os
import sys
from datetime import date, datetime

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithRevisions(docs_base.DocsExamplesBase):

    def test_accept_revisions(self):

        #ExStart:AcceptAllRevisions
        doc = aw.Document()
        body = doc.first_section.body
        para = body.first_paragraph

        # Add text to the first paragraph, then add two more paragraphs.
        para.append_child(aw.Run(doc, "Paragraph 1. "))
        body.append_paragraph("Paragraph 2. ")
        body.append_paragraph("Paragraph 3. ")

        # We have three paragraphs, none of which registered as any type of revision
        # If we add/remove any content in the document while tracking revisions,
        # they will be displayed as such in the document and can be accepted/rejected.
        doc.start_track_revisions("John Doe", datetime.today())

        # This paragraph is a revision and will have the according "IsInsertRevision" flag set.
        para = body.append_paragraph("Paragraph 4. ")
        self.assertTrue(para.is_insert_revision)

        # Get the document's paragraph collection and remove a paragraph.
        paragraphs = body.paragraphs
        self.assertEqual(4, paragraphs.count)
        para = paragraphs[2]
        para.remove()

        # Since we are tracking revisions, the paragraph still exists in the document, will have the "IsDeleteRevision" set
        # and will be displayed as a revision in Microsoft Word, until we accept or reject all revisions.
        self.assertEqual(4, paragraphs.count)
        self.assertTrue(para.is_delete_revision)

        # The delete revision paragraph is removed once we accept changes.
        doc.accept_all_revisions()
        self.assertEqual(3, paragraphs.count)

        # Stopping the tracking of revisions makes this text appear as normal text.
        # Revisions are not counted when the document is changed.
        doc.stop_track_revisions()

        # Save the document.
        doc.save(docs_base.artifacts_dir + "WorkingWithRevisions.accept_revisions.docx")
        #ExEnd:AcceptAllRevisions


    def test_get_revision_types(self):

        #ExStart:GetRevisionTypes
        doc = aw.Document(docs_base.my_dir + "Revisions.docx")

        paragraphs = doc.first_section.body.paragraphs
        for i in range(0, paragraphs.count):

            if paragraphs[i].is_move_from_revision:
                print(f"The paragraph {i} has been moved (deleted).")
            if paragraphs[i].is_move_to_revision:
                print(f"The paragraph {i} has been moved (inserted).")

        #ExEnd:GetRevisionTypes


    def test_get_revision_groups(self):

        #ExStart:GetRevisionGroups
        doc = aw.Document(docs_base.my_dir + "Revisions.docx")

        for group in doc.revisions.groups:

            print(f"{group.author}, {group.revision_type}:")
            print(group.text)

        #ExEnd:GetRevisionGroups


    def test_remove_comments_in_pdf(self):

        #ExStart:RemoveCommentsInPDF
        doc = aw.Document(docs_base.my_dir + "Revisions.docx")

        # Do not render the comments in PDF.
        doc.layout_options.comment_display_mode = aw.layout.CommentDisplayMode.HIDE

        doc.save(docs_base.artifacts_dir + "WorkingWithRevisions.remove_comments_in_pdf.pdf")
        #ExEnd:RemoveCommentsInPDF


    def test_show_revisions_in_balloons(self):

        #ExStart:ShowRevisionsInBalloons
        #ExStart:SetMeasurementUnit
        #ExStart:SetRevisionBarsPosition
        doc = aw.Document(docs_base.my_dir + "Revisions.docx")

        # Renders insert revisions inline, delete and format revisions in balloons.
        doc.layout_options.revision_options.show_in_balloons = aw.layout.ShowInBalloons.FORMAT_AND_DELETE
        doc.layout_options.revision_options.measurement_unit = aw.MeasurementUnits.INCHES
        # Renders revision bars on the right side of a page.
        doc.layout_options.revision_options.revision_bars_position = aw.drawing.HorizontalAlignment.RIGHT

        doc.save(docs_base.artifacts_dir + "WorkingWithRevisions.show_revisions_in_balloons.pdf")
        #ExEnd:SetRevisionBarsPosition
        #ExEnd:SetMeasurementUnit
        #ExEnd:ShowRevisionsInBalloons


    def test_get_revision_group_details(self):

        #ExStart:GetRevisionGroupDetails
        doc = aw.Document(docs_base.my_dir + "Revisions.docx")

        for revision in doc.revisions:

            group_text = "Revision group text: " + revision.group.text if revision.group != None else "Revision has no group"

            print(f"Type: {revision.revision_type}")
            print(f"Author: {revision.author}")
            print(f"Date: {revision.date_time}")
            print(f"Revision text: {revision.parent_node.to_string(aw.SaveFormat.TEXT)}")
            print(group_text)

        #ExEnd:GetRevisionGroupDetails


    def test_access_revised_version(self):

        #ExStart:AccessRevisedVersion
        doc = aw.Document(docs_base.my_dir + "Revisions.docx")
        doc.update_list_labels()

        # Switch to the revised version of the document.
        doc.revisions_view = aw.RevisionsView.FINAL

        for revision in doc.revisions:

            if revision.parent_node.node_type == aw.NodeType.PARAGRAPH:

                paragraph = revision.parent_node.as_paragraph()
                if paragraph.is_list_item:

                    print(paragraph.list_label.label_string)
                    print(paragraph.list_format.list_level)


        #ExEnd:AccessRevisedVersion


    def test_move_node_in_tracked_document(self):

        #ExStart:MoveNodeInTrackedDocument
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Paragraph 1")
        builder.writeln("Paragraph 2")
        builder.writeln("Paragraph 3")
        builder.writeln("Paragraph 4")
        builder.writeln("Paragraph 5")
        builder.writeln("Paragraph 6")
        body = doc.first_section.body
        print(f"Paragraph count: {body.paragraphs.count}")

        # Start tracking revisions.
        doc.start_track_revisions("Author", datetime(2020, 12, 23, 14, 0, 0))

        # Generate revisions when moving a node from one location to another.
        node = body.paragraphs[3]
        end_node = body.paragraphs[5].next_sibling
        reference_node = body.paragraphs[0]
        while (node != end_node):

            next_node = node.next_sibling
            body.insert_before(node, reference_node)
            node = next_node


        # Stop the process of tracking revisions.
        doc.stop_track_revisions()

        # There are 3 additional paragraphs in the move-from range.
        print("Paragraph count: 0", body.paragraphs.count)
        doc.save(docs_base.artifacts_dir + "WorkingWithRevisions.move_node_in_tracked_document.docx")
        #ExEnd:MoveNodeInTrackedDocument


    def test_shape_revision(self):

        #ExStart:ShapeRevision
        doc = aw.Document()

        # Insert an inline shape without tracking revisions.
        self.assertFalse(doc.track_revisions)
        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.CUBE)
        shape.wrap_type = aw.drawing.WrapType.INLINE
        shape.width = 100
        shape.height = 100
        doc.first_section.body.first_paragraph.append_child(shape)

        # Start tracking revisions and then insert another shape.
        doc.start_track_revisions("John Doe")
        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.SUN)
        shape.wrap_type = aw.drawing.WrapType.INLINE
        shape.width = 100.0
        shape.height = 100.0
        doc.first_section.body.first_paragraph.append_child(shape)

        # Get the document's shape collection which includes just the two shapes we added.
        shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)
        self.assertEqual(2, shapes.count)

        # Remove the first shape.
        shape0 = shapes[0].as_shape()
        shape0.remove()

        # Because we removed that shape while changes were being tracked, the shape counts as a delete revision.
        self.assertEqual(aw.drawing.ShapeType.CUBE, shape0.shape_type)
        self.assertTrue(shape0.is_delete_revision)

        # And we inserted another shape while tracking changes, so that shape will count as an insert revision.
        shape1 = shapes[1].as_shape()
        self.assertEqual(aw.drawing.ShapeType.SUN, shape1.shape_type)
        self.assertTrue(shape1.is_insert_revision)

        # The document has one shape that was moved, but shape move revisions will have two instances of that shape.
        # One will be the shape at its arrival destination and the other will be the shape at its original location.
        doc = aw.Document(docs_base.my_dir + "Revision shape.docx")

        shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)
        self.assertEqual(2, shapes.count)

        # This is the move to revision, also the shape at its arrival destination.
        shape0 = shapes[0].as_shape()
        self.assertFalse(shape0.is_move_from_revision)
        self.assertTrue(shape0.is_move_to_revision)

        # This is the move from revision, which is the shape at its original location.
        shape1 = shapes[1].as_shape()
        self.assertTrue(shape1.is_move_from_revision)
        self.assertFalse(shape1.is_move_to_revision)
        #ExEnd:ShapeRevision


if __name__ == '__main__':
    unittest.main()
