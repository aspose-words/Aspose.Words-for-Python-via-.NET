import unittest

import api_example_base as aeb

import aspose.words as aw
from document_helper import DocumentHelper


class ExEditableRange(aeb.ApiExampleBase):

    def test_create_and_remove(self):
        # ExStart
        # ExFor:DocumentBuilder.start_editable_range
        # ExFor:DocumentBuilder.end_editable_range
        # ExFor:EditableRange
        # ExFor:EditableRange.editable_range_end
        # ExFor:EditableRange.editable_range_start
        # ExFor:EditableRange.id
        # ExFor:EditableRange.remove
        # ExFor:EditableRangeEnd.editable_range_start
        # ExFor:EditableRangeEnd.id
        # ExFor:EditableRangeEnd.node_type
        # ExFor:EditableRangeStart.editable_range
        # ExFor:EditableRangeStart.id
        # ExFor:EditableRangeStart.node_type
        # ExSummary:Shows how to work with an editable range.
        doc = aw.Document()
        doc.protect(aw.ProtectionType.READ_ONLY, "MyPassword")

        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world! Since we have set the document's protection level to read-only," +
                        " we cannot edit this paragraph without the password.")

        # Editable ranges allow us to leave parts of protected documents open for editing.
        editableRangeStart = builder.start_editable_range()
        builder.writeln("This paragraph is inside an editable range, and can be edited.")
        editableRangeEnd = builder.end_editable_range()

        # A well-formed editable range has a start node, and end node.
        # These nodes have matching IDs and encompass editable nodes.
        editableRange = editableRangeStart.editable_range

        self.assertEqual(editableRangeStart.id, editableRange.id)
        self.assertEqual(editableRangeEnd.id, editableRange.id)

        # Different parts of the editable range link to each other.
        self.assertEqual(editableRangeStart.id, editableRange.editable_range_start.id)
        self.assertEqual(editableRangeStart.id, editableRangeEnd.editable_range_start.id)
        self.assertEqual(editableRange.id, editableRangeStart.editable_range.id)
        self.assertEqual(editableRangeEnd.id, editableRange.editable_range_end.id)

        # We can access the node types of each part like this. The editable range itself is not a node,
        # but an entity which consists of a start, an end, and their enclosed contents.
        self.assertEqual(aw.NodeType.EDITABLE_RANGE_START, editableRangeStart.node_type)
        self.assertEqual(aw.NodeType.EDITABLE_RANGE_END, editableRangeEnd.node_type)

        builder.writeln("This paragraph is outside the editable range, and cannot be edited.")

        doc.save(aeb.artifacts_dir + "EditableRange.create_and_remove.docx")

        # Remove an editable range. All the nodes that were inside the range will remain intact.
        editableRange.remove()
        # ExEnd

        self.assertEqual(
            "Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" +
            "This paragraph is inside an editable range, and can be edited.\r" +
            "This paragraph is outside the editable range, and cannot be edited.", doc.get_text().strip())
        self.assertEqual(0, doc.get_child_nodes(aw.NodeType.EDITABLE_RANGE_START, True).count)

        doc = aw.Document(aeb.artifacts_dir + "EditableRange.create_and_remove.docx")

        self.assertEqual(aw.ProtectionType.READ_ONLY, doc.protection_type)
        self.assertEqual(
            "Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" +
            "This paragraph is inside an editable range, and can be edited.\r" +
            "This paragraph is outside the editable range, and cannot be edited.", doc.get_text().strip())

        # editableRange = (doc.get_child(aw.NodeType.EDITABLE_RANGE_START, 0, True).as_editeble_range_start()).editable_range
        #
        # TestUtil.verify_editable_range(0, string.empty, EditorType.unspecified, editableRange)

    def test_nested(self):
        # ExStart
        # ExFor:DocumentBuilder.start_editable_range
        # ExFor:DocumentBuilder.end_editable_range(EditableRangeStart)
        # ExFor:EditableRange.editor_group
        # ExSummary:Shows how to create nested editable ranges.
        doc = aw.Document()
        doc.protect(aw.ProtectionType.READ_ONLY, "MyPassword")

        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world! Since we have set the document's protection level to read-only, " +
                        "we cannot edit this paragraph without the password.")

        # Create two nested editable ranges.
        outerEditableRangeStart = builder.start_editable_range()
        builder.writeln("This paragraph inside the outer editable range and can be edited.")

        innerEditableRangeStart = builder.start_editable_range()
        builder.writeln("This paragraph inside both the outer and inner editable ranges and can be edited.")

        # Currently, the document builder's node insertion cursor is in more than one ongoing editable range.
        # When we want to end an editable range in this situation,
        # we need to specify which of the ranges we wish to end by passing its EditableRangeStart node.
        builder.end_editable_range(innerEditableRangeStart)

        builder.writeln("This paragraph inside the outer editable range and can be edited.")

        builder.end_editable_range(outerEditableRangeStart)

        builder.writeln("This paragraph is outside any editable ranges, and cannot be edited.")

        # If a region of text has two overlapping editable ranges with specified groups,
        # the combined group of users excluded by both groups are prevented from editing it.
        outerEditableRangeStart.editable_range.editor_group = aw.EditorType.EVERYONE
        innerEditableRangeStart.editable_range.editor_group = aw.EditorType.CONTRIBUTORS

        doc.save(aeb.artifacts_dir + "EditableRange.nested.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "EditableRange.nested.docx")

        self.assertEqual(
            "Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" +
            "This paragraph inside the outer editable range and can be edited.\r" +
            "This paragraph inside both the outer and inner editable ranges and can be edited.\r" +
            "This paragraph inside the outer editable range and can be edited.\r" +
            "This paragraph is outside any editable ranges, and cannot be edited.", doc.get_text().strip())

        # editableRange = (doc.get_child(aw.NodeType.EDITABLE_RANGE_START, 0, True).as_editable_range()).editable_range
        #
        # TestUtil.verify_editable_range(0, string.empty, EditorType.everyone, editableRange)
        #
        # editableRange = ((EditableRangeStart)doc.get_child(NodeType.editable_range_start, 1, true)).editable_range
        #
        # TestUtil.verify_editable_range(1, string.empty, EditorType.contributors, editableRange)

    # ExStart

    # ExFor:EditableRange
    # ExFor:EditableRange.editor_group
    # ExFor:EditableRange.single_user
    # ExFor:EditableRangeEnd
    # ExFor:EditableRangeEnd.accept(DocumentVisitor)
    # ExFor:EditableRangeStart
    # ExFor:EditableRangeStart.accept(DocumentVisitor)
    # ExFor:EditorType
    # ExSummary:Shows how to limit the editing rights of editable ranges to a specific group/user.
    # def test_visitor(self) :
    #
    #     doc = aw.Document()
    #     doc.protect(ProtectionType.read_only, "MyPassword")
    #
    #     builder = aw.DocumentBuilder(doc)
    #     builder.writeln("Hello world! Since we have set the document's protection level to read-only," +
    #                     " we cannot edit this paragraph without the password.")
    #
    #     # When we write-protect documents, editable ranges allow us to pick specific areas that users may edit.
    #     # There are two mutually exclusive ways to narrow down the list of allowed editors.
    #     # 1 -  Specify a user:
    #     EditableRange editableRange = builder.start_editable_range().editable_range
    #     editableRange.single_user = "john.doe@myoffice.com"
    #     builder.writeln($"This paragraph is inside the first editable range, can only be edited by editableRange.single_user.")
    #     builder.end_editable_range()
    #
    #     self.assertEqual(EditorType.unspecified, editableRange.editor_group)
    #
    #     # 2 -  Specify a group that allowed users are associated with:
    #     editableRange = builder.start_editable_range().editable_range
    #     editableRange.editor_group = EditorType.administrators
    #     builder.writeln($"This paragraph is inside the first editable range, can only be edited by editableRange.editor_group.")
    #     builder.end_editable_range()
    #
    #     self.assertEqual(string.empty, editableRange.single_user)
    #
    #     builder.writeln("This paragraph is outside the editable range, and cannot be edited by anybody.")
    #
    #     # Print details and contents of every editable range in the document.
    #     EditableRangePrinter editableRangePrinter = new EditableRangePrinter()
    #
    #     doc.accept(editableRangePrinter)
    #
    #     print(editableRangePrinter.to_text())
    #
    #
    # # <summary>
    # # Collects properties and contents of visited editable ranges in a string.
    # # </summary>
    # public class EditableRangePrinter : DocumentVisitor
    #
    #     public EditableRangePrinter()
    #
    #         mBuilder = new StringBuilder()
    #
    #
    #     public string ToText()
    #
    #         return mBuilder.to_string()
    #
    #
    #     public void Reset()
    #
    #         mBuilder.clear()
    #         mInsideEditableRange = false
    #
    #
    #     # <summary>
    #     # Called when an EditableRangeStart node is encountered in the document.
    #     # </summary>
    #     public override VisitorAction VisitEditableRangeStart(EditableRangeStart editableRangeStart)
    #
    #         mBuilder.append_line(" -- Editable range found! -- ")
    #         mBuilder.append_line("\tID:\t\t" + editableRangeStart.id)
    #         if (editableRangeStart.editable_range.single_user == string.empty)
    #             mBuilder.append_line("\tGroup:\t" + editableRangeStart.editable_range.editor_group)
    #         else
    #             mBuilder.append_line("\tUser:\t" + editableRangeStart.editable_range.single_user)
    #         mBuilder.append_line("\tContents:")
    #
    #         mInsideEditableRange = true
    #
    #         return VisitorAction.continue
    #
    #
    #     # <summary>
    #     # Called when an EditableRangeEnd node is encountered in the document.
    #     # </summary>
    #     public override VisitorAction VisitEditableRangeEnd(EditableRangeEnd editableRangeEnd)
    #
    #         mBuilder.append_line(" -- End of editable range --\n")
    #
    #         mInsideEditableRange = false
    #
    #         return VisitorAction.continue
    #
    #
    #     # <summary>
    #     # Called when a Run node is encountered in the document. This visitor only records runs that are inside editable ranges.
    #     # </summary>
    #     public override VisitorAction VisitRun(Run run)
    #
    #         if (mInsideEditableRange) mBuilder.append_line("\t\"" + run.text + "\"")
    #
    #         return VisitorAction.continue
    #
    #
    #     private bool mInsideEditableRange
    #     private readonly StringBuilder mBuilder
    #
    # #ExEnd

    # def test_incorrect_structure_exception(self) :
    #
    #     doc = aw.Document()
    #
    #     builder = aw.DocumentBuilder(doc)
    #
    #     # Assert that isn't valid structure for the current document.
    #     Assert.that(() => builder.end_editable_range(), Throws.type_of<InvalidOperationException>())
    #
    #     builder.start_editable_range()

    def test_incorrect_structure_do_not_added(self):
        doc = DocumentHelper().create_document_fill_with_dummy_text()
        builder = aw.DocumentBuilder(doc)

        startRange1 = builder.start_editable_range()

        builder.writeln("EditableRange_1_1")
        builder.writeln("EditableRange_1_2")

        startRange1.editable_range.editor_group = aw.EditorType.EVERYONE
        doc = DocumentHelper.save_open(doc)

        # Assert that it's not valid structure and editable ranges aren't added to the current document.
        startNodes = doc.get_child_nodes(aw.NodeType.EDITABLE_RANGE_START, True)
        self.assertEqual(0, startNodes.count)

        endNodes = doc.get_child_nodes(aw.NodeType.EDITABLE_RANGE_END, True)
        self.assertEqual(0, endNodes.count)


if __name__ == '__main__':
    unittest.main()
