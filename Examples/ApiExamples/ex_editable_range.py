import unittest

import aspose.words as aw
from api_example_base import ApiExampleBase, my_dir, artifacts_dir
from document_helper import DocumentHelper
from testutil import TestUtil

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class ExEditableRange(ApiExampleBase):

    def test_create_and_remove(self):
        #ExStart
        #ExFor:DocumentBuilder.start_editable_range
        #ExFor:DocumentBuilder.end_editable_range
        #ExFor:EditableRange
        #ExFor:EditableRange.editable_range_end
        #ExFor:EditableRange.editable_range_start
        #ExFor:EditableRange.id
        #ExFor:EditableRange.remove
        #ExFor:EditableRangeEnd.editable_range_start
        #ExFor:EditableRangeEnd.id
        #ExFor:EditableRangeEnd.node_type
        #ExFor:EditableRangeStart.editable_range
        #ExFor:EditableRangeStart.id
        #ExFor:EditableRangeStart.node_type
        #ExSummary:Shows how to work with an editable range.
        doc = aw.Document()
        doc.protect(aw.ProtectionType.READ_ONLY, "MyPassword")

        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world! Since we have set the document's protection level to read-only," +
                        " we cannot edit this paragraph without the password.")

        # Editable ranges allow us to leave parts of protected documents open for editing.
        editable_range_start = builder.start_editable_range()
        builder.writeln("This paragraph is inside an editable range, and can be edited.")
        editable_range_end = builder.end_editable_range()

        # A well-formed editable range has a start node, and end node.
        # These nodes have matching IDs and encompass editable nodes.
        editable_range = editable_range_start.editable_range

        self.assertEqual(editable_range_start.id, editable_range.id)
        self.assertEqual(editable_range_end.id, editable_range.id)

        # Different parts of the editable range link to each other.
        self.assertEqual(editable_range_start.id, editable_range.editable_range_start.id)
        self.assertEqual(editable_range_start.id, editable_range_end.editable_range_start.id)
        self.assertEqual(editable_range.id, editable_range_start.editable_range.id)
        self.assertEqual(editable_range_end.id, editable_range.editable_range_end.id)

        # We can access the node types of each part like this. The editable range itself is not a node,
        # but an entity which consists of a start, an end, and their enclosed contents.
        self.assertEqual(aw.NodeType.EDITABLE_RANGE_START, editable_range_start.node_type)
        self.assertEqual(aw.NodeType.EDITABLE_RANGE_END, editable_range_end.node_type)

        builder.writeln("This paragraph is outside the editable range, and cannot be edited.")

        doc.save(ARTIFACTS_DIR + "EditableRange.create_and_remove.docx")

        # Remove an editable range. All the nodes that were inside the range will remain intact.
        editable_range.remove()
        #ExEnd

        self.assertEqual(
            "Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" +
            "This paragraph is inside an editable range, and can be edited.\r" +
            "This paragraph is outside the editable range, and cannot be edited.", doc.get_text().strip())
        self.assertEqual(0, doc.get_child_nodes(aw.NodeType.EDITABLE_RANGE_START, True).count)

        doc = aw.Document(ARTIFACTS_DIR + "EditableRange.create_and_remove.docx")

        self.assertEqual(aw.ProtectionType.READ_ONLY, doc.protection_type)
        self.assertEqual(
            "Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" +
            "This paragraph is inside an editable range, and can be edited.\r" +
            "This paragraph is outside the editable range, and cannot be edited.", doc.get_text().strip())

        editable_range = doc.get_child(aw.NodeType.EDITABLE_RANGE_START, 0, True).as_editable_range_start().editable_range
        
        TestUtil.verify_editable_range(self, 0, "", aw.EditorType.UNSPECIFIED, editable_range)

    def test_nested(self):
        #ExStart
        #ExFor:DocumentBuilder.start_editable_range
        #ExFor:DocumentBuilder.end_editable_range(EditableRangeStart)
        #ExFor:EditableRange.editor_group
        #ExSummary:Shows how to create nested editable ranges.
        doc = aw.Document()
        doc.protect(aw.ProtectionType.READ_ONLY, "MyPassword")

        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world! Since we have set the document's protection level to read-only, " +
                        "we cannot edit this paragraph without the password.")

        # Create two nested editable ranges.
        outer_editable_range_start = builder.start_editable_range()
        builder.writeln("This paragraph inside the outer editable range and can be edited.")

        inner_editable_range_start = builder.start_editable_range()
        builder.writeln("This paragraph inside both the outer and inner editable ranges and can be edited.")

        # Currently, the document builder's node insertion cursor is in more than one ongoing editable range.
        # When we want to end an editable range in this situation,
        # we need to specify which of the ranges we wish to end by passing its EditableRangeStart node.
        builder.end_editable_range(inner_editable_range_start)

        builder.writeln("This paragraph inside the outer editable range and can be edited.")

        builder.end_editable_range(outer_editable_range_start)

        builder.writeln("This paragraph is outside any editable ranges, and cannot be edited.")

        # If a region of text has two overlapping editable ranges with specified groups,
        # the combined group of users excluded by both groups are prevented from editing it.
        outer_editable_range_start.editable_range.editor_group = aw.EditorType.EVERYONE
        inner_editable_range_start.editable_range.editor_group = aw.EditorType.CONTRIBUTORS

        doc.save(ARTIFACTS_DIR + "EditableRange.nested.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "EditableRange.nested.docx")

        self.assertEqual(
            "Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" +
            "This paragraph inside the outer editable range and can be edited.\r" +
            "This paragraph inside both the outer and inner editable ranges and can be edited.\r" +
            "This paragraph inside the outer editable range and can be edited.\r" +
            "This paragraph is outside any editable ranges, and cannot be edited.", doc.get_text().strip())

        editable_range = doc.get_child(aw.NodeType.EDITABLE_RANGE_START, 0, True).as_editable_range_start().editable_range
        
        TestUtil.verify_editable_range(self, 0, "", aw.EditorType.EVERYONE, editable_range)
        
        editable_range = doc.get_child(aw.NodeType.EDITABLE_RANGE_START, 1, True).as_editable_range_start().editable_range
        
        TestUtil.verify_editable_range(self, 1, "", aw.EditorType.CONTRIBUTORS, editable_range)

    #ExStart
    #ExFor:EditableRange
    #ExFor:EditableRange.editor_group
    #ExFor:EditableRange.single_user
    #ExFor:EditableRangeEnd
    #ExFor:EditableRangeEnd.accept(DocumentVisitor)
    #ExFor:EditableRangeStart
    #ExFor:EditableRangeStart.accept(DocumentVisitor)
    #ExFor:EditorType
    #ExSummary:Shows how to limit the editing rights of editable ranges to a specific group/user.
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
    #     EditableRange editable_range = builder.start_editable_range().editable_range
    #     editable_range.single_user = "john.doe@myoffice.com"
    #     builder.writeln($"This paragraph is inside the first editable range, can only be edited by editable_range.single_user.")
    #     builder.end_editable_range()
    #
    #     self.assertEqual(EditorType.unspecified, editable_range.editor_group)
    #
    #     # 2 -  Specify a group that allowed users are associated with:
    #     editable_range = builder.start_editable_range().editable_range
    #     editable_range.editor_group = EditorType.administrators
    #     builder.writeln($"This paragraph is inside the first editable range, can only be edited by editable_range.editor_group.")
    #     builder.end_editable_range()
    #
    #     self.assertEqual(string.empty, editable_range.single_user)
    #
    #     builder.writeln("This paragraph is outside the editable range, and cannot be edited by anybody.")
    #
    #     # Print details and contents of every editable range in the document.
    #     EditableRangePrinter editable_rangePrinter = new EditableRangePrinter()
    #
    #     doc.accept(editable_rangePrinter)
    #
    #     print(editable_rangePrinter.to_text())
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
    #     public override VisitorAction VisitEditableRangeStart(EditableRangeStart editable_range_start)
    #
    #         mBuilder.append_line(" -- Editable range found! -- ")
    #         mBuilder.append_line("\tID:\t\t" + editable_range_start.id)
    #         if (editable_range_start.editable_range.single_user == string.empty)
    #             mBuilder.append_line("\tGroup:\t" + editable_range_start.editable_range.editor_group)
    #         else
    #             mBuilder.append_line("\tUser:\t" + editable_range_start.editable_range.single_user)
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
    #     public override VisitorAction VisitEditableRangeEnd(EditableRangeEnd editable_range_end)
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

    def test_incorrect_structure_exception(self):
    
        doc = aw.Document()
    
        builder = aw.DocumentBuilder(doc)
    
        # Assert that isn't valid structure for the current document.
        with self.assertRaises(Exception):
            builder.end_editable_range()
    
        builder.start_editable_range()

    def test_incorrect_structure_do_not_added(self):
        doc = DocumentHelper.create_document_fill_with_dummy_text()
        builder = aw.DocumentBuilder(doc)

        start_range1 = builder.start_editable_range()

        builder.writeln("EditableRange_1_1")
        builder.writeln("EditableRange_1_2")

        start_range1.editable_range.editor_group = aw.EditorType.EVERYONE
        doc = DocumentHelper.save_open(doc)

        # Assert that it's not valid structure and editable ranges aren't added to the current document.
        start_nodes = doc.get_child_nodes(aw.NodeType.EDITABLE_RANGE_START, True)
        self.assertEqual(0, start_nodes.count)

        end_nodes = doc.get_child_nodes(aw.NodeType.EDITABLE_RANGE_END, True)
        self.assertEqual(0, end_nodes.count)
