# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from document_helper import DocumentHelper
import aspose.words as aw
import document_helper
import test_util
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR

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
        doc.protect(type=aw.ProtectionType.READ_ONLY, password='MyPassword')
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln("Hello world! Since we have set the document's protection level to read-only," + ' we cannot edit this paragraph without the password.')
        # Editable ranges allow us to leave parts of protected documents open for editing.
        editable_range_start = builder.start_editable_range()
        builder.writeln('This paragraph is inside an editable range, and can be edited.')
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
        builder.writeln('This paragraph is outside the editable range, and cannot be edited.')
        doc.save(file_name=ARTIFACTS_DIR + 'EditableRange.CreateAndRemove.docx')
        # Remove an editable range. All the nodes that were inside the range will remain intact.
        editable_range.remove()
        #ExEnd
        self.assertEqual("Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" + 'This paragraph is inside an editable range, and can be edited.\r' + 'This paragraph is outside the editable range, and cannot be edited.', doc.get_text().strip())
        self.assertEqual(0, doc.get_child_nodes(aw.NodeType.EDITABLE_RANGE_START, True).count)
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'EditableRange.CreateAndRemove.docx')
        self.assertEqual(aw.ProtectionType.READ_ONLY, doc.protection_type)
        self.assertEqual("Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" + 'This paragraph is inside an editable range, and can be edited.\r' + 'This paragraph is outside the editable range, and cannot be edited.', doc.get_text().strip())
        editable_range = doc.get_child(aw.NodeType.EDITABLE_RANGE_START, 0, True).as_editable_range_start().editable_range
        test_util.TestUtil.verify_editable_range(0, '', aw.EditorType.UNSPECIFIED, editable_range)

    def test_nested(self):
        #ExStart
        #ExFor:DocumentBuilder.start_editable_range
        #ExFor:DocumentBuilder.end_editable_range(EditableRangeStart)
        #ExFor:EditableRange.editor_group
        #ExSummary:Shows how to create nested editable ranges.
        doc = aw.Document()
        doc.protect(type=aw.ProtectionType.READ_ONLY, password='MyPassword')
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln("Hello world! Since we have set the document's protection level to read-only, " + 'we cannot edit this paragraph without the password.')
        # Create two nested editable ranges.
        outer_editable_range_start = builder.start_editable_range()
        builder.writeln('This paragraph inside the outer editable range and can be edited.')
        inner_editable_range_start = builder.start_editable_range()
        builder.writeln('This paragraph inside both the outer and inner editable ranges and can be edited.')
        # Currently, the document builder's node insertion cursor is in more than one ongoing editable range.
        # When we want to end an editable range in this situation,
        # we need to specify which of the ranges we wish to end by passing its EditableRangeStart node.
        builder.end_editable_range(inner_editable_range_start)
        builder.writeln('This paragraph inside the outer editable range and can be edited.')
        builder.end_editable_range(outer_editable_range_start)
        builder.writeln('This paragraph is outside any editable ranges, and cannot be edited.')
        # If a region of text has two overlapping editable ranges with specified groups,
        # the combined group of users excluded by both groups are prevented from editing it.
        outer_editable_range_start.editable_range.editor_group = aw.EditorType.EVERYONE
        inner_editable_range_start.editable_range.editor_group = aw.EditorType.CONTRIBUTORS
        doc.save(file_name=ARTIFACTS_DIR + 'EditableRange.Nested.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'EditableRange.Nested.docx')
        self.assertEqual("Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" + 'This paragraph inside the outer editable range and can be edited.\r' + 'This paragraph inside both the outer and inner editable ranges and can be edited.\r' + 'This paragraph inside the outer editable range and can be edited.\r' + 'This paragraph is outside any editable ranges, and cannot be edited.', doc.get_text().strip())
        editable_range = doc.get_child(aw.NodeType.EDITABLE_RANGE_START, 0, True).as_editable_range_start().editable_range
        test_util.TestUtil.verify_editable_range(0, '', aw.EditorType.EVERYONE, editable_range)
        editable_range = doc.get_child(aw.NodeType.EDITABLE_RANGE_START, 1, True).as_editable_range_start().editable_range
        test_util.TestUtil.verify_editable_range(1, '', aw.EditorType.CONTRIBUTORS, editable_range)

    def test_incorrect_structure_do_not_added(self):
        doc = document_helper.DocumentHelper.create_document_fill_with_dummy_text()
        builder = aw.DocumentBuilder(doc=doc)
        start_range1 = builder.start_editable_range()
        builder.writeln('EditableRange_1_1')
        builder.writeln('EditableRange_1_2')
        start_range1.editable_range.editor_group = aw.EditorType.EVERYONE
        doc = document_helper.DocumentHelper.save_open(doc)
        # Assert that it's not valid structure and editable ranges aren't added to the current document.
        start_nodes = doc.get_child_nodes(aw.NodeType.EDITABLE_RANGE_START, True)
        self.assertEqual(0, start_nodes.count)
        end_nodes = doc.get_child_nodes(aw.NodeType.EDITABLE_RANGE_END, True)
        self.assertEqual(0, end_nodes.count)

    def test_incorrect_structure_exception(self):
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Assert that isn't valid structure for the current document.
        with self.assertRaises(Exception):
            builder.end_editable_range()
        builder.start_editable_range()