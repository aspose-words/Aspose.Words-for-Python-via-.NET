# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import re

import aspose.words as aw

from api_example_base import ApiExampleBase, ARTIFACTS_DIR

class ExRenameMergeFields(ApiExampleBase):
    """Shows how to rename merge fields in a Word document."""

    def test_rename(self):
        """Finds all merge fields in a Word document and changes their names."""

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Dear ")
        builder.insert_field("MERGEFIELD  FirstName ")
        builder.write(" ")
        builder.insert_field("MERGEFIELD  LastName ")
        builder.writeln(",")
        builder.insert_field("MERGEFIELD  CustomGreeting ")

        # Select all field start nodes so we can find the MERGEFIELDs.
        field_starts = doc.get_child_nodes(aw.NodeType.FIELD_START, True)
        for field_start in field_starts:
            field_start = field_start.as_field_start()

            if field_start.field_type == aw.fields.FieldType.FIELD_MERGE_FIELD:
                merge_field = MergeField(field_start)
                merge_field.name = merge_field.name + "_Renamed"

        doc.save(ARTIFACTS_DIR + "RenameMergeFields.rename.docx")


class MergeField:
    """Represents a facade object for a merge field in a Microsoft Word document."""

    def __init__(self, field_start: aw.fields.FieldStart):

        if field_start.field_type != aw.fields.FieldType.FIELD_MERGE_FIELD:
            raise ValueError("Field start type must be FieldMergeField.")

        self.field_start = field_start

        # Find the field separator node.
        self.field_separator = MergeField.find_next_sibling(self.field_start, aw.NodeType.FIELD_SEPARATOR)
        if self.field_separator is None:
            raise Exception("Cannot find field separator.")

        # Find the field end node. Normally field end will always be found, but in the example document
        # there happens to be a paragraph break included in the hyperlink and this puts the field end
        # in the next paragraph. It will be much more complicated to handle fields which span several
        # paragraphs correctly, but in this case allowing field end to be null is enough for our purposes.
        self.field_end = MergeField.find_next_sibling(self.field_separator, aw.NodeType.FIELD_END)

    @property
    def name(self) -> str:
        """Gets the name of the merge field."""
        return MergeField.get_text_same_parent(self.field_separator.next_sibling, self.field_end).strip("«»")

    @name.setter
    def name(self, value: str):
        """Sets the name of the merge field."""

        # Merge field name is stored in the field result which is a Run
        # node between field separator and field end.
        field_result = self.field_separator.next_sibling.as_run()
        field_result.text = f"«{value}»"

        # But sometimes the field result can consist of more than one run, delete these runs.
        MergeField.remove_same_parent(field_result.next_sibling, self.field_end)

        self.update_field_code(value)

    def update_field_code(self, field_name: str):

        # Field code is stored in a Run node between field start and field separator.
        field_code = self.field_start.next_sibling.as_run()
        match = re.match(
            r"\s*(?P<start>MERGEFIELD\s|)(\s|)(?P<name>\S+)\s+",
            field_code.text)

        new_field_code = f" {match.group('start')}{field_name} "
        field_code.text = new_field_code

        # But sometimes the field code can consist of more than one run, delete these runs.
        MergeField.remove_same_parent(field_code.next_sibling, self.field_separator)

    @staticmethod
    def find_next_sibling(start_node: aw.Node, node_type: aw.NodeType) -> aw.Node:
        """Goes through siblings starting from the start node until it finds a node of the specified type or null."""

        node = start_node
        while node is not None:
            if node.node_type == node_type:
                return node
            node = node.next_sibling

        return None

    @staticmethod
    def get_text_same_parent(start_node: aw.Node, end_node: aw.Node) -> str:
        """Retrieves text from start up to but not including the end node."""

        if end_node is not None and start_node.parent_node != end_node.parent_node:
            raise ValueError("Start and end nodes are expected to have the same parent.")

        text = ""
        child = start_node
        while child != end_node:
            text += child.get_text()
            child = child.next_sibling

        return text

    @staticmethod
    def remove_same_parent(start_node: aw.Node, end_node: aw.Node):
        """Removes nodes from start up to but not including the end node.
        Start and end are assumed to have the same parent."""

        if end_node is not None and start_node.parent_node != end_node.parent_node:
            raise ValueError("Start and end nodes are expected to have the same parent.")

        cur_child = start_node
        while cur_child is not None and cur_child != end_node:
            next_child = cur_child.next_sibling
            cur_child.remove()
            cur_child = next_child
