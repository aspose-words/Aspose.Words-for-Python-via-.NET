# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import re

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

#ExStart
#ExFor:NodeList
#ExFor:FieldStart
#ExSummary:Shows how to find all hyperlinks in a Word document, and then change their URLs and display names.
class ExReplaceHyperlinks(ApiExampleBase):

    def test_fields(self):

        doc = aw.Document(MY_DIR + "Hyperlinks.docx")

        # Hyperlinks in a Word documents are fields. To begin looking for hyperlinks, we must first find all the fields.
        # Use the "select_nodes" method to find all the fields in the document via an XPath.
        field_starts = doc.select_nodes("//FieldStart")

        for field_start in field_starts:
            field_start = field_start.as_field_start()

            if field_start.field_type == aw.fields.FieldType.FIELD_HYPERLINK:
                hyperlink = Hyperlink(field_start)

                # Hyperlinks that link to bookmarks do not have URLs.
                if hyperlink.is_local:
                    continue

                # Give each URL hyperlink a new URL and name.
                hyperlink.target = "http://www.aspose.com"
                hyperlink.name = "Aspose - The .NET & Java Component Publisher"

        doc.save(ARTIFACTS_DIR + "ReplaceHyperlinks.fields.docx")


class Hyperlink:
    """HYPERLINK fields contain and display hyperlinks in the document body. A field in Aspose.Words
    consists of several nodes, and it might be difficult to work with all those nodes directly.
    This implementation will work only if the hyperlink code and name each consist of only one Run node.

    The node structure for fields is as follows:

    [FieldStart][Run - field code][FieldSeparator][Run - field result][FieldEnd]

    Below are two example field codes of HYPERLINK fields:
    HYPERLINK "url"
    HYPERLINK \\l "bookmark name"

    A field's "result" property contains text that the field displays in the document body to the user."""
    def __init__(self, field_start: aw.fields.FieldStart):

        if field_start is None:
            raise ValueError("field_start")
        if field_start.field_type != aw.fields.FieldType.FIELD_HYPERLINK:
            raise ValueError("Field start type must be FieldHyperlink.")

        self.field_start = field_start

        # Find the field separator node.
        self.field_separator = Hyperlink.find_next_sibling(self.field_start, aw.NodeType.FIELD_SEPARATOR)
        if self.field_separator is None:
            raise Exception("Cannot find field separator.")

        # Normally, we can always find the field's end node, but the example document
        # contains a paragraph break inside a hyperlink, which puts the field end
        # in the next paragraph. It will be much more complicated to handle fields which span several
        # paragraphs correctly. In this case allowing field end to be null is enough.
        self.field_end = Hyperlink.find_next_sibling(self.field_separator, aw.NodeType.FIELD_END)

        # Field code looks something like "HYPERLINK "http:\\www.myurl.com"", but it can consist of several runs.
        field_code = Hyperlink.get_text_same_parent(self.field_start.next_sibling, self.field_separator)

        pattern = r"""
        \S+        # One or more non spaces HYPERLINK or other word in other languages.
        \s+        # One or more spaces.
        (?:""\s+)? # Non-capturing optional "" and one or more spaces.
        (\\l\s+)?  # Optional \l flag followed by one or more spaces.
        "          # One apostrophe.
        ([^"]+)    # One or more characters, excluding the apostrophe (hyperlink target).
        "          # One closing apostrophe.
        """

        match = re.match(pattern, field_code.strip(), re.VERBOSE)

        # The hyperlink is local if \l is present in the field code.
        self._is_local = len(match.group(2)) > 0
        self._target = match.groups(3)

    @property
    def name(self) -> str:
        """Gets the display name of the hyperlink."""
        return Hyperlink.GetTextSameParent(self.field_separator, self.field_end)

    @name.setter
    def name(self, value: str):
        """Sets the display name of the hyperlink."""

        # Hyperlink display name is stored in the field result, which is a Run
        # node between field separator and field end.
        field_result = self.field_separator.next_sibling.as_run()
        field_result.text = value

        # If the field result consists of more than one run, delete these runs.
        Hyperlink.remove_same_parent(field_result.next_sibling, self.field_end)

    @property
    def target(self) -> str:
        """Gets the target URL or bookmark name of the hyperlink."""
        return self._target

    @target.setter
    def target(self, value: str) -> str:
        """Sets the target URL or bookmark name of the hyperlink."""
        self._target = value
        self.update_field_code()

    @property
    def is_local(self) -> bool:
        """True if the hyperlinks target is a bookmark inside the document. False if the hyperlink is a URL."""
        return self._is_local

    @is_local.setter
    def is_local(self, value: bool):
        self._is_local = value
        self.update_field_code()

    def update_field_code(self):

        # A field's field code is in a Run node between the field's start node and field separator.
        field_code = self.field_start.next_sibling.as_run()
        field_code.text = 'HYPERLINK {0}"{1}"'.format(
            "\\l " if self.is_local else "", self.target)

        # If the field code consists of more than one run, delete these runs.
        Hyperlink.remove_same_parent(field_code.next_sibling, self.field_separator)

    @staticmethod
    def find_next_sibling(start_node: aw.Node, node_type: aw.NodeType) -> aw.Node:
        """Goes through siblings starting from the start node until it finds a node of the specified type or None."""

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

        text = ''
        child = start_node
        while child != end_node:
            text += child.get_text()
            child = child.next_sibling

        return text

    @staticmethod
    def remove_same_parent(start_node: aw.Node, end_node: aw.Node):
        """Removes nodes from start up to but not including the end node.
        Assumes that the start and end nodes have the same parent."""

        if end_node is not None and start_node.parent_node != end_node.parent_node:
            raise ValueError("Start and end nodes are expected to have the same parent.")

        cur_child = start_node
        while cur_child is not None and cur_child != end_node:
            next_child = cur_child.next_sibling
            cur_child.remove()
            cur_child = next_child

#ExEnd
