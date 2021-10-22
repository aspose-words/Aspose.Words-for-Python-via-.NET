import unittest
import os
import sys
import aspose.words as aw

class ExtractContentHelper():

    #ExStart:CommonExtractContent
    @staticmethod
    def extract_content(start_node: aw.Node, end_node: aw.Node, is_inclusive: bool):

        # First, check that the nodes passed to this method are valid for use.
        ExtractContentHelper.verify_parameter_nodes(start_node, end_node)

        # Create a list to store the extracted nodes.
        nodes = []

        # If either marker is part of a comment, including the comment itself, we need to move the pointer
        # forward to the Comment Node found after the CommentRangeEnd node.
        if end_node.node_type == aw.NodeType.COMMENT_RANGE_END and is_inclusive:

            node = ExtractContentHelper.find_next_node(aw.NodeType.COMMENT, end_node.next_sibling)
            if node != None:
                end_node = node

        # Keep a record of the original nodes passed to this method to split marker nodes if needed.
        original_start_node = start_node
        original_end_node = end_node

        # Extract content based on block-level nodes (paragraphs and tables). Traverse through parent nodes to find them.
        # We will split the first and last nodes' content, depending if the marker nodes are inline.
        start_node = ExtractContentHelper.get_ancestor_in_body(start_node)
        end_node = ExtractContentHelper.get_ancestor_in_body(end_node)

        is_extracting = True
        is_starting_node = True
        # The current node we are extracting from the document.
        curr_node = start_node

        # Begin extracting content. Process all block-level nodes and specifically split the first
        # and last nodes when needed, so paragraph formatting is retained.
        # Method is a little more complicated than a regular extractor as we need to factor
        # in extracting using inline nodes, fields, bookmarks, etc. to make it useful.
        while is_extracting:

            # Clone the current node and its children to obtain a copy.
            clone_node = curr_node.clone(True)
            is_ending_node = curr_node == end_node

            if is_starting_node or is_ending_node:

                # We need to process each marker separately, so pass it off to a separate method instead.
                # End should be processed at first to keep node indexes.
                if is_ending_node:
                    # !isStartingNode: don't add the node twice if the markers are the same node.
                    ExtractContentHelper.process_marker(clone_node, nodes, original_end_node, curr_node, is_inclusive, False, not is_starting_node, False)
                    is_extracting = False

                # Conditional needs to be separate as the block level start and end markers, maybe the same node.
                if is_starting_node:
                    ExtractContentHelper.process_marker(clone_node, nodes, original_start_node, curr_node, is_inclusive, True, True, False)
                    is_starting_node = False

            else:
                # Node is not a start or end marker, simply add the copy to the list.
                nodes.append(clone_node)

            # Move to the next node and extract it. If the next node is None,
            # the rest of the content is found in a different section.
            if curr_node.next_sibling == None and is_extracting:
                # Move to the next section.
                next_section = curr_node.get_ancestor(aw.NodeType.SECTION).next_sibling.as_section()
                curr_node = next_section.body.first_child

            else:
                # Move to the next node in the body.
                curr_node = curr_node.next_sibling

        # For compatibility with mode with inline bookmarks, add the next paragraph (empty).
        if is_inclusive and original_end_node == end_node and not original_end_node.is_composite:
            ExtractContentHelper.include_next_paragraph(end_node, nodes)

        # Return the nodes between the node markers.
        return nodes

    #ExEnd:CommonExtractContent
    @staticmethod
    def paragraphs_by_style_name(doc: aw.Document, style_name: str):

        # Create an array to collect paragraphs of the specified style.
        paragraphs_with_style = []

        paragraphs = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)

        # Look through all paragraphs to find those with the specified style.
        for paragraph in paragraphs:
            paragraph = paragraph.as_paragraph()
            if paragraph.paragraph_format.style.name == style_name:
                paragraphs_with_style.append(paragraph)

        return paragraphs_with_style


    #ExStart:CommonGenerateDocument
    @staticmethod
    def generate_document(src_doc: aw.Document, nodes):

        dst_doc = aw.Document()
        # Remove the first paragraph from the empty document.
        dst_doc.first_section.body.remove_all_children()

        # Import each node from the list into the new document. Keep the original formatting of the node.
        importer = aw.NodeImporter(src_doc, dst_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        for node in nodes:
            import_node = importer.import_node(node, True)
            dst_doc.first_section.body.append_child(import_node)

        return dst_doc

    #ExEnd:CommonGenerateDocument

    #ExStart:CommonExtractContentHelperMethods
    @staticmethod
    def verify_parameter_nodes(start_node: aw.Node, end_node: aw.Node):

        # The order in which these checks are done is important.
        if start_node == None:
            raise ValueError("Start node cannot be None")
        if end_node == None:
            raise ValueError("End node cannot be None")

        if start_node.document != end_node.document:
            raise ValueError("Start node and end node must belong to the same document")

        if start_node.get_ancestor(aw.NodeType.BODY) == None or end_node.get_ancestor(aw.NodeType.BODY) == None:
            raise ValueError("Start node and end node must be a child or descendant of a body")

        # Check the end node is after the start node in the DOM tree.
        # First, check if they are in different sections, then if they're not,
        # check their position in the body of the same section.
        start_section = start_node.get_ancestor(aw.NodeType.SECTION).as_section()
        end_section = end_node.get_ancestor(aw.NodeType.SECTION).as_section()

        start_index = start_section.parent_node.index_of(start_section)
        end_index = end_section.parent_node.index_of(end_section)

        if start_index == end_index:

            if (start_section.body.index_of(ExtractContentHelper.get_ancestor_in_body(start_node)) >
                end_section.body.index_of(ExtractContentHelper.get_ancestor_in_body(end_node))):
                raise ValueError("The end node must be after the start node in the body")

        elif start_index > end_index:
            raise ValueError("The section of end node must be after the section start node")


    @staticmethod
    def find_next_node(node_type: aw.NodeType, from_node: aw.Node):

        if from_node == None or from_node.node_type == node_type:
            return from_node

        if from_node.is_composite:

            node = ExtractContentHelper.find_next_node(node_type, from_node.as_composite_node().first_child)
            if node != None:
                return node

        return ExtractContentHelper.find_next_node(node_type, from_node.next_sibling)


    @staticmethod
    def is_inline(node: aw.Node):

        # Test if the node is a descendant of a Paragraph or Table node and is not a paragraph
        # or a table a paragraph inside a comment class that is decent of a paragraph is possible.
        return ((node.get_ancestor(aw.NodeType.PARAGRAPH) != None or node.get_ancestor(aw.NodeType.TABLE) != None) and
                not (node.node_type == aw.NodeType.PARAGRAPH or node.node_type == aw.NodeType.TABLE))


    @staticmethod
    def process_marker(clone_node: aw.Node, nodes, node: aw.Node, block_level_ancestor: aw.Node,
        is_inclusive: bool, is_start_marker: bool, can_add: bool, force_add: bool):

        # If we are dealing with a block-level node, see if it should be included and add it to the list.
        if node == block_level_ancestor:
            if can_add and is_inclusive:
                nodes.append(clone_node)
            return


        # cloneNode is a clone of blockLevelNode. If node != blockLevelNode, blockLevelAncestor
        # is the node's ancestor that means it is a composite node.
        #assert(cloneNode.is_composite)

        # If a marker is a FieldStart node check if it's to be included or not.
        # We assume for simplicity that the FieldStart and FieldEnd appear in the same paragraph.
        if node.node_type == aw.NodeType.FIELD_START:
            # If the marker is a start node and is not included, skip to the end of the field.
            # If the marker is an end node and is to be included, then move to the end field so the field will not be removed.
            if is_start_marker and not is_inclusive or not is_start_marker and is_inclusive:
                while node.next_sibling != None and node.node_type != aw.NodeType.FIELD_END:
                    node = node.next_sibling

        # Support a case if the marker node is on the third level of the document body or lower.
        node_branch =  ExtractContentHelper.fill_self_and_parents(node, block_level_ancestor)

        # Process the corresponding node in our cloned node by index.
        current_clone_node = clone_node
        for i in range(len(node_branch) - 1, 0):

            current_node = node_branch[i]
            node_index = current_node.parent_node.index_of(current_node)
            current_clone_node = current_clone_node.as_composite_node.child_nodes[node_index]

            ExtractContentHelper.remove_nodes_outside_of_range(current_clone_node, is_inclusive or (i > 0), is_start_marker)


        # After processing, the composite node may become empty if it has doesn't include it.
        if can_add and (force_add or clone_node.as_composite_node().has_child_nodes):
            nodes.append(clone_node)


    @staticmethod
    def remove_nodes_outside_of_range(marker_node: aw.Node, is_inclusive: bool, is_start_marker: bool):

        is_processing = True
        is_removing = is_start_marker
        next_node = marker_node.parent_node.first_child

        while is_processing and next_node != None:

            current_node = next_node
            is_skip = False

            if current_node == marker_node:
                if is_start_marker:
                    is_processing = False
                    if is_inclusive:
                        is_removing = False
                else:
                    is_removing = True
                    if is_inclusive:
                        is_skip = True

            next_node = next_node.next_sibling
            if is_removing and not is_skip:
                current_node.remove()


    @staticmethod
    def fill_self_and_parents(node: aw.Node, till_node: aw.Node):

        list = []
        current_node = node

        while current_node != till_node:
            list.append(current_node)
            current_node = current_node.parent_node

        return list

    @staticmethod
    def include_next_paragraph(node: aw.Node, nodes):

        paragraph = ExtractContentHelper.find_next_node(aw.NodeType.PARAGRAPH, node.next_sibling).as_paragraph()
        if paragraph != None:

            # Move to the first child to include paragraphs without content.
            marker_node = paragraph.first_child if paragraph.has_child_nodes else paragraph
            root_node = ExtractContentHelper.get_ancestor_in_body(paragraph)

            ExtractContentHelper.process_marker(root_node.clone(True), nodes, marker_node, root_node,
                marker_node == paragraph, False, True, True)


    @staticmethod
    def get_ancestor_in_body(start_node: aw.Node):

        while start_node.parent_node.node_type != aw.NodeType.BODY:
            start_node = start_node.parent_node
        return start_node

    #ExEnd:CommonExtractContentHelperMethods
