import unittest
import os
import sys
import aspose.words as aw

class ExtractContentHelper():

    #ExStart:CommonExtractContent
    @staticmethod
    def extract_content(startNode : aw.Node, endNode : aw.Node, isInclusive : bool) :
        
        # First, check that the nodes passed to this method are valid for use.
        ExtractContentHelper.verify_parameter_nodes(startNode, endNode)

        # Create a list to store the extracted nodes.
        nodes = []

        # If either marker is part of a comment, including the comment itself, we need to move the pointer
        # forward to the Comment Node found after the CommentRangeEnd node.
        if (endNode.node_type == aw.NodeType.COMMENT_RANGE_END and isInclusive) :
            
            node = ExtractContentHelper.find_next_node(aw.NodeType.COMMENT, endNode.next_sibling)
            if (node != None) :
                endNode = node

        # Keep a record of the original nodes passed to this method to split marker nodes if needed.
        originalStartNode = startNode
        originalEndNode = endNode

        # Extract content based on block-level nodes (paragraphs and tables). Traverse through parent nodes to find them.
        # We will split the first and last nodes' content, depending if the marker nodes are inline.
        startNode = ExtractContentHelper.get_ancestor_in_body(startNode)
        endNode = ExtractContentHelper.get_ancestor_in_body(endNode)

        isExtracting = True
        isStartingNode = True
        # The current node we are extracting from the document.
        currNode = startNode

        # Begin extracting content. Process all block-level nodes and specifically split the first
        # and last nodes when needed, so paragraph formatting is retained.
        # Method is a little more complicated than a regular extractor as we need to factor
        # in extracting using inline nodes, fields, bookmarks, etc. to make it useful.
        while (isExtracting) :
            
            # Clone the current node and its children to obtain a copy.
            cloneNode = currNode.clone(True)
            isEndingNode = currNode == endNode

            if (isStartingNode or isEndingNode) :
                
                # We need to process each marker separately, so pass it off to a separate method instead.
                # End should be processed at first to keep node indexes.
                if (isEndingNode) :
                    # !isStartingNode: don't add the node twice if the markers are the same node.
                    ExtractContentHelper.process_marker(cloneNode, nodes, originalEndNode, currNode, isInclusive, False, not isStartingNode, False)
                    isExtracting = False

                # Conditional needs to be separate as the block level start and end markers, maybe the same node.
                if (isStartingNode) :
                    ExtractContentHelper.process_marker(cloneNode, nodes, originalStartNode, currNode, isInclusive, True, True, False)
                    isStartingNode = False
                
            else :
                # Node is not a start or end marker, simply add the copy to the list.
                nodes.append(cloneNode)

            # Move to the next node and extract it. If the next node is None,
            # the rest of the content is found in a different section.
            if (currNode.next_sibling == None and isExtracting) :
                # Move to the next section.
                nextSection = currNode.get_ancestor(aw.NodeType.SECTION).next_sibling.as_section()
                currNode = nextSection.body.first_child
                
            else :
                # Move to the next node in the body.
                currNode = currNode.next_sibling
                
        # For compatibility with mode with inline bookmarks, add the next paragraph (empty).
        if (isInclusive and originalEndNode == endNode and not originalEndNode.is_composite) :
            ExtractContentHelper.include_next_paragraph(endNode, nodes)

        # Return the nodes between the node markers.
        return nodes
        
    #ExEnd:CommonExtractContent
    @staticmethod
    def paragraphs_by_style_name(doc : aw.Document, styleName : str) :
        
        # Create an array to collect paragraphs of the specified style.
        paragraphsWithStyle = []
            
        paragraphs = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
            
        # Look through all paragraphs to find those with the specified style.
        for paragraph in paragraphs :
            paragraph = paragraph.as_paragraph()
            if (paragraph.paragraph_format.style.name == styleName) :
                paragraphsWithStyle.append(paragraph)
            
        return paragraphsWithStyle
        

    #ExStart:CommonGenerateDocument
    @staticmethod
    def generate_document(srcDoc : aw.Document, nodes) :
        
        dstDoc = aw.Document()
        # Remove the first paragraph from the empty document.
        dstDoc.first_section.body.remove_all_children()

        # Import each node from the list into the new document. Keep the original formatting of the node.
        importer = aw.NodeImporter(srcDoc, dstDoc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        for node in nodes :
            importNode = importer.import_node(node, True)
            dstDoc.first_section.body.append_child(importNode)
            
        return dstDoc
        
    #ExEnd:CommonGenerateDocument

    #ExStart:CommonExtractContentHelperMethods
    @staticmethod
    def verify_parameter_nodes(startNode : aw.Node, endNode : aw.Node) :
        
        # The order in which these checks are done is important.
        if (startNode == None):
            raise ValueError("Start node cannot be None")
        if (endNode == None) :
            raise ValueError("End node cannot be None")

        if (startNode.document != endNode.document) :
            raise ValueError("Start node and end node must belong to the same document")

        if (startNode.get_ancestor(aw.NodeType.BODY) == None or endNode.get_ancestor(aw.NodeType.BODY) == None) :
            raise ValueError("Start node and end node must be a child or descendant of a body")

        # Check the end node is after the start node in the DOM tree.
        # First, check if they are in different sections, then if they're not,
        # check their position in the body of the same section.
        startSection = startNode.get_ancestor(aw.NodeType.SECTION).as_section()
        endSection = endNode.get_ancestor(aw.NodeType.SECTION).as_section()

        startIndex = startSection.parent_node.index_of(startSection)
        endIndex = endSection.parent_node.index_of(endSection)

        if (startIndex == endIndex) :
            
            if (startSection.body.index_of(ExtractContentHelper.get_ancestor_in_body(startNode)) >
                endSection.body.index_of(ExtractContentHelper.get_ancestor_in_body(endNode))):
                raise ValueError("The end node must be after the start node in the body")
            
        elif (startIndex > endIndex) :
            raise ValueError("The section of end node must be after the section start node")
        

    @staticmethod
    def find_next_node(nodeType : aw.NodeType, fromNode : aw.Node) :
        
        if (fromNode == None or fromNode.node_type == nodeType) :
            return fromNode

        if (fromNode.is_composite) :
            
            node = ExtractContentHelper.find_next_node(nodeType, fromNode.as_composite_node().first_child)
            if (node != None) :
                return node

        return ExtractContentHelper.find_next_node(nodeType, fromNode.next_sibling)
        

    @staticmethod
    def is_inline(node : aw.Node) :
        
        # Test if the node is a descendant of a Paragraph or Table node and is not a paragraph
        # or a table a paragraph inside a comment class that is decent of a paragraph is possible.
        return ((node.get_ancestor(aw.NodeType.PARAGRAPH) != None or node.get_ancestor(aw.NodeType.TABLE) != None) and
                not (node.node_type == aw.NodeType.PARAGRAPH or node.node_type == aw.NodeType.TABLE))
        

    @staticmethod
    def process_marker(cloneNode : aw.Node, nodes, node : aw.Node, blockLevelAncestor : aw.Node,
        isInclusive : bool, isStartMarker : bool, canAdd : bool, forceAdd : bool) :
        
        # If we are dealing with a block-level node, see if it should be included and add it to the list.
        if (node == blockLevelAncestor) :
            if (canAdd and isInclusive) :
                nodes.append(cloneNode)
            return
            

        # cloneNode is a clone of blockLevelNode. If node != blockLevelNode, blockLevelAncestor
        # is the node's ancestor that means it is a composite node.
        #assert(cloneNode.is_composite)

        # If a marker is a FieldStart node check if it's to be included or not.
        # We assume for simplicity that the FieldStart and FieldEnd appear in the same paragraph.
        if (node.node_type == aw.NodeType.FIELD_START) :
            # If the marker is a start node and is not included, skip to the end of the field.
            # If the marker is an end node and is to be included, then move to the end field so the field will not be removed.
            if (isStartMarker and not isInclusive or not isStartMarker and isInclusive) :
                while (node.next_sibling != None and node.node_type != aw.NodeType.FIELD_END) :
                    node = node.next_sibling
                
        # Support a case if the marker node is on the third level of the document body or lower.
        nodeBranch =  ExtractContentHelper.fill_self_and_parents(node, blockLevelAncestor)

        # Process the corresponding node in our cloned node by index.
        currentCloneNode = cloneNode
        for i in range(len(nodeBranch) - 1, 0) :
            
            currentNode = nodeBranch[i]
            nodeIndex = currentNode.parent_node.index_of(currentNode)
            currentCloneNode = currentCloneNode.as_composite_node.child_nodes[nodeIndex]

            ExtractContentHelper.remove_nodes_outside_of_range(currentCloneNode, isInclusive or (i > 0), isStartMarker)
            

        # After processing, the composite node may become empty if it has doesn't include it.
        if (canAdd and (forceAdd or cloneNode.as_composite_node().has_child_nodes)) :
            nodes.append(cloneNode)
        

    @staticmethod
    def remove_nodes_outside_of_range(markerNode : aw.Node, isInclusive : bool, isStartMarker : bool) :
        
        isProcessing = True
        isRemoving = isStartMarker
        nextNode = markerNode.parent_node.first_child

        while (isProcessing and nextNode != None) :
            
            currentNode = nextNode
            isSkip = False

            if (currentNode == markerNode) :
                if (isStartMarker) :
                    isProcessing = False
                    if (isInclusive) :
                        isRemoving = False
                else :
                    isRemoving = True
                    if (isInclusive) :
                        isSkip = True

            nextNode = nextNode.next_sibling
            if (isRemoving and not isSkip) :
                currentNode.remove()
            
        
    @staticmethod
    def fill_self_and_parents(node : aw.Node, tillNode : aw.Node) :
        
        list = []
        currentNode = node

        while (currentNode != tillNode) :
            list.append(currentNode)
            currentNode = currentNode.parent_node
            
        return list
        
    @staticmethod
    def include_next_paragraph(node : aw.Node, nodes) :
        
        paragraph = ExtractContentHelper.find_next_node(aw.NodeType.PARAGRAPH, node.next_sibling).as_paragraph()
        if (paragraph != None) :
            
            # Move to the first child to include paragraphs without content.
            markerNode = paragraph.first_child if paragraph.has_child_nodes else paragraph
            rootNode = ExtractContentHelper.get_ancestor_in_body(paragraph)

            ExtractContentHelper.process_marker(rootNode.clone(True), nodes, markerNode, rootNode,
                markerNode == paragraph, False, True, True)
            
        
    @staticmethod
    def get_ancestor_in_body(startNode : aw.Node) :
        
        while (startNode.parent_node.node_type != aw.NodeType.BODY) :
            startNode = startNode.parent_node
        return startNode
        
    #ExEnd:CommonExtractContentHelperMethods
    