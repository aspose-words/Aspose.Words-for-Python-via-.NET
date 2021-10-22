import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw
import aspose.pydrawing as drawing

class WorkingWithNode(docs_base.DocsExamplesBase):

    def test_use_node_type(self) :

        #ExStart:UseNodeType
        doc = aw.Document()

        type = doc.node_type
        #ExEnd:UseNodeType


    def test_get_parent_node(self) :

        #ExStart:GetParentNode
        doc = aw.Document()

        # The section is the first child node of the document.
        section = doc.first_child

        # The section's parent node is the document.
        print(f"Section parent is the document: {doc == section.parent_node}")
        #ExEnd:GetParentNode


    def test_owner_document(self) :

        #ExStart:OwnerDocument
        doc = aw.Document()

        # Creating a new node of any type requires a document passed into the constructor.
        para = aw.Paragraph(doc)

        # The new paragraph node does not yet have a parent.
        print(f"Paragraph has no parent node: {para.parent_node == None}")

        # But the paragraph node knows its document.
        print(f"Both nodes' documents are the same: {para.document == doc}")

        # The fact that a node always belongs to a document allows us to access and modify
        # properties that reference the document-wide data, such as styles or lists.
        para.paragraph_format.style_name = "Heading 1"

        # Now add the paragraph to the main text of the first section.
        doc.first_section.body.append_child(para)

        # The paragraph node is now a child of the Body node.
        print(f"Paragraph has a parent node: {para.parent_node != None}")
        #ExEnd:OwnerDocument


    def test_enumerate_child_nodes(self) :

        #ExStart:EnumerateChildNodes
        doc = aw.Document()
        paragraph = doc.get_child(aw.NodeType.PARAGRAPH, 0, True).as_paragraph()

        children = paragraph.child_nodes
        for child in children :

            # A paragraph may contain children of various types such as runs, shapes, and others.
            if child.node_type == aw.NodeType.RUN :

                run = child.as_run()
                print(run.text)


        #ExEnd:EnumerateChildNodes


    #ExStart:RecurseAllNodes
    def test_recurse_all_nodes(self) :

        doc = aw.Document(docs_base.my_dir + "Paragraphs.docx")

        # Invoke the recursive function that will walk the tree.
        self.traverse_all_nodes(doc)


    # <summary>
    # A simple function that will walk through all children of a specified node recursively
    # and print the type of each node to the screen.
    # </summary>
    def traverse_all_nodes(self, parent_node) :

        # This is the most efficient way to loop through immediate children of a node.
        for child_node in parent_node.child_nodes :

            print(aw.Node.node_type_to_string(child_node.node_type))

            # Recurse into the node if it is a composite node.
            if child_node.is_composite :
                self.traverse_all_nodes(child_node.as_composite_node())


    #ExEnd:RecurseAllNodes

    def test_typed_access(self) :

        #ExStart:TypedAccess
        doc = aw.Document()

        section = doc.first_section
        body = section.body

        # Quick typed access to all Table child nodes contained in the Body.
        tables = body.tables

        for table in tables :

            # Quick typed access to the first row of the table.
            if table.first_row != None :
                table.first_row.remove()

            # Quick typed access to the last row of the table.
            if table.last_row != None :
                table.last_row.remove()

        #ExEnd:TypedAccess


    def test_create_and_add_paragraph_node(self) :

        #ExStart:CreateAndAddParagraphNode
        doc = aw.Document()

        para = aw.Paragraph(doc)

        section = doc.last_section
        section.body.append_child(para)
        #ExEnd:CreateAndAddParagraphNode

    def test_change_run_color(self) :

        doc = aw.Document(docs_base.my_dir + "Document.docx")

        # Get the first Run node and cast it to Run object.
        run = doc.get_child(aw.NodeType.RUN, 0, True).as_shape()

        # Make changes to the run
        run.font.color = drawing.Color.red

        # Save the result
        doc.save(docs_base.artifacts_dir + "WorkingWithNode.change_run_color.docx")


if __name__ == '__main__':
    unittest.main()
