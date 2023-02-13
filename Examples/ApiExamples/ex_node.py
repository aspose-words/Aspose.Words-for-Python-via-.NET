# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import io
import unittest

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, IMAGE_DIR

class ExNode(ApiExampleBase):

    def test_clone_composite_node(self):

        #ExStart
        #ExFor:Node
        #ExFor:Node.clone
        #ExSummary:Shows how to clone a composite node.
        doc = aw.Document()
        para = doc.first_section.body.first_paragraph
        para.append_child(aw.Run(doc, "Hello world!"))

        # Below are two ways of cloning a composite node.
        # 1 -  Create a clone of a node, and create a clone of each of its child nodes as well.
        clone_with_children = para.clone(True)

        self.assertTrue(clone_with_children.as_composite_node().has_child_nodes)
        self.assertEqual("Hello world!", clone_with_children.get_text().strip())

        # 2 -  Create a clone of a node just by itself without any children.
        clone_without_children = para.clone(False)

        self.assertFalse(clone_without_children.as_composite_node().has_child_nodes)
        self.assertEqual("", clone_without_children.get_text().strip())
        #ExEnd

    def test_get_parent_node(self):

        #ExStart
        #ExFor:Node.parent_node
        #ExSummary:Shows how to access a node's parent node.
        doc = aw.Document()
        para = doc.first_section.body.first_paragraph

        # Append a child Run node to the document's first paragraph.
        run = aw.Run(doc, "Hello world!")
        para.append_child(run)

        # The paragraph is the parent node of the run node. We can trace this lineage
        # all the way to the document node, which is the root of the document's node tree.
        self.assertEqual(para, run.parent_node)
        self.assertEqual(doc.first_section.body, para.parent_node)
        self.assertEqual(doc.first_section, doc.first_section.body.parent_node)
        self.assertEqual(doc, doc.first_section.parent_node)
        #ExEnd

    def test_owner_document(self):

        #ExStart
        #ExFor:Node.document
        #ExFor:Node.parent_node
        #ExSummary:Shows how to create a node and set its owning document.
        doc = aw.Document()
        para = aw.Paragraph(doc)
        para.append_child(aw.Run(doc, "Hello world!"))

        # We have not yet appended this paragraph as a child to any composite node.
        self.assertIsNone(para.parent_node)

        # If a node is an appropriate child node type of another composite node,
        # we can attach it as a child only if both nodes have the same owner document.
        # The owner document is the document we passed to the node's constructor.
        # We have not attached this paragraph to the document, so the document does not contain its text.
        self.assertEqual(para.document, doc)
        self.assertEqual("", doc.get_text().strip())

        # Since the document owns this paragraph, we can apply one of its styles to the paragraph's contents.
        para.paragraph_format.style = doc.styles.get_by_name("Heading 1")

        # Add this node to the document, and then verify its contents.
        doc.first_section.body.append_child(para)

        self.assertEqual(doc.first_section.body, para.parent_node)
        self.assertEqual("Hello world!", doc.get_text().strip())
        #ExEnd

        self.assertEqual(doc, para.document)
        self.assertIsNotNone(para.parent_node)

    def test_child_nodes_enumerate(self):

        #ExStart
        #ExFor:Node
        #ExFor:Node.custom_node_id
        #ExFor:NodeType
        #ExFor:CompositeNode
        #ExFor:CompositeNode.get_child
        #ExFor:CompositeNode.child_nodes
        #ExFor:CompositeNode.__iter__
        #ExFor:NodeCollection.count
        #ExFor:NodeCollection.__getitem__
        #ExSummary:Shows how to traverse through a composite node's collection of child nodes.
        doc = aw.Document()

        # Add two runs and one shape as child nodes to the first paragraph of this document.
        paragraph = doc.get_child(aw.NodeType.PARAGRAPH, 0, True).as_paragraph()
        paragraph.append_child(aw.Run(doc, "Hello world! "))

        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.RECTANGLE)
        shape.width = 200
        shape.height = 200
        # Note that the 'custom_node_id' is not saved to an output file and exists only during the node lifetime.
        shape.custom_node_id = 100
        shape.wrap_type = aw.drawing.WrapType.INLINE
        paragraph.append_child(shape)

        paragraph.append_child(aw.Run(doc, "Hello again!"))

        # Iterate through the paragraph's collection of immediate children,
        # and print any runs or shapes that we find within.
        children = paragraph.child_nodes

        self.assertEqual(3, paragraph.child_nodes.count)

        for child in children:
            if child.node_type == aw.NodeType.RUN:
                print("Run contents:")
                print(f"\t\"{child.get_text().strip()}\"")

            elif child.node_type == aw.NodeType.SHAPE:
                child_shape = child.as_shape()
                print("Shape:")
                print(f"\t{child_shape.shape_type}, {child_shape.width}x{child_shape.height}")
                self.assertEqual(100, shape.custom_node_id) #ExSkip

        #ExEnd

        self.assertEqual(aw.NodeType.RUN, paragraph.get_child(aw.NodeType.RUN, 0, True).node_type)
        self.assertEqual("Hello world! Hello again!", doc.get_text().strip())

    #ExStart
    #ExFor:Node.next_sibling
    #ExFor:CompositeNode.first_child
    #ExFor:Node.is_composite
    #ExFor:CompositeNode.is_composite
    #ExFor:Node.node_type_to_string
    #ExFor:Paragraph.node_type
    #ExFor:Table.node_type
    #ExFor:Node.node_type
    #ExFor:Footnote.node_type
    #ExFor:FormField.node_type
    #ExFor:SmartTag.node_type
    #ExFor:Cell.node_type
    #ExFor:Row.node_type
    #ExFor:Document.node_type
    #ExFor:Comment.node_type
    #ExFor:Run.node_type
    #ExFor:Section.node_type
    #ExFor:SpecialChar.node_type
    #ExFor:Shape.node_type
    #ExFor:FieldEnd.node_type
    #ExFor:FieldSeparator.node_type
    #ExFor:FieldStart.node_type
    #ExFor:BookmarkStart.node_type
    #ExFor:CommentRangeEnd.node_type
    #ExFor:BuildingBlock.node_type
    #ExFor:GlossaryDocument.node_type
    #ExFor:BookmarkEnd.node_type
    #ExFor:GroupShape.node_type
    #ExFor:CommentRangeStart.node_type
    #ExSummary:Shows how to traverse a composite node's tree of child nodes.
    def test_recurse_children(self):

        doc = aw.Document(MY_DIR + "Paragraphs.docx")

        # Any node that can contain child nodes, such as the document itself, is composite.
        self.assertTrue(doc.is_composite)

        # Invoke the recursive function that will go through and print all the child nodes of a composite node.
        ExNode.traverse_all_nodes(doc, 0)

    @staticmethod
    def traverse_all_nodes(parent_node: aw.CompositeNode, depth: int):
        """Recursively traverses a node tree while printing the type of each node
        with an indent depending on depth as well as the contents of all inline nodes."""

        child_node = parent_node.first_child
        while child_node is not None:

            print("\t" * depth + aw.Node.node_type_to_string(child_node.node_type), end="")

            # Recurse into the node if it is a composite node. Otherwise, print its contents if it is an inline node.
            if child_node.is_composite:
                print()
                ExNode.traverse_all_nodes(child_node.as_composite_node(), depth + 1)

            elif child_node is aw.Inline:
                print(f" - \"{child_node.get_text().strip()}\"")

            else:
                print()

            child_node = child_node.next_sibling

    #ExEnd

    def test_remove_nodes(self):

        #ExStart
        #ExFor:Node
        #ExFor:Node.node_type
        #ExFor:Node.remove
        #ExSummary:Shows how to remove all child nodes of a specific type from a composite node.
        doc = aw.Document(MY_DIR + "Tables.docx")

        self.assertEqual(2, doc.get_child_nodes(aw.NodeType.TABLE, True).count)

        cur_node = doc.first_section.body.first_child

        while cur_node is not None:

            # Save the next sibling node as a variable in case we want to move to it after deleting this node.
            next_node = cur_node.next_sibling

            # A section body can contain Paragraph and Table nodes.
            # If the node is a Table, remove it from the parent.
            if cur_node.node_type == aw.NodeType.TABLE:
                cur_node.remove()

            cur_node = next_node

        self.assertEqual(0, doc.get_child_nodes(aw.NodeType.TABLE, True).count)
        #ExEnd

    def test_enum_next_sibling(self):

        #ExStart
        #ExFor:CompositeNode.first_child
        #ExFor:Node.next_sibling
        #ExFor:Node.node_type_to_string
        #ExFor:Node.node_type
        #ExSummary:Shows how to use a node's next_sibling property to enumerate through its immediate children.
        doc = aw.Document(MY_DIR + "Paragraphs.docx")

        node = doc.first_section.body.first_child
        while node is not None:
            print()
            print("Node type:", aw.Node.node_type_to_string(node.node_type))

            contents = node.get_text().strip()
            print("This node contains no text" if contents == "" else f'Contents: "{node.get_text().strip()}"')

            node = node.next_sibling
        #ExEnd

    def test_typed_access(self):

        #ExStart
        #ExFor:Story.tables
        #ExFor:Table.first_row
        #ExFor:Table.last_row
        #ExFor:TableCollection
        #ExSummary:Shows how to remove the first and last rows of all tables in a document.
        doc = aw.Document(MY_DIR + "Tables.docx")

        tables = doc.first_section.body.tables

        self.assertEqual(5, tables[0].rows.count)
        self.assertEqual(4, tables[1].rows.count)

        for table in tables:
            table = table.as_table()

            if table.first_row is not None:
                table.first_row.remove()

            if table.last_row is not None:
                table.last_row.remove()

        self.assertEqual(3, tables[0].rows.count)
        self.assertEqual(2, tables[1].rows.count)
        #ExEnd

    def test_remove_child(self):

        #ExStart
        #ExFor:CompositeNode.last_child
        #ExFor:Node.previous_sibling
        #ExFor:CompositeNode.remove_child
        #ExSummary:Shows how to use of methods of Node and CompositeNode to remove a section before the last section in the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Section 1 text.")
        builder.insert_break(aw.BreakType.SECTION_BREAK_CONTINUOUS)
        builder.writeln("Section 2 text.")

        # Both sections are siblings of each other.
        last_section = doc.last_child.as_section()
        first_section = last_section.previous_sibling.as_section()

        # Remove a section based on its sibling relationship with another section.
        if last_section.previous_sibling is not None:
            doc.remove_child(first_section)

        # The section we removed was the first one, leaving the document with only the second.
        self.assertEqual("Section 2 text.", doc.get_text().strip())
        #ExEnd

    def test_select_composite_nodes(self):

        #ExStart
        #ExFor:CompositeNode.select_single_node
        #ExFor:CompositeNode.select_nodes
        #ExFor:NodeList.__iter__
        #ExFor:NodeList.to_array
        #ExSummary:Shows how to select certain nodes by using an XPath expression.
        doc = aw.Document(MY_DIR + "Tables.docx")

        # This expression will extract all paragraph nodes,
        # which are descendants of any table node in the document.
        node_list = doc.select_nodes("//Table//Paragraph")

        # Iterate through the list with an enumerator and print the contents of every paragraph in each cell of the table.
        index = 0

        for node in node_list:
            print(f'Table paragraph index {index}, contents: "{node.get_text().strip()}"')
            index += 1

        # This expression will select any paragraphs that are direct children of any Body node in the document.
        node_list = doc.select_nodes("//Body/Paragraph")

        # We can treat the list as an array.
        self.assertEqual(4, len(node_list.to_array()))

        # Use "select_single_node" to select the first result of the same expression as above.
        node = doc.select_single_node("//Body/Paragraph")

        self.assertIsInstance(node.as_paragraph(), aw.Paragraph)
        #ExEnd

    def test_node_is_inside_field(self):

        #ExStart
        #ExFor:CompositeNode.select_nodes
        #ExSummary:Shows how to use an XPath expression to test whether a node is inside a field.
        doc = aw.Document(MY_DIR + "Mail merge destination - Northwind employees.docx")

        # The NodeList that results from this XPath expression will contain all nodes we find inside a field.
        # However, FieldStart and FieldEnd nodes can be on the list if there are nested fields in the path.
        # Currently does not find rare fields in which the FieldCode or FieldResult spans across multiple paragraphs.
        result_list = doc.select_nodes("//FieldStart/following-sibling::node()[following-sibling::FieldEnd]")

        # Check if the specified run is one of the nodes that are inside the field.
        for node in result_list:
            if node.node_type == aw.NodeType.RUN:
                print("Contents of the first Run node that's part of a field:", node.get_text().strip())
                break
        #ExEnd

    def test_create_and_add_paragraph_node(self):

        doc = aw.Document()

        para = aw.Paragraph(doc)

        section = doc.last_section
        section.body.append_child(para)

    def test_remove_smart_tags_from_composite_node(self):

        #ExStart
        #ExFor:CompositeNode.remove_smart_tags
        #ExSummary:Removes all smart tags from descendant nodes of a composite node.
        doc = aw.Document(MY_DIR + "Smart tags.doc")

        self.assertEqual(8, doc.get_child_nodes(aw.NodeType.SMART_TAG, True).count)

        doc.remove_smart_tags()

        self.assertEqual(0, doc.get_child_nodes(aw.NodeType.SMART_TAG, True).count)
        #ExEnd

    def test_get_index_of_node(self):

        #ExStart
        #ExFor:CompositeNode.index_of
        #ExSummary:Shows how to get the index of a given child node from its parent.
        doc = aw.Document(MY_DIR + "Rendering.docx")

        body = doc.first_section.body

        # Retrieve the index of the last paragraph in the body of the first section.
        self.assertEqual(24, body.child_nodes.index_of(body.last_paragraph))
        #ExEnd

    def test_convert_node_to_html_with_default_options(self):

        #ExStart
        #ExFor:Node.to_string(SaveFormat)
        #ExFor:Node.to_string(SaveOptions)
        #ExSummary:Exports the content of a node to String in HTML format.
        doc = aw.Document(MY_DIR + "Document.docx")

        node = doc.last_section.body.last_paragraph

        # When we call the "to_string" method using the html SaveFormat overload,
        # it converts the node's contents to their raw html representation.
        self.assertEqual('<p style="margin-top:0pt; margin-bottom:8pt; line-height:108%; font-size:12pt">' +
                         '<span style="font-family:\'Times New Roman\'">Hello World!</span>' +
                         '</p>', node.to_string(aw.SaveFormat.HTML))

        # We can also modify the result of this conversion using a SaveOptions object.
        save_options = aw.saving.HtmlSaveOptions()
        save_options.export_relative_font_size = True

        self.assertEqual('<p style="margin-top:0pt; margin-bottom:8pt; line-height:108%">' +
                         '<span style="font-family:\'Times New Roman\'">Hello World!</span>' +
                         '</p>', node.to_string(save_options))
        #ExEnd

    def test_typed_node_collection_to_array(self):

        #ExStart
        #ExFor:ParagraphCollection.to_array
        #ExSummary:Shows how to create an array from a NodeCollection.
        doc = aw.Document(MY_DIR + "Paragraphs.docx")

        paras = doc.first_section.body.paragraphs.to_array()

        self.assertEqual(22, len(paras))
        #ExEnd

    def test_node_enumeration_hot_remove(self):

        #ExStart
        #ExFor:ParagraphCollection.to_array
        #ExSummary:Shows how to use "hot remove" to remove a node during enumeration.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("The first paragraph")
        builder.writeln("The second paragraph")
        builder.writeln("The third paragraph")
        builder.writeln("The fourth paragraph")

        # Remove a node from the collection in the middle of an enumeration.
        for para in doc.first_section.body.paragraphs:
            para = para.as_paragraph()
            if "third" in para.range.text:
                para.remove()

        self.assertNotIn("The third paragraph", doc.get_text())
        #ExEnd

    @unittest.skip("aspose.words.Document.create_navigator method isn't supported yet")
    #ExStart
    #ExFor:CompositeNode.create_navigator
    #ExSummary:Shows how to create an XPathNavigator, and then use it to traverse and read nodes.
    def test_node_xpath_navigator(self):

        doc = aw.Document()
        navigator = doc.create_navigator()

        if navigator is not None:
            self.assertEqual("Document", navigator.name)
            self.assertEqual(False, navigator.move_to_next())
            self.assertEqual(1, navigator.select_children(aw.XPathNodeType.ALL).count)

            # The document tree has the document, first section,
            # body, and first paragraph as nodes, with each being an only child of the previous.
            # We can add a few more to give the tree some branches for the navigator to traverse.
            doc_builder = aw.DocumentBuilder(doc)
            doc_builder.write("Section 1, Paragraph 1. ")
            doc_builder.insert_paragraph()
            doc_builder.write("Section 1, Paragraph 2. ")
            doc.append_child(aw.Section(doc))
            doc_builder.move_to_section(1)
            doc_builder.write("Section 2, Paragraph 1. ")

            # Use our navigator to print a map of all the nodes in the document to the console.
            string_stream = io.StringIO()
            ExNode.map_document(navigator, string_stream, 0)
            print(string_stream.getvalue())
            self._test_node_xpath_navigator(string_stream.getvalue(), doc) #ExSkip

    @staticmethod
    def map_document(navigator, string_stream: io.StringIO, depth: int):
        """Traverses all children of a composite node and map the structure in the style of a directory tree.
        The amount of space indentation indicates depth relative to the initial node.
        Prints the text contents of the current node only if it is a Run."""

        while True:
            string_stream.write(" " * depth)
            string_stream.write(navigator.name + ": ")

            if navigator.name == "Run":
                string_stream.append(navigator.value)

            string_stream.append("\n")

            if navigator.has_children:
                navigator.move_to_first_child()
                ExNode.map_document(navigator, string_stream, depth + 1)
                navigator.move_to_parent()

            if not navigator.move_to_next():
                break

    #ExEnd

    def _test_node_xpath_navigator(self, navigator_result: str, doc: aw.Document):

        for run in doc.get_child_nodes(aw.NodeType.RUN, True):
            run = run.as_run()
            self.assertTrue(navigator_result.contains(run.get_text().strip()))

    ##ExStart
    ##ExFor:NodeChangingAction
    ##ExFor:NodeChangingArgs.action
    ##ExFor:NodeChangingArgs.new_parent
    ##ExFor:NodeChangingArgs.old_parent
    ##ExSummary:Shows how to use a NodeChangingCallback to monitor changes to the document tree in real-time as we edit it.
    #def test_node_changing_callback(self):

    #    doc = aw.Document()
    #    doc.node_changing_callback = ExNode.NodeChangingPrinter()

    #    builder = aw.DocumentBuilder(doc)
    #    builder.writeln("Hello world!")
    #    builder.start_table()
    #    builder.insert_cell()
    #    builder.write("Cell 1")
    #    builder.insert_cell()
    #    builder.write("Cell 2")
    #    builder.end_table()

    #    builder.insert_image(drawing.Image.from_file(IMAGE_DIR + "Logo.jpg"))

    #    builder.current_paragraph.parent_node.remove_all_children()

    #class NodeChangingPrinter(aw.INodeChangingCallback):
    #    """Prints every node insertion/removal as it takes place in the document."""

    #    def node_inserting(self, args: aw.NodeChangingArgs):

    #        self.assertEqual(aw.NodeChangingAction.INSERT, args.action)
    #        self.assertIsNone(args.old_parent)

    #    def node_inserted(self, args: aw.NodeChangingArgs):

    #        self.assertEqual(aw.NodeChangingAction.INSERT, args.action)
    #        self.assertIsNotNone(args.new_parent)

    #        print("Inserted node:")
    #        print(f"\tType:\t{args.node.node_type}")

    #        if args.node.get_text().strip() != "":
    #            print(f"\tText:\t\"{args.node.get_text().strip()}\"")

    #        print(f"\tHash:\t{args.node.get_hash_code()}")
    #        print(f"\tParent:\t{args.new_parent.node_type} ({args.new_parent.get_hash_code()})")

    #    def node_removing(self, args: aw.NodeChangingArgs):

    #        self.assertEqual(aw.NodeChangingAction.REMOVE, args.action)

    #    def node_removed(self, args: aw.NodeChangingArgs):

    #        self.assertEqual(aw.NodeChangingAction.REMOVE, args.action)
    #        self.assertIsNone(args.new_parent)

    #        print(f"Removed node: {args.node.node_type} ({args.node.get_hash_code()})")

    ##ExEnd

    def test_node_collection(self):

        #ExStart
        #ExFor:NodeCollection.contains(Node)
        #ExFor:NodeCollection.insert(int,Node)
        #ExFor:NodeCollection.remove(Node)
        #ExSummary:Shows how to work with a NodeCollection.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Add text to the document by inserting Runs using a DocumentBuilder.
        builder.write("Run 1. ")
        builder.write("Run 2. ")

        # Every invocation of the "write" method creates a new Run,
        # which then appears in the parent Paragraph's RunCollection.
        runs = doc.first_section.body.first_paragraph.runs

        self.assertEqual(2, runs.count)

        # We can also insert a node into the RunCollection manually.
        new_run = aw.Run(doc, "Run 3. ")
        runs.insert(3, new_run)

        self.assertTrue(runs.contains(new_run))
        self.assertEqual("Run 1. Run 2. Run 3.", doc.get_text().strip())

        # Access individual runs and remove them to remove their text from the document.
        run = runs[1]
        runs.remove(run)

        self.assertEqual("Run 1. Run 3.", doc.get_text().strip())
        self.assertIsNotNone(run)
        self.assertFalse(runs.contains(run))
        #ExEnd

    @unittest.skip("drawing.Image type isn't supported yet")
    def test_node_list(self):

        #ExStart
        #ExFor:NodeList.count
        #ExFor:NodeList.__getitem__(int)
        #ExSummary:Shows how to use XPaths to navigate a NodeList.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert some nodes with a DocumentBuilder.
        builder.writeln("Hello world!")

        builder.start_table()
        builder.insert_cell()
        builder.write("Cell 1")
        builder.insert_cell()
        builder.write("Cell 2")
        builder.end_table()

        builder.insert_image(drawing.Image.from_file(IMAGE_DIR + "Logo.jpg"))

        # Our document contains three Run nodes.
        node_list = doc.select_nodes("//Run")

        self.assertEqual(3, node_list.count)
        self.assertTrue(any(node.get_text().strip() == "Hello world!" for node in node_list))
        self.assertTrue(any(node.get_text().strip() == "Cell 1" for node in node_list))
        self.assertTrue(any(node.get_text().strip() == "Cell 2" for node in node_list))

        # Use a double forward slash to select all Run nodes
        # that are indirect descendants of a Table node, which would be the runs inside the two cells we inserted.
        node_list = doc.select_nodes("//Table//Run")

        self.assertEqual(2, node_list.count)
        self.assertTrue(any(node.get_text().strip() == "Cell 1" for node in node_list))
        self.assertTrue(any(node.get_text().strip() == "Cell 2" for node in node_list))

        # Single forward slashes specify direct descendant relationships,
        # which we skipped when we used double slashes.
        self.assertEqual(
            doc.select_nodes("//Table//Run"),
            doc.select_nodes("//Table/Row/Cell/Paragraph/Run"))

        # Access the shape that contains the image we inserted.
        node_list = doc.select_nodes("//Shape")

        self.assertEqual(1, node_list.count)

        shape = node_list[0].as_shape()
        self.assertTrue(shape.has_image)
        #ExEnd
