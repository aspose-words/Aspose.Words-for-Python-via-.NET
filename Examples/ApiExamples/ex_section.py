# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, IMAGE_DIR

class ExSection(ApiExampleBase):

    def test_protect(self):

        #ExStart
        #ExFor:Document.protect(ProtectionType)
        #ExFor:ProtectionType
        #ExFor:Section.protected_for_forms
        #ExSummary:Shows how to turn off protection for a section.
        doc = aw.Document()

        builder = aw.DocumentBuilder(doc)
        builder.writeln("Section 1. Hello world!")
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)

        builder.writeln("Section 2. Hello again!")
        builder.write("Please enter text here: ")
        builder.insert_text_input("TextInput1", aw.fields.TextFormFieldType.REGULAR, "", "Placeholder text", 0)

        # Apply write protection to every section in the document.
        doc.protect(aw.ProtectionType.ALLOW_ONLY_FORM_FIELDS)

        # Turn off write protection for the first section.
        doc.sections[0].protected_for_forms = False

        # In this output document, we will be able to edit the first section freely,
        # and we will only be able to edit the contents of the form field in the second section.
        doc.save(ARTIFACTS_DIR + "Section.protect.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Section.protect.docx")

        self.assertFalse(doc.sections[0].protected_for_forms)
        self.assertTrue(doc.sections[1].protected_for_forms)

    def test_add_remove(self):

        #ExStart
        #ExFor:Document.sections
        #ExFor:Section.clone
        #ExFor:SectionCollection
        #ExFor:NodeCollection.remove_at(int)
        #ExSummary:Shows how to add and remove sections in a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Section 1")
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
        builder.write("Section 2")

        self.assertEqual("Section 1\u000cSection 2", doc.get_text().strip())

        # Delete the first section from the document.
        doc.sections.remove_at(0)

        self.assertEqual("Section 2", doc.get_text().strip())

        # Append a copy of what is now the first section to the end of the document.
        last_section_idx = doc.sections.count - 1
        new_section = doc.sections[last_section_idx].clone()
        doc.sections.add(new_section)

        self.assertEqual("Section 2\u000cSection 2", doc.get_text().strip())
        #ExEnd

    def test_first_and_last(self):

        #ExStart
        #ExFor:Document.first_section
        #ExFor:Document.last_section
        #ExSummary:Shows how to create a new section with a document builder.
        doc = aw.Document()

        # A blank document contains one section by default,
        # which contains child nodes that we can edit.
        self.assertEqual(1, doc.sections.count)

        # Use a document builder to add text to the first section.
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world!")

        # Create a second section by inserting a section break.
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)

        self.assertEqual(2, doc.sections.count)

        # Each section has its own page setup settings.
        # We can split the text in the second section into two columns.
        # This will not affect the text in the first section.
        doc.last_section.page_setup.text_columns.set_count(2)
        builder.writeln("Column 1.")
        builder.insert_break(aw.BreakType.COLUMN_BREAK)
        builder.writeln("Column 2.")

        self.assertEqual(1, doc.first_section.page_setup.text_columns.count)
        self.assertEqual(2, doc.last_section.page_setup.text_columns.count)

        doc.save(ARTIFACTS_DIR + "Section.first_and_last.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Section.first_and_last.docx")

        self.assertEqual(1, doc.first_section.page_setup.text_columns.count)
        self.assertEqual(2, doc.last_section.page_setup.text_columns.count)

    def test_create_manually(self):

        #ExStart
        #ExFor:Node.get_text
        #ExFor:CompositeNode.remove_all_children
        #ExFor:CompositeNode.append_child
        #ExFor:Section
        #ExFor:Section.__init__
        #ExFor:Section.page_setup
        #ExFor:PageSetup.section_start
        #ExFor:PageSetup.paper_size
        #ExFor:SectionStart
        #ExFor:PaperSize
        #ExFor:Body
        #ExFor:Body.__init__
        #ExFor:Paragraph
        #ExFor:Paragraph.__init__
        #ExFor:Paragraph.paragraph_format
        #ExFor:ParagraphFormat
        #ExFor:ParagraphFormat.style_name
        #ExFor:ParagraphFormat.alignment
        #ExFor:ParagraphAlignment
        #ExFor:Run
        #ExFor:Run.__init__(DocumentBase)
        #ExFor:Run.text
        #ExFor:Inline.font
        #ExSummary:Shows how to construct an Aspose.Words document by hand.
        doc = aw.Document()

        # A blank document contains one section, one body and one paragraph.
        # Call the "remove_all_children" method to remove all those nodes,
        # and end up with a document node with no children.
        doc.remove_all_children()

        # This document now has no composite child nodes that we can add content to.
        # If we wish to edit it, we will need to repopulate its node collection.
        # First, create a new section, and then append it as a child to the root document node.
        section = aw.Section(doc)
        doc.append_child(section)

        # Set some page setup properties for the section.
        section.page_setup.section_start = aw.SectionStart.NEW_PAGE
        section.page_setup.paper_size = aw.PaperSize.LETTER

        # A section needs a body, which will contain and display all its contents
        # on the page between the section's header and footer.
        body = aw.Body(doc)
        section.append_child(body)

        # Create a paragraph, set some formatting properties, and then append it as a child to the body.
        para = aw.Paragraph(doc)

        para.paragraph_format.style_name = "Heading 1"
        para.paragraph_format.alignment = aw.ParagraphAlignment.CENTER

        body.append_child(para)

        # Finally, add some content to do the document. Create a run,
        # set its appearance and contents, and then append it as a child to the paragraph.
        run = aw.Run(doc)
        run.text = "Hello World!"
        run.font.color = drawing.Color.red
        para.append_child(run)

        self.assertEqual("Hello World!", doc.get_text().strip())

        doc.save(ARTIFACTS_DIR + "Section.create_manually.docx")
        #ExEnd

    def test_ensure_minimum(self):

        #ExStart
        #ExFor:NodeCollection.add
        #ExFor:Section.ensure_minimum
        #ExFor:SectionCollection.__getitem__(int)
        #ExSummary:Shows how to prepare a new section node for editing.
        doc = aw.Document()

        # A blank document comes with a section, which has a body, which in turn has a paragraph.
        # We can add contents to this document by adding elements such as text runs, shapes, or tables to that paragraph.
        self.assertEqual(aw.NodeType.SECTION, doc.get_child(aw.NodeType.ANY, 0, True).node_type)
        self.assertEqual(aw.NodeType.BODY, doc.sections[0].get_child(aw.NodeType.ANY, 0, True).node_type)
        self.assertEqual(aw.NodeType.PARAGRAPH, doc.sections[0].body.get_child(aw.NodeType.ANY, 0, True).node_type)

        # If we add a new section like this, it will not have a body, or any other child nodes.
        doc.sections.add(aw.Section(doc))

        self.assertEqual(0, doc.sections[1].get_child_nodes(aw.NodeType.ANY, True).count)

        # Run the "ensure_minimum" method to add a body and a paragraph to this section to begin editing it.
        doc.last_section.ensure_minimum()

        self.assertEqual(aw.NodeType.BODY, doc.sections[1].get_child(aw.NodeType.ANY, 0, True).node_type)
        self.assertEqual(aw.NodeType.PARAGRAPH, doc.sections[1].body.get_child(aw.NodeType.ANY, 0, True).node_type)

        doc.sections[0].body.first_paragraph.append_child(aw.Run(doc, "Hello world!"))

        self.assertEqual("Hello world!", doc.get_text().strip())
        #ExEnd

    def test_body_ensure_minimum(self):

        #ExStart
        #ExFor:Section.body
        #ExFor:Body.ensure_minimum
        #ExSummary:Clears main text from all sections from the document leaving the sections themselves.
        doc = aw.Document()

        # A blank document contains one section, one body and one paragraph.
        # Call the "remove_all_children" method to remove all those nodes,
        # and end up with a document node with no children.
        doc.remove_all_children()

        # This document now has no composite child nodes that we can add content to.
        # If we wish to edit it, we will need to repopulate its node collection.
        # First, create a new section, and then append it as a child to the root document node.
        section = aw.Section(doc)
        doc.append_child(section)

        # A section needs a body, which will contain and display all its contents
        # on the page between the section's header and footer.
        body = aw.Body(doc)
        section.append_child(body)

        # This body has no children, so we cannot add runs to it yet.
        self.assertEqual(0, doc.first_section.body.get_child_nodes(aw.NodeType.ANY, True).count)

        # Call the "ensure_minimum" to make sure that this body contains at least one empty paragraph.
        body.ensure_minimum()

        # Now, we can add runs to the body, and get the document to display them.
        body.first_paragraph.append_child(aw.Run(doc, "Hello world!"))

        self.assertEqual("Hello world!", doc.get_text().strip())
        #ExEnd

    def test_body_child_nodes(self):

        #ExStart
        #ExFor:Body.node_type
        #ExFor:HeaderFooter.node_type
        #ExFor:Document.first_section
        #ExSummary:Shows how to iterate through the children of a composite node.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Section 1")
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.write("Primary header")
        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)
        builder.write("Primary footer")

        section = doc.first_section

        # A Section is a composite node and can contain child nodes,
        # but only if those child nodes are of a "BODY" or "HEADER_FOOTER" node type.
        for node in section:
            if node.node_type == aw.NodeType.BODY:
                body = node.as_body()
                print("Body:")
                print(f"\t\"{body.get_text().strip()}\"")
            elif node.node_type == aw.NodeType.HEADER_FOOTER:
                header_footer = node.as_header_footer()
                print(f"HeaderFooter type: {header_footer.header_footer_type}:")
                print(f"\t\"{header_footer.get_text().strip()}\"")
            else:
                raise Exception("Unexpected node type in a section.")

        #ExEnd

    def test_clear(self):

        #ExStart
        #ExFor:NodeCollection.clear
        #ExSummary:Shows how to remove all sections from a document.
        doc = aw.Document(MY_DIR + "Document.docx")

        # This document has one section with a few child nodes containing and displaying all the document's contents.
        self.assertEqual(1, doc.sections.count)
        self.assertEqual(17, doc.sections[0].get_child_nodes(aw.NodeType.ANY, True).count)
        self.assertEqual("Hello World!\r\rHello Word!\r\r\rHello World!", doc.get_text().strip())

        # Clear the collection of sections, which will remove all of the document's children.
        doc.sections.clear()

        self.assertEqual(0, doc.get_child_nodes(aw.NodeType.ANY, True).count)
        self.assertEqual("", doc.get_text().strip())
        #ExEnd

    def test_prepend_append_content(self):

        #ExStart
        #ExFor:Section.append_content
        #ExFor:Section.prepend_content
        #ExSummary:Shows how to append the contents of a section to another section.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Section 1")
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
        builder.write("Section 2")
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
        builder.write("Section 3")

        section = doc.sections[2]

        self.assertEqual("Section 3" + aw.ControlChar.SECTION_BREAK, section.get_text())

        # Insert the contents of the first section to the beginning of the third section.
        section_to_prepend = doc.sections[0]
        section.prepend_content(section_to_prepend)

        # Insert the contents of the second section to the end of the third section.
        section_to_append = doc.sections[1]
        section.append_content(section_to_append)

        # The "prepend_content" and "append_content" methods did not create any new sections.
        self.assertEqual(3, doc.sections.count)
        self.assertEqual("Section 1" + aw.ControlChar.PARAGRAPH_BREAK +
                         "Section 3" + aw.ControlChar.PARAGRAPH_BREAK +
                         "Section 2" + aw.ControlChar.SECTION_BREAK, section.get_text())
        #ExEnd

    def test_clear_content(self):

        #ExStart
        #ExFor:Section.clear_content
        #ExSummary:Shows how to clear the contents of a section.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Hello world!")

        self.assertEqual("Hello world!", doc.get_text().strip())
        self.assertEqual(1, doc.first_section.body.paragraphs.count)

        # Running the "clear_content" method will remove all the section contents
        # but leave a blank paragraph to add content again.
        doc.first_section.clear_content()

        self.assertEqual("", doc.get_text().strip())
        self.assertEqual(1, doc.first_section.body.paragraphs.count)
        #ExEnd

    def test_clear_headers_footers(self):

        #ExStart
        #ExFor:Section.clear_headers_footers
        #ExSummary:Shows how to clear the contents of all headers and footers in a section.
        doc = aw.Document()

        self.assertEqual(0, doc.first_section.headers_footers.count)

        builder = aw.DocumentBuilder(doc)

        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.writeln("This is the primary header.")
        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)
        builder.writeln("This is the primary footer.")

        self.assertEqual(2, doc.first_section.headers_footers.count)

        self.assertEqual("This is the primary header.", doc.first_section.headers_footers[aw.HeaderFooterType.HEADER_PRIMARY].get_text().strip())
        self.assertEqual("This is the primary footer.", doc.first_section.headers_footers[aw.HeaderFooterType.FOOTER_PRIMARY].get_text().strip())

        # Empty all the headers and footers in this section of all their contents.
        # The headers and footers themselves will still be present but will have nothing to display.
        doc.first_section.clear_headers_footers()

        self.assertEqual(2, doc.first_section.headers_footers.count)

        self.assertEqual("", doc.first_section.headers_footers[aw.HeaderFooterType.HEADER_PRIMARY].get_text().strip())
        self.assertEqual("", doc.first_section.headers_footers[aw.HeaderFooterType.FOOTER_PRIMARY].get_text().strip())
        #ExEnd

    def test_delete_header_footer_shapes(self):

        #ExStart
        #ExFor:Section.delete_header_footer_shapes
        #ExSummary:Shows how to remove all shapes from all headers footers in a section.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create a primary header with a shape.
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, 100, 100)

        # Create a primary footer with an image.
        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)
        builder.insert_image(IMAGE_DIR + "Logo Icon.ico")

        self.assertEqual(1, doc.first_section.headers_footers[aw.HeaderFooterType.HEADER_PRIMARY].get_child_nodes(aw.NodeType.SHAPE, True).count)
        self.assertEqual(1, doc.first_section.headers_footers[aw.HeaderFooterType.FOOTER_PRIMARY].get_child_nodes(aw.NodeType.SHAPE, True).count)

        # Remove all shapes from the headers and footers in the first section.
        doc.first_section.delete_header_footer_shapes()

        self.assertEqual(0, doc.first_section.headers_footers[aw.HeaderFooterType.HEADER_PRIMARY].get_child_nodes(aw.NodeType.SHAPE, True).count)
        self.assertEqual(0, doc.first_section.headers_footers[aw.HeaderFooterType.FOOTER_PRIMARY].get_child_nodes(aw.NodeType.SHAPE, True).count)
        #ExEnd

    def test_sections_clone_section(self):

        doc = aw.Document(MY_DIR + "Document.docx")
        clone_section = doc.sections[0].clone()

    def test_sections_import_section(self):

        src_doc = aw.Document(MY_DIR + "Document.docx")
        dst_doc = aw.Document()

        source_section = src_doc.sections[0]
        new_section = dst_doc.import_node(source_section, True).as_section()
        dst_doc.sections.add(new_section)

    def test_migrate_from2_x_import_section(self):

        src_doc = aw.Document()
        dst_doc = aw.Document()

        source_section = src_doc.sections[0]
        new_section = dst_doc.import_node(source_section, True).as_section()
        dst_doc.sections.add(new_section)

    def test_modify_page_setup_in_all_sections(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Section 1")
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
        builder.write("Section 2")

        # It is important to understand that a document can contain many sections,
        # and each section has its page setup. In this case, we want to modify them all.
        for section in doc:
            section = section.as_section()
            section.page_setup.paper_size = aw.PaperSize.LETTER

        doc.save(ARTIFACTS_DIR + "Section.modify_page_setup_in_all_sections.doc")

    def test_culture_info_page_setup_defaults(self):

        #Thread.current_thread.current_culture = CultureInfo("en-us")

        doc_en = aw.Document()

        # Assert that page defaults comply with current culture info.
        section_en = doc_en.sections[0]
        self.assertEqual(72.0, section_en.page_setup.left_margin) # 2.54 cm
        self.assertEqual(72.0, section_en.page_setup.right_margin) # 2.54 cm
        self.assertEqual(72.0, section_en.page_setup.top_margin) # 2.54 cm
        self.assertEqual(72.0, section_en.page_setup.bottom_margin) # 2.54 cm
        self.assertEqual(36.0, section_en.page_setup.header_distance) # 1.27 cm
        self.assertEqual(36.0, section_en.page_setup.footer_distance) # 1.27 cm
        self.assertEqual(36.0, section_en.page_setup.text_columns.spacing) # 1.27 cm

        # Change the culture and assert that the page defaults are changed.
        #Thread.current_thread.current_culture = CultureInfo("de-de")

        doc_de = aw.Document()

        section_de = doc_de.sections[0]
        self.assertEqual(70.85, section_de.page_setup.left_margin) # 2.5 cm
        self.assertEqual(70.85, section_de.page_setup.right_margin) # 2.5 cm
        self.assertEqual(70.85, section_de.page_setup.top_margin) # 2.5 cm
        self.assertEqual(56.7, section_de.page_setup.bottom_margin) # 2 cm
        self.assertEqual(35.4, section_de.page_setup.header_distance) # 1.25 cm
        self.assertEqual(35.4, section_de.page_setup.footer_distance) # 1.25 cm
        self.assertEqual(35.4, section_de.page_setup.text_columns.spacing) # 1.25 cm

        # Change page defaults.
        section_de.page_setup.left_margin = 90 # 3.17 cm
        section_de.page_setup.right_margin = 90 # 3.17 cm
        section_de.page_setup.top_margin = 72 # 2.54 cm
        section_de.page_setup.bottom_margin = 72 # 2.54 cm
        section_de.page_setup.header_distance = 35.4 # 1.25 cm
        section_de.page_setup.footer_distance = 35.4 # 1.25 cm
        section_de.page_setup.text_columns.spacing = 35.4 # 1.25 cm

        doc_de = DocumentHelper.save_open(docDe)

        section_de_after = doc_de.sections[0]
        self.assertEqual(90.0, section_de_after.page_setup.left_margin) # 3.17 cm
        self.assertEqual(90.0, section_de_after.page_setup.right_margin) # 3.17 cm
        self.assertEqual(72.0, section_de_after.page_setup.top_margin) # 2.54 cm
        self.assertEqual(72.0, section_de_after.page_setup.bottom_margin) # 2.54 cm
        self.assertEqual(35.4, section_de_after.page_setup.header_distance) # 1.25 cm
        self.assertEqual(35.4, section_de_after.page_setup.footer_distance) # 1.25 cm
        self.assertEqual(35.4, section_de_after.page_setup.text_columns.spacing) # 1.25 cm
