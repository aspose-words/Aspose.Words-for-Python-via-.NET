import typing

import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithSection(DocsExamplesBase):

    def test_add_section(self):

        #ExStart:AddSection
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello1")
        builder.writeln("Hello2")

        section_to_add = aw.Section(doc)
        doc.sections.add(section_to_add)
        #ExEnd:AddSection

    def test_delete_section(self):

        #ExStart:DeleteSection
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello1")
        doc.append_child(aw.Section(doc))
        builder.writeln("Hello2")
        doc.append_child(aw.Section(doc))

        doc.sections.remove_at(0)
        #ExEnd:DeleteSection

    def test_delete_all_sections(self):

        #ExStart:DeleteAllSections
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello1")
        doc.append_child(aw.Section(doc))
        builder.writeln("Hello2")
        doc.append_child(aw.Section(doc))

        doc.sections.clear()
        #ExEnd:DeleteAllSections

    def test_append_section_content(self):

        #ExStart:AppendSectionContent
        #GistId:000cda3bfe9679c09bfd03617bd1f9e8
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Section 1")
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
        builder.write("Section 2")
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
        builder.write("Section 3")

        section = doc.sections[2]

        # Insert the contents of the first section to the beginning of the third section.
        section_to_prepend = doc.sections[0]
        section.prepend_content(section_to_prepend)

        # Insert the contents of the second section to the end of the third section.
        section_to_append = doc.sections[1]
        section.append_content(section_to_append)
        #ExEnd:AppendSectionContent

    def test_clone_section(self):

        #ExStart:CloneSection
        #GistId:000cda3bfe9679c09bfd03617bd1f9e8
        doc = aw.Document(MY_DIR + "Document.docx")
        clone_section = doc.sections[0].clone()
        #ExEnd:CloneSection

    def test_copy_section(self):

        #ExStart:CopySection
        #GistId:000cda3bfe9679c09bfd03617bd1f9e8
        src_doc = aw.Document(MY_DIR + "Document.docx")
        dst_doc = aw.Document()

        source_section = src_doc.sections[0]
        new_section = dst_doc.import_node(source_section, True).as_section()
        dst_doc.sections.add(new_section)

        dst_doc.save(ARTIFACTS_DIR + "WorkingWithSection.copy_section.docx")
        #ExEnd:CopySection

    def test_delete_header_footer_content(self):

        #ExStart:DeleteHeaderFooterContent
        #GistId:000cda3bfe9679c09bfd03617bd1f9e8
        doc = aw.Document(MY_DIR + "Document.docx")

        section = doc.sections[0]
        section.clear_headers_footers()
        #ExEnd:DeleteHeaderFooterContent

    def delete_header_footer_shapes(self):
        
        #ExStart:DeleteHeaderFooterShapes
        #GistId:000cda3bfe9679c09bfd03617bd1f9e8
        doc = aw.Document(MY_DIR + "Document.docx");

        section = doc.Sections[0];
        section.delete_header_footer_shapes();
        #ExEnd:DeleteHeaderFooterShapes
        

    def test_delete_section_content(self):

        #ExStart:DeleteSectionContent
        doc = aw.Document(MY_DIR + "Document.docx")

        section = doc.sections[0]
        section.clear_content()
        #ExEnd:DeleteSectionContent

    def test_modify_page_setup_in_all_sections(self):

        #ExStart:ModifyPageSetupInAllSections
        #GistId:000cda3bfe9679c09bfd03617bd1f9e8
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Section 1")
        doc.append_child(aw.Section(doc))
        builder.writeln("Section 2")
        doc.append_child(aw.Section(doc))
        builder.writeln("Section 3")
        doc.append_child(aw.Section(doc))
        builder.writeln("Section 4")

        # It is important to understand that a document can contain many sections,
        # and each section has its page setup. In this case, we want to modify them all.
        for child in doc:
            child.as_section().page_setup.paper_size = aw.PaperSize.LETTER

        doc.save(ARTIFACTS_DIR + "WorkingWithSection.modify_page_setup_in_all_sections.doc")
        #ExEnd:ModifyPageSetupInAllSections

    def test_sections_access_by_index(self):

        #ExStart:SectionsAccessByIndex
        doc = aw.Document(MY_DIR + "Document.docx")

        section = doc.sections[0]
        section.page_setup.left_margin = 90 # 3.17 cm
        section.page_setup.right_margin = 90 # 3.17 cm
        section.page_setup.top_margin = 72 # 2.54 cm
        section.page_setup.bottom_margin = 72 # 2.54 cm
        section.page_setup.header_distance = 35.4 # 1.25 cm
        section.page_setup.footer_distance = 35.4 # 1.25 cm
        section.page_setup.text_columns.spacing = 35.4 # 1.25 cm
        #ExEnd:SectionsAccessByIndex


    def test_section_child_nodes(self):
        #ExStart:SectionChildNodes
        #GistId:000cda3bfe9679c09bfd03617bd1f9e8
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Section 1")
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.write("Primary header")
        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)
        builder.write("Primary footer")

        section = doc.first_section

        # A Section is a composite node and can contain child nodes,
        # but only if those child nodes are of a "Body" or "HeaderFooter" node type.
        for node in section.child_nodes:
            if node.node_type == aw.NodeType.BODY:
                body = node.as_body()
                print("Body:")
                print(f"\t\"{body.get_text().strip()}\"")

            if node.node_type == aw.NodeType.HEADER_FOOTER:
                header_footer = node.as_header_footer()
                print(f"HeaderFooter type: {header_footer.header_footer_type};")
                print(f"\t\"{header_footer.get_text().strip()}\"")

        #ExEnd: SectionChildNodes

    def test_ensure_minimum(self):
        #ExStart:EnsureMinimum
        #GistId:000cda3bfe9679c09bfd03617bd1f9e8
        doc = aw.Document()

        # If we add a new section like this, it will not have a body, or any other child nodes.
        doc.sections.add(aw.Section(doc))

        # Run the "EnsureMinimum" method to add a body and a paragraph to this section to begin editing it.
        doc.last_section.ensure_minimum()

        doc.sections[0].body.first_paragraph.append_child(aw.Run(doc, "Hello world!"))
        # ExEnd: EnsureMinimum

#ExStart:InsertSectionBreaks
#GistId:000cda3bfe9679c09bfd03617bd1f9e8
    def test_insert_section_breaks(self):
        doc = aw.Document(MY_DIR + "Footnotes and endnotes.docx");
        builder = aw.DocumentBuilder(doc);

        paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True);        
        topicStartParas = []

        for para in paras:        
            style = para.paragraph_format.style_identifier;
            if style == aw.StyleIdentifier.HEADING1:
                topicStartParas.Add(para);
        
        for para in topicStartParas:            
            section = para.ParentSection;

            # Insert section break if the paragraph is not at the beginning of a section already.
            if para != section.body.first_paragraph:                
                builder.MoveTo(para.first_child);
                builder.InsertBreak(aw.BreakType.SECTION_BREAK_NEW_PAGE);

                # This is the paragraph that was inserted at the end of the now old section.
                # We don't really need the extra paragraph, we just needed the section.
                section.body.last_paragraph.remove();
#ExEnd:InsertSectionBreaks