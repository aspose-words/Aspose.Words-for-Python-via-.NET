import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class RemoveContent(docs_base.DocsExamplesBase):

        def test_remove_page_breaks(self) :

            #ExStart:OpenFromFile
            doc = aw.Document(docs_base.my_dir + "Document.docx")
            #ExEnd:OpenFromFile

            # In Aspose.words section breaks are represented as separate Section nodes in the document.
            # To remove these separate sections, the sections are combined.
            self.remove_page_breaks(doc)
            self.remove_section_breaks(doc)

            doc.save(docs_base.artifacts_dir + "RemoveContent.remove_page_breaks.docx")


        #ExStart:RemovePageBreaks
        @staticmethod
        def remove_page_breaks(doc : aw.Document) :

            paragraphs = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)

            for para in paragraphs :
                para = para.as_paragraph()

                # If the paragraph has a page break before the set, then clear it.
                if (para.paragraph_format.page_break_before) :
                    para.paragraph_format.page_break_before = False

                # Check all runs in the paragraph for page breaks and remove them.
                for run in para.runs :
                    run = run.as_run()
                    if (run.text.find(aw.ControlChar.PAGE_BREAK) >= 0) :
                        run.text = run.text.replace(aw.ControlChar.PAGE_BREAK, "")


        #ExEnd:RemovePageBreaks

        #ExStart:RemoveSectionBreaks
        @staticmethod
        def remove_section_breaks(doc : aw.Document) :

            # Loop through all sections starting from the section that precedes the last one and moving to the first section.
            for i in range(doc.sections.count - 2, 0) :

                # Copy the content of the current section to the beginning of the last section.
                doc.last_section.prepend_content(doc.sections[i])
                # Remove the copied section.
                doc.sections[i].remove()


        #ExEnd:RemoveSectionBreaks

        def test_remove_footers(self) :

            #ExStart:RemoveFooters
            doc = aw.Document(docs_base.my_dir + "Header and footer types.docx")

            for section in doc :

                section = section.as_section()
                # Up to three different footers are possible in a section (for first, even and odd pages)
                # we check and delete all of them.
                footer = section.headers_footers.get_by_header_footer_type(aw.HeaderFooterType.FOOTER_FIRST)
                if (footer != None) :
                    footer.remove()

                # Primary footer is the footer used for odd pages.
                footer = section.headers_footers.get_by_header_footer_type(aw.HeaderFooterType.FOOTER_PRIMARY)
                if (footer != None) :
                    footer.remove()

                footer = section.headers_footers.get_by_header_footer_type(aw.HeaderFooterType.FOOTER_EVEN)
                if (footer != None) :
                    footer.remove()


            doc.save(docs_base.artifacts_dir + "RemoveContent.remove_footers.docx")
            #ExEnd:RemoveFooters


        #ExStart:RemoveTOCFromDocument
        def test_remove_toc(self) :

            doc = aw.Document(docs_base.my_dir + "Table of contents.docx")

            # Remove the first table of contents from the document.
            self.remove_table_of_contents(doc, 0)

            doc.save(docs_base.artifacts_dir + "RemoveContent.remove_toc.doc")


        # <summary>
        # Removes the specified table of contents field from the document.
        # </summary>
        # <param name="doc">The document to remove the field from.</param>
        # <param name="index">The zero-based index of the TOC to remove.</param>
        @staticmethod
        def remove_table_of_contents(doc : aw.Document, index : int) :

            # Store the FieldStart nodes of TOC fields in the document for quick access.
            field_starts = []
            # This is a list to store the nodes found inside the specified TOC. They will be removed at the end of this method.
            node_list = []

            for start in doc.get_child_nodes(aw.NodeType.FIELD_START, True) :
                start = start.as_field_start()
                if (start.field_type == aw.fields.FieldType.FIELD_TOC) :
                    field_starts.append(start)

            # Ensure the TOC specified by the passed index exists.
            if (index > len(field_starts) - 1) :
                raise IndexError("TOC index is out of range")

            is_removing = True

            current_node = field_starts[index]
            while (is_removing) :

                # It is safer to store these nodes and delete them all at once later.
                node_list.append(current_node)
                current_node = current_node.next_pre_order(doc)

                # Once we encounter a FieldEnd node of type FieldTOC,
                # we know we are at the end of the current TOC and stop here.
                if (current_node.node_type == aw.NodeType.FIELD_END) :

                    field_end = current_node.as_field_end()
                    if (field_end.field_type == aw.fields.FieldType.FIELD_TOC) :
                        is_removing = False

            for node in node_list :
                node.remove()

        #ExEnd:RemoveTOCFromDocument


if __name__ == '__main__':
        unittest.main()
