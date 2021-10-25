import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class JoinAndAppendDocuments(docs_base.DocsExamplesBase):

    def test_simple_append_document(self):

        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")

        # Append the source document to the destination document using no extra options.
        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.simple_append_document.docx")


    def test_append_document(self):

        #ExStart:AppendDocumentManually
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")

        # Loop through all sections in the source document.
        # Section nodes are immediate children of the Document node so we can just enumerate the Document.
        for src_section in src_doc:

            # Because we are copying a section from one document to another,
            # it is required to import the Section node into the destination document.
            # This adjusts any document-specific references to styles, lists, etc.
            #
            # Importing a node creates a copy of the original node, but the copy
            # ss ready to be inserted into the destination document.
            dst_section = dst_doc.import_node(src_section, True, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

            # Now the new section node can be appended to the destination document.
            dst_doc.append_child(dst_section)


        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.append_document.docx")
        #ExEnd:AppendDocumentManually


    def test_append_document_to_blank(self):

        #ExStart:AppendDocumentToBlank
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document()

        # The destination document is not empty, often causing a blank page to appear before the appended document.
        # This is due to the base document having an empty section and the new document being started on the next page.
        # Remove all content from the destination document before appending.
        dst_doc.remove_all_children()
        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.append_document_to_blank.docx")
        #ExEnd:AppendDocumentToBlank


    def test_append_with_import_format_options(self):

        #ExStart:AppendWithImportFormatOptions
        src_doc = aw.Document(docs_base.my_dir + "Document source with list.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Document destination with list.docx")

        # Specify that if numbering clashes in source and destination documents,
        # then numbering from the source document will be used.
        options = aw.ImportFormatOptions()
        options.keep_source_numbering = True

        dst_doc.append_document(src_doc, aw.ImportFormatMode.USE_DESTINATION_STYLES, options)
        #ExEnd:AppendWithImportFormatOptions


    def test_convert_num_page_fields(self):

        #ExStart:ConvertNumPageFields
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")

        # Restart the page numbering on the start of the source document.
        src_doc.first_section.page_setup.restart_page_numbering = True
        src_doc.first_section.page_setup.page_starting_number = 1

        # Append the source document to the end of the destination document.
        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        # After joining the documents the NUMPAGE fields will now display the total number of pages which
        # is undesired behavior. Call this method to fix them by replacing them with PAGEREF fields.
        self.convert_num_page_fields_to_page_ref(dst_doc)

        # This needs to be called in order to update the new fields with page numbers.
        dst_doc.update_page_layout()

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.convert_num_page_fields.docx")
        #ExEnd:ConvertNumPageFields


    #ExStart:ConvertNumPageFieldsToPageRef
    def convert_num_page_fields_to_page_ref(self, doc: aw.Document):

        # This is the prefix for each bookmark, which signals where page numbering restarts.
        # The underscore "_" at the start inserts this bookmark as hidden in MS Word.
        bookmark_prefix = "_SubDocumentEnd"
        num_pages_field_name = "NUMPAGES"
        page_ref_field_name = "PAGEREF"

        # Defines the number of page restarts encountered and, therefore,
        # the number of "sub" documents found within this document.
        sub_document_count = 0

        builder = aw.DocumentBuilder(doc)

        for section in doc.sections:

            section = section.as_section()
            # This section has its page numbering restarted to treat this as the start of a sub-document.
            # Any PAGENUM fields in this inner document must be converted to special PAGEREF fields to correct numbering.
            if section.page_setup.restart_page_numbering:

                # Don't do anything if this is the first section of the document.
                # This part of the code will insert the bookmark marking the end of the previous sub-document so,
                # therefore, it does not apply to the first section in the document.
                if section != doc.first_section:

                    # Get the previous section and the last node within the body of that section.
                    prev_section = section.previous_sibling.as_section()
                    last_node = prev_section.body.last_child

                    builder.move_to(last_node)

                    # This bookmark represents the end of the sub-document.
                    builder.start_bookmark(bookmark_prefix + str(sub_document_count))
                    builder.end_bookmark(bookmark_prefix + str(sub_document_count))

                    # Increase the sub-document count to insert the correct bookmarks.
                    sub_document_count += 1


            # The last section needs the ending bookmark to signal that it is the end of the current sub-document.
            if section == doc.last_section:

                # Insert the bookmark at the end of the body of the last section.
                # Don't increase the count this time as we are just marking the end of the document.
                last_node = doc.last_section.body.last_child

                builder.move_to(last_node)
                builder.start_bookmark(bookmark_prefix + str(sub_document_count))
                builder.end_bookmark(bookmark_prefix + str(sub_document_count))


            # Iterate through each NUMPAGES field in the section and replace it with a PAGEREF field
            # referring to the bookmark of the current sub-document. This bookmark is positioned at the end
            # of the sub-document but does not exist yet. It is inserted when a section with restart page numbering
            # or the last section is encountered.
            nodes = section.get_child_nodes(aw.NodeType.FIELD_START, True).to_array()

            for field_start in nodes:

                field_start = field_start.as_field_start()
                if field_start.field_type == aw.fields.FieldType.FIELD_NUM_PAGES:

                    field_code = self.get_field_code(field_start)
                    # Since the NUMPAGES field does not take any additional parameters,
                    # we can assume the field's remaining part. Code after the field name is the switches.
                    # We will use these to help recreate the NUMPAGES field as a PAGEREF field.
                    field_switches = field_code.replace(num_pages_field_name, "").strip()

                    # Inserting the new field directly at the FieldStart node of the original field will cause
                    # the new field not to pick up the original field's formatting. To counter this,
                    # insert the field just before the original field if a previous run cannot be found,
                    # we are forced to use the FieldStart node.
                    previous_node = field_start.previous_sibling if (field_start.previous_sibling is not None) else field_start

                    # Insert a PAGEREF field at the same position as the field.
                    builder.move_to(previous_node)

                    new_field = builder.insert_field(f" {page_ref_field_name} bookmarkPrefixsubDocumentCount {field_switches} ")

                    # The field will be inserted before the referenced node. Move the node before the field instead.
                    previous_node.parent_node.insert_before(previous_node, new_field.start)

                    # Remove the original NUMPAGES field from the document.
                    self.remove_field(field_start)


    #ExEnd:ConvertNumPageFieldsToPageRef

    #ExStart:GetRemoveField
    @staticmethod
    def remove_field(field_start: aw.fields.FieldStart):

        is_removing = True

        current_node = field_start
        while current_node is not None and is_removing:

            if current_node.node_type == aw.NodeType.FIELD_END:
                is_removing = False

            next_node = current_node.next_pre_order(current_node.document)
            current_node.remove()
            current_node = next_node


    @staticmethod
    def get_field_code(field_start: aw.fields.FieldStart):

        builder = ""

        node = field_start
        while ((node is not None) and (node.node_type != aw.NodeType.FIELD_SEPARATOR) and (node.node_type != aw.NodeType.FIELD_END)):
            if node.node_type == aw.NodeType.RUN:
                builder += node.get_text()
            node = node.next_pre_order(node.document)

        return builder

    #ExEnd:GetRemoveField

    def test_different_page_setup(self):

        #ExStart:DifferentPageSetup
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")

        # Set the source document to continue straight after the end of the destination document.
        src_doc.first_section.page_setup.section_start = aw.SectionStart.CONTINUOUS

        # Restart the page numbering on the start of the source document.
        src_doc.first_section.page_setup.restart_page_numbering = True
        src_doc.first_section.page_setup.page_starting_number = 1

        # To ensure this does not happen when the source document has different page setup settings, make sure the
        # settings are identical between the last section of the destination document.
        # If there are further continuous sections that follow on in the source document,
        # this will need to be repeated for those sections.
        src_doc.first_section.page_setup.page_width = dst_doc.last_section.page_setup.page_width
        src_doc.first_section.page_setup.page_height = dst_doc.last_section.page_setup.page_height
        src_doc.first_section.page_setup.orientation = dst_doc.last_section.page_setup.orientation

        # Iterate through all sections in the source document.
        for para in src_doc.get_child_nodes(aw.NodeType.PARAGRAPH, True):
            para = para.as_paragraph()
            para.paragraph_format.keep_with_next = True


        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.different_page_setup.docx")
        #ExEnd:DifferentPageSetup


    def test_join_continuous(self):

        #ExStart:JoinContinuous
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")

        # Make the document appear straight after the destination documents content.
        src_doc.first_section.page_setup.section_start = aw.SectionStart.CONTINUOUS
        # Append the source document using the original styles found in the source document.
        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.join_continuous.docx")
        #ExEnd:JoinContinuous


    def test_join_new_page(self):

        #ExStart:JoinNewPage
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")

        # Set the appended document to start on a new page.
        src_doc.first_section.page_setup.section_start = aw.SectionStart.NEW_PAGE
        # Append the source document using the original styles found in the source document.
        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.join_new_page.docx")
        #ExEnd:JoinNewPage


    def test_keep_source_formatting(self):

        #ExStart:KeepSourceFormatting
        dst_doc = aw.Document()
        dst_doc.first_section.body.append_paragraph("Destination document text. ")

        src_doc = aw.Document()
        src_doc.first_section.body.append_paragraph("Source document text. ")

        # Append the source document to the destination document.
        # Pass format mode to retain the original formatting of the source document when importing it.
        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.keep_source_formatting.docx")
        #ExEnd:KeepSourceFormatting


    def test_keep_source_together(self):

        #ExStart:KeepSourceTogether
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Document destination with list.docx")

        # Set the source document to appear straight after the destination document's content.
        src_doc.first_section.page_setup.section_start = aw.SectionStart.CONTINUOUS

        for para in src_doc.get_child_nodes(aw.NodeType.PARAGRAPH, True):
            para = para.as_paragraph()
            para.paragraph_format.keep_with_next = True


        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.keep_source_together.docx")
        #ExEnd:KeepSourceTogether


    def test_list_keep_source_formatting(self):

        #ExStart:ListKeepSourceFormatting
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Document destination with list.docx")

        # Append the content of the document so it flows continuously.
        src_doc.first_section.page_setup.section_start = aw.SectionStart.CONTINUOUS

        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.list_keep_source_formatting.docx")
        #ExEnd:ListKeepSourceFormatting


    def test_list_use_destination_styles(self):

        #ExStart:ListUseDestinationStyles
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Document destination with list.docx")

        # Set the source document to continue straight after the end of the destination document.
        src_doc.first_section.page_setup.section_start = aw.SectionStart.CONTINUOUS

        # Keep track of the lists that are created.
        new_lists = {}

        for para in src_doc.get_child_nodes(aw.NodeType.PARAGRAPH, True):
            para = para.as_paragraph()
            if para.is_list_item:

                list_id = para.list_format.list.list_id

                # Check if the destination document contains a list with this ID already. If it does, then this may
                # cause the two lists to run together. Create a copy of the list in the source document instead.
                if dst_doc.lists.get_list_by_list_id(list_id) is not None:

                    # A newly copied list already exists for this ID, retrieve the stored list,
                    # and use it on the current paragraph.
                    if new_lists.contains_key(list_id):
                        current_list = new_lists[list_id]
                    else:
                        # Add a copy of this list to the document and store it for later reference.
                        current_list = src_doc.lists.add_copy(para.list_format.list)
                        new_lists.add(list_id, current_list)

                    # Set the list of this paragraph to the copied list.
                    para.list_format.list = current_list


        # Append the source document to end of the destination document.
        dst_doc.append_document(src_doc, aw.ImportFormatMode.USE_DESTINATION_STYLES)

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.list_use_destination_styles.docx")
        #ExEnd:ListUseDestinationStyles


    def test_restart_page_numbering(self):

        #ExStart:RestartPageNumbering
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")

        src_doc.first_section.page_setup.section_start = aw.SectionStart.NEW_PAGE
        src_doc.first_section.page_setup.restart_page_numbering = True

        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.restart_page_numbering.docx")
        #ExEnd:RestartPageNumbering


    def test_update_page_layout(self):

        #ExStart:UpdatePageLayout
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")

        # If the destination document is rendered to PDF, image etc.
        # or UpdatePageLayout is called before the source document. Is appended,
        # then any changes made after will not be reflected in the rendered output
        dst_doc.update_page_layout()

        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        # For the changes to be updated to rendered output, UpdatePageLayout must be called again.
        # If not called again, the appended document will not appear in the output of the next rendering.
        dst_doc.update_page_layout()

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.update_page_layout.docx")
        #ExEnd:UpdatePageLayout


    def test_use_destination_styles(self):

        #ExStart:UseDestinationStyles
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")

        # Append the source document using the styles of the destination document.
        dst_doc.append_document(src_doc, aw.ImportFormatMode.USE_DESTINATION_STYLES)

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.use_destination_styles.docx")
        #ExEnd:UseDestinationStyles


    def test_smart_style_behavior(self):

        #ExStart:SmartStyleBehavior
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")
        builder = aw.DocumentBuilder(dst_doc)

        builder.move_to_document_end()
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        options = aw.ImportFormatOptions()
        options.smart_style_behavior = True

        builder.insert_document(src_doc, aw.ImportFormatMode.USE_DESTINATION_STYLES, options)
        builder.document.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.smart_style_behavior.docx")
        #ExEnd:SmartStyleBehavior


    def test_insert_document_with_builder(self):

        #ExStart:InsertDocumentWithBuilder
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")
        builder = aw.DocumentBuilder(dst_doc)

        builder.move_to_document_end()
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        builder.insert_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
        builder.document.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.insert_document_with_builder.docx")
        #ExEnd:InsertDocumentWithBuilder


    def test_keep_source_numbering(self):

        #ExStart:KeepSourceNumbering
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")

        # Keep source list formatting when importing numbered paragraphs.
        import_format_options = aw.ImportFormatOptions()
        import_format_options.keep_source_numbering = True

        importer = aw.NodeImporter(src_doc, dst_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, import_format_options)

        src_paras = src_doc.first_section.body.paragraphs

        for src_para in src_paras:
            imported_node = importer.import_node(src_para, False)
            dst_doc.first_section.body.append_child(imported_node)


        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.keep_source_numbering.docx")
        #ExEnd:KeepSourceNumbering


    def test_ignore_text_boxes(self):

        #ExStart:IgnoreTextBoxes
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")

        # Keep the source text boxes formatting when importing.
        import_format_options = aw.ImportFormatOptions()
        import_format_options.ignore_text_boxes = False

        importer = aw.NodeImporter(src_doc, dst_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, import_format_options)

        src_paras = src_doc.first_section.body.paragraphs
        for src_para in src_paras:
            imported_node = importer.import_node(src_para, True)
            dst_doc.first_section.body.append_child(imported_node)


        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.ignore_text_boxes.docx")
        #ExEnd:IgnoreTextBoxes


    def test_ignore_header_footer(self):

        #ExStart:IgnoreHeaderFooter
        src_document = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_document = aw.Document(docs_base.my_dir + "Northwind traders.docx")

        import_format_options = aw.ImportFormatOptions()
        import_format_options.ignore_header_footer = False

        dst_document.append_document(src_document, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, import_format_options)

        dst_document.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.ignore_header_footer.docx")
        #ExEnd:IgnoreHeaderFooter


    def test_link_headers_footers(self):

        #ExStart:LinkHeadersFooters
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")

        # Set the appended document to appear on a new page.
        src_doc.first_section.page_setup.section_start = aw.SectionStart.NEW_PAGE
        # Link the headers and footers in the source document to the previous section.
        # This will override any headers or footers already found in the source document.
        src_doc.first_section.headers_footers.link_to_previous(True)

        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.link_headers_footers.docx")
        #ExEnd:LinkHeadersFooters


    def test_remove_source_headers_footers(self):

        #ExStart:RemoveSourceHeadersFooters
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")

        # Remove the headers and footers from each of the sections in the source document.
        for section in src_doc.sections:
            section.as_section().clear_headers_footers()

        # Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting
        # for HeadersFooters can still be set. This will cause the headers and footers to continue from the destination
        # document. This should set to False to avoid this behavior.
        src_doc.first_section.headers_footers.link_to_previous(False)

        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.remove_source_headers_footers.docx")
        #ExEnd:RemoveSourceHeadersFooters


    def test_unlink_headers_footers(self):

        #ExStart:UnlinkHeadersFooters
        src_doc = aw.Document(docs_base.my_dir + "Document source.docx")
        dst_doc = aw.Document(docs_base.my_dir + "Northwind traders.docx")

        # Unlink the headers and footers in the source document to stop this
        # from continuing the destination document's headers and footers.
        src_doc.first_section.headers_footers.link_to_previous(False)

        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        dst_doc.save(docs_base.artifacts_dir + "JoinAndAppendDocuments.unlink_headers_footers.docx")
        #ExEnd:UnlinkHeadersFooters


if __name__ == '__main__':
    unittest.main()
