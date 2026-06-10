# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import aspose.words.mailmerging
import system_helper
import unittest
from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExNodeImporter(ApiExampleBase):

    def test_keep_source_numbering(self):
        for keep_source_numbering in [False, True]:
            #ExStart
            #ExFor:ImportFormatOptions.keep_source_numbering
            #ExFor:NodeImporter.__init__(DocumentBase,DocumentBase,ImportFormatMode,ImportFormatOptions)
            #ExSummary:Shows how to resolve list numbering clashes in source and destination documents.
            # Open a document with a custom list numbering scheme, and then clone it.
            # Since both have the same numbering format, the formats will clash if we import one document into the other.
            src_doc = aw.Document(file_name=MY_DIR + 'Custom list numbering.docx')
            dst_doc = src_doc.clone()
            # When we import the document's clone into the original and then append it,
            # then the two lists with the same list format will join.
            # If we set the "KeepSourceNumbering" flag to "false", then the list from the document clone
            # that we append to the original will carry on the numbering of the list we append it to.
            # This will effectively merge the two lists into one.
            # If we set the "KeepSourceNumbering" flag to "true", then the document clone
            # list will preserve its original numbering, making the two lists appear as separate lists.
            import_format_options = aw.ImportFormatOptions()
            import_format_options.keep_source_numbering = keep_source_numbering
            importer = aw.NodeImporter(src_doc=src_doc, dst_doc=dst_doc, import_format_mode=aw.ImportFormatMode.KEEP_DIFFERENT_STYLES, import_format_options=import_format_options)
            for paragraph in src_doc.first_section.body.paragraphs:
                paragraph = paragraph.as_paragraph()
                imported_node = importer.import_node(paragraph, True)
                dst_doc.first_section.body.append_child(imported_node)
            dst_doc.update_list_labels()
            if keep_source_numbering:
                self.assertEqual('6. Item 1\r\n' + '7. Item 2 \r\n' + '8. Item 3\r\n' + '9. Item 4\r\n' + '6. Item 1\r\n' + '7. Item 2 \r\n' + '8. Item 3\r\n' + '9. Item 4', dst_doc.first_section.body.to_string(save_format=aw.SaveFormat.TEXT).strip())
            else:
                self.assertEqual('6. Item 1\r\n' + '7. Item 2 \r\n' + '8. Item 3\r\n' + '9. Item 4\r\n' + '10. Item 1\r\n' + '11. Item 2 \r\n' + '12. Item 3\r\n' + '13. Item 4', dst_doc.first_section.body.to_string(save_format=aw.SaveFormat.TEXT).strip())
            #ExEnd
    #ExStart
    #ExFor:Paragraph.is_end_of_section
    #ExFor:NodeImporter
    #ExFor:NodeImporter.__init__(DocumentBase,DocumentBase,ImportFormatMode)
    #ExFor:NodeImporter.import_node(Node,bool)
    #ExSummary:Shows how to insert the contents of one document to a bookmark in another document (InsertDocument).

    @staticmethod
    def insert_document(insertion_destination, doc_to_insert):
        if insertion_destination.node_type == aw.NodeType.PARAGRAPH or insertion_destination.node_type == aw.NodeType.TABLE:
            destination_parent = insertion_destination.parent_node
            importer = aw.NodeImporter(src_doc=doc_to_insert, dst_doc=insertion_destination.document, import_format_mode=aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
            # Loop through all block-level nodes in the section's body,
            # then clone and insert every node that is not the last empty paragraph of a section.
            for src_section in filter(lambda a: a is not None, map(lambda b: system_helper.linq.Enumerable.of_type(lambda x: x.as_section(), b), list(doc_to_insert.sections))):
                for src_node in src_section.body:
                    if src_node.node_type == aw.NodeType.PARAGRAPH:
                        para = src_node.as_paragraph()
                        if para.is_end_of_section and (not para.has_child_nodes):
                            continue
                    new_node = importer.import_node(src_node, True)
                    destination_parent.insert_after(new_node, insertion_destination)
                    insertion_destination = new_node
        else:
            raise Exception()
    #ExEnd

    class InsertDocumentAtMailMergeHandler(aw.mailmerging.IFieldMergingCallback):

        def field_merging(self, args):
            if args.document_field_name == 'Document_1':
                builder = aw.DocumentBuilder(doc=args.document)
                builder.move_to_merge_field(field_name=args.document_field_name)
                sub_doc = aw.Document(args.field_value)
                ExNodeImporter._insert_document(builder.current_paragraph, sub_doc)
                if not builder.current_paragraph.has_child_nodes:
                    builder.current_paragraph.remove()
                args.text = None

        def image_field_merging(self, args):
            pass