# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExNodeImporter(ApiExampleBase):

    def test_keep_source_numbering(self):
        for keep_source_numbering in (False, True):
            with self.subTest(keep_source_numbering=keep_source_numbering):
                #ExStart
                #ExFor:ImportFormatOptions.keep_source_numbering
                #ExFor:NodeImporter.__init__(DocumentBase,DocumentBase,ImportFormatMode,ImportFormatOptions)
                #ExSummary:Shows how to resolve list numbering clashes in source and destination documents.
                # Open a document with a custom list numbering scheme, and then clone it.
                # Since both have the same numbering format, the formats will clash if we import one document into the other.
                src_doc = aw.Document(MY_DIR + 'Custom list numbering.docx')
                dst_doc = src_doc.clone()
                # When we import the document's clone into the original and then append it,
                # then the two lists with the same list format will join.
                # If we set the "keep_source_numbering" flag to "False", then the list from the document clone
                # that we append to the original will carry on the numbering of the list we append it to.
                # This will effectively merge the two lists into one.
                # If we set the "keep_source_numbering" flag to "True", then the document clone
                # list will preserve its original numbering, making the two lists appear as separate lists.
                import_format_options = aw.ImportFormatOptions()
                import_format_options.keep_source_numbering = keep_source_numbering
                importer = aw.NodeImporter(src_doc, dst_doc, aw.ImportFormatMode.KEEP_DIFFERENT_STYLES, import_format_options)
                for paragraph in src_doc.first_section.body.paragraphs:
                    paragraph = paragraph.as_paragraph()
                    imported_node = importer.import_node(paragraph, True)
                    dst_doc.first_section.body.append_child(imported_node)
                dst_doc.update_list_labels()
                if keep_source_numbering:
                    self.assertEqual('6. Item 1\r\n' + '7. Item 2 \r\n' + '8. Item 3\r\n' + '9. Item 4\r\n' + '6. Item 1\r\n' + '7. Item 2 \r\n' + '8. Item 3\r\n' + '9. Item 4', dst_doc.first_section.body.to_string(aw.SaveFormat.TEXT).strip())
                else:
                    self.assertEqual('6. Item 1\r\n' + '7. Item 2 \r\n' + '8. Item 3\r\n' + '9. Item 4\r\n' + '10. Item 1\r\n' + '11. Item 2 \r\n' + '12. Item 3\r\n' + '13. Item 4', dst_doc.first_section.body.to_string(aw.SaveFormat.TEXT).strip())

    def test_insert_at_bookmark(self):
        #ExStart
        #ExFor:Paragraph.is_end_of_section
        #ExFor:NodeImporter
        #ExFor:NodeImporter.__init__(DocumentBase,DocumentBase,ImportFormatMode)
        #ExFor:NodeImporter.import_node(Node,bool)
        #ExSummary:Shows how to insert the contents of one document to a bookmark in another document.

        def insert_at_bookmark():
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc)
            builder.start_bookmark('InsertionPoint')
            builder.write('We will insert a document here: ')
            builder.end_bookmark('InsertionPoint')
            doc_to_insert = aw.Document()
            builder = aw.DocumentBuilder(doc_to_insert)
            builder.write('Hello world!')
            doc_to_insert.save(ARTIFACTS_DIR + 'NodeImporter.insert_at_bookmark.docx')
            bookmark = doc.range.bookmarks.get_by_name('InsertionPoint')
            insert_document(bookmark.bookmark_start.parent_node, doc_to_insert)
            self.assertEqual('We will insert a document here: ' + '\rHello world!', doc.get_text().strip())

        def insert_document(insertion_destination: aw.Node, doc_to_insert: aw.Document):
            """Inserts the contents of a document after the specified node."""
            if insertion_destination.node_type == aw.NodeType.PARAGRAPH or insertion_destination.node_type == aw.NodeType.TABLE:
                destination_parent = insertion_destination.parent_node
                importer = aw.NodeImporter(doc_to_insert, insertion_destination.document, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
                # Loop through all block-level nodes in the section's body,
                # then clone and insert every node that is not the last empty paragraph of a section.
                for src_section in doc_to_insert.sections:
                    src_section = src_section.as_section()
                    for src_node in src_section.body:
                        if src_node.node_type == aw.NodeType.PARAGRAPH:
                            para = src_node.as_paragraph()
                            if para.is_end_of_section and (not para.has_child_nodes):
                                continue
                        new_node = importer.import_node(src_node, True)
                        destination_parent.insert_after(new_node, insertion_destination)
                        insertion_destination = new_node
            else:
                raise Exception('The destination node should be either a paragraph or table.')
        #ExEnd
        insert_at_bookmark()