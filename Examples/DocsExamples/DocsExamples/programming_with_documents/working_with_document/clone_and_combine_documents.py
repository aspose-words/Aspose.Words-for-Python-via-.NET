import io

import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class CloneAndCombineDocuments(DocsExamplesBase):

    def test_cloning_document(self):

        #ExStart:CloneDocument
        #GistId:e510e7e7b1fd08239ef592aa440675c1
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("This is the original document before applying the clone method")

        # Clone the document.
        clone = doc.clone().as_document()

        # Edit the cloned document.
        builder = aw.DocumentBuilder(clone)
        builder.write("Section 1")
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
        builder.write("Section 2")

        # This shows what is in the document originally. The document has two sections.
        self.assertEqual(
            "Section 1\x0cSection 2This is the original document before applying the clone method",
            clone.get_text().strip())

        # Duplicate the last section and append the copy to the end of the document.
        last_section_idx = clone.sections.count - 1
        new_section = clone.sections[last_section_idx].clone()
        clone.sections.add(new_section)

        # Check what the document contains after we changed it.
        self.assertEqual(
            "Section 1\x0cSection 2This is the original document before applying the clone method" +
            "\r\x0cSection 2This is the original document before applying the clone method",
            clone.get_text().strip())
        clone.save(ARTIFACTS_DIR + "CloneAndCombineDocuments.cloning_document.docx")
        #ExEnd:CloneDocument

    def test_insert_document_at_replace(self):

        #ExStart:InsertDocumentAtReplace
        #GistId:ef9ae5b20ccd67e59f9d49d32b534661
        main_doc = aw.Document(MY_DIR + "Document insertion 1.docx")

        options = aw.replacing.FindReplaceOptions()
        options.direction = aw.replacing.FindReplaceDirection.BACKWARD
        options.replacing_callback = InsertDocumentAtReplaceHandler()

        main_doc.range.replace_regex(r"\[MY_DOCUMENT\]", "", options)
        main_doc.save(ARTIFACTS_DIR + "CloneAndCombineDocuments.insert_document_at_replace.docx")
        #ExEnd:InsertDocumentAtReplace

    def test_insert_document_at_mail_merge(self):

        #ExStart:InsertDocumentAtMailMerge
        #GistId:ef9ae5b20ccd67e59f9d49d32b534661
        main_doc = aw.Document(MY_DIR + "Document insertion 1.docx")

        main_doc.mail_merge.field_merging_callback = InsertDocumentAtMailMergeHandler()
        # The main document has a merge field in it called "Document_1".
        # The corresponding data for this field contains a fully qualified path to the document.
        # That should be inserted to this field.
        main_doc.mail_merge.execute(["Document_1"], [MY_DIR + "Document insertion 2.docx"])

        main_doc.save(ARTIFACTS_DIR + "CloneAndCombineDocuments.insert_document_at_mail_merge.doc")
        #ExEnd:InsertDocumentAtMailMerge

    def test_insert_document_at_bookmark(self):

        #ExStart:InsertDocumentAtBookmark
        #GistId:ffc2b4de06eabf9183a3ed2aa34e939d
        main_doc = aw.Document(MY_DIR + "Document insertion 1.docx")
        sub_doc = aw.Document(MY_DIR + "Document insertion 2.docx")

        bookmark = main_doc.range.bookmarks.get_by_name("insertionPlace")
        self.insert_document(bookmark.bookmark_start.parent_node, sub_doc)

        main_doc.save(ARTIFACTS_DIR + "CloneAndCombineDocuments.insert_document_at_bookmark.docx")
        #ExEnd:InsertDocumentAtBookmark

    #ExStart:InsertDocumentAsNodes
    #GistId:ffc2b4de06eabf9183a3ed2aa34e939d
    @staticmethod
    def insert_document(insertion_destination: aw.Node, doc_to_insert: aw.Document):
        """Inserts content of the external document after the specified node.
        Section breaks and section formatting of the inserted document are ignored.

       :param insertion_destination: Node in the destination document after which the content
            Should be inserted. This node should be a block level node (paragraph or table).
       :param doc_to_insert: The document to insert.
        """

        if insertion_destination.node_type not in (aw.NodeType.PARAGRAPH, aw.NodeType.TABLE):
            raise ValueError("The destination node should be either a paragraph or table.")

        destination_parent = insertion_destination.parent_node

        importer = aw.NodeImporter(doc_to_insert, insertion_destination.document, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        # Loop through all block-level nodes in the section's body,
        # then clone and insert every node that is not the last empty paragraph of a section.
        for src_section in doc_to_insert.sections:
            for src_node in src_section.as_section().body.get_child_nodes(aw.NodeType.ANY, False):
                if src_node.node_type == aw.NodeType.PARAGRAPH:
                    para = src_node.as_paragraph()
                    if para.is_end_of_section and not para.has_child_nodes:
                        continue

                new_node = importer.import_node(src_node, True)

                destination_parent.insert_after(new_node, insertion_destination)
                insertion_destination = new_node
    #ExEnd:InsertDocumentAsNodes

    #ExStart:InsertDocumentWithSectionFormatting
    @staticmethod
    def insert_document_with_section_formatting(insert_after_node: aw.Node, src_doc: aw.Document):
        """Inserts content of the external document after the specified node.

       :param insert_after_node: Node in the destination document after which the content
            Should be inserted. This node should be a block level node (paragraph or table).
       :param src_doc: The document to insert."""

        if insert_after_node.node_type not in (aw.NodeType.PARAGRAPH, aw.NodeType.TABLE):
            raise ValueError("The destination node should be either a paragraph or table.")

        dst_doc = insert_after_node.document.as_document()
        # To retain section formatting, split the current section into two at the marker node and then import the content
        # from srcDoc as whole sections. The section of the node to which the insert marker node belongs.
        current_section = insert_after_node.get_ancestor(aw.NodeType.SECTION).as_section()

        # Don't clone the content inside the section, we just want the properties of the section retained.
        clone_section = current_section.clone(False).as_section()

        # However, make sure the clone section has a body but no empty first paragraph.
        clone_section.ensure_minimum()
        clone_section.body.first_paragraph.remove()

        insert_after_node.document.insert_after(clone_section, current_section)

        # Append all nodes after the marker node to the new section. This will split the content at the section level at.
        # The marker so the sections from the other document can be inserted directly.
        current_node = insert_after_node.next_sibling
        while current_node is not None:
            next_node = current_node.next_sibling
            clone_section.body.append_child(current_node)
            current_node = next_node

        # This object will be translating styles and lists during the import.
        importer = aw.NodeImporter(src_doc, dst_doc, aw.ImportFormatMode.USE_DESTINATION_STYLES)

        for src_section in src_doc.sections:
            new_node = importer.import_node(src_section, True)

            dst_doc.insert_after(new_node, current_section)
            current_section = new_node.as_section()

    #ExEnd:InsertDocumentWithSectionFormatting

#ExStart:InsertDocumentAtMailMergeHandler
#GistId:ef9ae5b20ccd67e59f9d49d32b534661
class InsertDocumentAtMailMergeHandler(aw.mailmerging.IFieldMergingCallback):
    """This handler makes special processing for the "Document_1" field.
    The field value contains the path to load the document.
    We load the document and insert it into the current merge field."""

    def field_merging(self, args: aw.mailmerging.FieldMergingArgs):
        if args.document_field_name == "Document_1":
            # Use document builder to navigate to the merge field with the specified name.
            builder = aw.DocumentBuilder(args.document)
            builder.move_to_merge_field(args.document_field_name)

            # The name of the document to load and insert is stored in the field value.
            sub_doc = aw.Document(args.field_value)

            CloneAndCombineDocuments.insert_document(builder.current_paragraph, sub_doc)

            # The paragraph that contained the merge field might be empty now, and you probably want to delete it.
            if not builder.current_paragraph.has_child_nodes:
                builder.current_paragraph.remove()

            # Indicate to the mail merge engine that we have inserted what we wanted.
            args.text = None

    def image_field_merging(self, args: aw.mailmerging.ImageFieldMergingArgs):
        # Do nothing.
        pass
#ExEnd:InsertDocumentAtMailMergeHandler

#ExStart:InsertDocumentAtMailMergeBlobHandler
class InsertDocumentAtMailMergeBlobHandler(aw.mailmerging.IFieldMergingCallback):
    """This handler makes special processing for the "Document_1" field.
    The field value contains the BLOB of the document to load.
    We load the document and insert it into the current merge field."""

    def field_merging(self, args: aw.mailmerging.FieldMergingArgs):
        if args.document_field_name == "Document_1":
            builder = aw.DocumentBuilder(args.document)
            builder.move_to_merge_field(args.document_field_name)

            stream = io.BytesIO(args.field_value)
            sub_doc = aw.Document(stream)

            CloneAndCombineDocuments.insert_document(builder.current_paragraph, sub_doc)

            # The paragraph that contained the merge field might be empty now, and you probably want to delete it.
            if not builder.current_paragraph.has_child_nodes:
                builder.current_paragraph.remove()

            args.text = None

    def image_field_merging(self, args: aw.mailmerging.ImageFieldMergingArgs):
        # Do nothing.
        pass
#ExEnd:InsertDocumentAtMailMergeBlobHandler

#ExStart:InsertDocumentAtReplaceHandler
#GistId:ef9ae5b20ccd67e59f9d49d32b534661
class InsertDocumentAtReplaceHandler(aw.replacing.IReplacingCallback):

    def replacing(self, args: aw.replacing.ReplacingArgs):
        sub_doc = aw.Document(MY_DIR + "Document insertion 2.docx")

        # Insert a document after the paragraph, containing the match text.
        para = args.match_node.parent_node.as_paragraph()
        CloneAndCombineDocuments.insert_document(para, sub_doc)

        # Remove the paragraph with the match text.
        para.remove()
        return aw.replacing.ReplaceAction.SKIP
#ExEnd:InsertDocumentAtReplaceHandler
