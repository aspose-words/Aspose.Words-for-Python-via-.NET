import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithBookmarks(docs_base.DocsExamplesBase):

    def test_access_bookmarks(self):

        #ExStart:AccessBookmarks
        doc = aw.Document(docs_base.my_dir + "Bookmarks.docx")

        # By index:
        bookmark1 = doc.range.bookmarks[0]
        # By name:
        bookmark2 = doc.range.bookmarks.get_by_name("MyBookmark3")
        #ExEnd:AccessBookmarks


    def test_update_bookmark_data(self):

        #ExStart:UpdateBookmarkData
        doc = aw.Document(docs_base.my_dir + "Bookmarks.docx")

        bookmark = doc.range.bookmarks.get_by_name("MyBookmark1")

        name = bookmark.name
        text = bookmark.text

        bookmark.name = "RenamedBookmark"
        bookmark.text = "This is a new bookmarked text."
        #ExEnd:UpdateBookmarkData


    def test_bookmark_table_columns(self):

        #ExStart:BookmarkTable
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_table()

        builder.insert_cell()

        builder.start_bookmark("MyBookmark")

        builder.write("This is row 1 cell 1")

        builder.insert_cell()
        builder.write("This is row 1 cell 2")

        builder.end_row()

        builder.insert_cell()
        builder.writeln("This is row 2 cell 1")

        builder.insert_cell()
        builder.writeln("This is row 2 cell 2")

        builder.end_row()
        builder.end_table()

        builder.end_bookmark("MyBookmark")
        #ExEnd:BookmarkTable

        #ExStart:BookmarkTableColumns
        for bookmark in doc.range.bookmarks:

            print("Bookmark: " + bookmark.name + " (Column)" if bookmark.is_column else "")

            if bookmark.is_column:

                row = bookmark.bookmark_start.get_ancestor(aw.NodeType.ROW).as_row()
                if bookmark.first_column < row.cells.count:
                    print(row.cells[bookmark.first_column].get_text().trim_end(aw.ControlChar.CELL_CHAR))


        #ExEnd:BookmarkTableColumns


    def test_copy_bookmarked_text(self):

        src_doc = aw.Document(docs_base.my_dir + "Bookmarks.docx")

        # This is the bookmark whose content we want to copy.
        src_bookmark = src_doc.range.bookmarks.get_by_name("MyBookmark1")

        # We will be adding to this document.
        dst_doc = aw.Document()

        # Let's say we will be appended to the end of the body of the last section.
        dst_node = dst_doc.last_section.body

        # If you import multiple times without a single context, it will result in many styles created.
        importer = aw.NodeImporter(src_doc, dst_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        self.append_bookmarked_text(importer, src_bookmark, dst_node)

        dst_doc.save(docs_base.artifacts_dir + "WorkingWithBookmarks.copy_bookmarked_text.docx")


    # <summary>
    # Copies content of the bookmark and adds it to the end of the specified node.
    # The destination node can be in a different document.
    # </summary>
    # <param name="importer">Maintains the import context.</param>
    # <param name="srcBookmark">The input bookmark.</param>
    # <param name="dstNode">Must be a node that can contain paragraphs (such as a Story).</param>
    @staticmethod
    def append_bookmarked_text(importer: aw.NodeImporter, src_bookmark: aw.Bookmark, dst_node: aw.CompositeNode):

        # This is the paragraph that contains the beginning of the bookmark.
        start_para = src_bookmark.bookmark_start.parent_node.as_paragraph()

        # This is the paragraph that contains the end of the bookmark.
        end_para = src_bookmark.bookmark_end.parent_node.as_paragraph()

        if start_para is None or end_para is None:
            raise RuntimeError("Parent of the bookmark start or end is not a paragraph, cannot handle this scenario yet.")

        # Limit ourselves to a reasonably simple scenario.
        if start_para.parent_node != end_para.parent_node:
            raise RuntimeError("Start and end paragraphs have different parents, cannot handle this scenario yet.")

        # We want to copy all paragraphs from the start paragraph up to (and including) the end paragraph,
        # therefore the node at which we stop is one after the end paragraph.
        end_node = end_para.next_sibling

        cur_node = start_para
        while cur_node != end_node:

            # This creates a copy of the current node and imports it (makes it valid) in the context
            # of the destination document. Importing means adjusting styles and list identifiers correctly.
            new_node = importer.import_node(cur_node, True)
            dst_node.append_child(new_node)
            cur_node = cur_node.next_sibling


    def test_create_bookmark(self):

        #ExStart:CreateBookmark
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_bookmark("My Bookmark")
        builder.writeln("Text inside a bookmark.")

        builder.start_bookmark("Nested Bookmark")
        builder.writeln("Text inside a NestedBookmark.")
        builder.end_bookmark("Nested Bookmark")

        builder.writeln("Text after Nested Bookmark.")
        builder.end_bookmark("My Bookmark")

        options = aw.saving.PdfSaveOptions()
        options.outline_options.bookmarks_outline_levels.add("My Bookmark", 1)
        options.outline_options.bookmarks_outline_levels.add("Nested Bookmark", 2)

        doc.save(docs_base.artifacts_dir + "WorkingWithBookmarks.create_bookmark.pdf", options)
        #ExEnd:CreateBookmark


    def test_show_hide_bookmarks(self):

        #ExStart:ShowHideBookmarks
        doc = aw.Document(docs_base.my_dir + "Bookmarks.docx")

        self.show_hide_bookmarked_content(doc, "MyBookmark1", False)

        doc.save(docs_base.artifacts_dir + "WorkingWithBookmarks.show_hide_bookmarks.docx")
        #ExEnd:ShowHideBookmarks


    #ExStart:ShowHideBookmarkedContent
    @staticmethod
    def show_hide_bookmarked_content(doc: aw.Document, bookmark_name: str, show_hide: bool):

        bookmark = doc.range.bookmarks.get_by_name(bookmark_name)

        builder = aw.DocumentBuilder(doc)
        builder.move_to_document_end()

        # IF "MERGEFIELD bookmark" = "True" "" ""
        field = builder.insert_field('IF "', None)
        builder.move_to(field.start.next_sibling)
        builder.insert_field("MERGEFIELD " + bookmark_name + "", None)
        builder.write('" = "True" ')
        builder.write('"')
        builder.write('"')
        builder.write(' ""')

        current_node = field.start
        flag = True
        while current_node is not None and flag:

            if current_node.node_type == aw.NodeType.RUN:
                if current_node.to_string(aw.SaveFormat.TEXT).strip() == '"':
                    flag = False

            next_node = current_node.next_sibling

            bookmark.bookmark_start.parent_node.insert_before(current_node, bookmark.bookmark_start)
            current_node = next_node


        end_node = bookmark.bookmark_end
        flag = True
        while current_node is not None and flag:

            if current_node.node_type == aw.NodeType.FIELD_END:
                flag = False

            next_node = current_node.next_sibling

            bookmark.bookmark_end.parent_node.insert_after(current_node, end_node)
            end_node = current_node
            current_node = next_node


        doc.mail_merge.execute([ bookmark_name ], [ show_hide ])

    #ExEnd:ShowHideBookmarkedContent

    def test_untangle_row_bookmarks(self):

        doc = aw.Document(docs_base.my_dir + "Table column bookmarks.docx")

        # This performs the custom task of putting the row bookmark ends into the same row with the bookmark starts.
        self.untangle(doc)

        # Now we can easily delete rows by a bookmark without damaging any other row's bookmarks.
        self.delete_row_by_bookmark(doc, "ROW2")

        # This is just to check that the other bookmark was not damaged.
        if doc.range.bookmarks.get_by_name("ROW1").bookmark_end is None:
            raise RuntimeError("Wrong, the end of the bookmark was deleted.")

        doc.save(docs_base.artifacts_dir + "WorkingWithBookmarks.untangle_row_bookmarks.docx")


    @staticmethod
    def untangle(doc: aw.Document):

        for bookmark in doc.range.bookmarks:

            # Get the parent row of both the bookmark and bookmark end node.
            row1 = bookmark.bookmark_start.get_ancestor(aw.NodeType.ROW)
            row2 = bookmark.bookmark_end.get_ancestor(aw.NodeType.ROW)

            # If both rows are found okay, and the bookmark start and end are contained in adjacent rows,
            # move the bookmark end node to the end of the last paragraph in the top row's last cell.
            if row1 is not None and row2 is not None and row1.next_sibling == row2:
                row1.as_row().last_cell.last_paragraph.append_child(bookmark.bookmark_end)


    @staticmethod
    def delete_row_by_bookmark(doc: aw.Document, bookmark_name: str):

        bookmark = doc.range.bookmarks.get_by_name(bookmark_name)

        if bookmark is not None:
            row = bookmark.bookmark_start.get_ancestor(aw.NodeType.ROW)
            if row is not None:
                row.remove()


if __name__ == '__main__':
    unittest.main()
