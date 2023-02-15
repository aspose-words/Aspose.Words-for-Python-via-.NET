# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR
from document_helper import DocumentHelper

class ExBookmarks(ApiExampleBase):

    def test_insert(self):

        #ExStart
        #ExFor:Bookmark.name
        #ExSummary:Shows how to insert a bookmark.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # A valid bookmark has a name, a BookmarkStart, and a BookmarkEnd node.
        # Any whitespace in the names of bookmarks will be converted to underscores if we open the saved document with Microsoft Word.
        # If we highlight the bookmark's name in Microsoft Word via Insert -> Links -> Bookmark, and press "Go To",
        # the cursor will jump to the text enclosed between the BookmarkStart and BookmarkEnd nodes.
        builder.start_bookmark("My Bookmark")
        builder.write("Contents of MyBookmark.")
        builder.end_bookmark("My Bookmark")

        # Bookmarks are stored in this collection.
        self.assertEqual("My Bookmark", doc.range.bookmarks[0].name)

        doc.save(ARTIFACTS_DIR + "Bookmarks.insert.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Bookmarks.insert.docx")

        self.assertEqual("My Bookmark", doc.range.bookmarks[0].name)

    ##ExStart
    ##ExFor:Bookmark
    ##ExFor:Bookmark.name
    ##ExFor:Bookmark.text
    ##ExFor:Bookmark.bookmark_start
    ##ExFor:Bookmark.bookmark_end
    ##ExFor:BookmarkStart
    ##ExFor:BookmarkStart.__init__
    ##ExFor:BookmarkEnd
    ##ExFor:BookmarkEnd.__init__
    ##ExFor:BookmarkStart.accept(DocumentVisitor)
    ##ExFor:BookmarkEnd.accept(DocumentVisitor)
    ##ExFor:BookmarkStart.bookmark
    ##ExFor:BookmarkStart.get_text
    ##ExFor:BookmarkStart.name
    ##ExFor:BookmarkEnd.name
    ##ExFor:BookmarkCollection
    ##ExFor:BookmarkCollection.__getitem__(int)
    ##ExFor:BookmarkCollection.__getitem__(str)
    ##ExFor:BookmarkCollection.__iter__
    ##ExFor:Range.bookmarks
    ##ExFor:DocumentVisitor.visit_bookmark_start
    ##ExFor:DocumentVisitor.visit_bookmark_end
    ##ExSummary:Shows how to add bookmarks and update their contents.
    #def test_create_update_and_print_bookmarks(self):

    #    # Create a document with three bookmarks, then use a custom document visitor implementation to print their contents.
    #    doc = ExBookmarks.create_document_with_bookmarks(3)
    #    bookmarks = doc.range.bookmarks
    #    self.assertEqual(3, bookmarks.count) #ExSkip

    #    ExBookmarks.print_all_bookmark_info(bookmarks)

    #    # Bookmarks can be accessed in the bookmark collection by index or name, and their names can be updated.
    #    bookmarks[0].name = f"{bookmarks[0].name}_new_name"
    #    bookmarks.get_by_name("MyBookmark_2").text = f"Updated text contents of {bookmarks[1].name}"

    #    # Print all bookmarks again to see updated values.
    #    ExBookmarks.print_all_bookmark_info(bookmarks)

    #@staticmethod
    #def create_document_with_bookmarks(number_of_bookmarks: int) -> aw.Document:
    #    """Create a document with a given number of bookmarks."""

    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    for i in range(1, number_of_bookmarks + 1):
    #        bookmark_name = f"MyBookmark_{i}"

    #        builder.write("Text before bookmark.")
    #        builder.start_bookmark(bookmark_name)
    #        builder.write(f"Text inside {bookmark_name}.")
    #        builder.end_bookmark(bookmark_name)
    #        builder.writeln("Text after bookmark.")

    #    return doc

    #@staticmethod
    #def print_all_bookmark_info(bookmarks: aw.BookmarkCollection):
    #    """Use an iterator and a visitor to print info of every bookmark in the collection."""

    #    bookmark_visitor = ExBookmarks.BookmarkInfoPrinter()

    #    # Get each bookmark in the collection to accept a visitor that will print its contents.
    #    for bookmark in bookmarks:
    #        bookmark.bookmark_start.accept(bookmark_visitor)
    #        bookmark.bookmark_end.accept(bookmark_visitor)

    #        print(bookmark.bookmark_start.get_text())

    #class BookmarkInfoPrinter(aw.DocumentVisitor):
    #    """Prints contents of every visited bookmark to the console."""

    #    def visit_bookmark_start(self, bookmark_start: aw.BookmarkStart) -> aw.VisitorAction:

    #        print(f"BookmarkStart name: \"{bookmark_start.Name}\", Contents: \"{bookmark_start.bookmark.text}\"")
    #        return aw.VisitorAction.CONTINUE

    #    def VisitBookmarkEnd(self, bookmark_end: aw.BookmarkEnd) -> aw.VisitorAction:

    #        print(f"BookmarkEnd name: \"{bookmark_end.Name}\"")
    #        return aw.VisitorAction.CONTINUE

    ##ExEnd

    def test_table_column_bookmarks(self):

        #ExStart
        #ExFor:Bookmark.is_column
        #ExFor:Bookmark.first_column
        #ExFor:Bookmark.last_column
        #ExSummary:Shows how to get information about table column bookmarks.
        doc = aw.Document(MY_DIR + "Table column bookmarks.doc")

        for bookmark in doc.range.bookmarks:

            # If a bookmark encloses columns of a table, it is a table column bookmark, and its "is_column" flag set to True.
            print(f"Bookmark: {bookmark.name}{' (Column)' if bookmark.is_column else ''}")
            if bookmark.is_column:
                row = bookmark.bookmark_start.get_ancestor(aw.NodeType.ROW)
                if row is aw.tables.Row and bookmark.first_column < row.cells.count:
                    # Print the contents of the first and last columns enclosed by the bookmark.
                    print(row.cells[bookmark.first_column].get_text().rstrip(aw.ControlChar.CELL_CHAR))
                    print(row.cells[bookmark.last_column].get_text().rstrip(aw.ControlChar.CELL_CHAR))

        #ExEnd

        doc = DocumentHelper.save_open(doc)

        first_table_column_bookmark = doc.range.bookmarks.get_by_name("FirstTableColumnBookmark")
        second_table_column_bookmark = doc.range.bookmarks.get_by_name("SecondTableColumnBookmark")

        self.assertTrue(first_table_column_bookmark.is_column)
        self.assertEqual(1, first_table_column_bookmark.first_column)
        self.assertEqual(3, first_table_column_bookmark.last_column)

        self.assertTrue(second_table_column_bookmark.is_column)
        self.assertEqual(0, second_table_column_bookmark.first_column)
        self.assertEqual(3, second_table_column_bookmark.last_column)

    def test_remove(self):

        #ExStart
        #ExFor:BookmarkCollection.clear
        #ExFor:BookmarkCollection.count
        #ExFor:BookmarkCollection.remove(Bookmark)
        #ExFor:BookmarkCollection.remove(str)
        #ExFor:BookmarkCollection.remove_at
        #ExFor:Bookmark.remove
        #ExSummary:Shows how to remove bookmarks from a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert five bookmarks with text inside their boundaries.
        for i in range(1, 6):
            bookmark_name = f"MyBookmark_{i}"

            builder.start_bookmark(bookmark_name)
            builder.write(f"Text inside {bookmark_name}.")
            builder.end_bookmark(bookmark_name)
            builder.insert_break(aw.BreakType.PARAGRAPH_BREAK)

        # This collection stores bookmarks.
        bookmarks = doc.range.bookmarks

        self.assertEqual(5, bookmarks.count)

        # There are several ways of removing bookmarks.
        # 1 -  Calling the bookmark's "remove" method:
        bookmarks.get_by_name("MyBookmark_1").remove()

        self.assertFalse(any(b for b in bookmarks if b.name == "MyBookmark_1"))

        # 2 -  Passing the bookmark to the collection's "remove" method:
        bookmark = doc.range.bookmarks[0]
        doc.range.bookmarks.remove(bookmark)

        self.assertFalse(any(b for b in bookmarks if b.name == "MyBookmark_2"))

        # 3 -  Removing a bookmark from the collection by name:
        doc.range.bookmarks.remove("MyBookmark_3")

        self.assertFalse(any(b for b in bookmarks if b.name == "MyBookmark_3"))

        # 4 -  Removing a bookmark at an index in the bookmark collection:
        doc.range.bookmarks.remove_at(0)

        self.assertFalse(any(b for b in bookmarks if b.name == "MyBookmark_4"))

        # We can clear the entire bookmark collection.
        bookmarks.clear()

        # The text that was inside the bookmarks is still present in the document.
        self.assertListEqual([], list(bookmarks))
        self.assertEqual("Text inside MyBookmark_1.\r" +
                         "Text inside MyBookmark_2.\r" +
                         "Text inside MyBookmark_3.\r" +
                         "Text inside MyBookmark_4.\r" +
                         "Text inside MyBookmark_5.", doc.get_text().strip())
        #ExEnd
