# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from document_helper import DocumentHelper
import os
import aspose.words as aw
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExBookmarks(ApiExampleBase):

    def test_insert(self):
        #ExStart
        #ExFor:Bookmark.name
        #ExSummary:Shows how to insert a bookmark.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # A valid bookmark has a name, a BookmarkStart, and a BookmarkEnd node.
        # Any whitespace in the names of bookmarks will be converted to underscores if we open the saved document with Microsoft Word.
        # If we highlight the bookmark's name in Microsoft Word via Insert -> Links -> Bookmark, and press "Go To",
        # the cursor will jump to the text enclosed between the BookmarkStart and BookmarkEnd nodes.
        builder.start_bookmark('My Bookmark')
        builder.write('Contents of MyBookmark.')
        builder.end_bookmark('My Bookmark')
        # Bookmarks are stored in this collection.
        self.assertEqual('My Bookmark', doc.range.bookmarks[0].name)
        doc.save(file_name=ARTIFACTS_DIR + 'Bookmarks.Insert.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Bookmarks.Insert.docx')
        self.assertEqual('My Bookmark', doc.range.bookmarks[0].name)

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
        builder = aw.DocumentBuilder(doc=doc)
        # Insert five bookmarks with text inside their boundaries.
        i = 1
        while i <= 5:
            bookmark_name = 'MyBookmark_' + str(i)
            builder.start_bookmark(bookmark_name)
            builder.write(f'Text inside {bookmark_name}.')
            builder.end_bookmark(bookmark_name)
            builder.insert_break(aw.BreakType.PARAGRAPH_BREAK)
            i += 1
        # This collection stores bookmarks.
        bookmarks = doc.range.bookmarks
        self.assertEqual(5, bookmarks.count)
        # There are several ways of removing bookmarks.
        # 1 -  Calling the bookmark's Remove method:
        bookmarks.get_by_name('MyBookmark_1').remove()
        self.assertFalse(any([b.name == 'MyBookmark_1' for b in bookmarks]))
        # 2 -  Passing the bookmark to the collection's Remove method:
        bookmark = doc.range.bookmarks[0]
        doc.range.bookmarks.remove(bookmark=bookmark)
        self.assertFalse(any([b.name == 'MyBookmark_2' for b in bookmarks]))
        # 3 -  Removing a bookmark from the collection by name:
        doc.range.bookmarks.remove(bookmark_name='MyBookmark_3')
        self.assertFalse(any([b.name == 'MyBookmark_3' for b in bookmarks]))
        # 4 -  Removing a bookmark at an index in the bookmark collection:
        doc.range.bookmarks.remove_at(0)
        self.assertFalse(any([b.name == 'MyBookmark_4' for b in bookmarks]))
        # We can clear the entire bookmark collection.
        bookmarks.clear()
        # The text that was inside the bookmarks is still present in the document.
        self.assertEqual(0, bookmarks.count)
        self.assertEqual('Text inside MyBookmark_1.\r' + 'Text inside MyBookmark_2.\r' + 'Text inside MyBookmark_3.\r' + 'Text inside MyBookmark_4.\r' + 'Text inside MyBookmark_5.', doc.get_text().strip())
        #ExEnd

    def test_table_column_bookmarks(self):
        #ExStart
        #ExFor:Bookmark.is_column
        #ExFor:Bookmark.first_column
        #ExFor:Bookmark.last_column
        #ExSummary:Shows how to get information about table column bookmarks.
        doc = aw.Document(MY_DIR + 'Table column bookmarks.doc')
        for bookmark in doc.range.bookmarks:
            # If a bookmark encloses columns of a table, it is a table column bookmark, and its "is_column" flag set to True.
            print(f"Bookmark: {bookmark.name}{(' (Column)' if bookmark.is_column else '')}")
            if bookmark.is_column:
                row = bookmark.bookmark_start.get_ancestor(aw.NodeType.ROW)
                if row is aw.tables.Row and bookmark.first_column < row.cells.count:
                    # Print the contents of the first and last columns enclosed by the bookmark.
                    print(row.cells[bookmark.first_column].get_text().rstrip(aw.ControlChar.CELL_CHAR))
                    print(row.cells[bookmark.last_column].get_text().rstrip(aw.ControlChar.CELL_CHAR))
        #ExEnd
        doc = DocumentHelper.save_open(doc)
        first_table_column_bookmark = doc.range.bookmarks.get_by_name('FirstTableColumnBookmark')
        second_table_column_bookmark = doc.range.bookmarks.get_by_name('SecondTableColumnBookmark')
        self.assertTrue(first_table_column_bookmark.is_column)
        self.assertEqual(1, first_table_column_bookmark.first_column)
        self.assertEqual(3, first_table_column_bookmark.last_column)
        self.assertTrue(second_table_column_bookmark.is_column)
        self.assertEqual(0, second_table_column_bookmark.first_column)
        self.assertEqual(3, second_table_column_bookmark.last_column)