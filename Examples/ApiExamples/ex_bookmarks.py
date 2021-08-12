import unittest

import api_example_base as aeb
from document_helper import DocumentHelper

import aspose.words as aw

class ExBookmarks(aeb.ApiExampleBase):
    
    def test_insert(self) :
        
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

        doc.save(aeb.artifacts_dir + "Bookmarks.insert.docx")
        #ExEnd

        doc = aw.Document(aeb.artifacts_dir + "Bookmarks.insert.docx")

        self.assertEqual("My Bookmark", doc.range.bookmarks[0].name)
        

    #ExStart
    #ExFor:Bookmark
    #ExFor:Bookmark.name
    #ExFor:Bookmark.text
    #ExFor:Bookmark.bookmark_start
    #ExFor:Bookmark.bookmark_end
    #ExFor:BookmarkStart
    #ExFor:BookmarkStart.#ctor
    #ExFor:BookmarkEnd
    #ExFor:BookmarkEnd.#ctor
    #ExFor:BookmarkStart.accept(DocumentVisitor)
    #ExFor:BookmarkEnd.accept(DocumentVisitor)
    #ExFor:BookmarkStart.bookmark
    #ExFor:BookmarkStart.get_text
    #ExFor:BookmarkStart.name
    #ExFor:BookmarkEnd.name
    #ExFor:BookmarkCollection
    #ExFor:BookmarkCollection.item(Int32)
    #ExFor:BookmarkCollection.item(String)
    #ExFor:BookmarkCollection.get_enumerator
    #ExFor:Range.bookmarks
    #ExFor:DocumentVisitor.visit_bookmark_start 
    #ExFor:DocumentVisitor.visit_bookmark_end
    #ExSummary:Shows how to add bookmarks and update their contents.
    def test_create_update_and_print_bookmarks(self) :
        
        # Create a document with three bookmarks, then use a custom document visitor implementation to print their contents.
        doc = self.create_document_with_bookmarks(3)
        bookmarks = doc.range.bookmarks
        self.assertEqual(3, bookmarks.count) #ExSkip

        self.print_all_bookmark_info(bookmarks)
            
        # Bookmarks can be accessed in the bookmark collection by index or name, and their names can be updated.
        bookmarks[0].name = f"{bookmarks[0].name}___new_name"
        bookmarks[1].text = f"Updated text contents of {bookmarks[1].name}"

        # Print all bookmarks again to see updated values.
        self.print_all_bookmark_info(bookmarks)
        

    # 
    # Create a document with a given number of bookmarks.
    # 
    @staticmethod
    def create_document_with_bookmarks(numberOfBookmarks) :
        
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        for i in range(1,numberOfBookmarks+1) :
            
            bookmarkName = f"MyBookmark_{i}"

            builder.write("Text before bookmark.")
            builder.start_bookmark(bookmarkName)
            builder.write(f"Text inside {bookmarkName}.")
            builder.end_bookmark(bookmarkName)
            builder.writeln("Text after bookmark.")
            

        return doc
        

    # 
    # Use an iterator and a visitor to print info of every bookmark in the collection.
    #
    @staticmethod
    def print_all_bookmark_info(bookmarks) :
        
        for bookmark in bookmarks :
            print(f"Bookmark name : {bookmark.name}, Contents: {bookmark.text}")
                   
               

    def test_table_column_bookmarks(self) :
        
        #ExStart
        #ExFor:Bookmark.is_column
        #ExFor:Bookmark.first_column
        #ExFor:Bookmark.last_column
        #ExSummary:Shows how to get information about table column bookmarks.
        doc =aw.Document(aeb.my_dir + "Table column bookmarks.doc")

        for bookmark in doc.range.bookmarks :
            
            # If a bookmark encloses columns of a table, it is a table column bookmark, and its IsColumn flag set to true.
            print(f"Bookmark: {bookmark.name}" + "(Column)" if bookmark.is_column else "")
#            if bookmark.is_column :                                                                                    #ExSkip 
#                                                                                                                       #ExSkip
#                if bookmark.bookmark_start.get_ancestor(aw.NodeType.ROW) is Row row &&                                 #ExSkip
#                    bookmark.first_column < row.cells.count)                                                           #ExSkip            
#                                                                                                                       #ExSkip 
#                    # Print the contents of the first and last columns enclosed by the bookmark.                       #ExSkip
#                    Console.write_line(row.cells[bookmark.first_column].get_text().trim_end(ControlChar.cell_char))    #ExSkip
#                    Console.write_line(row.cells[bookmark.last_column].get_text().trim_end(ControlChar.cell_char))     #ExSkip
                    
                
            
        #ExEnd

        doc = DocumentHelper.save_open(doc)

        # Currently bookmarks can be accessed only by index.
        firstTableColumnBookmark = doc.range.bookmarks[0] #["FirstTableColumnBookmark"] 
        secondTableColumnBookmark = doc.range.bookmarks[1] #["SecondTableColumnBookmark"]

        self.assertTrue(firstTableColumnBookmark.is_column)
        self.assertEqual(1, firstTableColumnBookmark.first_column)
        self.assertEqual(3, firstTableColumnBookmark.last_column)

        self.assertTrue(secondTableColumnBookmark.is_column)
        self.assertEqual(0, secondTableColumnBookmark.first_column)
        self.assertEqual(3, secondTableColumnBookmark.last_column)
        

    def test_remove(self) :
        
        #ExStart
        #ExFor:BookmarkCollection.clear
        #ExFor:BookmarkCollection.count
        #ExFor:BookmarkCollection.remove(Bookmark)
        #ExFor:BookmarkCollection.remove(String)
        #ExFor:BookmarkCollection.remove_at
        #ExFor:Bookmark.remove
        #ExSummary:Shows how to remove bookmarks from a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert five bookmarks with text inside their boundaries.
        for i in range(1,6) :
            
            bookmarkName = f"MyBookmark_{i}"

            builder.start_bookmark(bookmarkName)
            builder.write(f"Text inside {bookmarkName}.")
            builder.end_bookmark(bookmarkName)
            builder.insert_break(aw.BreakType.PARAGRAPH_BREAK)
            

        # This collection stores bookmarks.
        bookmarks = doc.range.bookmarks

        self.assertEqual(5, bookmarks.count)

        # There are several ways of removing bookmarks.
        # 1 -  Calling the bookmark's Remove method:
        bookmarks[0].remove() # MyBookmark_1

        self.assertFalse(self.has_bookmark("MyBookmark_1", bookmarks))

        # 2 -  Passing the bookmark to the collection's Remove method:
        bookmark = doc.range.bookmarks[0]
        doc.range.bookmarks.remove(bookmark)

        self.assertFalse(self.has_bookmark("MyBookmark_2", bookmarks))
            
        # 3 -  Removing a bookmark from the collection by name:
        doc.range.bookmarks.remove("MyBookmark_3")

        self.assertFalse(self.has_bookmark("MyBookmark_3", bookmarks))

        # 4 -  Removing a bookmark at an index in the bookmark collection:
        doc.range.bookmarks.remove_at(0)

        self.assertFalse(self.has_bookmark("MyBookmark_4", bookmarks))

        # We can clear the entire bookmark collection.
        bookmarks.clear()

        # The text that was inside the bookmarks is still present in the document.
        self.assertEqual(0, bookmarks.count)
        self.assertEqual("Text inside MyBookmark_1.\r" +
                        "Text inside MyBookmark_2.\r" +
                        "Text inside MyBookmark_3.\r" +
                        "Text inside MyBookmark_4.\r" +
                        "Text inside MyBookmark_5.", doc.get_text().strip())
        #ExEnd
    
    @staticmethod
    def has_bookmark(bookmarkName: str, bookmarks) :
     
        for b in bookmarks :
            if b.name == bookmarkName :
                return True
        
        return False
        
if __name__ == '__main__':
    unittest.main()    
