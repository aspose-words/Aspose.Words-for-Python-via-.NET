import unittest

import api_example_base as aeb
from document_helper import DocumentHelper

import aspose.words as aw

class ExBookmarksOutlineLevelCollection(aeb.ApiExampleBase):
    
    def test_bookmark_levels(self) :
        
        #ExStart
        #ExFor:BookmarksOutlineLevelCollection
        #ExFor:BookmarksOutlineLevelCollection.add(String, Int32)
        #ExFor:BookmarksOutlineLevelCollection.clear
        #ExFor:BookmarksOutlineLevelCollection.contains(System.string)
        #ExFor:BookmarksOutlineLevelCollection.count
        #ExFor:BookmarksOutlineLevelCollection.index_of_key(System.string)
        #ExFor:BookmarksOutlineLevelCollection.item(System.int_32)
        #ExFor:BookmarksOutlineLevelCollection.item(System.string)
        #ExFor:BookmarksOutlineLevelCollection.remove(System.string)
        #ExFor:BookmarksOutlineLevelCollection.remove_at(System.int_32)
        #ExFor:OutlineOptions.bookmarks_outline_levels
        #ExSummary:Shows how to set outline levels for bookmarks.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a bookmark with another bookmark nested inside it.
        builder.start_bookmark("Bookmark 1")
        builder.writeln("Text inside Bookmark 1.")

        builder.start_bookmark("Bookmark 2")
        builder.writeln("Text inside Bookmark 1 and 2.")
        builder.end_bookmark("Bookmark 2")

        builder.writeln("Text inside Bookmark 1.")
        builder.end_bookmark("Bookmark 1")

        # Insert another bookmark.
        builder.start_bookmark("Bookmark 3")
        builder.writeln("Text inside Bookmark 3.")
        builder.end_bookmark("Bookmark 3")

        # When saving to .pdf, bookmarks can be accessed via a drop-down menu and used as anchors by most readers.
        # Bookmarks can also have numeric values for outline levels,
        # enabling lower level outline entries to hide higher-level child entries when collapsed in the reader.
        pdfSaveOptions = aw.saving.PdfSaveOptions()
        outlineLevels = pdfSaveOptions.outline_options.bookmarks_outline_levels

        outlineLevels.add("Bookmark 1", 1)
        outlineLevels.add("Bookmark 2", 2)
        outlineLevels.add("Bookmark 3", 3)

        self.assertEqual(3, outlineLevels.count)
        self.assertTrue(outlineLevels.contains("Bookmark 1"))
        self.assertEqual(1, outlineLevels[0])
        self.assertEqual(2, outlineLevels[outlineLevels.index_of_key("Bookmark 2")])
        self.assertEqual(2, outlineLevels.index_of_key("Bookmark 3"))

        # We can remove two elements so that only the outline level designation for "Bookmark 1" is left.
        outlineLevels.remove_at(2)
        outlineLevels.remove("Bookmark 2")

        # There are nine outline levels. Their numbering will be optimized during the save operation.
        # In this case, levels "5" and "9" will become "2" and "3".
        outlineLevels.add("Bookmark 2", 5)
        outlineLevels.add("Bookmark 3", 9)

        doc.save(aeb.artifacts_dir + "BookmarksOutlineLevelCollection.bookmark_levels.pdf", pdfSaveOptions)

        # Emptying this collection will preserve the bookmarks and put them all on the same outline level.
        outlineLevels.clear()
        #ExEnd
        
    
if __name__ == '__main__':
    unittest.main()    
