import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class ExBookmarksOutlineLevelCollection(ApiExampleBase):

    def test_bookmark_levels(self):

        #ExStart
        #ExFor:BookmarksOutlineLevelCollection
        #ExFor:BookmarksOutlineLevelCollection.add(str,int)
        #ExFor:BookmarksOutlineLevelCollection.clear
        #ExFor:BookmarksOutlineLevelCollection.contains(str)
        #ExFor:BookmarksOutlineLevelCollection.count
        #ExFor:BookmarksOutlineLevelCollection.index_of_key(str)
        #ExFor:BookmarksOutlineLevelCollection.__getitem__(int)
        #ExFor:BookmarksOutlineLevelCollection.__getitem__(str)
        #ExFor:BookmarksOutlineLevelCollection.remove(str)
        #ExFor:BookmarksOutlineLevelCollection.remove_at(int)
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
        pdf_save_options = aw.saving.PdfSaveOptions()
        outline_levels = pdf_save_options.outline_options.bookmarks_outline_levels

        outline_levels.add("Bookmark 1", 1)
        outline_levels.add("Bookmark 2", 2)
        outline_levels.add("Bookmark 3", 3)

        self.assertEqual(3, outline_levels.count)
        self.assertTrue(outline_levels.contains("Bookmark 1"))
        self.assertEqual(1, outline_levels[0])
        self.assertEqual(2, outline_levels.get_by_name("Bookmark 2"))
        self.assertEqual(2, outline_levels.index_of_key("Bookmark 3"))

        # We can remove two elements so that only the outline level designation for "Bookmark 1" is left.
        outline_levels.remove_at(2)
        outline_levels.remove("Bookmark 2")

        # There are nine outline levels. Their numbering will be optimized during the save operation.
        # In this case, levels "5" and "9" will become "2" and "3".
        outline_levels.add("Bookmark 2", 5)
        outline_levels.add("Bookmark 3", 9)

        doc.save(ARTIFACTS_DIR + "BookmarksOutlineLevelCollection.bookmark_levels.pdf", pdf_save_options)

        # Emptying this collection will preserve the bookmarks and put them all on the same outline level.
        outline_levels.clear()
        #ExEnd

        #bookmark_editor = aspose.pdf.facades.PdfBookmarkEditor()
        #bookmark_editor.bind_pdf(ARTIFACTS_DIR + "BookmarksOutlineLevelCollection.bookmark_levels.pdf")

        #bookmarks = bookmark_editor.extract_bookmarks()

        #self.assertEqual(3, bookmarks.count)
        #self.assertEqual("Bookmark 1", bookmarks[0].title)
        #self.assertEqual("Bookmark 2", bookmarks[1].title)
        #self.assertEqual("Bookmark 3", bookmarks[2].title)

