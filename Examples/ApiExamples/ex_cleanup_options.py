# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase

class ExCleanupOptions(ApiExampleBase):

    def test_remove_unused_resources(self):

        #ExStart
        #ExFor:Document.cleanup(CleanupOptions)
        #ExFor:CleanupOptions
        #ExFor:CleanupOptions.unused_lists
        #ExFor:CleanupOptions.unused_styles
        #ExFor:CleanupOptions.unused_builtin_styles
        #ExSummary:Shows how to remove all unused custom styles from a document.
        doc = aw.Document()

        doc.styles.add(aw.StyleType.LIST, "MyListStyle1")
        doc.styles.add(aw.StyleType.LIST, "MyListStyle2")
        doc.styles.add(aw.StyleType.CHARACTER, "MyParagraphStyle1")
        doc.styles.add(aw.StyleType.CHARACTER, "MyParagraphStyle2")

        # Combined with the built-in styles, the document now has eight styles.
        # A custom style is marked as "used" while there is any text within the document
        # formatted in that style. This means that the 4 styles we added are currently unused.
        self.assertEqual(8, doc.styles.count)

        # Apply a custom character style, and then a custom list style. Doing so will mark them as "used".
        builder = aw.DocumentBuilder(doc)
        builder.font.style = doc.styles.get_by_name("MyParagraphStyle1")
        builder.writeln("Hello world!")

        builder.list_format.list = doc.lists.add(doc.styles.get_by_name("MyListStyle1"))
        builder.writeln("Item 1")
        builder.writeln("Item 2")

        # Now, there is one unused character style and one unused list style.
        # The "cleanup" method, when configured with a CleanupOptions object, can target unused styles and remove them.
        cleanup_options = aw.CleanupOptions()
        cleanup_options.unused_lists = True
        cleanup_options.unused_styles = True
        cleanup_options.unused_builtin_styles = True

        doc.cleanup(cleanup_options)

        self.assertEqual(4, doc.styles.count)

        # Removing every node that a custom style is applied to marks it as "unused" again.
        # Rerun the "cleanup" method to remove them.
        doc.first_section.body.remove_all_children()
        doc.cleanup(cleanup_options)

        self.assertEqual(2, doc.styles.count)
        #ExEnd

    def test_remove_duplicate_styles(self):

        #ExStart
        #ExFor:CleanupOptions.duplicate_style
        #ExSummary:Shows how to remove duplicated styles from the document.
        doc = aw.Document()

        # Add two styles to the document with identical properties,
        # but different names. The second style is considered a duplicate of the first.
        my_style = doc.styles.add(aw.StyleType.PARAGRAPH, "MyStyle1")
        my_style.font.size = 14
        my_style.font.name = "Courier New"
        my_style.font.color = drawing.Color.blue

        duplicate_style = doc.styles.add(aw.StyleType.PARAGRAPH, "MyStyle2")
        duplicate_style.font.size = 14
        duplicate_style.font.name = "Courier New"
        duplicate_style.font.color = drawing.Color.blue

        self.assertEqual(6, doc.styles.count)

        # Apply both styles to different paragraphs within the document.
        builder = aw.DocumentBuilder(doc)
        builder.paragraph_format.style_name = my_style.name
        builder.writeln("Hello world!")

        builder.paragraph_format.style_name = duplicate_style.name
        builder.writeln("Hello again!")

        paragraphs = doc.first_section.body.paragraphs

        self.assertEqual(my_style, paragraphs[0].paragraph_format.style)
        self.assertEqual(duplicate_style, paragraphs[1].paragraph_format.style)

        # Configure a CleanOptions object, then call the "cleanup" method to substitute all duplicate styles
        # with the original and remove the duplicates from the document.
        cleanup_options = aw.CleanupOptions()
        cleanup_options.duplicate_style = True

        doc.cleanup(cleanup_options)

        self.assertEqual(5, doc.styles.count)
        self.assertEqual(my_style, paragraphs[0].paragraph_format.style)
        self.assertEqual(my_style, paragraphs[1].paragraph_format.style)
        #ExEnd
