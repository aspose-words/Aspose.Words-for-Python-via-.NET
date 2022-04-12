# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR
from document_helper import DocumentHelper

class ExStyles(ApiExampleBase):

    def test_styles(self):

        #ExStart
        #ExFor:DocumentBase.styles
        #ExFor:Style.document
        #ExFor:Style.name
        #ExFor:Style.is_heading
        #ExFor:Style.is_quick_style
        #ExFor:Style.next_paragraph_style_name
        #ExFor:Style.styles
        #ExFor:Style.type
        #ExFor:StyleCollection.document
        #ExFor:StyleCollection.__iter__
        #ExSummary:Shows how to access a document's style collection.
        doc = aw.Document()

        self.assertEqual(4, doc.styles.count)

        # Enumerate and list all the styles that a document created using Aspose.Words contains by default.
        for cur_style in doc.styles:
            print(f"Style name:\t\"{cur_style.name}\", of type \"{cur_style.type}\"")
            print(f"\tSubsequent style:\t{cur_style.next_paragraph_style_name}")
            print(f"\tIs heading:\t\t\t{cur_style.is_heading}")
            print(f"\tIs QuickStyle:\t\t{cur_style.is_quick_style}")

            self.assertEqual(doc, cur_style.document)

        #ExEnd

    def test_create_style(self):

        #ExStart
        #ExFor:Style.font
        #ExFor:Style
        #ExFor:Style.remove
        #ExSummary:Shows how to create and apply a custom style.
        doc = aw.Document()

        style = doc.styles.add(aw.StyleType.PARAGRAPH, "MyStyle")
        style.font.name = "Times New Roman"
        style.font.size = 16
        style.font.color = drawing.Color.navy

        builder = aw.DocumentBuilder(doc)

        # Apply one of the styles from the document to the paragraph that the document builder is creating.
        builder.paragraph_format.style = doc.styles.get_by_name("MyStyle")
        builder.writeln("Hello world!")

        first_paragraph_style = doc.first_section.body.first_paragraph.paragraph_format.style

        self.assertEqual(style, first_paragraph_style)

        # Remove our custom style from the document's styles collection.
        doc.styles.get_by_name("MyStyle").remove()

        first_paragraph_style = doc.first_section.body.first_paragraph.paragraph_format.style

        # Any text that used a removed style reverts to the default formatting.
        self.assertFalse(any(s.name == "MyStyle" for s in doc.styles))
        self.assertEqual("Times New Roman", first_paragraph_style.font.name)
        self.assertEqual(12.0, first_paragraph_style.font.size)
        self.assertEqual(drawing.Color.empty().to_argb(), first_paragraph_style.font.color.to_argb())
        #ExEnd

    def test_style_collection(self):

        #ExStart
        #ExFor:StyleCollection.add(StyleType,str)
        #ExFor:StyleCollection.count
        #ExFor:StyleCollection.default_font
        #ExFor:StyleCollection.default_paragraph_format
        #ExFor:StyleCollection.__getitem__(StyleIdentifier)
        #ExFor:StyleCollection.__getitem__(int)
        #ExSummary:Shows how to add a Style to a document's styles collection.
        doc = aw.Document()
        styles = doc.styles

        # Set default parameters for new styles that we may later add to this collection.
        styles.default_font.name = "Courier New"

        # If we add a style of the "StyleType.PARAGRAPH", the collection will apply the values of
        # its "default_paragraph_format" property to the style's "paragraph_format" property.
        styles.default_paragraph_format.first_line_indent = 15.0

        # Add a style, and then verify that it has the default settings.
        styles.add(aw.StyleType.PARAGRAPH, "MyStyle")

        self.assertEqual("Courier New", styles[4].font.name)
        self.assertEqual(15.0, styles.get_by_name("MyStyle").paragraph_format.first_line_indent)
        #ExEnd

    def test_remove_styles_from_style_gallery(self):

        #ExStart
        #ExFor:StyleCollection.clear_quick_style_gallery
        #ExSummary:Shows how to remove styles from Style Gallery panel.
        doc = aw.Document()

        # Note that remove styles work only with DOCX format for now.
        doc.styles.clear_quick_style_gallery()

        doc.save(ARTIFACTS_DIR + "Styles.remove_styles_from_style_gallery.docx")
        #ExEnd

    def test_change_tocs_tab_stops(self):

        #ExStart
        #ExFor:TabStop
        #ExFor:ParagraphFormat.tab_stops
        #ExFor:Style.style_identifier
        #ExFor:TabStopCollection.remove_by_position
        #ExFor:TabStop.alignment
        #ExFor:TabStop.position
        #ExFor:TabStop.leader
        #ExSummary:Shows how to modify the position of the right tab stop in TOC related paragraphs.
        doc = aw.Document(MY_DIR + "Table of contents.docx")

        # Iterate through all paragraphs with TOC result-based styles; this is any style between TOC and TOC9.
        for para in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True):
            para = para.as_paragraph()
            if (para.paragraph_format.style.style_identifier >= aw.StyleIdentifier.TOC1 and
                para.paragraph_format.style.style_identifier <= aw.StyleIdentifier.TOC9):

                # Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
                tab = para.paragraph_format.tab_stops[0]

                # Replace the first default tab, stop with a custom tab stop.
                para.paragraph_format.tab_stops.remove_by_position(tab.position)
                para.paragraph_format.tab_stops.add(tab.position - 50, tab.alignment, tab.leader)

        doc.save(ARTIFACTS_DIR + "Styles.change_tocs_tab_stops.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Styles.change_tocs_tab_stops.docx")

        for para in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True):
            para = para.as_paragraph()
            if (para.paragraph_format.style.style_identifier >= aw.StyleIdentifier.TOC1 and
                para.paragraph_format.style.style_identifier <= aw.StyleIdentifier.TOC9):

                tab_stop = para.get_effective_tab_stops()[0]
                self.assertEqual(400.8, tab_stop.position)
                self.assertEqual(aw.TabAlignment.RIGHT, tab_stop.alignment)
                self.assertEqual(aw.TabLeader.DOTS, tab_stop.leader)

    def test_copy_style_same_document(self):

        #ExStart
        #ExFor:StyleCollection.add_copy
        #ExFor:Style.name
        #ExSummary:Shows how to clone a document's style.
        doc = aw.Document()

        # The add_copy method creates a copy of the specified style and
        # automatically generates a new name for the style, such as "Heading 1_0".
        new_style = doc.styles.add_copy(doc.styles.get_by_name("Heading 1"))

        # Use the style's "name" property to change the style's identifying name.
        new_style.name = "My Heading 1"

        # Our document now has two identical looking styles with different names.
        # Changing settings of one of the styles do not affect the other.
        new_style.font.color = drawing.Color.red

        self.assertEqual("My Heading 1", new_style.name)
        self.assertEqual("Heading 1", doc.styles.get_by_name("Heading 1").name)

        self.assertEqual(doc.styles.get_by_name("Heading 1").type, new_style.type)
        self.assertEqual(doc.styles.get_by_name("Heading 1").font.name, new_style.font.name)
        self.assertEqual(doc.styles.get_by_name("Heading 1").font.size, new_style.font.size)
        self.assertNotEqual(doc.styles.get_by_name("Heading 1").font.color, new_style.font.color)
        #ExEnd

    def test_copy_style_different_document(self):

        #ExStart
        #ExFor:StyleCollection.add_copy
        #ExSummary:Shows how to import a style from one document into a different document.
        src_doc = aw.Document()

        # Create a custom style for the source document.
        src_style = src_doc.styles.add(aw.StyleType.PARAGRAPH, "MyStyle")
        src_style.font.color = drawing.Color.red

        # Import the source document's custom style into the destination document.
        dst_doc = aw.Document()
        new_style = dst_doc.styles.add_copy(src_style)

        # The imported style has an appearance identical to its source style.
        self.assertEqual("MyStyle", new_style.name)
        self.assertEqual(drawing.Color.red.to_argb(), new_style.font.color.to_argb())
        #ExEnd

    def test_default_styles(self):

        doc = aw.Document()

        doc.styles.default_font.name = "PMingLiU"
        doc.styles.default_font.bold = True

        doc.styles.default_paragraph_format.space_after = 20
        doc.styles.default_paragraph_format.alignment = aw.ParagraphAlignment.RIGHT

        doc = DocumentHelper.save_open(doc)

        self.assertTrue(doc.styles.default_font.bold)
        self.assertEqual("PMingLiU", doc.styles.default_font.name)
        self.assertEqual(20, doc.styles.default_paragraph_format.space_after)
        self.assertEqual(aw.ParagraphAlignment.RIGHT, doc.styles.default_paragraph_format.alignment)

    def test_paragraph_style_bulleted_list(self):

        #ExStart
        #ExFor:StyleCollection
        #ExFor:DocumentBase.styles
        #ExFor:Style
        #ExFor:Font
        #ExFor:Style.font
        #ExFor:Style.paragraph_format
        #ExFor:Style.list_format
        #ExFor:ParagraphFormat.style
        #ExSummary:Shows how to create and use a paragraph style with list formatting.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create a custom paragraph style.
        style = doc.styles.add(aw.StyleType.PARAGRAPH, "MyStyle1")
        style.font.size = 24
        style.font.name = "Verdana"
        style.paragraph_format.space_after = 12

        # Create a list and make sure the paragraphs that use this style will use this list.
        style.list_format.list = doc.lists.add(aw.lists.ListTemplate.BULLET_DEFAULT)
        style.list_format.list_level_number = 0

        # Apply the paragraph style to the document builder's current paragraph, and then add some text.
        builder.paragraph_format.style = style
        builder.writeln("Hello World: MyStyle1, bulleted list.")

        # Change the document builder's style to one that has no list formatting and write another paragraph.
        builder.paragraph_format.style = doc.styles.get_by_name("Normal")
        builder.writeln("Hello World: Normal.")

        builder.document.save(ARTIFACTS_DIR + "Styles.paragraph_style_bulleted_list.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Styles.paragraph_style_bulleted_list.docx")

        style = doc.styles.get_by_name("MyStyle1")

        self.assertEqual("MyStyle1", style.name)
        self.assertEqual(24, style.font.size)
        self.assertEqual("Verdana", style.font.name)
        self.assertEqual(12.0, style.paragraph_format.space_after)

    def test_style_aliases(self):

        #ExStart
        #ExFor:Style.aliases
        #ExFor:Style.base_style_name
        #ExFor:Style.__eq__(Style)
        #ExFor:Style.linked_style_name
        #ExSummary:Shows how to use style aliases.
        doc = aw.Document(MY_DIR + "Style with alias.docx")

        # This document contains a style named "MyStyle,MyStyle Alias 1,MyStyle Alias 2".
        # If a style's name has multiple values separated by commas, each clause is a separate alias.
        style = doc.styles.get_by_name("MyStyle")
        self.assertSequenceEqual(["MyStyle Alias 1", "MyStyle Alias 2"], style.aliases)
        self.assertEqual("Title", style.base_style_name)
        self.assertEqual("MyStyle Char", style.linked_style_name)

        # We can reference a style using its alias, as well as its name.
        self.assertEqual(doc.styles.get_by_name("MyStyle Alias 1"), doc.styles.get_by_name("MyStyle Alias 2"))

        builder = aw.DocumentBuilder(doc)
        builder.move_to_document_end()
        builder.paragraph_format.style = doc.styles.get_by_name("MyStyle Alias 1")
        builder.writeln("Hello world!")
        builder.paragraph_format.style = doc.styles.get_by_name("MyStyle Alias 2")
        builder.write("Hello again!")

        self.assertEqual(
            doc.first_section.body.paragraphs[0].paragraph_format.style,
            doc.first_section.body.paragraphs[1].paragraph_format.style)
        #ExEnd
