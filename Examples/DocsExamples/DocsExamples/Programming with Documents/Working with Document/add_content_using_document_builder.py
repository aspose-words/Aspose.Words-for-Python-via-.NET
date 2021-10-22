import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw
import aspose.pydrawing as drawing

class AddContentUsingDocumentBuilder(docs_base.DocsExamplesBase):

    def test_document_builder_insert_bookmark(self) :

        #ExStart:DocumentBuilderInsertBookmark
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_bookmark("FineBookmark")
        builder.writeln("This is just a fine bookmark.")
        builder.end_bookmark("FineBookmark")

        doc.save(docs_base.artifacts_dir + "WorkingWithBookmarks.document_builder_insert_bookmark.docx")
        #ExEnd:DocumentBuilderInsertBookmark


    def test_build_table(self) :

        #ExStart:BuildTable
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        table.auto_fit(aw.tables.AutoFitBehavior.FIXED_COLUMN_WIDTHS)

        builder.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.CENTER
        builder.write("This is row 1 cell 1")

        builder.insert_cell()
        builder.write("This is row 1 cell 2")

        builder.end_row()

        builder.insert_cell()

        builder.row_format.height = 100
        builder.row_format.height_rule = aw.HeightRule.EXACTLY
        builder.cell_format.orientation = aw.TextOrientation.UPWARD
        builder.writeln("This is row 2 cell 1")

        builder.insert_cell()
        builder.cell_format.orientation = aw.TextOrientation.DOWNWARD
        builder.writeln("This is row 2 cell 2")

        builder.end_row()
        builder.end_table()

        doc.save(docs_base.artifacts_dir + "AddContentUsingDocumentBuilder.build_table.docx")
        #ExEnd:BuildTable


    def test_insert_horizontal_rule(self) :

        #ExStart:InsertHorizontalRule
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Insert a horizontal rule shape into the document.")
        builder.insert_horizontal_rule()

        doc.save(docs_base.artifacts_dir + "AddContentUsingDocumentBuilder.insert_horizontal_rule.docx")
        #ExEnd:InsertHorizontalRule


    def test_horizontal_rule_format(self) :

        #ExStart:HorizontalRuleFormat
        builder = aw.DocumentBuilder()

        shape = builder.insert_horizontal_rule()

        horizontal_rule_format = shape.horizontal_rule_format
        horizontal_rule_format.alignment = aw.drawing.HorizontalRuleAlignment.CENTER
        horizontal_rule_format.width_percent = 70
        horizontal_rule_format.height = 3
        horizontal_rule_format.color = drawing.Color.blue
        horizontal_rule_format.no_shade = True

        builder.document.save(docs_base.artifacts_dir + "AddContentUsingDocumentBuilder.horizontal_rule_format.docx")
        #ExEnd:HorizontalRuleFormat


    def test_insert_break(self) :

        #ExStart:InsertBreak
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("This is page 1.")
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        builder.writeln("This is page 2.")
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        builder.writeln("This is page 3.")

        doc.save(docs_base.artifacts_dir + "AddContentUsingDocumentBuilder.insert_break.docx")
        #ExEnd:InsertBreak


    def test_insert_text_input_form_field(self) :

        #ExStart:InsertTextInputFormField
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_text_input("TextInput", aw.fields.TextFormFieldType.REGULAR, "", "Hello", 0)

        doc.save(docs_base.artifacts_dir + "AddContentUsingDocumentBuilder.insert_text_input_form_field.docx")
        #ExEnd:InsertTextInputFormField


    def test_insert_check_box_form_field(self) :

        #ExStart:InsertCheckBoxFormField
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_check_box("CheckBox", True, True, 0)

        doc.save(docs_base.artifacts_dir + "AddContentUsingDocumentBuilder.insert_check_box_form_field.docx")
        #ExEnd:InsertCheckBoxFormField


    def test_insert_combo_box_form_field(self) :

        #ExStart:InsertComboBoxFormField
        items =  [ "One", "Two", "Three" ]

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_combo_box("DropDown", items, 0)

        doc.save(docs_base.artifacts_dir + "AddContentUsingDocumentBuilder.insert_combo_box_form_field.docx")
        #ExEnd:InsertComboBoxFormField


    def test_insert_html(self) :

        #ExStart:InsertHtml
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_html(
            "<P align='right'>Paragraph right</P>" +
            "<b>Implicit paragraph left</b>" +
            "<div align='center'>Div center</div>" +
            "<h1 align='left'>Heading 1 left.</h1>")

        doc.save(docs_base.artifacts_dir + "AddContentUsingDocumentBuilder.insert_html.docx")
        #ExEnd:InsertHtml


    def test_insert_hyperlink(self) :

        #ExStart:InsertHyperlink
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Please make sure to visit ")
        builder.font.color = drawing.Color.blue
        builder.font.underline = aw.Underline.SINGLE

        builder.insert_hyperlink("Aspose Website", "http:#www.aspose.com", False)

        builder.font.clear_formatting()
        builder.write(" for more information.")

        doc.save(docs_base.artifacts_dir + "AddContentUsingDocumentBuilder.insert_hyperlink.docx")
        #ExEnd:InsertHyperlink


    def test_insert_table_of_contents(self) :

        #ExStart:InsertTableOfContents
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_table_of_contents("\\o \"1-3\" \\h \\z \\u")

        # Start the actual document content on the second page.
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1

        builder.writeln("Heading 1")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING2

        builder.writeln("Heading 1.1")
        builder.writeln("Heading 1.2")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1

        builder.writeln("Heading 2")
        builder.writeln("Heading 3")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING2

        builder.writeln("Heading 3.1")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING3

        builder.writeln("Heading 3.1.1")
        builder.writeln("Heading 3.1.2")
        builder.writeln("Heading 3.1.3")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING2

        builder.writeln("Heading 3.2")
        builder.writeln("Heading 3.3")

        #ExStart:UpdateFields
        # The newly inserted table of contents will be initially empty.
        # It needs to be populated by updating the fields in the document.
        doc.update_fields()
        #ExEnd:UpdateFields

        doc.save(docs_base.artifacts_dir + "AddContentUsingDocumentBuilder.insert_table_of_contents.docx")
        #ExEnd:InsertTableOfContents


    def test_insert_inline_image(self) :

        #ExStart:InsertInlineImage
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_image(docs_base.images_dir + "Transparent background logo.png")

        doc.save(docs_base.artifacts_dir + "AddContentUsingDocumentBuilder.insert_inline_image.docx")
        #ExEnd:InsertInlineImage


    def test_insert_floating_image(self) :

        #ExStart:InsertFloatingImage
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_image(docs_base.images_dir + "Transparent background logo.png",
            aw.drawing.RelativeHorizontalPosition.MARGIN,
            100,
            aw.drawing.RelativeVerticalPosition.MARGIN,
            100,
            200,
            100,
            aw.drawing.WrapType.SQUARE)

        doc.save(docs_base.artifacts_dir + "AddContentUsingDocumentBuilder.insert_floating_image.docx")
        #ExEnd:InsertFloatingImage


    def test_insert_paragraph(self) :

        #ExStart:InsertParagraph
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        font = builder.font
        font.size = 16
        font.bold = True
        font.color = drawing.Color.blue
        font.name = "Arial"
        font.underline = aw.Underline.DASH

        paragraphFormat = builder.paragraph_format
        paragraphFormat.first_line_indent = 8
        paragraphFormat.alignment = aw.ParagraphAlignment.JUSTIFY
        paragraphFormat.keep_together = True

        builder.writeln("A whole paragraph.")

        doc.save(docs_base.artifacts_dir + "AddContentUsingDocumentBuilder.insert_paragraph.docx")
        #ExEnd:InsertParagraph


    def test_insert_tc_field(self) :

        #ExStart:InsertTCField
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field("TC \"Entry Text\" \\f t")

        doc.save(docs_base.artifacts_dir + "AddContentUsingDocumentBuilder.insert_tc_field.docx")
        #ExEnd:InsertTCField


    #def test_insert_tc_fields_at_text(self) :

    #    #ExStart:InsertTCFieldsAtText
    #    doc = aw.Document()

    #    options = aw.replacing.FindReplaceOptions()
    #    options.apply_font.highlight_color = drawing.Color.dark_orange
    #    options.replacing_callback = new InsertTCFieldHandler("Chapter 1", "\\l 1")

    #    doc.range.replace(new Regex("The Beginning"), "", options)
    #    #ExEnd:InsertTCFieldsAtText


    ##ExStart:InsertTCFieldHandler
    #public sealed class InsertTCFieldHandler : IReplacingCallback

    #    # Store the text and switches to be used for the TC fields.
    #    private readonly string mFieldText
    #    private readonly string mFieldSwitches

    #    # <summary>
    #    # The display text and switches to use for each TC field. Display name can be an empty string or None.
    #    # </summary>
    #    public InsertTCFieldHandler(string text, string switches)

    #        mFieldText = text
    #        mFieldSwitches = switches


    #    ReplaceAction IReplacingCallback.replacing(ReplacingArgs args)

    #        DocumentBuilder builder = new DocumentBuilder((Document) args.match_node.document)
    #        builder.move_to(args.match_node)

    #        # If the user-specified text to be used in the field as display text, then use that,
    #        # otherwise use the match string as the display text.
    #        string insertText = !string.is_None_or_empty(mFieldText) ? mFieldText : args.match.value

    #        builder.insert_field($"TC \"insertText\" mFieldSwitches")

    #        return ReplaceAction.skip


    ##ExEnd:InsertTCFieldHandler

    def test_cursor_position(self) :

        #ExStart:CursorPosition
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        curNode = builder.current_node
        curParagraph = builder.current_paragraph
        #ExEnd:CursorPosition

        print("\nCursor move to paragraph: " + curParagraph.get_text())


    def test_move_to_node(self) :

        #ExStart:MoveToNode
        #ExStart:MoveToBookmark
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Start a bookmark and add content to it using a DocumentBuilder.
        builder.start_bookmark("MyBookmark")
        builder.writeln("Bookmark contents.")
        builder.end_bookmark("MyBookmark")

        # The node that the DocumentBuilder is currently at is past the boundaries of the bookmark.
        self.assertEqual(doc.range.bookmarks[0].bookmark_end, builder.current_paragraph.first_child)

        # If we wish to revise the content of our bookmark with the DocumentBuilder, we can move back to it like this.
        builder.move_to_bookmark("MyBookmark")

        # Now we're located between the bookmark's BookmarkStart and BookmarkEnd nodes, so any text the builder adds will be within it.
        self.assertEqual(doc.range.bookmarks[0].bookmark_start, builder.current_paragraph.first_child)

        # We can move the builder to an individual node,
        # which in this case will be the first node of the first paragraph, like this.
        builder.move_to(doc.first_section.body.first_paragraph.get_child_nodes(aw.NodeType.ANY, False)[0])
        #ExEnd:MoveToBookmark

        self.assertEqual(aw.NodeType.BOOKMARK_START, builder.current_node.node_type)
        self.assertTrue(builder.is_at_start_of_paragraph)

        # A shorter way of moving the very start/end of a document is with these methods.
        builder.move_to_document_end()
        self.assertTrue(builder.is_at_end_of_paragraph)
        builder.move_to_document_start()
        self.assertTrue(builder.is_at_start_of_paragraph)
        #ExEnd:MoveToNode


    def test_move_to_document_start_end(self) :

        #ExStart:MoveToDocumentStartEnd
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Move the cursor position to the beginning of your document.
        builder.move_to_document_start()
        print("\nThis is the beginning of the document.")

        # Move the cursor position to the end of your document.
        builder.move_to_document_end()
        print("\nThis is the end of the document.")
        #ExEnd:MoveToDocumentStartEnd


    def test_move_to_section(self) :

        #ExStart:MoveToSection
        doc = aw.Document()
        doc.append_child(aw.Section(doc))

        # Move a DocumentBuilder to the second section and add text.
        builder = aw.DocumentBuilder(doc)
        builder.move_to_section(1)
        builder.writeln("Text added to the 2nd section.")

        # Create document with paragraphs.
        doc = aw.Document(docs_base.my_dir + "Paragraphs.docx")
        paragraphs = doc.first_section.body.paragraphs
        self.assertEqual(22, paragraphs.count)

        # When we create a DocumentBuilder for a document, its cursor is at the very beginning of the document by default,
        # and any content added by the DocumentBuilder will just be prepended to the document.
        builder = aw.DocumentBuilder(doc)
        self.assertEqual(0, paragraphs.index_of(builder.current_paragraph))

        # You can move the cursor to any position in a paragraph.
        builder.move_to_paragraph(2, 10)
        self.assertEqual(2, paragraphs.index_of(builder.current_paragraph))
        builder.writeln("This is a new third paragraph. ")
        self.assertEqual(3, paragraphs.index_of(builder.current_paragraph))
        #ExEnd:MoveToSection


    def test_move_to_headers_footers(self) :

        #ExStart:MoveToHeadersFooters
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Specify that we want headers and footers different for first, even and odd pages.
        builder.page_setup.different_first_page_header_footer = True
        builder.page_setup.odd_and_even_pages_header_footer = True

        # Create the headers.
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_FIRST)
        builder.write("Header for the first page")
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_EVEN)
        builder.write("Header for even pages")
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.write("Header for all other pages")

        # Create two pages in the document.
        builder.move_to_section(0)
        builder.writeln("Page1")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln("Page2")

        doc.save(docs_base.artifacts_dir + "AddContentUsingDocumentBuilder.move_to_headers_footers.docx")
        #ExEnd:MoveToHeadersFooters


    def test_move_to_paragraph(self) :

        #ExStart:MoveToParagraph
        doc = aw.Document(docs_base.my_dir + "Paragraphs.docx")
        builder = aw.DocumentBuilder(doc)

        builder.move_to_paragraph(2, 0)
        builder.writeln("This is the 3rd paragraph.")
        #ExEnd:MoveToParagraph


    def test_move_to_table_cell(self) :

        #ExStart:MoveToTableCell
        doc = aw.Document(docs_base.my_dir + "Tables.docx")
        builder = aw.DocumentBuilder(doc)

        # Move the builder to row 3, cell 4 of the first table.
        builder.move_to_cell(0, 2, 3, 0)
        builder.write("\nCell contents added by DocumentBuilder")
        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        self.assertEqual(table.rows[2].cells[3], builder.current_node.parent_node.parent_node)
        self.assertEqual("Cell contents added by DocumentBuilderCell 3 contents\a", table.rows[2].cells[3].get_text().strip())
        #ExEnd:MoveToTableCell


    def test_move_to_bookmark_end(self) :

        #ExStart:MoveToBookmarkEnd
        doc = aw.Document(docs_base.my_dir + "Bookmarks.docx")
        builder = aw.DocumentBuilder(doc)

        builder.move_to_bookmark("MyBookmark1", False, True)
        builder.writeln("This is a bookmark.")
        #ExEnd:MoveToBookmarkEnd


    def test_move_to_merge_field(self) :

        #ExStart:MoveToMergeField
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a field using the DocumentBuilder and add a run of text after it.
        field = builder.insert_field("MERGEFIELD field")
        builder.write(" Text after the field.")

        # The builder's cursor is currently at end of the document.
        self.assertIsNone(builder.current_node)
        # We can move the builder to a field like this, placing the cursor at immediately after the field.
        builder.move_to_field(field, True)

        # Note that the cursor is at a place past the FieldEnd node of the field, meaning that we are not actually inside the field.
        # If we wish to move the DocumentBuilder to inside a field,
        # we will need to move it to a field's FieldStart or FieldSeparator node using the DocumentBuilder.move_to() method.
        self.assertEqual(field.end, builder.current_node.previous_sibling)
        builder.write(" Text immediately after the field.")
        #ExEnd:MoveToMergeField


if __name__ == '__main__':
    unittest.main()
