import locale
import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing
import datetime

import api_example_base as aeb
from document_helper import DocumentHelper


class ExDocumentBuilder(aeb.ApiExampleBase):

    def test_write_and_font(self):
        # ExStart
        # ExFor:Font.size
        # ExFor:Font.bold
        # ExFor:Font.name
        # ExFor:Font.color
        # ExFor:Font.underline
        # ExFor:DocumentBuilder.#ctor
        # ExSummary:Shows how to insert formatted text using DocumentBuilder.

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Specify font formatting, then add text.
        font = builder.font
        font.size = 16
        font.bold = True
        font.color = drawing.Color.blue
        font.name = "Courier New"
        font.underline = aw.Underline.DASH

        builder.write("Hello world!")
        # ExEnd

        doc = DocumentHelper.save_open(builder.document)
        firstRun = doc.first_section.body.paragraphs[0].runs[0]

        self.assertEqual("Hello world!", firstRun.get_text().strip())
        self.assertEqual(16.0, firstRun.font.size)
        self.assertTrue(firstRun.font.bold)
        self.assertEqual("Courier New", firstRun.font.name)
        self.assertEqual(drawing.Color.blue.to_argb(), firstRun.font.color.to_argb())
        self.assertEqual(aw.Underline.DASH, firstRun.font.underline)

    def test_headers_and_footers(self):
        # ExStart
        # ExFor:DocumentBuilder
        # ExFor:DocumentBuilder.#ctor(Document)
        # ExFor:DocumentBuilder.move_to_header_footer
        # ExFor:DocumentBuilder.move_to_section
        # ExFor:DocumentBuilder.insert_break
        # ExFor:DocumentBuilder.writeln
        # ExFor:HeaderFooterType
        # ExFor:PageSetup.different_first_page_header_footer
        # ExFor:PageSetup.odd_and_even_pages_header_footer
        # ExFor:BreakType
        # ExSummary:Shows how to create headers and footers in a document using DocumentBuilder.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Specify that we want different headers and footers for first, even and odd pages.
        builder.page_setup.different_first_page_header_footer = True
        builder.page_setup.odd_and_even_pages_header_footer = True

        # Create the headers, then add three pages to the document to display each header type.
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_FIRST)
        builder.write("Header for the first page")
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_EVEN)
        builder.write("Header for even pages")
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.write("Header for all other pages")

        builder.move_to_section(0)
        builder.writeln("Page1")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln("Page2")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln("Page3")

        doc.save(aeb.artifacts_dir + "DocumentBuilder.headers_and_footers.docx")
        # ExEnd

        headers_footers = aw.Document(
            aeb.artifacts_dir + "DocumentBuilder.headers_and_footers.docx").first_section.headers_footers

        self.assertEqual(3, headers_footers.count)
        self.assertEqual("Header for the first page",
                         headers_footers[aw.HeaderFooterType.HEADER_FIRST].get_text().strip()) # AttributeError: 'NoneType' object has no attribute 'get_text'
        self.assertEqual("Header for even pages",
                         headers_footers[aw.HeaderFooterType.HEADER_EVEN].get_text().strip())  # True
        self.assertEqual("Header for all other pages",
                         headers_footers[aw.HeaderFooterType.HEADER_PRIMARY].get_text().strip())  # True

    def test_merge_fields(self):
        # ExStart
        # ExFor:DocumentBuilder.insert_field(String)
        # ExFor:DocumentBuilder.move_to_merge_field(String, Boolean, Boolean)
        # ExSummary:Shows how to insert fields, and move the document builder's cursor to them.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.insert_field("MERGEFIELD MyMergeField1 \\* MERGEFORMAT")
        builder.insert_field("MERGEFIELD MyMergeField2 \\* MERGEFORMAT")

        # Move the cursor to the first MERGEFIELD.
        builder.move_to_merge_field("MyMergeField1", True, False)

        # Note that the cursor is placed immediately after the first MERGEFIELD, and before the second.
        self.assertEqual(doc.range.fields[1].start, builder.current_node)
        self.assertEqual(doc.range.fields[0].end, builder.current_node.previous_sibling)

        # If we wish to edit the field's field code or contents using the builder,
        # its cursor would need to be inside a field.
        # To place it inside a field, we would need to call the document builder's MoveTo method
        # and pass the field's start or separator node as an argument.
        builder.write(" Text between our merge fields. ")

        doc.save(aeb.artifacts_dir + "DocumentBuilder.merge_fields.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.merge_fields.docx")

        self.assertEqual("\u0013MERGEFIELD MyMergeField1 \\* MERGEFORMAT\u0014«MyMergeField1»\u0015" +
                         " Text between our merge fields. " +
                         "\u0013MERGEFIELD MyMergeField2 \\* MERGEFORMAT\u0014«MyMergeField2»\u0015",
                         doc.get_text().strip())
        self.assertEqual(2, doc.range.fields.count)
        # TestUtil.verify_field(FieldType.field_merge_field, "MERGEFIELD MyMergeField1 \\* MERGEFORMAT", "«MyMergeField1»", doc.range.fields[0]) # TestUtil hasn't been done yet
        # TestUtil.verify_field(FieldType.field_merge_field, "MERGEFIELD MyMergeField2 \\* MERGEFORMAT", "«MyMergeField2»", doc.range.fields[1]) # TestUtil hasn't been done yet
        print("TestUtil hasn't been done yet")

    def test_insert_horizontal_rule(self):
        # ExStart
        # ExFor:DocumentBuilder.insert_horizontal_rule
        # ExFor:ShapeBase.is_horizontal_rule
        # ExFor:Shape.horizontal_rule_format
        # ExFor:HorizontalRuleFormat
        # ExFor:HorizontalRuleFormat.alignment
        # ExFor:HorizontalRuleFormat.width_percent
        # ExFor:HorizontalRuleFormat.height
        # ExFor:HorizontalRuleFormat.color
        # ExFor:HorizontalRuleFormat.no_shade
        # ExSummary:Shows how to insert a horizontal rule shape, and customize its formatting.

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        shape = builder.insert_horizontal_rule()

        horizontalRuleFormat = shape.horizontal_rule_format
        horizontalRuleFormat.alignment = aw.drawing.HorizontalRuleAlignment.CENTER
        horizontalRuleFormat.width_percent = 70
        horizontalRuleFormat.height = 3
        horizontalRuleFormat.color = drawing.Color.blue
        horizontalRuleFormat.no_shade = True

        self.assertTrue(shape.is_horizontal_rule)
        self.assertTrue(shape.horizontal_rule_format.no_shade)
        # ExEnd

        doc = DocumentHelper.save_open(doc)
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True)
        shape = shape.as_shape()

        self.assertEqual(aw.drawing.HorizontalRuleAlignment.CENTER, shape.horizontal_rule_format.alignment)
        self.assertEqual(70, shape.horizontal_rule_format.width_percent)
        self.assertEqual(3, shape.horizontal_rule_format.height)
        self.assertEqual(drawing.Color.blue.to_argb(), shape.horizontal_rule_format.color.to_argb())

    #    [Test(Description = "Checking the boundary conditions of WidthPercent and Height properties")]
    #    public void HorizontalRuleFormatExceptions()
    #
    #        DocumentBuilder builder = aw.DocumentBuilder()
    #        Shape shape = builder.insert_horizontal_rule()
    #
    #        HorizontalRuleFormat horizontalRuleFormat = shape.horizontal_rule_format
    #        horizontalRuleFormat.width_percent = 1
    #        horizontalRuleFormat.width_percent = 100
    #        Assert.that(() => horizontalRuleFormat.width_percent = 0, Throws.type_of<ArgumentOutOfRangeException>())
    #        Assert.that(() => horizontalRuleFormat.width_percent = 101, Throws.type_of<ArgumentOutOfRangeException>())
    #
    #        horizontalRuleFormat.height = 0
    #        horizontalRuleFormat.height = 1584
    #        Assert.that(() => horizontalRuleFormat.height = -1, Throws.type_of<ArgumentOutOfRangeException>())
    #        Assert.that(() => horizontalRuleFormat.height = 1585, Throws.type_of<ArgumentOutOfRangeException>())

    def test_insert_hyperlink(self):
        # ExStart
        # ExFor:DocumentBuilder.insert_hyperlink
        # ExFor:Font.clear_formatting
        # ExFor:Font.color
        # ExFor:Font.underline
        # ExFor:Underline
        # ExSummary:Shows how to insert a hyperlink field.

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("For more information, please visit the ")

        # Insert a hyperlink and emphasize it with custom formatting.
        # The hyperlink will be a clickable piece of text which will take us to the location specified in the URL.
        builder.font.color = drawing.Color.blue
        builder.font.underline = aw.Underline.SINGLE
        builder.insert_hyperlink("Google website", "https:#www.google.com", False)
        builder.font.clear_formatting()
        builder.writeln(".")

        # Ctrl + left clicking the link in the text in Microsoft Word will take us to the URL via a new web browser window.
        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_hyperlink.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_hyperlink.docx")

        hyperlink = doc.range.fields[0]
        # TestUtil.verify_web_response_status_code(HttpStatusCode.ok, hyperlink.address)
        #
        fieldContents = hyperlink.start.next_sibling.as_run
        #
        self.assertEqual(drawing.Color.blue.to_argb(), fieldContents.font.color.to_argb())
        self.assertEqual(aw.Underline.SINGLE, fieldContents.font.underline)
        self.assertEqual("HYPERLINK \"https:#www.google.com\"", fieldContents.get_text().strip())

    def test_push_pop_font(self):
        # ExStart
        # ExFor:DocumentBuilder.push_font
        # ExFor:DocumentBuilder.pop_font
        # ExFor:DocumentBuilder.insert_hyperlink
        # ExSummary:Shows how to use a document builder's formatting stack.

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Set up font formatting, then write the text that goes before the hyperlink.
        builder.font.name = "Arial"
        builder.font.size = 24
        builder.write("To visit Google, hold Ctrl and click ")

        # Preserve our current formatting configuration on the stack.
        builder.push_font()

        # Alter the builder's current formatting by applying a new style.
        builder.font.style_identifier = aw.StyleIdentifier.HYPERLINK
        builder.insert_hyperlink("here", "http:#www.google.com", False)

        self.assertEqual(drawing.Color.blue.to_argb(), builder.font.color.to_argb())
        self.assertEqual(aw.Underline.SINGLE, builder.font.underline)

        # Restore the font formatting that we saved earlier and remove the element from the stack.
        builder.pop_font()

        # self.assertEqual(drawing.Color.empty.to_argb(), builder.font.color.to_argb())
        self.assertEqual(aw.Underline.NONE, builder.font.underline)

        builder.write(". We hope you enjoyed the example.")

        doc.save(aeb.artifacts_dir + "DocumentBuilder.push_pop_font.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.push_pop_font.docx")
        runs = doc.first_section.body.first_paragraph.runs

        self.assertEqual(4, runs.count)

        self.assertEqual("To visit Google, hold Ctrl and click", runs[0].get_text().strip())
        self.assertEqual(". We hope you enjoyed the example.", runs[3].get_text().strip())
        self.assertEqual(runs[0].font.color, runs[3].font.color)
        self.assertEqual(runs[0].font.underline, runs[3].font.underline)

        self.assertEqual("here", runs[2].get_text().strip())
        self.assertEqual(drawing.Color.blue.to_argb(), runs[2].font.color.to_argb())
        self.assertEqual(aw.Underline.SINGLE, runs[2].font.underline)
        self.assertNotEqual(runs[0].font.color, runs[2].font.color)
        self.assertNotEqual(runs[0].font.underline, runs[2].font.underline)

    #        TestUtil.verify_web_response_status_code(HttpStatusCode.ok, ((FieldHyperlink)doc.range.fields[0]).address)
    print("TestUtil hasn't been done yet")

    # if NET462 || JAVA
    def test_insert_watermark(self):
        # ExStart
        # ExFor:DocumentBuilder.move_to_header_footer
        # ExFor:PageSetup.page_width
        # ExFor:PageSetup.page_height
        # ExFor:WrapType
        # ExFor:RelativeHorizontalPosition
        # ExFor:RelativeVerticalPosition
        # ExSummary:Shows how to insert an image, and use it as a watermark.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert the image into the header so that it will be visible on every page.
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        shape = builder.insert_image(aeb.image_dir + "Transparent background logo.png")
        shape.wrap_type = aw.drawing.WrapType.NONE
        shape.behind_text = True

        # Place the image at the center of the page.
        shape.relative_horizontal_position = aw.drawing.RelativeHorizontalPosition.PAGE
        shape.relative_vertical_position = aw.drawing.RelativeVerticalPosition.PAGE
        shape.left = (builder.page_setup.page_width - shape.width) / 2
        shape.top = (builder.page_setup.page_height - shape.height) / 2

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_watermark.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_watermark.docx")
        # shape = doc.first_section.headers_footers[aw.HeaderFooterType.HEADER_PRIMARY].get_child(aw.NodeType.SHAPE, 0, True).as_shape
        #
        # # TestUtil.verify_image_in_shape(400, 400, ImageType.png, shape)
        # self.assertEqual(aw.WrapType.NONE, shape.wrap_type)
        # self.assertTrue(shape.behind_text)
        # self.assertEqual(aw.RelativeHorizontalPosition.PAGE, shape.relative_horizontal_position)
        # self.assertEqual(aw.RelativeVerticalPosition.PAGE, shape.relative_vertical_position)
        # self.assertEqual((doc.first_section.page_setup.page_width - shape.width) / 2, shape.left)
        # self.assertEqual((doc.first_section.page_setup.page_height - shape.height) / 2, shape.top)

    # def test_insert_ole_object(self) :
    #
    #    #ExStart
    #    #ExFor:DocumentBuilder.insert_ole_object(String, Boolean, Boolean, Stream)
    #    #ExFor:DocumentBuilder.insert_ole_object(String, String, Boolean, Boolean, Stream)
    #    #ExFor:DocumentBuilder.insert_ole_object_as_icon(String, Boolean, String, String)
    #    #ExSummary:Shows how to insert an OLE object into a document.
    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)
    #
    #    # OLE objects are links to files in our local file system that can be opened by other installed applications.
    #    # Double clicking these shapes will launch the application, and then use it to open the linked object.
    #    # There are three ways of using the InsertOleObject method to insert these shapes and configure their appearance.
    #    # 1 -  Image taken from the local file system:
    #    imageStream = io.BytesIO(aeb.image_dir + "Logo.jpg")
    #
    #        # If 'presentation' is omitted and 'asIcon' is set, this overloaded method selects
    #        # the icon according to the file extension and uses the filename for the icon caption.
    #        builder.insert_ole_object(aeb.MyDir + "Spreadsheet.xlsx", False, False, imageStream)
    #
    #
    #    # If 'presentation' is omitted and 'asIcon' is set, this overloaded method selects
    #    # the icon according to 'progId' and uses the filename for the icon caption.
    #    # 2 -  Icon based on the application that will open the object:
    #    builder.insert_ole_object(aeb.MyDir + "Spreadsheet.xlsx", "Excel.sheet", False, True, null)
    #
    #    # If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
    #    # the icon according to 'progId' and uses the predefined icon caption.
    #    # 3 -  Image icon that's 32 x 32 pixels or smaller from the local file system, with a custom caption:
    #    builder.insert_ole_object_as_icon(aeb.MyDir + "Presentation.pptx", False, aeb.ImageDir + "Logo icon.ico",
    #        "Double click to view presentation!")
    #
    #    doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_ole_object.docx")
    #    #ExEnd
    #
    #    doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_ole_object.docx")
    #    Shape shape = (Shape)doc.get_child(NodeType.shape,0, True)
    #
    #    self.assertEqual(ShapeType.ole_object, shape.shape_type)
    #    self.assertEqual("Excel.sheet.12", shape.ole_format.prog_id)
    #    self.assertEqual(".xlsx", shape.ole_format.suggested_extension)
    #
    #    shape = (Shape)doc.get_child(NodeType.shape, 1, True)
    #
    #    self.assertEqual(ShapeType.ole_object, shape.shape_type)
    #    self.assertEqual("Package", shape.ole_format.prog_id)
    #    self.assertEqual(".xlsx", shape.ole_format.suggested_extension)
    #
    #    shape = (Shape)doc.get_child(NodeType.shape, 2, True)
    #
    #    self.assertEqual(ShapeType.ole_object, shape.shape_type)
    #    self.assertEqual("PowerPoint.show.12", shape.ole_format.prog_id)
    #    self.assertEqual(".pptx", shape.ole_format.suggested_extension)

    # elif NETCOREAPP2_1 || __MOBILE__
    # def test_insert_watermark_net_standard_2(self) :
    #
    #     #ExStart
    #     #ExFor:DocumentBuilder.move_to_header_footer
    #     #ExFor:PageSetup.page_width
    #     #ExFor:PageSetup.page_height
    #     #ExFor:WrapType
    #     #ExFor:RelativeHorizontalPosition
    #     #ExFor:RelativeVerticalPosition
    #     #ExSummary:Shows how to insert an image, and use it as a watermark (.net_standard 2.0).
    #     doc = aw.Document()
    #     builder = aw.DocumentBuilder(doc)
    #
    #     # Insert the image into the header so that it will be visible on every page.
    #     builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
    #
    #     builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
    #     shape = builder.insert_image(aeb.ImageDir + "Transparent background logo.png")
    #     shape.wrap_type = WrapType.none
    #     shape.behind_text = True
    #
    #     # Place the image at the center of the page.
    #     shape.relative_horizontal_position = RelativeHorizontalPosition.page
    #     shape.relative_vertical_position = RelativeVerticalPosition.page
    #     shape.left = (builder.page_setup.page_width - shape.width) / 2
    #     shape.top = (builder.page_setup.page_height - shape.height) / 2
    #
    #
    #     doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_watermark_net_standard_2.docx")
    # ExEnd
    #
    #        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_watermark_net_standard_2.docx")
    #        Shape outShape = (Shape)doc.first_section.headers_footers[aw.HeaderFooterType.HEADER_PRIMARY].get_child(NodeType.shape, 0, True)
    #
    #        TestUtil.verify_image_in_shape(400, 400, ImageType.png, outShape)
    #        self.assertEqual(WrapType.none, outShape.wrap_type)
    #        self.assertTrue(outShape.behind_text)
    #        self.assertEqual(RelativeHorizontalPosition.page, outShape.relative_horizontal_position)
    #        self.assertEqual(RelativeVerticalPosition.page, outShape.relative_vertical_position)
    #        self.assertEqual((doc.first_section.page_setup.page_width - outShape.width) / 2, outShape.left)
    #        self.assertEqual((doc.first_section.page_setup.page_height - outShape.height) / 2, outShape.top)
    #
    # # endif

    def test_insert_html(self):
        # ExStart
        # ExFor:DocumentBuilder.insert_html(String)
        # ExSummary:Shows how to use a document builder to insert html content into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        html = ("<p align='right'>Paragraph right</p>"
                "<b>Implicit paragraph left</b>"
                "<div align='center'>Div center</div>"
                "<h1 align='left'>Heading 1 left.</h1>")

        builder.insert_html(html)

        # Inserting HTML code parses the formatting of each element into equivalent document text formatting.
        paragraphs = doc.first_section.body.paragraphs

        self.assertEqual("Paragraph right", paragraphs[0].get_text().strip())
        self.assertEqual(aw.ParagraphAlignment.RIGHT, paragraphs[0].paragraph_format.alignment)

        self.assertEqual("Implicit paragraph left", paragraphs[1].get_text().strip())
        self.assertEqual(aw.ParagraphAlignment.LEFT, paragraphs[1].paragraph_format.alignment)
        self.assertTrue(paragraphs[1].runs[0].font.bold)

        self.assertEqual("Div center", paragraphs[2].get_text().strip())
        self.assertEqual(aw.ParagraphAlignment.CENTER, paragraphs[2].paragraph_format.alignment)

        self.assertEqual("Heading 1 left.", paragraphs[3].get_text().strip())
        self.assertEqual("Heading 1", paragraphs[3].paragraph_format.style.name)

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_html.docx")
        # ExEnd

    def test_insert_html_with_formatting(self):
        # ExStart
        # ExFor:DocumentBuilder.insert_html(String, Boolean)
        # ExSummary:Shows how to apply a document builder's formatting while inserting HTML content.

        for use_builder_formatting in (False, True):
            with self.subTest(use_builder_formatting=use_builder_formatting):
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Set a text alignment for the builder, insert an HTML paragraph with a specified alignment, and one without.
                builder.paragraph_format.alignment = aw.ParagraphAlignment.DISTRIBUTED
                builder.insert_html("<p align='right'>Paragraph 1.</p><p>Paragraph 2.</p>", use_builder_formatting)

                paragraphs = doc.first_section.body.paragraphs

                # The first paragraph has an alignment specified. When InsertHtml parses the HTML code,
                # the paragraph alignment value found in the HTML code always supersedes the document builder's value.
                self.assertEqual("Paragraph 1.", paragraphs[0].get_text().strip())
                self.assertEqual(aw.ParagraphAlignment.RIGHT, paragraphs[0].paragraph_format.alignment)

                # The second paragraph has no alignment specified. It can have its alignment value filled in
                # by the builder's value depending on the flag we passed to the InsertHtml method.
                self.assertEqual("Paragraph 2.", paragraphs[1].get_text().strip())
                self.assertEqual(
                    aw.ParagraphAlignment.DISTRIBUTED if use_builder_formatting else aw.ParagraphAlignment.LEFT,
                    paragraphs[1].paragraph_format.alignment)

                doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_html_with_formatting_" + str(
                    use_builder_formatting) + ".docx")
        # ExEnd

    def test_math_ml(self):
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        math_ml = "<math xmlns=\"http:#www.w_3.org/1998/Math/MathML\"><mrow><msub><mi>a</mi><mrow><mn>1</mn></mrow></msub><mo>+</mo><msub><mi>b</mi><mrow><mn>1</mn></mrow></msub></mrow></math>"

        builder.insert_html(math_ml)

        doc.save(aeb.artifacts_dir + "DocumentBuilder.MathML.docx")
        doc.save(aeb.artifacts_dir + "DocumentBuilder.MathML.pdf")

        self.assertTrue(DocumentHelper.compare_docs(
            aeb.golds_dir + "DocumentBuilder.MathML Gold.docx",
            aeb.artifacts_dir + "DocumentBuilder.MathML.docx"))

    def test_insert_text_and_bookmark(self):
        # ExStart
        # ExFor:DocumentBuilder.start_bookmark
        # ExFor:DocumentBuilder.end_bookmark
        # ExSummary:Shows how create a bookmark.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # A valid bookmark needs to have document body text enclosed by
        # BookmarkStart and BookmarkEnd nodes created with a matching bookmark name.
        builder.start_bookmark("MyBookmark")
        builder.writeln("Hello world!")
        builder.end_bookmark("MyBookmark")

        self.assertEqual(1, doc.range.bookmarks.count)
        self.assertEqual("MyBookmark", doc.range.bookmarks[0].name)
        self.assertEqual("Hello world!", doc.range.bookmarks[0].text.strip())
        # ExEnd

    def test_create_column_bookmark(self):
        # ExStart
        # ExFor:DocumentBuilder.start_column_bookmark
        # ExFor:DocumentBuilder.end_column_bookmark
        # ExSummary:Shows how to create a column bookmark.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_table()

        builder.insert_cell()
        # Cells 1,2,4,5 will be bookmarked.
        builder.start_column_bookmark("MyBookmark_1")
        # Badly formed bookmarks or bookmarks with duplicate names will be ignored when the document is saved.
        builder.start_column_bookmark("MyBookmark_1")
        builder.start_column_bookmark("BadStartBookmark")
        builder.write("Cell 1")

        builder.insert_cell()
        builder.write("Cell 2")

        builder.insert_cell()
        builder.write("Cell 3")

        builder.end_row()

        builder.insert_cell()
        builder.write("Cell 4")

        builder.insert_cell()
        builder.write("Cell 5")
        builder.end_column_bookmark("MyBookmark_1")
        builder.end_column_bookmark("MyBookmark_1")

        # Assert.Throws(typeof(InvalidOperationException), () => builder.EndColumnBookmark("BadEndBookmark")); # ExSkip # ???

        builder.insert_cell()
        builder.write("Cell 6")

        builder.end_row()
        builder.end_table()

        doc.save(aeb.artifacts_dir + "Bookmarks.create_column_bookmark.docx")
        # ExEnd

    @unittest.skip("How to convert a aspose.words.fields.DropDownItemCollection to array ? (line 604)")
    def test_create_form(self):
        # ExStart
        # ExFor:TextFormFieldType
        # ExFor:DocumentBuilder.insert_text_input
        # ExFor:DocumentBuilder.insert_combo_box
        # ExSummary:Shows how to create form fields.
        builder = aw.DocumentBuilder()

        # Form fields are objects in the document that the user can interact with by being prompted to enter values.
        # We can create them using a document builder, and below are two ways of doing so.
        # 1 -  Basic text input:
        builder.insert_text_input("My text input", aw.fields.TextFormFieldType.REGULAR,
                                  "", "Enter your name here", 30)

        # 2 -  Combo box with prompt text, and a range of possible values:
        items = "-- Select your favorite footwear --", "Sneakers", "Oxfords", "Flip-flops", "Other"

        builder.insert_paragraph()
        builder.insert_combo_box("My combo box", items, 0)

        builder.document.save(aeb.artifacts_dir + "DocumentBuilder.create_form.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.create_form.docx")
        form_field = doc.range.form_fields[0]

        self.assertEqual("My text input", form_field.name)
        self.assertEqual(aw.fields.TextFormFieldType.REGULAR, form_field.text_input_type)
        self.assertEqual("Enter your name here", form_field.result)

        form_field = doc.range.form_fields[1]

        self.assertEqual("My combo box", form_field.name)
        self.assertEqual(aw.fields.TextFormFieldType.REGULAR, form_field.text_input_type)
        self.assertEqual("-- Select your favorite footwear --", form_field.result)
        self.assertEqual(0, form_field.drop_down_selected_index)
        self.assertEqual(["-- Select your favorite footwear --", "Sneakers", "Oxfords", "Flip-flops", "Other"],
                         form_field.drop_down_items.to_array())

    def test_insert_check_box(self):

        # ExStart
        # ExFor:DocumentBuilder.insert_check_box(string, bool, bool, int)
        # ExFor:DocumentBuilder.insert_check_box(String, bool, int)
        # ExSummary:Shows how to insert checkboxes into the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert checkboxes of varying sizes and default checked statuses.
        builder.write("Unchecked check box of a default size: ")
        builder.insert_check_box("", False, False, 0)
        builder.insert_paragraph()

        builder.write("Large checked check box: ")
        builder.insert_check_box("CheckBox_Default", True, True, 50)
        builder.insert_paragraph()

        # Form fields have a name length limit of 20 characters.
        builder.write("Very large checked check box: ")
        builder.insert_check_box("CheckBox_OnlyCheckedValue", True, 100)

        self.assertEqual("CheckBox_OnlyChecked", doc.range.form_fields[2].name)

        # We can interact with these check boxes in Microsoft Word by double clicking them.
        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_check_box.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_check_box.docx")

        form_fields = doc.range.form_fields

        self.assertEqual("", form_fields[0].name)
        self.assertEqual(False, form_fields[0].checked)
        self.assertEqual(False, form_fields[0].default)
        self.assertEqual(10, form_fields[0].check_box_size)

        self.assertEqual("CheckBox_Default", form_fields[1].name)
        self.assertEqual(True, form_fields[1].checked)
        self.assertEqual(True, form_fields[1].default)
        self.assertEqual(50, form_fields[1].check_box_size)

        self.assertEqual("CheckBox_OnlyChecked", form_fields[2].name)
        self.assertEqual(True, form_fields[2].checked)
        self.assertEqual(True, form_fields[2].default)
        self.assertEqual(100, form_fields[2].check_box_size)

    def test_insert_check_box_empty_name(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Checking that the checkbox insertion with an empty name working correctly
        builder.insert_check_box("", True, False, 1)
        builder.insert_check_box("", False, 1)

    def test_working_with_nodes(self):

        # ExStart
        # ExFor:DocumentBuilder.move_to(Node)
        # ExFor:DocumentBuilder.move_to_bookmark(String)
        # ExFor:DocumentBuilder.current_paragraph
        # ExFor:DocumentBuilder.current_node
        # ExFor:DocumentBuilder.move_to_document_start
        # ExFor:DocumentBuilder.move_to_document_end
        # ExFor:DocumentBuilder.is_at_end_of_paragraph
        # ExFor:DocumentBuilder.is_at_start_of_paragraph
        # ExSummary:Shows how to move a document builder's cursor to different nodes in a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create a valid bookmark, an entity that consists of nodes enclosed by a bookmark start node,
        # and a bookmark end node. 
        builder.start_bookmark("MyBookmark")
        builder.write("Bookmark contents.")
        builder.end_bookmark("MyBookmark")

        first_paragraph_nodes = doc.first_section.body.first_paragraph.child_nodes

        self.assertEqual(aw.NodeType.BOOKMARK_START, first_paragraph_nodes[0].node_type)
        self.assertEqual(aw.NodeType.RUN, first_paragraph_nodes[1].node_type)
        self.assertEqual("Bookmark contents.", first_paragraph_nodes[1].get_text().strip())
        self.assertEqual(aw.NodeType.BOOKMARK_END, first_paragraph_nodes[2].node_type)

        # The document builder's cursor is always ahead of the node that we last added with it.
        # If the builder's cursor is at the end of the document, its current node will be null.
        # The previous node is the bookmark end node that we last added.
        # Adding new nodes with the builder will append them to the last node.
        self.assertIsNone(builder.current_node)

        # If we wish to edit a different part of the document with the builder,
        # we will need to bring its cursor to the node we wish to edit.
        builder.move_to_bookmark("MyBookmark")

        # Moving it to a bookmark will move it to the first node within the bookmark start and end nodes, the enclosed run.
        self.assertEqual(first_paragraph_nodes[1], builder.current_node)

        # We can also move the cursor to an individual node like this.
        builder.move_to(doc.first_section.body.first_paragraph.get_child_nodes(aw.NodeType.ANY, False)[0])

        self.assertEqual(aw.NodeType.BOOKMARK_START, builder.current_node.node_type)
        self.assertEqual(doc.first_section.body.first_paragraph, builder.current_paragraph)
        self.assertTrue(builder.is_at_start_of_paragraph)

        # We can use specific methods to move to the start/end of a document.
        builder.move_to_document_end()

        self.assertTrue(builder.is_at_end_of_paragraph)

        builder.move_to_document_start()

        self.assertTrue(builder.is_at_start_of_paragraph)
        # ExEnd

    def test_fill_merge_fields(self):

        # ExStart
        # ExFor:DocumentBuilder.move_to_merge_field(String)
        # ExFor:DocumentBuilder.bold
        # ExFor:DocumentBuilder.italic
        # ExSummary:Shows how to fill MERGEFIELDs with data with a document builder instead of a mail merge.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert some MERGEFIELDS, which accept data from columns of the same name in a data source during a mail merge,
        # and then fill them manually.
        builder.insert_field(" MERGEFIELD Chairman ")
        builder.insert_field(" MERGEFIELD ChiefFinancialOfficer ")
        builder.insert_field(" MERGEFIELD ChiefTechnologyOfficer ")

        builder.move_to_merge_field("Chairman")
        builder.bold = True
        builder.writeln("John Doe")

        builder.move_to_merge_field("ChiefFinancialOfficer")
        builder.italic = True
        builder.writeln("Jane Doe")

        builder.move_to_merge_field("ChiefTechnologyOfficer")
        builder.italic = True
        builder.writeln("John Bloggs")

        doc.save(aeb.artifacts_dir + "DocumentBuilder.fill_merge_fields.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.fill_merge_fields.docx")
        paragraphs = doc.first_section.body.paragraphs

        self.assertTrue(paragraphs[0].runs[0].font.bold)
        self.assertEqual("John Doe", paragraphs[0].runs[0].get_text().strip())

        self.assertTrue(paragraphs[1].runs[0].font.italic)
        self.assertEqual("Jane Doe", paragraphs[1].runs[0].get_text().strip())

        self.assertTrue(paragraphs[2].runs[0].font.italic)
        self.assertEqual("John Bloggs", paragraphs[2].runs[0].get_text().strip())

    @unittest.skip(".as_field_toc ??? need new wrapper")
    def test_insert_toc(self):

        # ExStart
        # ExFor:DocumentBuilder.insert_table_of_contents
        # ExFor:Document.update_fields
        # ExFor:DocumentBuilder.#ctor(Document)
        # ExFor:ParagraphFormat.style_identifier
        # ExFor:DocumentBuilder.insert_break
        # ExFor:BreakType
        # ExSummary:Shows how to insert a Table of contents (TOC) into a document using heading styles as entries.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a table of contents for the first page of the document.
        # Configure the table to pick up paragraphs with headings of levels 1 to 3.
        # Also, set its entries to be hyperlinks that will take us
        # to the location of the heading when left-clicked in Microsoft Word.
        builder.insert_table_of_contents("\\o \"1-3\" \\h \\z \\u")
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # Populate the table of contents by adding paragraphs with heading styles.
        # Each such heading with a level between 1 and 3 will create an entry in the table.
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

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING4
        builder.writeln("Heading 3.1.3.1")
        builder.writeln("Heading 3.1.3.2")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING2
        builder.writeln("Heading 3.2")
        builder.writeln("Heading 3.3")

        # A table of contents is a field of a type that needs to be updated to show an up-to-date result.
        doc.update_fields()
        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_toc.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_toc.docx")

        table_of_contents = doc.range.fields[0]
        table_of_contents = table_of_contents.as_field_toc()

        self.assertEqual("1-3", table_of_contents.heading_level_range)
        self.assertTrue(table_of_contents.insert_hyperlinks)
        self.assertTrue(table_of_contents.hide_in_web_layout)
        self.assertTrue(table_of_contents.use_paragraph_outline_level)

    def test_insert_table(self):

        # ExStart
        # ExFor:DocumentBuilder
        # ExFor:DocumentBuilder.write
        # ExFor:DocumentBuilder.start_table
        # ExFor:DocumentBuilder.insert_cell
        # ExFor:DocumentBuilder.end_row
        # ExFor:DocumentBuilder.end_table
        # ExFor:DocumentBuilder.cell_format
        # ExFor:DocumentBuilder.row_format
        # ExFor:CellFormat
        # ExFor:CellFormat.fit_text
        # ExFor:CellFormat.width
        # ExFor:CellFormat.vertical_alignment
        # ExFor:CellFormat.shading
        # ExFor:CellFormat.orientation
        # ExFor:CellFormat.wrap_text
        # ExFor:RowFormat
        # ExFor:RowFormat.borders
        # ExFor:RowFormat.clear_formatting
        # ExFor:Shading.clear_formatting
        # ExSummary:Shows how to build a table with custom borders.

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_table()

        # Setting table formatting options for a document builder
        # will apply them to every row and cell that we add with it.
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER

        builder.cell_format.clear_formatting()
        builder.cell_format.width = 150
        builder.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.CENTER
        builder.cell_format.shading.background_pattern_color = drawing.Color.green_yellow
        builder.cell_format.wrap_text = False
        builder.cell_format.fit_text = True

        builder.row_format.clear_formatting()
        builder.row_format.height_rule = aw.HeightRule.EXACTLY
        builder.row_format.height = 50
        builder.row_format.borders.line_style = aw.LineStyle.ENGRAVE3_D
        builder.row_format.borders.color = drawing.Color.orange

        builder.insert_cell()
        builder.write("Row 1, Col 1")

        builder.insert_cell()
        builder.write("Row 1, Col 2")
        builder.end_row()

        # Changing the formatting will apply it to the current cell,
        # and any new cells that we create with the builder afterward.
        # This will not affect the cells that we have added previously.
        builder.cell_format.shading.clear_formatting()

        builder.insert_cell()
        builder.write("Row 2, Col 1")

        builder.insert_cell()
        builder.write("Row 2, Col 2")

        builder.end_row()

        # Increase row height to fit the vertical text.
        builder.insert_cell()
        builder.row_format.height = 150
        builder.cell_format.orientation = aw.TextOrientation.UPWARD
        builder.write("Row 3, Col 1")

        builder.insert_cell()
        builder.cell_format.orientation = aw.TextOrientation.DOWNWARD
        builder.write("Row 3, Col 2")

        builder.end_row()
        builder.end_table()

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_table.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_table.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual("Row 1, Col 1\a", table.rows[0].cells[0].get_text().strip())
        self.assertEqual("Row 1, Col 2\a", table.rows[0].cells[1].get_text().strip())
        self.assertEqual(aw.HeightRule.EXACTLY, table.rows[0].row_format.height_rule)
        self.assertEqual(50.0, table.rows[0].row_format.height)
        self.assertEqual(aw.LineStyle.ENGRAVE3_D, table.rows[0].row_format.borders.line_style)
        self.assertEqual(drawing.Color.orange.to_argb(), table.rows[0].row_format.borders.color.to_argb())

        for c in table.rows[0].cells:
            c = c.as_cell()
            self.assertEqual(150, c.cell_format.width)
            self.assertEqual(aw.tables.CellVerticalAlignment.CENTER, c.cell_format.vertical_alignment)
            self.assertEqual(drawing.Color.green_yellow.to_argb(), c.cell_format.shading.background_pattern_color.to_argb())
            self.assertFalse(c.cell_format.wrap_text)
            self.assertTrue(c.cell_format.fit_text)

            self.assertEqual(aw.ParagraphAlignment.CENTER, c.first_paragraph.paragraph_format.alignment)

        self.assertEqual("Row 2, Col 1\a", table.rows[1].cells[0].get_text().strip())
        self.assertEqual("Row 2, Col 2\a", table.rows[1].cells[1].get_text().strip())

        for c in table.rows[1].cells:
            c = c.as_cell()
            self.assertEqual(150, c.cell_format.width)
            self.assertEqual(aw.tables.CellVerticalAlignment.CENTER, c.cell_format.vertical_alignment)
            # self.assertEqual(drawing.Color.empty.to_argb(), c.cell_format.shading.background_pattern_color.to_argb()) # color is not supported
            self.assertFalse(c.cell_format.wrap_text)
            self.assertTrue(c.cell_format.fit_text)

            self.assertEqual(aw.ParagraphAlignment.CENTER, c.first_paragraph.paragraph_format.alignment)

        self.assertEqual(150, table.rows[2].row_format.height)

        self.assertEqual("Row 3, Col 1\a", table.rows[2].cells[0].get_text().strip())
        self.assertEqual(aw.TextOrientation.UPWARD, table.rows[2].cells[0].cell_format.orientation)
        self.assertEqual(aw.ParagraphAlignment.CENTER,
                         table.rows[2].cells[0].first_paragraph.paragraph_format.alignment)

        self.assertEqual("Row 3, Col 2\a", table.rows[2].cells[1].get_text().strip())
        self.assertEqual(aw.TextOrientation.DOWNWARD, table.rows[2].cells[1].cell_format.orientation)
        self.assertEqual(aw.ParagraphAlignment.CENTER,
                         table.rows[2].cells[1].first_paragraph.paragraph_format.alignment)

    def test_insert_table_with_style(self):

        # ExStart
        # ExFor:Table.style_identifier
        # ExFor:Table.style_options
        # ExFor:TableStyleOptions
        # ExFor:Table.auto_fit
        # ExFor:AutoFitBehavior
        # ExSummary:Shows how to build a new table while applying a style.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        table = builder.start_table()

        # We must insert at least one row before setting any table formatting.
        builder.insert_cell()

        # Set the table style used based on the style identifier.
        # Note that not all table styles are available when saving to .doc format.
        table.style_identifier = aw.StyleIdentifier.MEDIUM_SHADING1_ACCENT1

        # Partially apply the style to features of the table based on predicates, then build the table.
        table.style_options = aw.tables.TableStyleOptions.FIRST_COLUMN | aw.tables.TableStyleOptions.ROW_BANDS | aw.tables.TableStyleOptions.FIRST_ROW
        table.auto_fit(aw.tables.AutoFitBehavior.AUTO_FIT_TO_CONTENTS)

        builder.writeln("Item")
        builder.cell_format.right_padding = 40
        builder.insert_cell()
        builder.writeln("Quantity (kg)")
        builder.end_row()

        builder.insert_cell()
        builder.writeln("Apples")
        builder.insert_cell()
        builder.writeln("20")
        builder.end_row()

        builder.insert_cell()
        builder.writeln("Bananas")
        builder.insert_cell()
        builder.writeln("40")
        builder.end_row()

        builder.insert_cell()
        builder.writeln("Carrots")
        builder.insert_cell()
        builder.writeln("50")
        builder.end_row()

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_table_with_style.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_table_with_style.docx")

        doc.expand_table_styles_to_direct_formatting()

        self.assertEqual("Medium Shading 1 Accent 1", table.style.name)
        self.assertEqual(
            aw.tables.TableStyleOptions.FIRST_COLUMN | aw.tables.TableStyleOptions.ROW_BANDS | aw.tables.TableStyleOptions.FIRST_ROW,
            table.style_options)
        self.assertEqual(189, table.first_row.first_cell.cell_format.shading.background_pattern_color.b)
        self.assertEqual(drawing.Color.white.to_argb(), table.first_row.first_cell.first_paragraph.runs[0].font.color.to_argb()) # color is not supported
        self.assertNotEqual(drawing.Color.light_blue.to_argb(),
             table.last_row.first_cell.cell_format.shading.background_pattern_color.b)
        # self.assertEqual(Color.empty.to_argb(), table.last_row.first_cell.first_paragraph.runs[0].font.color.to_argb())

    def test_insert_table_set_heading_row(self):

        # ExStart
        # ExFor:RowFormat.heading_format
        # ExSummary:Shows how to build a table with rows that repeat on every page.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()

        # Any rows inserted while the "HeadingFormat" flag is set to "True"
        # will show up at the top of the table on every page that it spans.
        builder.row_format.heading_format = True
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        builder.cell_format.width = 100
        builder.insert_cell()
        builder.write("Heading row 1")
        builder.end_row()
        builder.insert_cell()
        builder.write("Heading row 2")
        builder.end_row()

        builder.cell_format.width = 50
        builder.paragraph_format.clear_formatting()
        builder.row_format.heading_format = False

        # Add enough rows for the table to span two pages.
        for i in range(0, 50):
            builder.insert_cell()
            builder.write("Row " + str(table.rows.count) + ", column 1.")
            builder.insert_cell()
            builder.write("Row " + str(table.rows.count) + ", column 2.")
            builder.end_row()

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_table_set_heading_row.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_table_set_heading_row.docx")
        table = doc.first_section.body.tables[0]

        for i in range(0, table.rows.count):
            self.assertEqual(i < 2, table.rows[i].row_format.heading_format)

    def test_insert_table_with_preferred_width(self):

        # ExStart
        # ExFor:Table.preferred_width
        # ExFor:aw.tables.PreferredWidth.from_percent
        # ExFor:aw.tables.PreferredWidth
        # ExSummary:Shows how to set a table to auto fit to 50% of the width of the page.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Cell #1")
        builder.insert_cell()
        builder.write("Cell #2")
        builder.insert_cell()
        builder.write("Cell #3")

        table.preferred_width = aw.tables.PreferredWidth.from_percent(50)

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_table_with_preferred_width.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_table_with_preferred_width.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(aw.tables.PreferredWidthType.PERCENT, table.preferred_width.type)
        self.assertEqual(50, table.preferred_width.value)

    @unittest.skip("Line 1135: preferred_width is equated to an enumerator, which is not the required result")
    def test_insert_cells_with_preferred_widths(self):

        # ExStart
        # ExFor:CellFormat.preferred_width
        # ExFor:aw.tables.PreferredWidth
        # ExFor:aw.tables.PreferredWidth.auto
        # ExFor:aw.tables.PreferredWidth.equals(aw.tables.PreferredWidth)
        # ExFor:aw.tables.PreferredWidth.equals(System.object)
        # ExFor:aw.tables.PreferredWidth.from_points
        # ExFor:aw.tables.PreferredWidth.from_percent
        # ExFor:aw.tables.PreferredWidth.get_hash_code
        # ExFor:aw.tables.PreferredWidth.to_string
        # ExSummary:Shows how to set a preferred width for table cells.

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        table = builder.start_table()

        # There are two ways of applying the "aw.tables.PreferredWidth" class to table cells.
        # 1 -  Set an absolute preferred width based on points:
        builder.insert_cell()
        builder.cell_format.preferred_width = aw.tables.PreferredWidth.from_points(40)
        builder.cell_format.shading.background_pattern_color = drawing.Color.light_yellow
        builder.writeln("Cell with a width of " + str(builder.cell_format.preferred_width) + ".")

        # 2 -  Set a relative preferred width based on percent of the table's width:
        builder.insert_cell()
        builder.cell_format.preferred_width = aw.tables.PreferredWidth.from_percent(20)
        builder.cell_format.shading.background_pattern_color = drawing.Color.light_blue
        builder.writeln("Cell with a width of " + str(builder.cell_format.preferred_width) + ".")

        builder.insert_cell()

        # A cell with no preferred width specified will take up the rest of the available space.
        builder.cell_format.preferred_width = aw.tables.PreferredWidth.AUTO

        # Each configuration of the "aw.tables.PreferredWidth" property creates a new object.
        self.assertNotEqual(table.first_row.cells[1].cell_format.preferred_width.get_hash_code(),
                            builder.cell_format.preferred_width.get_hash_code())

        builder.cell_format.shading.background_pattern_color = drawing.Color.light_green
        builder.writeln("Automatically sized cell.")

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_cells_with_preferred_widths.docx")
        # ExEnd

        self.assertEqual(100.0, aw.tables.PreferredWidth.from_percent(100).value)
        self.assertEqual(100.0, aw.tables.PreferredWidth.from_points(100).value)

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_cells_with_preferred_widths.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(aw.tables.PreferredWidthType.POINTS, table.first_row.cells[0].cell_format.preferred_width.type)
        self.assertEqual(40.0, table.first_row.cells[0].cell_format.preferred_width.value)
        self.assertEqual("Cell with a width of 800.\r\a", table.first_row.cells[0].get_text().strip())

        self.assertEqual(aw.tables.PreferredWidthType.PERCENT,
                         table.first_row.cells[1].cell_format.preferred_width.type)
        self.assertEqual(20.0, table.first_row.cells[1].cell_format.preferred_width.value)
        self.assertEqual("Cell with a width of 20%.\r\a", table.first_row.cells[1].get_text().strip())

        self.assertEqual(aw.tables.PreferredWidthType.AUTO, table.first_row.cells[2].cell_format.preferred_width.type)
        self.assertEqual(0.0, table.first_row.cells[2].cell_format.preferred_width.value)
        self.assertEqual("Automatically sized cell.\r\a", table.first_row.cells[2].get_text().strip())

    def test_insert_table_from_html(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert the table from HTML. Note that AutoFitSettings does not apply to tables
        # inserted from HTML.
        builder.insert_html("<table>" + "<tr>" + "<td>Row 1, Cell 1</td>" + "<td>Row 1, Cell 2</td>" + "</tr>" +
                            "<tr>" + "<td>Row 2, Cell 2</td>" + "<td>Row 2, Cell 2</td>" + "</tr>" + "</table>")

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_table_from_html.docx")

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_table_from_html.docx")

        self.assertEqual(1, doc.get_child_nodes(aw.NodeType.TABLE, True).count)
        self.assertEqual(2, doc.get_child_nodes(aw.NodeType.ROW, True).count)
        self.assertEqual(4, doc.get_child_nodes(aw.NodeType.CELL, True).count)

    def test_insert_nested_table(self):

        # ExStart
        # ExFor:Cell.first_paragraph
        # ExSummary:Shows how to create a nested table using a document builder.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Build the outer table.
        cell = builder.insert_cell()
        builder.writeln("Outer Table Cell 1")
        builder.insert_cell()
        builder.writeln("Outer Table Cell 2")
        builder.end_table()

        # Move to the first cell of the outer table, the build another table inside the cell.
        builder.move_to(cell.first_paragraph)
        builder.insert_cell()
        builder.writeln("Inner Table Cell 1")
        builder.insert_cell()
        builder.writeln("Inner Table Cell 2")
        builder.end_table()

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_nested_table.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_nested_table.docx")

        self.assertEqual(2, doc.get_child_nodes(aw.NodeType.TABLE, True).count)
        self.assertEqual(4, doc.get_child_nodes(aw.NodeType.CELL, True).count)
        self.assertEqual(1, cell.tables[0].count)
        self.assertEqual(2, cell.tables[0].first_row.cells.count)

    def test_create_table(self):

        # ExStart
        # ExFor:DocumentBuilder
        # ExFor:DocumentBuilder.write
        # ExFor:DocumentBuilder.insert_cell
        # ExSummary:Shows how to use a document builder to create a table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Start the table, then populate the first row with two cells.
        builder.start_table()
        builder.insert_cell()
        builder.write("Row 1, Cell 1.")
        builder.insert_cell()
        builder.write("Row 1, Cell 2.")

        # Call the builder's "EndRow" method to start a new row.
        builder.end_row()
        builder.insert_cell()
        builder.write("Row 2, Cell 1.")
        builder.insert_cell()
        builder.write("Row 2, Cell 2.")
        builder.end_table()

        doc.save(aeb.artifacts_dir + "DocumentBuilder.create_table.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.create_table.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(4, table.get_child_nodes(aw.NodeType.CELL, True).count)

        self.assertEqual("Row 1, Cell 1.\a", table.rows[0].cells[0].get_text().strip())
        self.assertEqual("Row 1, Cell 2.\a", table.rows[0].cells[1].get_text().strip())
        self.assertEqual("Row 2, Cell 1.\a", table.rows[1].cells[0].get_text().strip())
        self.assertEqual("Row 2, Cell 2.\a", table.rows[1].cells[1].get_text().strip())

    def test_build_formatted_table(self):

        # ExStart
        # ExFor:RowFormat.height
        # ExFor:RowFormat.height_rule
        # ExFor:Table.left_indent
        # ExFor:DocumentBuilder.paragraph_format
        # ExFor:DocumentBuilder.font
        # ExSummary:Shows how to create a formatted table using DocumentBuilder.

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        table.left_indent = 20

        # Set some formatting options for text and table appearance.
        builder.row_format.height = 40
        builder.row_format.height_rule = aw.HeightRule.AT_LEAST
        builder.cell_format.shading.background_pattern_color = drawing.Color.from_argb(198, 217, 241)

        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        builder.font.size = 16
        builder.font.name = "Arial"
        builder.font.bold = True

        # Configuring the formatting options in a document builder will apply them
        # to the current cell/row its cursor is in,
        # as well as any new cells and rows created using that builder.
        builder.write("Header Row,\n Cell 1")
        builder.insert_cell()
        builder.write("Header Row,\n Cell 2")
        builder.insert_cell()
        builder.write("Header Row,\n Cell 3")
        builder.end_row()

        # Reconfigure the builder's formatting objects for new rows and cells that we are about to make.
        # The builder will not apply these to the first row already created so that it will stand out as a header row.
        builder.cell_format.shading.background_pattern_color = drawing.Color.white
        builder.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.CENTER
        builder.row_format.height = 30
        builder.row_format.height_rule = aw.HeightRule.AUTO
        builder.insert_cell()
        builder.font.size = 12
        builder.font.bold = False

        builder.write("Row 1, Cell 1.")
        builder.insert_cell()
        builder.write("Row 1, Cell 2.")
        builder.insert_cell()
        builder.write("Row 1, Cell 3.")
        builder.end_row()
        builder.insert_cell()
        builder.write("Row 2, Cell 1.")
        builder.insert_cell()
        builder.write("Row 2, Cell 2.")
        builder.insert_cell()
        builder.write("Row 2, Cell 3.")
        builder.end_row()
        builder.end_table()

        doc.save(aeb.artifacts_dir + "DocumentBuilder.create_formatted_table.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.create_formatted_table.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(20.0, table.left_indent)

        self.assertEqual(aw.HeightRule.AT_LEAST, table.rows[0].row_format.height_rule)
        self.assertEqual(40.0, table.rows[0].row_format.height)

        for c in doc.get_child_nodes(aw.NodeType.CELL, True):

            c = c.as_cell()
            self.assertEqual(aw.ParagraphAlignment.CENTER, c.first_paragraph.paragraph_format.alignment)

            for r in c.first_paragraph.runs:

                r = r.as_run()
                self.assertEqual("Arial", r.font.name)

                if c.parent_row == table.first_row:

                    self.assertEqual(16, r.font.size)
                    self.assertTrue(r.font.bold)

                else:

                    self.assertEqual(12, r.font.size)
                    self.assertFalse(r.font.bold)


    def test_table_borders_and_shading(self):

        # ExStart
        # ExFor:Shading
        # ExFor:Table.set_borders
        # ExFor:BorderCollection.left
        # ExFor:BorderCollection.right
        # ExFor:BorderCollection.top
        # ExFor:BorderCollection.bottom
        # ExSummary:Shows how to apply border and shading color while building a table.

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Start a table and set a default color/thickness for its borders.
        table = builder.start_table()
        table.set_borders(aw.LineStyle.SINGLE, 2.0, drawing.Color.black)

        # Create a row with two cells with different background colors.
        builder.insert_cell()
        builder.cell_format.shading.background_pattern_color = drawing.Color.light_sky_blue
        builder.writeln("Row 1, Cell 1.")
        builder.insert_cell()
        builder.cell_format.shading.background_pattern_color = drawing.Color.orange
        builder.writeln("Row 1, Cell 2.")
        builder.end_row()

        # Reset cell formatting to disable the background colors
        # set a custom border thickness for all new cells created by the builder,
        # then build a second row.
        builder.cell_format.clear_formatting()
        builder.cell_format.borders.left.line_width = 4.0
        builder.cell_format.borders.right.line_width = 4.0
        builder.cell_format.borders.top.line_width = 4.0
        builder.cell_format.borders.bottom.line_width = 4.0

        builder.insert_cell()
        builder.writeln("Row 2, Cell 1.")
        builder.insert_cell()
        builder.writeln("Row 2, Cell 2.")

        doc.save(aeb.artifacts_dir + "DocumentBuilder.table_borders_and_shading.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.table_borders_and_shading.docx")
        table = doc.first_section.body.tables[0]

        for c in table.first_row:
            c = c.as_cell()
            self.assertEqual(0.5, c.cell_format.borders.top.line_width)
            self.assertEqual(0.5, c.cell_format.borders.bottom.line_width)
            self.assertEqual(0.5, c.cell_format.borders.left.line_width)
            self.assertEqual(0.5, c.cell_format.borders.right.line_width)

            # self.assertEqual(Color.empty.to_argb(), c.cell_format.borders.left.color.to_argb())
            self.assertEqual(aw.LineStyle.SINGLE, c.cell_format.borders.left.line_style)

        self.assertEqual(drawing.Color.light_sky_blue.to_argb(),
                         table.first_row.first_cell.cell_format.shading.background_pattern_color.to_argb())
        self.assertEqual(drawing.Color.orange.to_argb(),
                         table.first_row.cells[1].cell_format.shading.background_pattern_color.to_argb())

        for c in table.last_row:
            c = c.as_cell()
            self.assertEqual(4.0, c.cell_format.borders.top.line_width)
            self.assertEqual(4.0, c.cell_format.borders.bottom.line_width)
            self.assertEqual(4.0, c.cell_format.borders.left.line_width)
            self.assertEqual(4.0, c.cell_format.borders.right.line_width)

            # self.assertEqual(Color.empty.to_argb(), c.cell_format.borders.left.color.to_argb())
            self.assertEqual(aw.LineStyle.SINGLE, c.cell_format.borders.left.line_style)
            # self.assertEqual(Color.empty.to_argb(), c.cell_format.shading.background_pattern_color.to_argb())

    def test_set_preferred_type_convert_util(self):

        # ExStart
        # ExFor:aw.tables.PreferredWidth.from_points
        # ExSummary:Shows how to use unit conversion tools while specifying a preferred width for a cell.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.cell_format.preferred_width = aw.tables.PreferredWidth.from_points(aw.ConvertUtil.inch_to_point(3))
        builder.insert_cell()

        self.assertEqual(216.0, table.first_row.first_cell.cell_format.preferred_width.value)
        # ExEnd

    #@unittest.skip("TestUtil hasn't been done yet")
    def test_insert_hyperlink_to_local_bookmark(self):

        # ExStart
        # ExFor:DocumentBuilder.start_bookmark
        # ExFor:DocumentBuilder.end_bookmark
        # ExFor:DocumentBuilder.insert_hyperlink
        # ExSummary:Shows how to insert a hyperlink which references a local bookmark.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_bookmark("Bookmark1")
        builder.write("Bookmarked text. ")
        builder.end_bookmark("Bookmark1")
        builder.writeln("Text outside of the bookmark.")

        # Insert a HYPERLINK field that links to the bookmark. We can pass field switches
        # to the "InsertHyperlink" method as part of the argument containing the referenced bookmark's name.
        builder.font.color = drawing.Color.blue # colors are not supported
        builder.font.underline = aw.Underline.SINGLE
        builder.insert_hyperlink("Link to Bookmark1", "Bookmark1\" \o \"Hyperlink Tip", True)

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_hyperlink_to_local_bookmark.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_hyperlink_to_local_bookmark.docx")
        # FieldHyperlink hyperlink = (FieldHyperlink)doc.range.fields[0]

        # TestUtil.verify_field(aw.FieldType.FIELD_HYPERLINK, " HYPERLINK \\l \"Bookmark1\" \\o \"Hyperlink Tip\" ",
        #                       "Link to Bookmark1", aw.hyperlink)
        # self.assertEqual("Bookmark1", aw.hyperlink.sub_address)
        # self.assertEqual("Hyperlink Tip", aw.hyperlink.screen_tip)
        # self.assertTrue(doc.range.bookmarks.any(lambda b: b.name == "Bookmark1"))

    def test_cursor_position(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.write("Hello world!")

        # If the builder's cursor is at the end of the document,
        # there will be no nodes in front of it so that the current node will be null.
        self.assertIsNone(builder.current_node)

        self.assertEqual("Hello world!", builder.current_paragraph.get_text().strip())

        # Move to the beginning of the document and place the cursor at an existing node.
        builder.move_to_document_start()
        self.assertEqual(aw.NodeType.RUN, builder.current_node.node_type)

    def test_move_to(self):

        # ExStart
        # ExFor:Story.last_paragraph
        # ExFor:DocumentBuilder.move_to(Node)
        # ExSummary:Shows how to move a DocumentBuilder's cursor position to a specified node.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Run 1. ")

        # The document builder has a cursor, which acts as the part of the document
        # where the builder appends new nodes when we use its document construction methods.
        # This cursor functions in the same way as Microsoft Word's blinking cursor,
        # and it also always ends up immediately after any node that the builder just inserted.
        # To append content to a different part of the document,
        # we can move the cursor to a different node with the "MoveTo" method.
        self.assertEqual(doc.first_section.body.last_paragraph, builder.current_paragraph)  # ExSkip
        builder.move_to(doc.first_section.body.first_paragraph.runs[0])
        self.assertEqual(doc.first_section.body.first_paragraph, builder.current_paragraph)  # ExSkip

        # The cursor is now in front of the node that we moved it to.
        # Adding a second run will insert it in front of the first run.
        builder.writeln("Run 2. ")

        self.assertEqual("Run 2. \rRun 1.", doc.get_text().strip())

        # Move the cursor to the end of the document to continue appending text to the end as before.
        builder.move_to(doc.last_section.body.last_paragraph)
        builder.writeln("Run 3. ")

        self.assertEqual("Run 2. \rRun 1. \rRun 3.", doc.get_text().strip())
        self.assertEqual(doc.first_section.body.last_paragraph, builder.current_paragraph)  # ExSkip
        # ExEnd

    def test_move_to_paragraph(self):

        # ExStart
        # ExFor:DocumentBuilder.move_to_paragraph
        # ExSummary:Shows how to move a builder's cursor position to a specified paragraph.
        doc = aw.Document(aeb.my_dir + "Paragraphs.docx")
        paragraphs = doc.first_section.body.paragraphs

        self.assertEqual(22, paragraphs.count)

        # Create document builder to edit the document. The builder's cursor,
        # which is the point where it will insert new nodes when we call its document construction methods,
        # is currently at the beginning of the document.
        builder = aw.DocumentBuilder(doc)

        self.assertEqual(0, paragraphs.index_of(builder.current_paragraph))

        # Move that cursor to a different paragraph will place that cursor in front of that paragraph.
        builder.move_to_paragraph(2, 0)
        self.assertEqual(2, paragraphs.index_of(builder.current_paragraph))  # ExSkip

        # Any new content that we add will be inserted at that point.
        builder.writeln("This is a new third paragraph. ")
        # ExEnd

        self.assertEqual(3, paragraphs.index_of(builder.current_paragraph))

        doc = DocumentHelper.save_open(doc)

        self.assertEqual("This is a new third paragraph.", doc.first_section.body.paragraphs[2].get_text().strip())

    def test_move_to_cell(self):

        # ExStart
        # ExFor:DocumentBuilder.move_to_cell
        # ExSummary:Shows how to move a document builder's cursor to a cell in a table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create an empty 2x2 table.
        builder.start_table()
        builder.insert_cell()
        builder.insert_cell()
        builder.end_row()
        builder.insert_cell()
        builder.insert_cell()
        builder.end_table()

        # Because we have ended the table with the EndTable method,
        # the document builder's cursor is currently outside the table.
        # This cursor has the same function as Microsoft Word's blinking text cursor.
        # It can also be moved to a different location in the document using the builder's MoveTo methods.
        # We can move the cursor back inside the table to a specific cell.
        builder.move_to_cell(0, 1, 1, 0)
        builder.write("Column 2, cell 2.")

        doc.save(aeb.artifacts_dir + "DocumentBuilder.move_to_cell.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.move_to_cell.docx")

        table = doc.first_section.body.tables[0]

        self.assertEqual("Column 2, cell 2.\a", table.rows[1].cells[1].get_text().strip())

    def test_move_to_bookmark(self):

        # ExStart
        # ExFor:DocumentBuilder.move_to_bookmark(String, Boolean, Boolean)
        # ExSummary:Shows how to move a document builder's node insertion point cursor to a bookmark.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # A valid bookmark consists of a BookmarkStart node, a BookmarkEnd node with a
        # matching bookmark name somewhere afterward, and contents enclosed by those nodes.
        builder.start_bookmark("MyBookmark")
        builder.write("Hello world! ")
        builder.end_bookmark("MyBookmark")

        # There are 4 ways of moving a document builder's cursor to a bookmark.
        # If we are between the BookmarkStart and BookmarkEnd nodes, the cursor will be inside the bookmark.
        # This means that any text added by the builder will become a part of the bookmark.
        # 1 -  Outside of the bookmark, in front of the BookmarkStart node:
        self.assertTrue(builder.move_to_bookmark("MyBookmark", True, False))
        builder.write("1. ")

        self.assertEqual("Hello world! ", doc.range.bookmarks[
            0].text)  # Item properties can use only int, so we can't bookmarks["MyBookmark"]
        self.assertEqual("1. Hello world!", doc.get_text().strip())

        # 2 -  Inside the bookmark, right after the BookmarkStart node:
        self.assertTrue(builder.move_to_bookmark("MyBookmark", True, True))
        builder.write("2. ")

        self.assertEqual("2. Hello world! ", doc.range.bookmarks[
            0].text)  # Item properties can use only int, so we can't bookmarks["MyBookmark"]
        self.assertEqual("1. 2. Hello world!", doc.get_text().strip())

        # 2 -  Inside the bookmark, right in front of the BookmarkEnd node:
        self.assertTrue(builder.move_to_bookmark("MyBookmark", False, False))
        builder.write("3. ")

        self.assertEqual("2. Hello world! 3. ", doc.range.bookmarks[
            0].text)  # Item properties can use only int, so we can't bookmarks["MyBookmark"]
        self.assertEqual("1. 2. Hello world! 3.", doc.get_text().strip())

        # 4 -  Outside of the bookmark, after the BookmarkEnd node:
        self.assertTrue(builder.move_to_bookmark("MyBookmark", False, True))
        builder.write("4.")

        self.assertEqual("2. Hello world! 3. ", doc.range.bookmarks[
            0].text)  # Item properties can use only int, so we can't bookmarks["MyBookmark"]
        self.assertEqual("1. 2. Hello world! 3. 4.", doc.get_text().strip())
        # ExEnd

    def test_build_table(self):

        # ExStart
        # ExFor:Table
        # ExFor:DocumentBuilder.start_table
        # ExFor:DocumentBuilder.end_row
        # ExFor:DocumentBuilder.end_table
        # ExFor:DocumentBuilder.cell_format
        # ExFor:DocumentBuilder.row_format
        # ExFor:DocumentBuilder.write(String)
        # ExFor:DocumentBuilder.writeln(String)
        # ExFor:tables.CellVerticalAlignment
        # ExFor:CellFormat.orientation
        # ExFor:TextOrientation
        # ExFor:AutoFitBehavior
        # ExSummary:Shows how to build a formatted 2x2 table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.CENTER
        builder.write("Row 1, cell 1.")
        builder.insert_cell()
        builder.write("Row 1, cell 2.")
        builder.end_row()

        # While building the table, the document builder will apply its current RowFormat/CellFormat property values
        # to the current row/cell that its cursor is in and any new rows/cells as it creates them.
        self.assertEqual(aw.tables.CellVerticalAlignment.CENTER, table.rows[0].cells[0].cell_format.vertical_alignment)
        self.assertEqual(aw.tables.CellVerticalAlignment.CENTER, table.rows[0].cells[1].cell_format.vertical_alignment)

        builder.insert_cell()
        builder.row_format.height = 100
        builder.row_format.height_rule = aw.HeightRule.EXACTLY
        builder.cell_format.orientation = aw.TextOrientation.UPWARD
        builder.write("Row 2, cell 1.")
        builder.insert_cell()
        builder.cell_format.orientation = aw.TextOrientation.DOWNWARD
        builder.write("Row 2, cell 2.")
        builder.end_row()
        builder.end_table()

        # Previously added rows and cells are not retroactively affected by changes to the builder's formatting.
        self.assertEqual(0, table.rows[0].row_format.height)
        self.assertEqual(aw.HeightRule.AUTO, table.rows[0].row_format.height_rule)
        self.assertEqual(100, table.rows[1].row_format.height)
        self.assertEqual(aw.HeightRule.EXACTLY, table.rows[1].row_format.height_rule)
        self.assertEqual(aw.TextOrientation.UPWARD, table.rows[1].cells[0].cell_format.orientation)
        self.assertEqual(aw.TextOrientation.DOWNWARD, table.rows[1].cells[1].cell_format.orientation)

        doc.save(aeb.artifacts_dir + "DocumentBuilder.build_table.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.build_table.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(2, table.rows.count)
        self.assertEqual(2, table.rows[0].cells.count)
        self.assertEqual(2, table.rows[1].cells.count)

        self.assertEqual(0, table.rows[0].row_format.height)
        self.assertEqual(aw.HeightRule.AUTO, table.rows[0].row_format.height_rule)
        self.assertEqual(100, table.rows[1].row_format.height)
        self.assertEqual(aw.HeightRule.EXACTLY, table.rows[1].row_format.height_rule)

        self.assertEqual("Row 1, cell 1.\a", table.rows[0].cells[0].get_text().strip())
        self.assertEqual(aw.tables.CellVerticalAlignment.CENTER, table.rows[0].cells[0].cell_format.vertical_alignment)

        self.assertEqual("Row 1, cell 2.\a", table.rows[0].cells[1].get_text().strip())

        self.assertEqual("Row 2, cell 1.\a", table.rows[1].cells[0].get_text().strip())
        self.assertEqual(aw.TextOrientation.UPWARD, table.rows[1].cells[0].cell_format.orientation)

        self.assertEqual("Row 2, cell 2.\a", table.rows[1].cells[1].get_text().strip())
        self.assertEqual(aw.TextOrientation.DOWNWARD, table.rows[1].cells[1].cell_format.orientation)

    def test_table_cell_vertical_rotated_far_east_text_orientation(self):

        doc = aw.Document(aeb.my_dir + "Rotated cell text.docx")

        table = doc.first_section.body.tables[0]
        cell = table.first_row.first_cell

        self.assertEqual(aw.TextOrientation.VERTICAL_ROTATED_FAR_EAST, cell.cell_format.orientation)

        doc = DocumentHelper.save_open(doc)

        table = doc.first_section.body.tables[0]
        cell = table.first_row.first_cell

        self.assertEqual(aw.TextOrientation.VERTICAL_ROTATED_FAR_EAST, cell.cell_format.orientation)

    def test_insert_floating_image(self):

        # ExStart
        # ExFor:DocumentBuilder.insert_image(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        # ExSummary:Shows how to insert an image.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # There are two ways of using a document builder to source an image and then insert it as a floating shape.
        # 1 -  From a file in the local file system:
        builder.insert_image(aeb.image_dir + "Transparent background logo.png",
                             aw.drawing.RelativeHorizontalPosition.MARGIN, 100,
                             aw.drawing.RelativeVerticalPosition.MARGIN, 0, 200, 200, aw.drawing.WrapType.SQUARE)

        # 2 -  From a URL:
        builder.insert_image(aeb.aspose_logo_url, aw.drawing.RelativeHorizontalPosition.MARGIN, 100,
                             aw.drawing.RelativeVerticalPosition.MARGIN, 250, 200, 200, aw.drawing.WrapType.SQUARE)

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_floating_image.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_floating_image.docx")
        image = doc.get_child(aw.NodeType.SHAPE, 0, True)
        image = image.as_shape()

        # TestUtil.verify_image_in_shape(400, 400, ImageType.png, image)
        self.assertEqual(100.0, image.left)
        self.assertEqual(0.0, image.top)
        self.assertEqual(200.0, image.width)
        self.assertEqual(200.0, image.height)
        self.assertEqual(aw.drawing.WrapType.SQUARE, image.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, image.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, image.relative_vertical_position)

        image = doc.get_child(aw.NodeType.SHAPE, 1, True)
        image = image.as_shape()

        # TestUtil.verify_image_in_shape(320, 320, aw.drawing.ImageType.PNG, image)
        self.assertEqual(100.0, image.left)
        self.assertEqual(250.0, image.top)
        self.assertEqual(200.0, image.width)
        self.assertEqual(200.0, image.height)
        self.assertEqual(aw.drawing.WrapType.SQUARE, image.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, image.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, image.relative_vertical_position)

    def test_insert_image_original_size(self):

        # ExStart
        # ExFor:DocumentBuilder.insert_image(String, aw.drawing.RelativeHorizontalPosition, Double, aw.drawing.RelativeVerticalPosition, Double, Double, Double, aw.drawing.WrapType)
        # ExSummary:Shows how to insert an image from the local file system into a document while preserving its dimensions.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # The InsertImage method creates a floating shape with the passed image in its image data.
        # We can specify the dimensions of the shape can be passing them to this method.
        image_shape = builder.insert_image(aeb.image_dir + "Logo.jpg", aw.drawing.RelativeHorizontalPosition.MARGIN, 0,
                                           aw.drawing.RelativeVerticalPosition.MARGIN, 0, -1, -1,
                                           aw.drawing.WrapType.SQUARE)

        # Passing negative values as the intended dimensions will automatically define
        # the shape's dimensions based on the dimensions of its image.
        self.assertEqual(300.0, image_shape.width)
        self.assertEqual(300.0, image_shape.height)

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_image_original_size.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_image_original_size.docx")
        image_shape = doc.get_child(aw.NodeType.SHAPE, 0, True)
        image_shape = image_shape.as_shape()


        # TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(0.0, image_shape.left)
        self.assertEqual(0.0, image_shape.top)
        self.assertEqual(300.0, image_shape.width)
        self.assertEqual(300.0, image_shape.height)
        self.assertEqual(aw.drawing.WrapType.SQUARE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, image_shape.relative_vertical_position)

    def test_insert_text_input(self):

        # ExStart
        # ExFor:DocumentBuilder.insert_text_input
        # ExSummary:Shows how to insert a text input form field into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a form that prompts the user to enter text.
        builder.insert_text_input("TextInput", aw.fields.TextFormFieldType.REGULAR, "", "Enter your text here", 0)

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_text_input.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_text_input.docx")
        form_field = doc.range.form_fields[0]

        self.assertTrue(form_field.enabled)
        self.assertEqual("TextInput", form_field.name)
        self.assertEqual(0, form_field.max_length)
        self.assertEqual("Enter your text here", form_field.result)
        self.assertEqual(aw.fields.FieldType.FIELD_FORM_TEXT_INPUT, form_field.type)
        self.assertEqual("", form_field.text_input_format)
        self.assertEqual(aw.fields.TextFormFieldType.REGULAR, form_field.text_input_type)

    @unittest.skip("How to convert a aspose.words.fields.DropDownItemCollection to array ?")
    def test_insert_combo_box(self):

        # ExStart
        # ExFor:DocumentBuilder.insert_combo_box
        # ExSummary:Shows how to insert a combo box form field into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a form that prompts the user to pick one of the items from the menu.
        builder.write("Pick a fruit: ")
        items = ["Apple", "Banana", "Cherry"]
        builder.insert_combo_box("DropDown", items, 0)

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_combo_box.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_combo_box.docx")
        form_field = doc.range.form_fields[0]

        self.assertTrue(form_field.enabled)
        self.assertEqual("DropDown", form_field.name)
        self.assertEqual(0, form_field.drop_down_selected_index)
        self.assertEqual(items, form_field.drop_down_items)
        self.assertEqual(aw.fields.FieldType.FIELD_FORM_DROP_DOWN, form_field.type)

    # "WORDSNET-16868"
    @unittest.skip("Need some imports (uuid, datetime), need type casting (1907), some problems with SignOptions properties")
    def test_signature_line_provider_id(self):

        # ExStart
        # ExFor:SignatureLine.is_signed
        # ExFor:SignatureLine.is_valid
        # ExFor:SignatureLine.provider_id
        # ExFor:SignatureLineOptions.show_date
        # ExFor:SignatureLineOptions.email
        # ExFor:SignatureLineOptions.default_instructions
        # ExFor:SignatureLineOptions.instructions
        # ExFor:SignatureLineOptions.allow_comments
        # ExFor:DocumentBuilder.insert_signature_line(SignatureLineOptions)
        # ExFor:SignOptions.provider_id
        # ExSummary:Shows how to sign a document with a personal certificate and a signature line.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        signature_line_options = aw.SignatureLineOptions()
        signature_line_options.signer = "vderyushev"
        signature_line_options.signer_title = "QA"
        signature_line_options.email = "vderyushev@aspose.com"
        signature_line_options.show_date = True
        signature_line_options.default_instructions = False
        signature_line_options.instructions = "Please sign here."
        signature_line_options.allow_comments = True

        signature_line = builder.insert_signature_line(signature_line_options).signature_line
        # signature_line.provider_id = uuid.UUID("CF5A7BB4-8F3C-4756-9DF6-BEF7F13259A2") # import uuid ???

        self.assertFalse(signature_line.is_signed)
        self.assertFalse(signature_line.is_valid)

        doc.save(aeb.artifacts_dir + "DocumentBuilder.signature_line_provider_id.docx")

        sign_options = aw.digitalsignatures.SignOptions
        sign_options.signature_line_id = signature_line.id
        sign_options.provider_id = signature_line.provider_id
        sign_options.comments = "Document was signed by vderyushev"
        #                                     SignTime = datetime.now() # import datetime ???

        cert_holder = aw.digitalsignatures.CertificateHolder.create(aeb.my_dir + "morzal.pfx", "aw")

        aw.digitalsignatures.DigitalSignatureUtil.sign(
            aeb.artifacts_dir + "DocumentBuilder.signature_line_provider_id.docx",
            aeb.artifacts_dir + "DocumentBuilder.signature_line_provider_id.signed.docx", cert_holder, sign_options)

        # Re-open our saved document, and verify that the "IsSigned" and "IsValid" properties both equal "True",
        # indicating that the signature line contains a signature.
        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.signature_line_provider_id.signed.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True)
        signatureLine = shape.signature_line

        self.assertTrue(signatureLine.is_signed)
        self.assertTrue(signatureLine.is_valid)
        # ExEnd

        self.assertEqual("vderyushev", signatureLine.signer)
        self.assertEqual("QA", signatureLine.signer_title)
        self.assertEqual("vderyushev@aspose.com", signatureLine.email)
        self.assertTrue(signatureLine.show_date)
        self.assertFalse(signatureLine.default_instructions)
        self.assertEqual("Please sign here.", signatureLine.instructions)
        self.assertTrue(signatureLine.allow_comments)
        self.assertTrue(signatureLine.is_signed)
        self.assertTrue(signatureLine.is_valid)

        signatures = aw.digitalsignatures.DigitalSignatureUtil.load_signatures(
            aeb.artifacts_dir + "DocumentBuilder.signature_line_provider_id.signed.docx")

        self.assertEqual(1, signatures.count)
        self.assertTrue(signatures[0].is_valid)
        self.assertEqual("Document was signed by vderyushev", signatures[0].comments)
        # self.assertEqual(DateTime.today, signatures[0].sign_time.date) # import datetime ???
        self.assertEqual("CN=Morzal.me", signatures[0].issuer_name)
        self.assertEqual(aw.digitalsignatures.DigitalSignatureType.xml_dsig, signatures[0].signature_type)

    def test_signature_line_inline(self):

        # ExStart
        # ExFor:DocumentBuilder.insert_signature_line(SignatureLineOptions, aw.drawing.RelativeHorizontalPosition, Double, aw.drawing.RelativeVerticalPosition, Double, aw.drawing.WrapType)
        # ExSummary:Shows how to insert an inline signature line into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        options = aw.SignatureLineOptions()
        options.signer = "John Doe"
        options.signer_title = "Manager"
        options.email = "johndoe@aspose.com"
        options.show_date = True
        options.default_instructions = False
        options.instructions = "Please sign here."
        options.allow_comments = True

        builder.insert_signature_line(options, aw.drawing.RelativeHorizontalPosition.RIGHT_MARGIN, 2.0,
                                      aw.drawing.RelativeVerticalPosition.PAGE, 3.0, aw.drawing.WrapType.INLINE)

        # The signature line can be signed in Microsoft Word by double clicking it.
        doc.save(aeb.artifacts_dir + "DocumentBuilder.signature_line_inline.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.signature_line_inline.docx")

        shape = doc.get_child(aw.NodeType.SHAPE, 0, True)
        shape = shape.as_shape()
        signatureLine = shape.signature_line

        self.assertEqual("John Doe", signatureLine.signer)
        self.assertEqual("Manager", signatureLine.signer_title)
        self.assertEqual("johndoe@aspose.com", signatureLine.email)
        self.assertTrue(signatureLine.show_date)
        self.assertFalse(signatureLine.default_instructions)
        self.assertEqual("Please sign here.", signatureLine.instructions)
        self.assertTrue(signatureLine.allow_comments)
        self.assertFalse(signatureLine.is_signed)
        self.assertFalse(signatureLine.is_valid)

    def test_set_paragraph_formatting(self):

        # ExStart
        # ExFor:ParagraphFormat.right_indent
        # ExFor:ParagraphFormat.left_indent
        # ExSummary:Shows how to configure paragraph formatting to create off-center text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Center all text that the document builder writes, and set up indents.
        # The indent configuration below will create a body of text that will sit asymmetrically on the page.
        # The "center" that we align the text to will be the middle of the body of text, not the middle of the page.
        paragraph_format = builder.paragraph_format
        paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        paragraph_format.left_indent = 100
        paragraph_format.right_indent = 50
        paragraph_format.space_after = 25

        builder.writeln(
            "This paragraph demonstrates how left and right indentation affects word wrapping.")
        builder.writeln(
            "The space between the above paragraph and this one depends on the DocumentBuilder's paragraph format.")

        doc.save(aeb.artifacts_dir + "DocumentBuilder.set_paragraph_formatting.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.set_paragraph_formatting.docx")

        for paragraph in doc.first_section.body.paragraphs:
            paragraph = paragraph.as_paragraph()
            self.assertEqual(aw.ParagraphAlignment.CENTER, paragraph.paragraph_format.alignment)
            self.assertEqual(100.0, paragraph.paragraph_format.left_indent)
            self.assertEqual(50.0, paragraph.paragraph_format.right_indent)
            self.assertEqual(25.0, paragraph.paragraph_format.space_after)

    def test_set_cell_formatting(self):
        # ExStart
        # ExFor:DocumentBuilder.cell_format
        # ExFor:CellFormat.width
        # ExFor:CellFormat.left_padding
        # ExFor:CellFormat.right_padding
        # ExFor:CellFormat.top_padding
        # ExFor:CellFormat.bottom_padding
        # ExFor:DocumentBuilder.start_table
        # ExFor:DocumentBuilder.end_table
        # ExSummary:Shows how to format cells with a document builder.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Row 1, cell 1.")

        # Insert a second cell, and then configure cell text padding options.
        # The builder will apply these settings at its current cell, and any new cells creates afterwards.
        builder.insert_cell()

        cell_format = builder.cell_format
        cell_format.width = 250
        cell_format.left_padding = 30
        cell_format.right_padding = 30
        cell_format.top_padding = 30
        cell_format.bottom_padding = 30

        builder.write("Row 1, cell 2.")
        builder.end_row()
        builder.end_table()

        # The first cell was unaffected by the padding reconfiguration, and still holds the default values.
        self.assertEqual(0.0, table.first_row.cells[0].cell_format.width)
        self.assertEqual(5.4, table.first_row.cells[0].cell_format.left_padding)
        self.assertEqual(5.4, table.first_row.cells[0].cell_format.right_padding)
        self.assertEqual(0.0, table.first_row.cells[0].cell_format.top_padding)
        self.assertEqual(0.0, table.first_row.cells[0].cell_format.bottom_padding)

        self.assertEqual(250.0, table.first_row.cells[1].cell_format.width)
        self.assertEqual(30.0, table.first_row.cells[1].cell_format.left_padding)
        self.assertEqual(30.0, table.first_row.cells[1].cell_format.right_padding)
        self.assertEqual(30.0, table.first_row.cells[1].cell_format.top_padding)
        self.assertEqual(30.0, table.first_row.cells[1].cell_format.bottom_padding)

        # The first cell will still grow in the output document to match the size of its neighboring cell.
        doc.save(aeb.artifacts_dir + "DocumentBuilder.set_cell_formatting.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.set_cell_formatting.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(159.3, table.first_row.cells[0].cell_format.width)
        self.assertEqual(5.4, table.first_row.cells[0].cell_format.left_padding)
        self.assertEqual(5.4, table.first_row.cells[0].cell_format.right_padding)
        self.assertEqual(0.0, table.first_row.cells[0].cell_format.top_padding)
        self.assertEqual(0.0, table.first_row.cells[0].cell_format.bottom_padding)

        self.assertEqual(310.0, table.first_row.cells[1].cell_format.width)
        self.assertEqual(30.0, table.first_row.cells[1].cell_format.left_padding)
        self.assertEqual(30.0, table.first_row.cells[1].cell_format.right_padding)
        self.assertEqual(30.0, table.first_row.cells[1].cell_format.top_padding)
        self.assertEqual(30.0, table.first_row.cells[1].cell_format.bottom_padding)

    def test_set_row_formatting(self):

        # ExStart
        # ExFor:DocumentBuilder.row_format
        # ExFor:aw.HeightRule
        # ExFor:RowFormat.height
        # ExFor:RowFormat.height_rule
        # ExSummary:Shows how to format rows with a document builder.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Row 1, cell 1.")

        # Start a second row, and then configure its height. The builder will apply these settings to
        # its current row, as well as any new rows it creates afterwards.
        builder.end_row()

        row_format = builder.row_format
        row_format.height = 100
        row_format.height_rule = aw.HeightRule.EXACTLY

        builder.insert_cell()
        builder.write("Row 2, cell 1.")
        builder.end_table()

        # The first row was unaffected by the padding reconfiguration and still holds the default values.
        self.assertEqual(0.0, table.rows[0].row_format.height)
        self.assertEqual(aw.HeightRule.AUTO, table.rows[0].row_format.height_rule)

        self.assertEqual(100.0, table.rows[1].row_format.height)
        self.assertEqual(aw.HeightRule.EXACTLY, table.rows[1].row_format.height_rule)

        doc.save(aeb.artifacts_dir + "DocumentBuilder.set_row_formatting.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.set_row_formatting.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(0.0, table.rows[0].row_format.height)
        self.assertEqual(aw.HeightRule.AUTO, table.rows[0].row_format.height_rule)

        self.assertEqual(100.0, table.rows[1].row_format.height)
        self.assertEqual(aw.HeightRule.EXACTLY, table.rows[1].row_format.height_rule)

    def test_insert_footnote(self):

        # ExStart
        # ExFor:FootnoteType
        # ExFor:DocumentBuilder.insert_footnote(FootnoteType,String)
        # ExFor:DocumentBuilder.insert_footnote(FootnoteType,String,String)
        # ExSummary:Shows how to reference text with a footnote and an endnote.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert some text and mark it with a footnote with the IsAuto property set to "True" by default,
        # so the marker seen in the body text will be auto-numbered at "1",
        # and the footnote will appear at the bottom of the page.
        builder.write("This text will be referenced by a footnote.")
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Footnote comment regarding referenced text.")

        # Insert more text and mark it with an endnote with a custom reference mark,
        # which will be used in place of the number "2" and set "IsAuto" to False.
        builder.write("This text will be referenced by an endnote.")
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Endnote comment regarding referenced text.", "CustomMark")

        # Footnotes always appear at the bottom of their referenced text,
        # so this page break will not affect the footnote.
        # On the other hand, endnotes are always at the end of the document
        # so that this page break will push the endnote down to the next page.
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_footnote.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_footnote.docx")

        # TestUtil.verify_footnote(aw.notes.FootnoteType.FOOTNOTE, True, "",
        #                          "Footnote comment regarding referenced text.",
        #                          doc.get_child(aw.NodeType.FOOTNOTE, 0, True))
        # TestUtil.verify_footnote(aw.notes.FootnoteType.ENDNOTE, False, "CustomMark",
        #                          "CustomMark Endnote comment regarding referenced text.",
        #                          doc.get_child(aw.NodeType.FOOTNOTE, 1, True))

    def test_apply_borders_and_shading(self):

        # ExStart
        # ExFor:BorderCollection.item(BorderType)
        # ExFor:Shading
        # ExFor:TextureIndex
        # ExFor:ParagraphFormat.shading
        # ExFor:Shading.texture
        # ExFor:Shading.background_pattern_color
        # ExFor:Shading.foreground_pattern_color
        # ExSummary:Shows how to decorate text with borders and shading.

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        borders = builder.paragraph_format.borders
        borders.distance_from_text = 20
        borders[aw.BorderType.LEFT].line_style = aw.LineStyle.DOUBLE
        borders[aw.BorderType.RIGHT].line_style = aw.LineStyle.DOUBLE
        borders[aw.BorderType.TOP].line_style = aw.LineStyle.DOUBLE
        borders[aw.BorderType.BOTTOM].line_style = aw.LineStyle.DOUBLE

        shading = builder.paragraph_format.shading
        shading.texture = aw.TextureIndex.TEXTURE_DIAGONAL_CROSS
        shading.background_pattern_color = drawing.Color.light_coral
        shading.foreground_pattern_color = drawing.Color.light_salmon

        builder.write("This paragraph is formatted with a double border and shading.")
        doc.save(aeb.artifacts_dir + "DocumentBuilder.apply_borders_and_shading.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.apply_borders_and_shading.docx")
        borders = doc.first_section.body.first_paragraph.paragraph_format.borders

        self.assertEqual(20.0, borders.distance_from_text)
        self.assertEqual(aw.LineStyle.DOUBLE, borders[aw.BorderType.LEFT].line_style)
        self.assertEqual(aw.LineStyle.DOUBLE, borders[aw.BorderType.RIGHT].line_style)
        self.assertEqual(aw.LineStyle.DOUBLE, borders[aw.BorderType.TOP].line_style)
        self.assertEqual(aw.LineStyle.DOUBLE, borders[aw.BorderType.BOTTOM].line_style)

        self.assertEqual(aw.TextureIndex.TEXTURE_DIAGONAL_CROSS, shading.texture)
        self.assertEqual(drawing.Color.light_coral.to_argb(), shading.background_pattern_color.to_argb())
        self.assertEqual(drawing.Color.light_salmon.to_argb(), shading.foreground_pattern_color.to_argb())

    def test_delete_row(self):

        # ExStart
        # ExFor:DocumentBuilder.delete_row
        # ExSummary:Shows how to delete a row from a table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Row 1, cell 1.")
        builder.insert_cell()
        builder.write("Row 1, cell 2.")
        builder.end_row()
        builder.insert_cell()
        builder.write("Row 2, cell 1.")
        builder.insert_cell()
        builder.write("Row 2, cell 2.")
        builder.end_table()

        self.assertEqual(2, table.rows.count)

        # Delete the first row of the first table in the document.
        builder.delete_row(0, 0)

        self.assertEqual(1, table.rows.count)
        self.assertEqual("Row 2, cell 1.\aRow 2, cell 2.\a\a", table.get_text().strip())
        # ExEnd

    def test_append_document_and_resolve_styles(self):

        # ExStart
        # ExFor:Document.append_document(Document, ImportFormatMode, ImportFormatOptions)
        # ExSummary:Shows how to manage list style clashes while appending a document.
        # Load a document with text in a custom style and clone it.

        for keep_source_numbering in (False, True):
            with self.subTest(keep_source_numbering=keep_source_numbering):
                src_doc = aw.Document(aeb.my_dir + "Custom list numbering.docx")
                dst_doc = src_doc.clone()

                # We now have two documents, each with an identical style named "CustomStyle".
                # Change the text color for one of the styles to set it apart from the other.
                # dst_doc.styles["CustomStyle"].font.color = drawing.Color.dark_red              #  Item properties can use only int

                # If there is a clash of list styles, apply the list format of the source document.
                # Set the "keep_source_numbering" property to "False" to not import any list numbers into the destination document.
                # Set the "keep_source_numbering" property to "True" import all clashing
                # list style numbering with the same appearance that it had in the source document.
                options = aw.ImportFormatOptions()
                options.keep_source_numbering = keep_source_numbering

                # Joining two documents that have different styles that share the same name causes a style clash.
                # We can specify an import format mode while appending documents to resolve this clash.
                dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_DIFFERENT_STYLES, options)
                dst_doc.update_list_labels()

                dst_doc.save(aeb.artifacts_dir + "DocumentBuilder.append_document_and_resolve_styles.docx")
                # ExEnd

    def test_insert_document_and_resolve_styles(self):

        # ExStart
        # ExFor:Document.append_document(Document, ImportFormatMode, ImportFormatOptions)
        # ExSummary:Shows how to manage list style clashes while inserting a document.
        for keep_source_numbering in (False, True):
            with self.subTest(keep_source_numbering=keep_source_numbering):
                dst_doc = aw.Document()
                builder = aw.DocumentBuilder(dst_doc)
                builder.insert_break(aw.BreakType.PARAGRAPH_BREAK)

                dst_doc.lists.add(aw.lists.ListTemplate.NUMBER_DEFAULT)
                list = dst_doc.lists[0]

                builder.list_format.list = list

                for i in range(1, 16):
                    builder.write("List Item i\n")

                attach_doc = dst_doc.clone(True)

                # If there is a clash of list styles, apply the list format of the source document.
                # Set the "keep_source_numbering" property to "False" to not import any list numbers into the destination document.
                # Set the "keep_source_numbering" property to "True" import all clashing
                # list style numbering with the same appearance that it had in the source document.
                import_options = aw.ImportFormatOptions()
                import_options.keep_source_numbering = keep_source_numbering

                builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
                builder.insert_document(attach_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, import_options)

                dst_doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_document_and_resolve_styles.docx")
                # ExEnd

    def test_load_document_with_list_numbering(self):

        # ExStart
        # ExFor:Document.append_document(Document, ImportFormatMode, ImportFormatOptions)
        # ExSummary:Shows how to manage list style clashes while appending a clone of a document to itself.
        for keep_source_numbering in (False, True):
            with self.subTest(keep_source_numbering=keep_source_numbering):
                src_doc = aw.Document(aeb.my_dir + "List item.docx")
                dst_doc = aw.Document(aeb.my_dir + "List item.docx")

                # If there is a clash of list styles, apply the list format of the source document.
                # Set the "keep_source_numbering" property to "False" to not import any list numbers into the destination document.
                # Set the "keep_source_numbering" property to "True" import all clashing
                # list style numbering with the same appearance that it had in the source document.
                builder = aw.DocumentBuilder(dst_doc)
                builder.move_to_document_end()
                builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)

                options = aw.ImportFormatOptions()
                options.keep_source_numbering = keep_source_numbering
                builder.insert_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, options)

                dst_doc.update_list_labels()
                # ExEnd

    # @unittest.skip("No typecasting (lines 2365-2369)")
    def test_ignore_text_boxes(self):

        # ExStart
        # ExFor:ImportFormatOptions.ignore_text_boxes
        # ExSummary:Shows how to manage text box formatting while appending a document.
        # Create a document that will have nodes from another document inserted into it.
        for ignore_text_boxes in (False, True):
            with self.subTest(ignore_text_boxes=ignore_text_boxes):
                dst_doc = aw.Document()
                builder = aw.DocumentBuilder(dst_doc)

                builder.writeln("Hello world!")

                # Create another document with a text box, which we will import into the first document.
                src_doc = aw.Document()
                builder = aw.DocumentBuilder(src_doc)

                text_box = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 300, 100)
                builder.move_to(text_box.first_paragraph)
                builder.paragraph_format.style.font.name = "Courier New"
                builder.paragraph_format.style.font.size = 24
                builder.write("Textbox contents")

                # Set a flag to specify whether to clear or preserve text box formatting
                # while importing them to other documents.
                import_format_options = aw.ImportFormatOptions()
                import_format_options.ignore_text_boxes = ignore_text_boxes

                # Import the text box from the source document into the destination document,
                # and then verify whether we have preserved the styling of its text contents.
                importer = aw.NodeImporter(src_doc, dst_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING,
                                           import_format_options)
                imported_text_box = importer.import_node(text_box, True)
                imported_text_box = imported_text_box.as_shape()
                dst_doc.first_section.body.paragraphs[1].append_child(imported_text_box)

                if ignore_text_boxes:
                    self.assertEqual(12.0, imported_text_box.first_paragraph.runs[0].font.size)
                    self.assertEqual("Times New Roman", imported_text_box.first_paragraph.runs[0].font.name)
                else:
                    self.assertEqual(24.0, imported_text_box.first_paragraph.runs[0].font.size)
                    self.assertEqual("Courier New", imported_text_box.first_paragraph.runs[0].font.name)

                dst_doc.save(aeb.artifacts_dir + "DocumentBuilder.ignore_text_boxes.docx")
                # ExEnd

    def test_move_to_field(self):

        # ExStart
        # ExFor:DocumentBuilder.move_to_field
        # ExSummary:Shows how to move a document builder's node insertion point cursor to a specific field.
        for move_cursor_to_after_the_field in (False, True):
            with self.subTest(move_cursor_to_after_the_field=move_cursor_to_after_the_field):
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Insert a field using the DocumentBuilder and add a run of text after it.
                field = builder.insert_field(" AUTHOR \"John Doe\" ")

                # The builder's cursor is currently at end of the document.
                self.assertIsNone(builder.current_node)

                # Move the cursor to the field while specifying whether to place that cursor before or after the field.
                builder.move_to_field(field, move_cursor_to_after_the_field)

                # Note that the cursor is outside of the field in both cases.
                # This means that we cannot edit the field using the builder like this.
                # To edit a field, we can use the builder's MoveTo method on a field's FieldStart
                # or FieldSeparator node to place the cursor inside.
                if move_cursor_to_after_the_field:
                    self.assertIsNone(builder.current_node)
                    builder.write(" Text immediately after the field.")
                    self.assertEqual(
                        "\u0013 AUTHOR \"John Doe\" \u0014John Doe\u0015 Text immediately after the field.",
                        doc.get_text().strip())

                else:
                    self.assertEqual(field.start, builder.current_node)
                    builder.write("Text immediately before the field. ")
                    self.assertEqual(
                        "Text immediately before the field. \u0013 AUTHOR \"John Doe\" \u0014John Doe\u0015",
                        doc.get_text().strip())

                # ExEnd

    # def test_insert_ole_object_exception(self):
    #
    #     doc = aw.Document()
    #     builder = aw.DocumentBuilder(doc)
    #
    #     self.assertRaises(builder.insert_ole_object("", "checkbox", False, True, None),
    #                       Throws.type_of < ArgumentException > ())

    # @unittest.skip("Need to change locale language")
    def test_insert_pie_chart(self):

        # ExStart
        # ExFor:DocumentBuilder.insert_chart(ChartType, Double, Double)
        # ExSummary:Shows how to insert a pie chart into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        chart = builder.insert_chart(aw.drawing.charts.ChartType.PIE, aw.ConvertUtil.pixel_to_point(300),
                                     aw.ConvertUtil.pixel_to_point(300)).chart
        self.assertEqual(225.0, aw.ConvertUtil.pixel_to_point(300))  # ExSkip
        chart.series.clear()
        chart.series.add("My fruit",
                         ["Apples", "Bananas", "Cherries"],
                         [1.3, 2.2, 1.5])

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_pie_chart.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_pie_chart.docx")
        chart_shape = doc.get_child(aw.NodeType.SHAPE, 0, True)
        chart_shape = chart_shape.as_shape()

        self.assertEqual("Chart Title", chart_shape.chart.title.text)
        self.assertEqual(225.0, chart_shape.width)
        self.assertEqual(225.0, chart_shape.height)

    def test_insert_chart_relative_position(self):

        # ExStart
        # ExFor:DocumentBuilder.insert_chart(ChartType, aw.drawing.RelativeHorizontalPosition, Double, aw.drawing.RelativeVerticalPosition, Double, Double, Double, aw.drawing.WrapType)
        # ExSummary:Shows how to specify position and wrapping while inserting a chart.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_chart(aw.drawing.charts.ChartType.PIE,
                             aw.drawing.RelativeHorizontalPosition.MARGIN,
                             100,
                             aw.drawing.RelativeVerticalPosition.MARGIN,
                             100,
                             200,
                             100,
                             aw.drawing.WrapType.SQUARE)

        doc.save(aeb.artifacts_dir + "DocumentBuilder.inserted_chart_relative_position.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.inserted_chart_relative_position.docx")
        chart_shape = doc.get_child(aw.NodeType.SHAPE, 0, True)
        chart_shape = chart_shape.as_shape()

        self.assertEqual(100.0, chart_shape.top)
        self.assertEqual(100.0, chart_shape.left)
        self.assertEqual(200.0, chart_shape.width)
        self.assertEqual(100.0, chart_shape.height)
        self.assertEqual(aw.drawing.WrapType.SQUARE, chart_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, chart_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, chart_shape.relative_vertical_position)

    @unittest.skip("Aspose.Words date format does not match with datetime date format")
    def test_insert_field(self):
        locale.setlocale(locale.LC_ALL, 'en_US')
        # ExStart
        # ExFor:DocumentBuilder.insert_field(String)
        # ExFor:Field
        # ExFor:Field.result
        # ExFor:Field.get_field_code
        # ExFor:Field.type
        # ExFor:FieldType
        # ExSummary:Shows how to insert a field into a document using a field code.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field("DATE \\@ \"dddd, MMMM dd, yyyy\"")

        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field.type)
        self.assertEqual("DATE \\@ \"dddd, MMMM dd, yyyy\"", field.get_field_code())

        # This overload of the InsertField method automatically updates inserted fields.
        self.assertIs(datetime.datetime.strptime(field.result, '%A, %B %d, %Y'), datetime.now())
        # ExEnd

    def test_insert_field_and_update(self):

        # ExStart
        # ExFor:DocumentBuilder.insert_field(FieldType, Boolean)
        # ExFor:Field.update
        # ExSummary:Shows how to insert a field into a document using FieldType.

        print("TestUtil hasn't been done yet")

        for update_inserted_fields_immediately in (False, True):
            with self.subTest(update_inserted_fields_immediately=update_inserted_fields_immediately):
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Insert two fields while passing a flag which determines whether to update them as the builder inserts them.
                # In some cases, updating fields could be computationally expensive, and it may be a good idea to defer the update.
                doc.built_in_document_properties.author = "John Doe"
                builder.write("This document was written by ")
                builder.insert_field(aw.fields.FieldType.FIELD_AUTHOR, update_inserted_fields_immediately)

                builder.insert_paragraph()
                builder.write("\nThis is page ")
                builder.insert_field(aw.fields.FieldType.FIELD_PAGE, update_inserted_fields_immediately)

                self.assertEqual(" AUTHOR ", doc.range.fields[0].get_field_code())
                self.assertEqual(" PAGE ", doc.range.fields[1].get_field_code())

                if update_inserted_fields_immediately:
                    self.assertEqual("John Doe", doc.range.fields[0].result)
                    self.assertEqual("1", doc.range.fields[1].result)
                else:
                    self.assertEqual("", doc.range.fields[0].result)
                    self.assertEqual("", doc.range.fields[1].result)

                    # We will need to update these fields using the update methods manually.
                    doc.range.fields[0].update()

                    self.assertEqual("John Doe", doc.range.fields[0].result)

                    doc.update_fields()

                    self.assertEqual("1", doc.range.fields[1].result)

                # ExEnd

                doc = DocumentHelper.save_open(doc)

                self.assertEqual("This document was written by \u0013 AUTHOR \u0014John Doe\u0015" +
                                 "\r\rThis is page \u0013 PAGE \u00141\u0015", doc.get_text().strip())

                # TestUtil.verify_field(FieldType.field_author, " AUTHOR ", "John Doe", doc.range.fields[0])
                # TestUtil.verify_field(FieldType.field_page, " PAGE ", "1", doc.range.fields[1])

    # def test_field_result_formatting(self):
    #     # ExStart
    #     # ExFor:IFieldResultFormatter
    #     # ExFor:IFieldResultFormatter.format(Double, GeneralFormat)
    #     # ExFor:IFieldResultFormatter.format(String, GeneralFormat)
    #     # ExFor:IFieldResultFormatter.format_date_time(DateTime, String, CalendarType)
    #     # ExFor:IFieldResultFormatter.format_numeric(Double, String)
    #     # ExFor:FieldOptions.result_formatter
    #     # ExFor:CalendarType
    #     # ExSummary:Shows how to automatically apply a custom format to field results as the fields are updated.
    #     doc = aw.Document()
    #     builder = aw.DocumentBuilder(doc)
    #     formatter = FieldResultFormatter("$0", "Date: 0", "Item # 0:")
    #     doc.field_options.result_formatter = formatter
    #
    #     # Our field result formatter applies a custom format to newly created fields of three types of formats.
    #     # Field result formatters apply new formatting to fields as they are updated,
    #     # which happens as soon as we create them using this InsertField method overload.
    #     # 1 -  Numeric:
    #     builder.insert_field(" = 2 + 3 \\# $###")
    #
    #     self.assertEqual("$5", doc.range.fields[0].result)
    #     self.assertEqual(1, formatter.count_format_invocations(FieldResultFormatter.format_invocation_type.numeric))
    #
    #     # 2 -  Date/time:
    #     builder.insert_field("DATE \\@ \"d MMMM yyyy\"")
    #
    #     self.assertTrue(doc.range.fields[1].result.starts_with("Date: "))
    #     self.assertEqual(1, formatter.count_format_invocations(FieldResultFormatter.format_invocation_type.date_time))
    #
    #     # 3 -  General:
    #     builder.insert_field("QUOTE \"2\" \\* Ordinal")
    #
    #     self.assertEqual("Item # 2:", doc.range.fields[2].result)
    #     self.assertEqual(1, formatter.count_format_invocations(FieldResultFormatter.format_invocation_type.general))
    #
    #     formatter.print_format_invocations()
    #
    #
    # #/ <summary>
    # #/ When fields with formatting are updated, this formatter will override their formatting
    # #/ with a custom format, while tracking every invocation.
    # #/ </summary>
    # private class FieldResultFormatter : IFieldResultFormatter
    #
    #     public FieldResultFormatter(string numberFormat, string dateFormat, string generalFormat)
    #
    #         mNumberFormat = numberFormat
    #         mDateFormat = dateFormat
    #         mGeneralFormat = generalFormat
    #
    #
    #     public string FormatNumeric(double value, string format)
    #
    #         if (string.is_null_or_empty(mNumberFormat))
    #             return null
    #
    #         string newValue = String.format(mNumberFormat, value)
    #         FormatInvocations.add(new FormatInvocation(FormatInvocationType.numeric, value, format, newValue))
    #         return newValue
    #
    #
    #     public string FormatDateTime(DateTime value, string format, CalendarType calendarType)
    #
    #         if (string.is_null_or_empty(mDateFormat))
    #             return null
    #
    #         string newValue = String.format(mDateFormat, value)
    #         FormatInvocations.add(new FormatInvocation(FormatInvocationType.date_time, $"value (calendarType)", format, newValue))
    #         return newValue
    #
    #
    #     public string Format(string value, GeneralFormat format)
    #
    #         return Format((object)value, format)
    #
    #
    #     public string Format(double value, GeneralFormat format)
    #
    #         return Format((object)value, format)
    #
    #
    #     private string Format(object value, GeneralFormat format)
    #
    #         if (string.is_null_or_empty(mGeneralFormat))
    #             return null
    #
    #         string newValue = String.format(mGeneralFormat, value)
    #         FormatInvocations.add(new FormatInvocation(FormatInvocationType.general, value, format.to_string(), newValue))
    #         return newValue
    #
    #
    #     public int CountFormatInvocations(FormatInvocationType formatInvocationType)
    #
    #         if (formatInvocationType == FormatInvocationType.all)
    #             return FormatInvocations.count
    #
    #         return FormatInvocations.count(f => f.format_invocation_type == formatInvocationType)
    #
    #
    #     public void PrintFormatInvocations()
    #
    #         for (FormatInvocation f in FormatInvocations)
    #             Console.write_line($"Invocation type:\tf.format_invocation_type\n" +
    #                                 $"\tOriginal value:\t\tf.value\n" +
    #                                 $"\tOriginal format:\tf.original_format\n" +
    #                                 $"\tNew value:\t\t\tf.new_value\n")
    #
    #
    #     private readonly string mNumberFormat
    #     private readonly string mDateFormat
    #     private readonly string mGeneralFormat
    #     private List<FormatInvocation> FormatInvocations  get  = new List<FormatInvocation>()
    #
    #     private class FormatInvocation
    #
    #         public FormatInvocationType FormatInvocationType  get
    #         public object Value  get
    #         public string OriginalFormat  get
    #         public string NewValue  get
    #
    #         public FormatInvocation(FormatInvocationType formatInvocationType, object value, string originalFormat, string newValue)
    #
    #             Value = value
    #             FormatInvocationType = formatInvocationType
    #             OriginalFormat = originalFormat
    #             NewValue = newValue
    #
    #
    #
    #     public enum FormatInvocationType
    #
    #         Numeric, DateTime, General, All
    #
    #
    # #ExEnd
    #
    # [Test, Ignore("Failed")]
    # public void InsertVideoWithUrl()
    #
    #     #ExStart
    #     #ExFor:DocumentBuilder.insert_online_video(String, Double, Double)
    #     #ExSummary:Shows how to insert an online video into a document using a URL.
    #     doc = aw.Document()
    #     builder = aw.DocumentBuilder(doc)
    #
    #     builder.insert_online_video("https:#youtu.be/t_1LYZ102RA", 360, 270)
    #
    #     # We can watch the video from Microsoft Word by clicking on the shape.
    #     doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_video_with_url.docx")
    #     #ExEnd
    #
    #     doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_video_with_url.docx")
    #     Shape shape = (Shape)doc.get_child(NodeType.shape, 0, True)
    #
    #     TestUtil.verify_image_in_shape(480, 360, ImageType.jpeg, shape)
    #     TestUtil.verify_web_response_status_code(HttpStatusCode.ok, shape.h_ref)
    #
    #     self.assertEqual(360.0d, shape.width)
    #     self.assertEqual(270.0d, shape.height)
    #

    def test_insert_underline(self):

        # ExStart
        # ExFor:DocumentBuilder.underline
        # ExSummary:Shows how to format text inserted by a document builder.

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.underline = aw.Underline.DASH
        builder.font.color = drawing.Color.blue
        builder.font.size = 32

        # The builder applies formatting to its current paragraph and any new text added by it afterward.
        builder.writeln("Large, blue, and underlined text.")

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_underline.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_underline.docx")
        first_run = doc.first_section.body.first_paragraph.runs[0]

        self.assertEqual("Large, blue, and underlined text.", first_run.get_text().strip())
        self.assertEqual(aw.Underline.DASH, first_run.font.underline)
        self.assertEqual(drawing.Color.blue.to_argb(), first_run.font.color.to_argb())
        self.assertEqual(32.0, first_run.font.size)

    def test_current_story(self):

        # ExStart
        # ExFor:DocumentBuilder.current_story
        # ExSummary:Shows how to work with a document builder's current story.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # A Story is a type of node that has child Paragraph nodes, such as a Body.
        self.assertEqual(builder.current_story, doc.first_section.body)
        self.assertEqual(builder.current_story, builder.current_paragraph.parent_node)
        self.assertEqual(aw.StoryType.MAIN_TEXT, builder.current_story.story_type)

        builder.current_story.append_paragraph("Text added to current Story.")

        # A Story can also contain tables.
        table = builder.start_table()
        builder.insert_cell()
        builder.write("Row 1, cell 1")
        builder.insert_cell()
        builder.write("Row 1, cell 2")
        builder.end_table()

        self.assertTrue(builder.current_story.tables.contains(table))
        # ExEnd

        doc = DocumentHelper.save_open(doc)
        self.assertEqual(1, doc.first_section.body.tables.count)
        self.assertEqual("Row 1, cell 1\aRow 1, cell 2\a\a\rText added to current Story.",
                         doc.first_section.body.get_text().strip())

    @unittest.skip("Unsupported file format (?), webclient class is not supported")
    def test_insert_ole_objects(self):

        # ExStart
        # ExFor:DocumentBuilder.insert_ole_object(Stream, String, Boolean, Stream)
        # ExSummary:Shows how to use document builder to embed OLE objects in a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a Microsoft Excel spreadsheet from the local file system
        # into the document while keeping its default appearance.


        # doc_spreadsheet = aw.Document(aeb.my_dir + "Spreadsheet.xlsx")
        # outStream = io.BytesIO()
        # doc_spreadsheet.save(outStream, aw.SaveFormat.XLSX)
        #
        # docBytes = outStream.getbuffer()
        # inStream = io.BytesIO(docBytes)
        #
        # spreadsheetStream = aw.Document(inStream)
        # builder.writeln("Spreadsheet Ole object:")
        # # If 'presentation' is omitted and 'asIcon' is set, this overloaded method selects
        # # the icon according to 'progId' and uses the predefined icon caption.
        # builder.insert_ole_object(spreadsheetStream, "OleObject.xlsx", False, None)


        # # Insert a Microsoft Powerpoint presentation as an OLE object.
        # # This time, it will have an image downloaded from the web for an icon.
        # using (Stream powerpointStream = File.open(aeb.my_dir + "Presentation.pptx", FileMode.open))
        #
        #     using (WebClient webClient = new WebClient())
        #
        #         byte[] imgBytes = File.read_all_bytes(aeb.ImageDir + "Logo.jpg")
        #
        #         using (MemoryStream imageStream = new MemoryStream(imgBytes))
        #
        #             builder.insert_paragraph()
        #             builder.writeln("Powerpoint Ole object:")
        #             builder.insert_ole_object(powerpointStream, "OleObject.pptx", True, imageStream)




        # # Double-click these objects in Microsoft Word to open
        # # the linked files using their respective applications.
        # doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_ole_objects.docx")
        # #ExEnd
        #
        # doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_ole_objects.docx")
        #
        # self.assertEqual(2, doc.get_child_nodes(NodeType.shape, True).count)
        #
        # Shape shape = (Shape)doc.get_child(NodeType.shape, 0, True)
        # self.assertEqual("", shape.ole_format.icon_caption)
        # self.assertFalse(shape.ole_format.ole_icon)
        #
        # shape = (Shape)doc.get_child(NodeType.shape, 1, True)
        # self.assertEqual("Unknown", shape.ole_format.icon_caption)
        # self.assertTrue(shape.ole_format.ole_icon)

    # @unittest.skip("No such enumerator StyleIdentifier.heading_1 (line 2842)")
    def test_insert_style_separator(self):

        # ExStart
        # ExFor:DocumentBuilder.insert_style_separator
        # ExSummary:Shows how to work with style separators.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Each paragraph can only have one style.
        # The InsertStyleSeparator method allows us to work around this limitation.
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
        builder.write("This text is in a Heading style. ")
        builder.insert_style_separator()

        para_style = builder.document.styles.add(aw.StyleType.PARAGRAPH, "MyParaStyle")
        para_style.font.bold = False
        para_style.font.size = 8
        para_style.font.name = "Arial"

        builder.paragraph_format.style_name = para_style.name
        builder.write("This text is in a custom style. ")

        # Calling the InsertStyleSeparator method creates another paragraph,
        # which can have a different style to the previous. There will be no break between paragraphs.
        # The text in the output document will look like one paragraph with two styles.
        self.assertEqual(2, doc.first_section.body.paragraphs.count)
        self.assertEqual("Heading 1", doc.first_section.body.paragraphs[0].paragraph_format.style.name)
        self.assertEqual("MyParaStyle", doc.first_section.body.paragraphs[1].paragraph_format.style.name)

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_style_separator.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_style_separator.docx")

        self.assertEqual(2, doc.first_section.body.paragraphs.count)
        self.assertEqual("This text is in a Heading style. \r This text is in a custom style.",
                         doc.get_text().strip())
        self.assertEqual("Heading 1", doc.first_section.body.paragraphs[0].paragraph_format.style.name)
        self.assertEqual("MyParaStyle", doc.first_section.body.paragraphs[1].paragraph_format.style.name)
        self.assertEqual(" ", doc.first_section.body.paragraphs[1].runs[0].get_text())
        # TestUtil.doc_package_file_contains_string("w:rPr><w:vanish /><w:specVanish /></w:rPr>",
        #                                           aeb.artifacts_dir + "DocumentBuilder.insert_style_separator.docx",
        #                                           "document.xml")
        # TestUtil.doc_package_file_contains_string("<w:t xml:space=\"preserve\"> </w:t>",
        #                                           aeb.artifacts_dir + "DocumentBuilder.insert_style_separator.docx",
        #                                           "document.xml")

    @unittest.skip("Skipped in .NET tests: "
                   "Bug: does not insert headers and footers, all lists (bullets, numbering, multilevel) breaks")
    def test_insert_document(self):

        # ExStart
        # ExFor:DocumentBuilder.insert_document(Document, ImportFormatMode)
        # ExFor:ImportFormatMode
        # ExSummary:Shows how to insert a document into another document.
        doc = aw.Document(aeb.my_dir + "Document.docx")

        builder = aw.DocumentBuilder(doc)
        builder.move_to_document_end()
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        docToInsert = aw.Document(aeb.my_dir + "Formatted elements.docx")

        builder.insert_document(docToInsert, ImportFormatMode.keep_source_formatting)
        builder.document.save(aeb.artifacts_dir + "DocumentBuilder.insert_document.docx")
        # ExEnd

        self.assertEqual(29, doc.styles.count)
        self.assertTrue(DocumentHelper.compare_docs(aeb.artifacts_dir + "DocumentBuilder.insert_document.docx",
                                                    aeb.golds_dir + "DocumentBuilder.insert_document Gold.docx"))

    @unittest.skip("Item properties can use only int")
    def test_smart_style_behavior(self):
        # ExStart
        # ExFor:ImportFormatOptions
        # ExFor:ImportFormatOptions.smart_style_behavior
        # ExFor:DocumentBuilder.insert_document(Document, ImportFormatMode, ImportFormatOptions)
        # ExSummary:Shows how to resolve duplicate styles while inserting documents.

        dst_doc = aw.Document()
        builder = aw.DocumentBuilder(dst_doc)

        my_style = builder.document.styles.add(aw.StyleType.PARAGRAPH, "MyStyle")
        my_style.font.size = 14
        my_style.font.name = "Courier New"
        my_style.font.color = drawing.Color.blue

        builder.paragraph_format.style_name = my_style.name
        builder.writeln("Hello world!")

        # Clone the document and edit the clone's "my_style" style, so it is a different color than that of the original.
        # If we insert the clone into the original document, the two styles with the same name will cause a clash.
        src_doc = dst_doc.clone()
        src_doc.styles["MyStyle"].font.color = drawing.Color.red

        # When we enable SmartStyleBehavior and use the KeepSourceFormatting import format mode,
        # Aspose.words will resolve style clashes by converting source document styles.
        # with the same names as destination styles into direct paragraph attributes.
        options = aw.ImportFormatOptions()
        options.smart_style_behavior = True

        builder.insert_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, options)

        dst_doc.save(aeb.artifacts_dir + "DocumentBuilder.smart_style_behavior.docx")
        # ExEnd

        dst_doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.smart_style_behavior.docx")

        self.assertEqual(drawing.Color.blue.to_argb(), dst_doc.styles["MyStyle"].font.color.to_argb())
        self.assertEqual("MyStyle", dst_doc.first_section.body.paragraphs[0].paragraph_format.style.name)

        self.assertEqual("Normal", dst_doc.first_section.body.paragraphs[1].paragraph_format.style.name)
        self.assertEqual(14, dst_doc.first_section.body.paragraphs[1].runs[0].font.size)
        self.assertEqual("Courier New", dst_doc.first_section.body.paragraphs[1].runs[0].font.name)
        self.assertEqual(drawing.Color.red.to_argb(), dst_doc.first_section.body.paragraphs[1].runs[0].font.color.to_argb())

    def test_emphases_warning_source_markdown(self):

        doc = aw.Document(aeb.my_dir + "Emphases markdown warning.docx")

        warnings = aw.WarningInfoCollection()
        doc.warning_callback = warnings
        doc.save(aeb.artifacts_dir + "DocumentBuilder.emphases_warning_source_markdown.md")

        for warning_info in warnings:
            if warning_info.source == aw.WarningSource.MARKDOWN:
                self.assertEqual("The (*, 0:11) cannot be properly written into Markdown.", warning_info.description)

    def test_do_not_ignore_header_footer(self):

        # ExStart
        # ExFor:ImportFormatOptions.ignore_header_footer
        # ExSummary:Shows how to specifies ignoring or not source formatting of headers/footers content.
        dst_doc = aw.Document(aeb.my_dir + "Document.docx")
        src_doc = aw.Document(aeb.my_dir + "Header and footer types.docx")

        import_format_options = aw.ImportFormatOptions()
        import_format_options.ignore_header_footer = False

        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, import_format_options)

        dst_doc.save(aeb.artifacts_dir + "DocumentBuilder.do_not_ignore_header_footer.docx")
        # ExEnd

    """
    #if NET462 || NETCOREAPP2_1 || JAVA
    #/ <summary>
    #/ All markdown tests work with the same file. That's why we need order for them.
    #/ </summary>
    [Test, Order(1)]
    public void MarkdownDocumentEmphases()
        
        DocumentBuilder builder = aw.DocumentBuilder()
            
        # Bold and Italic are represented as Font.bold and Font.italic.
        builder.font.italic = True
        builder.writeln("This text will be italic")
            
        # Use clear formatting if we don't want to combine styles between paragraphs.
        builder.font.clear_formatting()
            
        builder.font.bold = True
        builder.writeln("This text will be bold")
            
        builder.font.clear_formatting()
            
        builder.font.italic = True
        builder.write("You ")
        builder.font.bold = True
        builder.write("can")
        builder.font.bold = False
        builder.writeln(" combine them")

        builder.font.clear_formatting()

        builder.font.strike_through = True
        builder.writeln("This text will be strikethrough")
            
        # Markdown treats asterisks (*), underscores (_) and tilde (~) as indicators of emphasis.
        builder.document.save(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        

    #/ <summary>
    #/ All markdown tests work with the same file. That's why we need order for them.
    #/ </summary>
    [Test, Order(2)]
    public void MarkdownDocumentInlineCode()
        
        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        builder = aw.DocumentBuilder(doc)
            
        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.paragraph_format.clear_formatting()
        builder.writeln("\n")
            
        # Style with name that starts from word InlineCode, followed by optional dot (.) and number of backticks (`).
        # If number of backticks is missed, then one backtick will be used by default.
        Style inlineCode1BackTicks = doc.styles.add(StyleType.character, "InlineCode")
        builder.font.style = inlineCode1BackTicks
        builder.writeln("Text with InlineCode style with one backtick")
            
        # Use optional dot (.) and number of backticks (`).
        # There will be 3 backticks.
        Style inlineCode3BackTicks = doc.styles.add(StyleType.character, "InlineCode.3")
        builder.font.style = inlineCode3BackTicks
        builder.writeln("Text with InlineCode style with 3 backticks")

        builder.document.save(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        

    #/ <summary>
    #/ All markdown tests work with the same file. That's why we need order for them.
    #/ </summary>
    [Test, Order(3)]
    [Description("WORDSNET-19850")]
    public void MarkdownDocumentHeadings()
        
        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        builder = aw.DocumentBuilder(doc)

        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.paragraph_format.clear_formatting()
        builder.writeln("\n")
            
        # By default, Heading styles in Word may have bold and italic formatting.
        # If we do not want text to be emphasized, set these properties explicitly to False.
        # Thus we can't use 'builder.font.clear_formatting()' because Bold/Italic will be set to True.
        builder.font.bold = False
        builder.font.italic = False
            
        # Create for one heading for each level.
        builder.paragraph_format.style_name = "Heading 1"
        builder.font.italic = True
        builder.writeln("This is an italic H1 tag")

        # Reset our styles from the previous paragraph to not combine styles between paragraphs.
        builder.font.bold = False
        builder.font.italic = False

        # Structure-enhanced text heading can be added through style inheritance.
        Style setextHeading1 = doc.styles.add(StyleType.paragraph, "SetextHeading1")
        builder.paragraph_format.style = setextHeading1
        doc.styles["SetextHeading1"].base_style_name = "Heading 1"
        builder.writeln("SetextHeading 1")
            
        builder.paragraph_format.style_name = "Heading 2"
        builder.writeln("This is an H2 tag")

        builder.font.bold = False
        builder.font.italic = False

        Style setextHeading2 = doc.styles.add(StyleType.paragraph, "SetextHeading2")
        builder.paragraph_format.style = setextHeading2
        doc.styles["SetextHeading2"].base_style_name = "Heading 2"
        builder.writeln("SetextHeading 2")
            
        builder.paragraph_format.style = doc.styles["Heading 3"]
        builder.writeln("This is an H3 tag")
            
        builder.font.bold = False
        builder.font.italic = False

        builder.paragraph_format.style = doc.styles["Heading 4"]
        builder.font.bold = True
        builder.writeln("This is an bold H4 tag")
            
        builder.font.bold = False
        builder.font.italic = False

        builder.paragraph_format.style = doc.styles["Heading 5"]
        builder.font.italic = True
        builder.font.bold = True
        builder.writeln("This is an italic and bold H5 tag")
            
        builder.font.bold = False
        builder.font.italic = False

        builder.paragraph_format.style = doc.styles["Heading 6"]
        builder.writeln("This is an H6 tag")
            
        doc.save(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        

    #/ <summary>
    #/ All markdown tests work with the same file. That's why we need order for them.
    #/ </summary>
    [Test, Order(4)]
    public void MarkdownDocumentBlockquotes()
        
        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        builder = aw.DocumentBuilder(doc)

        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.paragraph_format.clear_formatting()
        builder.writeln("\n")

        # By default, the document stores blockquote style for the first level.
        builder.paragraph_format.style_name = "Quote"
        builder.writeln("Blockquote")
            
        # Create styles for nested levels through style inheritance.
        Style quoteLevel2 = doc.styles.add(StyleType.paragraph, "Quote1")
        builder.paragraph_format.style = quoteLevel2
        doc.styles["Quote1"].base_style_name = "Quote"
        builder.writeln("1. Nested blockquote")
            
        Style quoteLevel3 = doc.styles.add(StyleType.paragraph, "Quote2")
        builder.paragraph_format.style = quoteLevel3
        doc.styles["Quote2"].base_style_name = "Quote1"
        builder.font.italic = True
        builder.writeln("2. Nested italic blockquote")
            
        Style quoteLevel4 = doc.styles.add(StyleType.paragraph, "Quote3")
        builder.paragraph_format.style = quoteLevel4
        doc.styles["Quote3"].base_style_name = "Quote2"
        builder.font.italic = False
        builder.font.bold = True
        builder.writeln("3. Nested bold blockquote")
            
        Style quoteLevel5 = doc.styles.add(StyleType.paragraph, "Quote4")
        builder.paragraph_format.style = quoteLevel5
        doc.styles["Quote4"].base_style_name = "Quote3"
        builder.font.bold = False
        builder.writeln("4. Nested blockquote")
            
        Style quoteLevel6 = doc.styles.add(StyleType.paragraph, "Quote5")
        builder.paragraph_format.style = quoteLevel6
        doc.styles["Quote5"].base_style_name = "Quote4"
        builder.writeln("5. Nested blockquote")
            
        Style quoteLevel7 = doc.styles.add(StyleType.paragraph, "Quote6")
        builder.paragraph_format.style = quoteLevel7
        doc.styles["Quote6"].base_style_name = "Quote5"
        builder.font.italic = True
        builder.font.bold = True
        builder.writeln("6. Nested italic bold blockquote")
            
        doc.save(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        

    #/ <summary>
    #/ All markdown tests work with the same file. That's why we need order for them.
    #/ </summary>
    [Test, Order(5)]
    public void MarkdownDocumentIndentedCode()
        
        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        builder = aw.DocumentBuilder(doc)

        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.writeln("\n")
        builder.paragraph_format.clear_formatting()
        builder.writeln("\n")

        Style indentedCode = doc.styles.add(StyleType.paragraph, "IndentedCode")
        builder.paragraph_format.style = indentedCode
        builder.writeln("This is an indented code")
            
        doc.save(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        

    #/ <summary>
    #/ All markdown tests work with the same file. That's why we need order for them.
    #/ </summary>
    [Test, Order(6)]
    public void MarkdownDocumentFencedCode()
        
        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        builder = aw.DocumentBuilder(doc)

        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.writeln("\n")
        builder.paragraph_format.clear_formatting()
        builder.writeln("\n")

        Style fencedCode = doc.styles.add(StyleType.paragraph, "FencedCode")
        builder.paragraph_format.style = fencedCode
        builder.writeln("This is a fenced code")

        Style fencedCodeWithInfo = doc.styles.add(StyleType.paragraph, "FencedCode.c#")
        builder.paragraph_format.style = fencedCodeWithInfo
        builder.writeln("This is a fenced code with info string")

        doc.save(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        

    #/ <summary>
    #/ All markdown tests work with the same file. That's why we need order for them.
    #/ </summary>
    [Test, Order(7)]
    public void MarkdownDocumentHorizontalRule()
        
        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        builder = aw.DocumentBuilder(doc)

        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.paragraph_format.clear_formatting()
        builder.writeln("\n")

        # Insert HorizontalRule that will be present in .md file as '-----'.
        builder.insert_horizontal_rule()
 
        builder.document.save(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        

    #/ <summary>
    #/ All markdown tests work with the same file. That's why we need order for them.
    #/ </summary>
    [Test, Order(8)]
    public void MarkdownDocumentBulletedList()
        
        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        builder = aw.DocumentBuilder(doc)

        # Prepare our created document for further work
        # and clear paragraph formatting not to use the previous styles.
        builder.move_to_document_end()
        builder.paragraph_format.clear_formatting()
        builder.writeln("\n")

        # Bulleted lists are represented using paragraph numbering.
        builder.list_format.apply_bullet_default()
        # There can be 3 types of bulleted lists.
        # The only diff in a numbering format of the very first level are ‘-’, ‘+’ or ‘*’ respectively.
        builder.list_format.list.list_levels[0].number_format = "-"
            
        builder.writeln("Item 1")
        builder.writeln("Item 2")
        builder.list_format.list_indent()
        builder.writeln("Item 2a")
        builder.writeln("Item 2b")
 
        builder.document.save(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        

    #/ <summary>
    #/ All markdown tests work with the same file. That's why we need order for them.
    #/ </summary>
    [Test, Order(9)]
    [TestCase("Italic", "Normal", True, False)]
    [TestCase("Bold", "Normal", False, True)]
    [TestCase("ItalicBold", "Normal", True, True)]
    [TestCase("Text with InlineCode style with one backtick", "InlineCode", False, False)]
    [TestCase("Text with InlineCode style with 3 backticks", "InlineCode.3", False, False)]
    [TestCase("This is an italic H1 tag", "Heading 1", True, False)]
    [TestCase("SetextHeading 1", "SetextHeading1", False, False)]
    [TestCase("This is an H2 tag", "Heading 2", False, False)]
    [TestCase("SetextHeading 2", "SetextHeading2", False, False)]
    [TestCase("This is an H3 tag", "Heading 3", False, False)]
    [TestCase("This is an bold H4 tag", "Heading 4", False, True)]
    [TestCase("This is an italic and bold H5 tag", "Heading 5", True, True)]
    [TestCase("This is an H6 tag", "Heading 6", False, False)]
    [TestCase("Blockquote", "Quote", False, False)]
    [TestCase("1. Nested blockquote", "Quote1", False, False)]
    [TestCase("2. Nested italic blockquote", "Quote2", True, False)]
    [TestCase("3. Nested bold blockquote", "Quote3", False, True)]
    [TestCase("4. Nested blockquote", "Quote4", False, False)]
    [TestCase("5. Nested blockquote", "Quote5", False, False)]
    [TestCase("6. Nested italic bold blockquote", "Quote6", True, True)]
    [TestCase("This is an indented code", "IndentedCode", False, False)]
    [TestCase("This is a fenced code", "FencedCode", False, False)]
    [TestCase("This is a fenced code with info string", "FencedCode.c#", False, False)]
    [TestCase("Item 1", "Normal", False, False)]
    public void LoadMarkdownDocumentAndAssertContent(string text, string styleName, bool isItalic, bool isBold)
        
        # Load created document from previous tests.
        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.markdown_document.md")
        ParagraphCollection paragraphs = doc.first_section.body.paragraphs
            
        for (Paragraph paragraph in paragraphs)
            
            if (paragraph.runs.count != 0)
                
                # Check that all document text has the necessary styles.
                if (paragraph.runs[0].text == text && !text.contains("InlineCode"))
                    
                    self.assertEqual(styleName, paragraph.paragraph_format.style.name)
                    self.assertEqual(isItalic, paragraph.runs[0].font.italic)
                    self.assertEqual(isBold, paragraph.runs[0].font.bold)
                    
                else if (paragraph.runs[0].text == text && text.contains("InlineCode"))
                    
                    self.assertEqual(styleName, paragraph.runs[0].font.style_name)
                    
                

            # Check that document also has a HorizontalRule present as a shape.
            NodeCollection shapesCollection = doc.first_section.body.get_child_nodes(NodeType.shape, True)
            Shape horizontalRuleShape = (Shape) shapesCollection[0]
                
            self.assertTrue(shapesCollection.count == 1)
            self.assertTrue(horizontalRuleShape.is_horizontal_rule)
            
        """

    def test_markdown_document_table_content_alignment(self):

        for table_content_alignment in (aw.saving.TableContentAlignment.LEFT,
                                        aw.saving.TableContentAlignment.RIGHT,
                                        aw.saving.TableContentAlignment.CENTER,
                                        aw.saving.TableContentAlignment.AUTO):
            with self.subTest(table_content_alignment=table_content_alignment):

                builder = aw.DocumentBuilder()

                builder.insert_cell()
                builder.paragraph_format.alignment = aw.ParagraphAlignment.RIGHT
                builder.write("Cell1")
                builder.insert_cell()
                builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
                builder.write("Cell2")

                save_options = aw.saving.MarkdownSaveOptions()
                save_options.table_content_alignment = table_content_alignment

                builder.document.save(
                    aeb.artifacts_dir + "DocumentBuilder.markdown_document_table_content_alignment.md", save_options)

                doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.markdown_document_table_content_alignment.md")
                table = doc.first_section.body.tables[0]

                if table_content_alignment == table_content_alignment.AUTO:
                    self.assertEqual(aw.ParagraphAlignment.RIGHT,
                                     table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                    self.assertEqual(aw.ParagraphAlignment.CENTER,
                                     table.first_row.cells[1].first_paragraph.paragraph_format.alignment)

                if table_content_alignment == table_content_alignment.LEFT:
                    self.assertEqual(aw.ParagraphAlignment.LEFT,
                                     table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                    self.assertEqual(aw.ParagraphAlignment.LEFT,
                                     table.first_row.cells[1].first_paragraph.paragraph_format.alignment)

                if table_content_alignment == table_content_alignment.CENTER:
                    self.assertEqual(aw.ParagraphAlignment.CENTER,
                                     table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                    self.assertEqual(aw.ParagraphAlignment.CENTER,
                                     table.first_row.cells[1].first_paragraph.paragraph_format.alignment)

                if table_content_alignment == table_content_alignment.RIGHT:
                    self.assertEqual(aw.ParagraphAlignment.RIGHT,
                                     table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                    self.assertEqual(aw.ParagraphAlignment.RIGHT,
                                     table.first_row.cells[1].first_paragraph.paragraph_format.alignment)


    # def test_rename_images(self):
    #     # ExStart
    #     # ExFor:MarkdownSaveOptions.image_saving_callback
    #     # ExFor:IImageSavingCallback
    #     # ExSummary:Shows how to rename the image name during saving into Markdown document.
    #     doc = aw.Document(aeb.my_dir + "Rendering.docx")
    #
    #     options = aw.MarkdownSaveOptions()
    #
    #     # If we convert a document that contains images into Markdown, we will end up with one Markdown file which links to several images.
    #     # Each image will be in the form of a file in the local file system.
    #     # There is also a callback that can customize the name and file system location of each image.
    #     options.image_saving_callback = SavedImageRename("DocumentBuilder.handle_document.md")
    #
    #     # The ImageSaving() method of our callback will be run at this time.
    #     doc.save(aeb.artifacts_dir + "DocumentBuilder.handle_document.md", options)
    #
    #     self.assertEqual(1,
    #         Directory.get_files(aeb.artifacts_dir)
    #             .where(s => s.starts_with(aeb.artifacts_dir + "DocumentBuilder.handle_document.md shape"))
    #             .count(f => f.ends_with(".jpeg")))
    #     self.assertEqual(8,
    #         Directory.get_files(aeb.artifacts_dir)
    #             .where(s => s.starts_with(aeb.artifacts_dir + "DocumentBuilder.handle_document.md shape"))
    #             .count(f => f.ends_with(".png")))
    #
    #
    # #/ <summary>
    # #/ Renames saved images that are produced when an Markdown document is saved.
    # #/ </summary>
    # public class SavedImageRename : IImageSavingCallback
    #
    #     public SavedImageRename(string outFileName)
    #
    #         mOutFileName = outFileName
    #
    #
    #     void IImageSavingCallback.image_saving(ImageSavingArgs args)
    #
    #         string imageFileName = $"mOutFileName shape ++mCount, of type args.current_shape.shape_type_path.get_extension(args.image_file_name)"
    #
    #         args.image_file_name = imageFileName
    #         args.image_stream = new FileStream(aeb.artifacts_dir + imageFileName, FileMode.create)
    #
    #         self.assertTrue(args.image_stream.can_write)
    #         self.assertTrue(args.is_image_available)
    #         self.assertFalse(args.keep_image_stream_open)
    #
    #
    #     private int mCount
    #     private readonly string mOutFileName
    #
    # #ExEnd

    # @unittest.skip("No type casting (line 3467), testUtil hadn't been done yet, no such property as HttpStatusCode")
    def test_insert_online_video(self) :
        
        #ExStart
        #ExFor:DocumentBuilder.insert_online_video(String, aw.drawing.RelativeHorizontalPosition, Double, aw.drawing.RelativeVerticalPosition, Double, Double, Double, aw.drawing.WrapType)
        #ExSummary:Shows how to insert an online video into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
            
        video_url = "https://vimeo.com/52477838"

        # Insert a shape that plays a video from the web when clicked in Microsoft Word.
        # This rectangular shape will contain an image based on the first frame of the linked video
        # and a "play button" visual prompt. The video has an aspect ratio of 16:9.
        # We will set the shape's size to that ratio, so the image does not appear stretched.
        builder.insert_online_video(video_url, aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN, 0,
            aw.drawing.RelativeVerticalPosition.TOP_MARGIN, 0, 320, 180, aw.drawing.WrapType.SQUARE)

        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_online_video.docx")
        #ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_online_video.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True)
        shape = shape.as_shape()

        # TestUtil.verify_image_in_shape(640, 360, aw.ImageType.JPEG, shape)

        self.assertEqual(320.0, shape.width)
        self.assertEqual(180.0, shape.height)
        self.assertEqual(0.0, shape.left)
        self.assertEqual(0.0, shape.top)
        self.assertEqual(aw.drawing.WrapType.SQUARE, shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.TOP_MARGIN, shape.relative_vertical_position)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN, shape.relative_horizontal_position)

        self.assertEqual("https://vimeo.com/52477838", shape.href)
        # TestUtil.verify_web_response_status_code(HttpStatusCode.OK, shape.h_ref)


    # def test_insert_online_video_custom_thumbnail(self) :
    #
    #     #ExStart
    #     #ExFor:DocumentBuilder.insert_online_video(String, String, Byte[], Double, Double)
    #     #ExFor:DocumentBuilder.insert_online_video(String, String, Byte[], aw.drawing.RelativeHorizontalPosition, Double, aw.drawing.RelativeVerticalPosition, Double, Double, Double, aw.drawing.WrapType)
    #     #ExSummary:Shows how to insert an online video into a document with a custom thumbnail.
    #     doc = aw.Document()
    #     builder = aw.DocumentBuilder(doc)
    #
    #     video_url = "https:#vimeo.com/52477838"
    #     videoEmbedCode =
    #         "<iframe src=\"https:#player.vimeo.com/video/52477838\" width=\"640\" height=\"360\" frameborder=\"0\" " +
    #         "title=\"Aspose\" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>"
    #
    #     byte[] thumbnailImageBytes = File.read_all_bytes(aeb.image_dir + "Logo.jpg")
    #
    #     using (MemoryStream stream = new MemoryStream(thumbnailImageBytes))
    #
    #         using (Image image = Image.from_stream(stream))
    #
    #             # Below are two ways of creating a shape with a custom thumbnail, which links to an online video
    #             # that will play when we click on the shape in Microsoft Word.
    #             # 1 -  Insert an inline shape at the builder's node insertion cursor:
    #             builder.insert_online_video(video_url, videoEmbedCode, thumbnailImageBytes, image.width, image.height)
    #
    #             builder.insert_break(aw.BreakType.PAGE_BREAK)
    #
    #             # 2 -  Insert a floating shape:
    #             double left = builder.page_setup.right_margin - image.width
    #             double top = builder.page_setup.bottom_margin - image.height
    #
    #             builder.insert_online_video(video_url, videoEmbedCode, thumbnailImageBytes,
    #                 aw.drawing.RelativeHorizontalPosition.right_margin, left, aw.drawing.RelativeVerticalPosition.bottom_margin, top,
    #                 image.width, image.height, aw.drawing.WrapType.square)
    #
    #
    #
    #     doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_online_video_custom_thumbnail.docx")
    #     #ExEnd
    #
    #     doc = aw.Document(aeb.artifacts_dir + "DocumentBuilder.insert_online_video_custom_thumbnail.docx")
    #     Shape shape = (Shape) doc.get_child(NodeType.shape, 0, True)
    #
    #     TestUtil.verify_image_in_shape(400, 400, ImageType.jpeg, shape)
    #     self.assertEqual(400.0d, shape.width)
    #     self.assertEqual(400.0d, shape.height)
    #     self.assertEqual(0.0d, shape.left)
    #     self.assertEqual(0.0d, shape.top)
    #     self.assertEqual(aw.drawing.WrapType.inline, shape.wrap_type)
    #     self.assertEqual(aw.drawing.RelativeVerticalPosition.paragraph, shape.relative_vertical_position)
    #     self.assertEqual(aw.drawing.RelativeHorizontalPosition.column, shape.relative_horizontal_position)
    #
    #     self.assertEqual("https:#vimeo.com/52477838", shape.h_ref)
    #
    #     shape = (Shape) doc.get_child(NodeType.shape, 1, True)
    #
    #     TestUtil.verify_image_in_shape(400, 400, ImageType.jpeg, shape)
    #     self.assertEqual(400.0d, shape.width)
    #     self.assertEqual(400.0d, shape.height)
    #     self.assertEqual(-329.15d, shape.left)
    #     self.assertEqual(-329.15d, shape.top)
    #     self.assertEqual(WrapType.square, shape.wrap_type)
    #     self.assertEqual(aw.drawing.RelativeVerticalPosition.bottom_margin, shape.relative_vertical_position)
    #     self.assertEqual(RelativeHorizontalPosition.right_margin, shape.relative_horizontal_position)
    #
    #     self.assertEqual("https:#vimeo.com/52477838", shape.h_ref)
    #
    #     ServicePointManager.security_protocol = SecurityProtocolType.tls_12
    #     TestUtil.verify_web_response_status_code(HttpStatusCode.ok, shape.h_ref)
        

    # @unittest.skip("Streams are not supported")
    def test_insert_ole_object_as_icon(self) :

        #ExStart
        #ExFor:DocumentBuilder.insert_ole_object_as_icon(String, String, Boolean, String, String)
        #ExFor:DocumentBuilder.insert_ole_object_as_icon(Stream, String, String, String)
        #ExSummary:Shows how to insert an embedded or linked OLE object as icon into the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
        # the icon according to 'progId' and uses the filename for the icon caption.
        builder.insert_ole_object_as_icon(aeb.my_dir + "Presentation.pptx", "Package", False, aeb.image_dir + "Logo icon.ico", "My embedded file")

        builder.insert_break(aw.BreakType.LINE_BREAK)

        # using (FileStream stream = new FileStream(aeb.my_dir + "Presentation.pptx", FileMode.open))
        #
        #     # If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
        #     # the icon according to the file extension and uses the filename for the icon caption.
        #     Shape shape = builder.insert_ole_object_as_icon(stream, "PowerPoint.application", aeb.ImageDir + "Logo icon.ico",
        #         "My embedded file stream")
        #
        #     OlePackage setOlePackage = shape.ole_format.ole_package
        #     setOlePackage.file_name = "Presentation.pptx"
        #     setOlePackage.display_name = "Presentation.pptx"


        doc.save(aeb.artifacts_dir + "DocumentBuilder.insert_ole_object_as_icon.docx")
        #ExEnd


if __name__ == '__main__':
    unittest.main()
