import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithMarkdown(DocsExamplesBase):

    def test_create_markdown_document(self):

        #ExStart:CreateMarkdownDocument
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Specify the "Heading 1" style for the paragraph.
        builder.paragraph_format.style_name = "Heading 1"
        builder.writeln("Heading 1")

        # Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder.paragraph_format.style_name = "Normal"

        # Insert horizontal rule.
        builder.insert_horizontal_rule()

        # Specify the ordered list.
        builder.insert_paragraph()
        builder.list_format.apply_number_default()

        # Specify the Italic emphasis for the text.
        builder.font.italic = True
        builder.writeln("Italic Text")
        builder.font.italic = False

        # Specify the Bold emphasis for the text.
        builder.font.bold = True
        builder.writeln("Bold Text")
        builder.font.bold = False

        # Specify the StrikeThrough emphasis for the text.
        builder.font.strike_through = True
        builder.writeln("StrikeThrough Text")
        builder.font.strike_through = False

        # Stop paragraphs numbering.
        builder.list_format.remove_numbers()

        # Specify the "Quote" style for the paragraph.
        builder.paragraph_format.style_name = "Quote"
        builder.writeln("A Quote block")

        # Specify nesting Quote.
        nested_quote = doc.styles.add(aw.StyleType.PARAGRAPH, "Quote1")
        nested_quote.base_style_name = "Quote"
        builder.paragraph_format.style_name = "Quote1"
        builder.writeln("A nested Quote block")

        # Reset paragraph style to Normal to stop Quote blocks.
        builder.paragraph_format.style_name = "Normal"

        # Specify a Hyperlink for the desired text.
        builder.font.bold = True
        # Note, the text of hyperlink can be emphasized.
        builder.insert_hyperlink("Aspose", "https:#www.aspose.com", False)
        builder.font.bold = False

        # Insert a simple table.
        builder.start_table()
        builder.insert_cell()
        builder.write("Cell1")
        builder.insert_cell()
        builder.write("Cell2")
        builder.end_table()

        # Save your document as a Markdown file.
        doc.save(ARTIFACTS_DIR + "WorkingWithMarkdown.create_markdown_document.md")
        #ExEnd:CreateMarkdownDocument


    def test_read_markdown_document(self):

        #ExStart:ReadMarkdownDocument
        doc = aw.Document(MY_DIR + "Quotes.md")

        # Let's remove Heading formatting from a Quote in the very last paragraph.
        paragraph = doc.first_section.body.last_paragraph
        paragraph.paragraph_format.style = doc.styles.get_by_name("Quote")

        doc.save(ARTIFACTS_DIR + "WorkingWithMarkdown.read_markdown_document.md")
        #ExEnd:ReadMarkdownDocument


    def test_emphases(self):

        #ExStart:Emphases
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Markdown treats asterisks (*) and underscores (_) as indicators of emphasis.")
        builder.write("You can write ")

        builder.font.bold = True
        builder.write("bold")

        builder.font.bold = False
        builder.write(" or ")

        builder.font.italic = True
        builder.write("italic")

        builder.font.italic = False
        builder.writeln(" text. ")

        builder.write("You can also write ")
        builder.font.bold = True

        builder.font.italic = True
        builder.write("BoldItalic")

        builder.font.bold = False
        builder.font.italic = False
        builder.write("text.")

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.emphases.md")
        #ExEnd:Emphases


    def test_headings(self):

        #ExStart:Headings
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # By default Heading styles in Word may have bold and italic formatting.
        # If we do not want the text to be emphasized, set these properties explicitly to false.
        builder.font.bold = False
        builder.font.italic = False

        builder.writeln("The following produces headings:")
        builder.paragraph_format.style = doc.styles.get_by_name("Heading 1")
        builder.writeln("Heading1")
        builder.paragraph_format.style = doc.styles.get_by_name("Heading 2")
        builder.writeln("Heading2")
        builder.paragraph_format.style = doc.styles.get_by_name("Heading 3")
        builder.writeln("Heading3")
        builder.paragraph_format.style = doc.styles.get_by_name("Heading 4")
        builder.writeln("Heading4")
        builder.paragraph_format.style = doc.styles.get_by_name("Heading 5")
        builder.writeln("Heading5")
        builder.paragraph_format.style = doc.styles.get_by_name("Heading 6")
        builder.writeln("Heading6")

        # Note that the emphases are also allowed inside Headings.
        builder.font.bold = True
        builder.paragraph_format.style = doc.styles.get_by_name("Heading 1")
        builder.writeln("Bold Heading1")

        doc.save(ARTIFACTS_DIR + "WorkingWithMarkdown.headings.md")
        #ExEnd:Headings


    def test_block_quotes(self):

        #ExStart:BlockQuotes
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("We support blockquotes in Markdown:")

        builder.paragraph_format.style = doc.styles.get_by_name("Quote")
        builder.writeln("Lorem")
        builder.writeln("ipsum")

        builder.paragraph_format.style = doc.styles.get_by_name("Normal")
        builder.writeln("The quotes can be of any level and can be nested:")

        quote_level3 = doc.styles.add(aw.StyleType.PARAGRAPH, "Quote2")
        builder.paragraph_format.style = quote_level3
        builder.writeln("Quote level 3")

        quote_level4 = doc.styles.add(aw.StyleType.PARAGRAPH, "Quote3")
        builder.paragraph_format.style = quote_level4
        builder.writeln("Nested quote level 4")

        builder.paragraph_format.style = doc.styles.get_by_name("Quote")
        builder.writeln()
        builder.writeln("Back to first level")

        quote_level1_with_heading = doc.styles.add(aw.StyleType.PARAGRAPH, "Quote Heading 3")
        builder.paragraph_format.style = quote_level1_with_heading
        builder.write("Headings are allowed inside Quotes")

        doc.save(ARTIFACTS_DIR + "WorkingWithMarkdown.block_quotes.md")
        #ExEnd:BlockQuotes


    def test_horizontal_rule(self):

        #ExStart:HorizontalRule
        builder = aw.DocumentBuilder()

        builder.writeln("We support Horizontal rules (Thematic breaks) in Markdown:")
        builder.insert_horizontal_rule()

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.horizontal_rule_example.md")
        #ExEnd:HorizontalRule

    def test_bold_text(self):

        #ExStart:BoldText
        # Use a document builder to add content to the document.
        builder = aw.DocumentBuilder()

        # Make the text Bold.
        builder.font.bold = True
        builder.writeln("This text will be Bold")

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.bold_text_example.md")
        #ExEnd:BoldText

    def test_italic_text(self):

        #ExStart:ItalicText
        # Use a document builder to add content to the document.
        builder = aw.DocumentBuilder()

        # Make the text Italic.
        builder.font.italic = True
        builder.writeln("This text will be Italic")

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.italic_text_example.md")
        #ExEnd:ItalicText

    def test_strikethrough_text(self):

        #ExStart:Strikethrough
        # Use a document builder to add content to the document.
        builder = aw.DocumentBuilder()

        # Make the text Strikethrough.
        builder.font.strike_through = True
        builder.writeln("This text will be Strikethrough")

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.strikethrough_text_example.md")
        #ExEnd:Strikethrough

    def test_inline_code(self):

        #ExStart:InlineCode
        # Use a document builder to add content to the document.
        builder = aw.DocumentBuilder()

        # Number of backticks is missed, one backtick will be used by default.
        inline_code1_back_ticks = builder.document.styles.add(aw.StyleType.CHARACTER, "InlineCode")
        builder.font.style = inline_code1_back_ticks
        builder.writeln("Text with InlineCode style with 1 backtick")

        # There will be 3 backticks.
        inline_code3_back_ticks = builder.document.styles.add(aw.StyleType.CHARACTER, "InlineCode.3")
        builder.font.style = inline_code3_back_ticks
        builder.writeln("Text with InlineCode style with 3 backtick")

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.inline_code_example.md")
        #ExEnd:InlineCode

    def test_autolink(self):

        #ExStart:Autolink
        # Use a document builder to add content to the document.
        builder = aw.DocumentBuilder()

        # Insert hyperlink.
        builder.insert_hyperlink("https://www.aspose.com", "https://www.aspose.com", False)
        builder.insert_hyperlink("email@aspose.com", "mailto:email@aspose.com", False)

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.autolink_example.md")
        #ExEnd:Autolink

    def test_link(self):

        #ExStart:Link
        # Use a document builder to add content to the document.
        builder = aw.DocumentBuilder()

        # Insert hyperlink.
        builder.insert_hyperlink("Aspose", "https://www.aspose.com", False)

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.link_example.md")
        #ExEnd:Link

    def test_image(self):

        #ExStart:Image
        # Use a document builder to add content to the document.
        builder = aw.DocumentBuilder()

        # Insert image.
        shape = aw.drawing.Shape(builder.document, aw.drawing.ShapeType.IMAGE)
        shape.wrap_type = aw.drawing.WrapType.INLINE
        shape.image_data.source_full_name = "/attachment/1456/pic001.png"
        shape.image_data.title = "title"
        builder.insert_node(shape)

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.image_example.md")
        #ExEnd:Image

    def test_setext_heading(self):

        #ExStart:SetextHeading
        # Use a document builder to add content to the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.paragraph_format.style_name = "Heading 1"
        builder.writeln("This is an H1 tag")

        # Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder.font.bold = False
        builder.font.italic = False

        setex_heading1 = doc.styles.add(aw.StyleType.PARAGRAPH, "SetexHeading1")
        builder.paragraph_format.style = setex_heading1
        doc.styles.get_by_name("SetexHeading1").base_style_name = "Heading 1"
        builder.writeln("Setex Heading level 1")

        builder.paragraph_format.style = doc.styles.get_by_name("Heading 3")
        builder.writeln("This is an H3 tag")

        # Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder.font.bold = False
        builder.font.italic = False

        setex_heading2 = doc.styles.add(aw.StyleType.PARAGRAPH, "SetexHeading2")
        builder.paragraph_format.style = setex_heading2
        doc.styles.get_by_name("SetexHeading2").base_style_name = "Heading 3"

        # Setex heading level will be reset to 2 if the base paragraph has a Heading level greater than 2.
        builder.writeln("Setex Heading level 2")

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.setext_heading_example.md")
        #ExEnd:SetextHeading

    def test_indented_code(self):

        #ExStart:IndentedCode
        # Use a document builder to add content to the document.
        builder = aw.DocumentBuilder()

        indented_code = builder.document.styles.add(aw.StyleType.PARAGRAPH, "IndentedCode")
        builder.paragraph_format.style = indented_code
        builder.writeln("This is an indented code")

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.indented_code_example.md")
        #ExEnd:IndentedCode

    def test_fenced_code(self):

        #ExStart:FencedCode
        # Use a document builder to add content to the document.
        builder = aw.DocumentBuilder()

        fenced_code = builder.document.styles.add(aw.StyleType.PARAGRAPH, "FencedCode")
        builder.paragraph_format.style = fenced_code
        builder.writeln("This is an fenced code")

        fenced_code_with_info = builder.document.styles.add(aw.StyleType.PARAGRAPH, "FencedCode.C#")
        builder.paragraph_format.style = fenced_code_with_info
        builder.writeln("This is a fenced code with info string")

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.fenced_code_example.md")
        #ExEnd:FencedCode

    def test_quote(self):

        #ExStart:Quote
        # Use a document builder to add content to the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # By default a document stores blockquote style for the first level.
        builder.paragraph_format.style_name = "Quote"
        builder.writeln("Blockquote")

        # Create styles for nested levels through style inheritance.
        quote_level2 = doc.styles.add(aw.StyleType.PARAGRAPH, "Quote1")
        builder.paragraph_format.style = quote_level2
        doc.styles.get_by_name("Quote1").base_style_name = "Quote"
        builder.writeln("1. Nested blockquote")

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.quote_example.md")
        #ExEnd:Quote

    def test_bulleted_list(self):

        #ExStart:BulletedList
        # Use a document builder to add content to the document.
        builder = aw.DocumentBuilder()

        builder.list_format.apply_bullet_default()
        builder.list_format.list.list_levels[0].number_format = "-"

        builder.writeln("Item 1")
        builder.writeln("Item 2")

        builder.list_format.list_indent()

        builder.writeln("Item 2a")
        builder.writeln("Item 2b")

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.bulleted_list_example.md")
        #ExEnd:BulletedList

    def test_ordered_list(self):

        #ExStart:OrderedList
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.list_format.apply_number_default()

        builder.writeln("Item 1")
        builder.writeln("Item 2")

        builder.list_format.list_indent()

        builder.writeln("Item 2a")
        builder.write("Item 2b")

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.ordered_list_example.md")
        #ExEnd:OrderedList

    def test_table(self):

        #ExStart:Table
        # Use a document builder to add content to the document.
        builder = aw.DocumentBuilder()

        # Add the first row.
        builder.insert_cell()
        builder.writeln("a")
        builder.insert_cell()
        builder.writeln("b")
        builder.end_row()

        # Add the second row.
        builder.insert_cell()
        builder.writeln("c")
        builder.insert_cell()
        builder.writeln("d")
        builder.end_table()

        builder.document.save(ARTIFACTS_DIR + "WorkingWithMarkdown.ordered_list_table.md")
        #ExEnd:Table


if __name__ == '__main__':
    unittest.main()
