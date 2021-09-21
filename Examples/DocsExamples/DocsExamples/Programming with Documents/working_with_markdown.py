import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithMarkdown(docs_base.DocsExamplesBase):
    
    def test_create_markdown_document(self) :
        
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
        nestedQuote = doc.styles.add(aw.StyleType.PARAGRAPH, "Quote1")
        nestedQuote.base_style_name = "Quote"
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
        doc.save(docs_base.artifacts_dir + "WorkingWithMarkdown.create_markdown_document.md")
        #ExEnd:CreateMarkdownDocument
        

    def test_read_markdown_document(self) :
        
        #ExStart:ReadMarkdownDocument
        doc = aw.Document(docs_base.my_dir + "Quotes.md")

        # Let's remove Heading formatting from a Quote in the very last paragraph.
        paragraph = doc.first_section.body.last_paragraph
        paragraph.paragraph_format.style = doc.styles["Quote"]

        doc.save(docs_base.artifacts_dir + "WorkingWithMarkdown.read_markdown_document.md")
        #ExEnd:ReadMarkdownDocument
        

    def test_emphases(self) :
        
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

        builder.document.save(docs_base.artifacts_dir + "WorkingWithMarkdown.emphases.md")
        #ExEnd:Emphases
        

    def test_headings(self) :
        
        #ExStart:Headings
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # By default Heading styles in Word may have bold and italic formatting.
        # If we do not want the text to be emphasized, set these properties explicitly to false.
        builder.font.bold = False
        builder.font.italic = False

        builder.writeln("The following produces headings:")
        builder.paragraph_format.style = doc.styles["Heading 1"]
        builder.writeln("Heading1")
        builder.paragraph_format.style = doc.styles["Heading 2"]
        builder.writeln("Heading2")
        builder.paragraph_format.style = doc.styles["Heading 3"]
        builder.writeln("Heading3")
        builder.paragraph_format.style = doc.styles["Heading 4"]
        builder.writeln("Heading4")
        builder.paragraph_format.style = doc.styles["Heading 5"]
        builder.writeln("Heading5")
        builder.paragraph_format.style = doc.styles["Heading 6"]
        builder.writeln("Heading6")

        # Note that the emphases are also allowed inside Headings.
        builder.font.bold = True
        builder.paragraph_format.style = doc.styles["Heading 1"]
        builder.writeln("Bold Heading1")

        doc.save(docs_base.artifacts_dir + "WorkingWithMarkdown.headings.md")
        #ExEnd:Headings
        

    def test_block_quotes(self) :
        
        #ExStart:BlockQuotes
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("We support blockquotes in Markdown:")
            
        builder.paragraph_format.style = doc.styles["Quote"]
        builder.writeln("Lorem")
        builder.writeln("ipsum")
            
        builder.paragraph_format.style = doc.styles["Normal"]
        builder.writeln("The quotes can be of any level and can be nested:")
            
        quoteLevel3 = doc.styles.add(aw.StyleType.PARAGRAPH, "Quote2")
        builder.paragraph_format.style = quoteLevel3
        builder.writeln("Quote level 3")
            
        quoteLevel4 = doc.styles.add(aw.StyleType.PARAGRAPH, "Quote3")
        builder.paragraph_format.style = quoteLevel4
        builder.writeln("Nested quote level 4")
            
        builder.paragraph_format.style = doc.styles["Quote"]
        builder.writeln()
        builder.writeln("Back to first level")
            
        quoteLevel1WithHeading = doc.styles.add(aw.StyleType.PARAGRAPH, "Quote Heading 3")
        builder.paragraph_format.style = quoteLevel1WithHeading
        builder.write("Headings are allowed inside Quotes")

        doc.save(docs_base.artifacts_dir + "WorkingWithMarkdown.block_quotes.md")
        #ExEnd:BlockQuotes
        

    def test_horizontal_rule(self) :
        
        #ExStart:HorizontalRule
        builder = aw.DocumentBuilder()

        builder.writeln("We support Horizontal rules (Thematic breaks) in Markdown:")
        builder.insert_horizontal_rule()

        builder.document.save(docs_base.artifacts_dir + "WorkingWithMarkdown.horizontal_rule_example.md")
        #ExEnd:HorizontalRule
        

    #def test_use_warning_source(self) :
        
    #    #ExStart:UseWarningSourceMarkdown
    #    doc = aw.Document(docs_base.my_dir + "Emphases markdown warning.docx")

    #    WarningInfoCollection warnings = new WarningInfoCollection()
    #    doc.warning_callback = warnings

    #    doc.save(docs_base.artifacts_dir + "WorkingWithMarkdown.use_warning_source.md")

    #    foreach (WarningInfo warningInfo in warnings)
            
    #        if (warningInfo.source == WarningSource.markdown)
    #            print(warningInfo.description)
            
    #    #ExEnd:UseWarningSourceMarkdown
        
    

if __name__ == '__main__':
    unittest.main()