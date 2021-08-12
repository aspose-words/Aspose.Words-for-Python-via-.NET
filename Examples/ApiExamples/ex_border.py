import unittest

import api_example_base as aeb
from document_helper import DocumentHelper

import aspose.words as aw

class ExBorder(aeb.ApiExampleBase):
    
    def test_font_border(self) :
        
        #ExStart
        #ExFor:Border
        #ExFor:Border.color
        #ExFor:Border.line_width
        #ExFor:Border.line_style
        #ExFor:Font.border
        #ExFor:LineStyle
        #ExFor:Font
        #ExFor:DocumentBuilder.font
        #ExFor:DocumentBuilder.write(String)
        #ExSummary:Shows how to insert a string surrounded by a border into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        #builder.font.border.color = Color.green
        builder.font.border.line_width = 2.5
        builder.font.border.line_style = aw.LineStyle.DASH_DOT_STROKER

        builder.write("Text surrounded by green border.")

        doc.save(aeb.ArtifactsDir + "Border.font_border.docx")
        #ExEnd

        doc = aw.Document(aeb.ArtifactsDir + "Border.font_border.docx")
        border = doc.first_section.body.first_paragraph.runs[0].font.border

        #self.assertEqual(Color.green.to_argb(), border.color.to_argb())
        self.assertEqual(2.5, border.line_width)
        self.assertEqual(aw.LineStyle.DASH_DOT_STROKER, border.line_style)
        

    def test_paragraph_top_border(self) :
        
        #ExStart
        #ExFor:BorderCollection
        #ExFor:Border
        #ExFor:BorderType
        #ExFor:ParagraphFormat.borders
        #ExSummary:Shows how to insert a paragraph with a top border.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        topBorder = builder.paragraph_format.borders.top
        #topBorder.color = Color.red
        topBorder.line_width = 4.0
        topBorder.line_style = aw.LineStyle.DASH_SMALL_GAP

        builder.writeln("Text with a red top border.")

        doc.save(aeb.ArtifactsDir + "Border.paragraph_top_border.docx")
        #ExEnd

        doc = aw.Document(aeb.ArtifactsDir + "Border.paragraph_top_border.docx")
        border = doc.first_section.body.first_paragraph.paragraph_format.borders.top

        #self.assertEqual(Color.red.to_argb(), border.color.to_argb())
        self.assertEqual(4.0, border.line_width)
        self.assertEqual(aw.LineStyle.DASH_SMALL_GAP, border.line_style)
        

    def test_clear_formatting(self) :
        
        #ExStart
        #ExFor:Border.clear_formatting
        #ExFor:Border.is_visible
        #ExSummary:Shows how to remove borders from a paragraph.
        doc = aw.Document(aeb.MyDir + "Borders.docx")

        # Each paragraph has an individual set of borders.
        # We can access the settings for the appearance of these borders via the paragraph format object.
        borders = doc.first_section.body.first_paragraph.paragraph_format.borders

        #self.assertEqual(Color.red.to_argb(), borders[0].color.to_argb())
        self.assertEqual(3.0, borders[0].line_width)
        self.assertEqual(aw.LineStyle.SINGLE, borders[0].line_style)
        self.assertTrue(borders[0].is_visible)

        # We can remove a border at once by running the ClearFormatting method. 
        # Running this method on every border of a paragraph will remove all its borders.
        for border in borders:
            border.clear_formatting()

        #self.assertEqual(Color.empty.to_argb(), borders[0].color.to_argb())
        self.assertEqual(0.0, borders[0].line_width)
        self.assertEqual(aw.LineStyle.NONE, borders[0].line_style)
        self.assertFalse(borders[0].is_visible)

        doc.save(aeb.ArtifactsDir + "Border.clear_formatting.docx")
        #ExEnd

        doc = aw.Document(aeb.ArtifactsDir + "Border.clear_formatting.docx")

        for testBorder in doc.first_section.body.first_paragraph.paragraph_format.borders :
            
            #self.assertEqual(Color.empty.to_argb(), testBorder.color.to_argb())
            self.assertEqual(0.0, testBorder.line_width)
            self.assertEqual(aw.LineStyle.NONE, testBorder.line_style)
            
        

    def test_shared_elements(self) :
        
        #ExStart
        #ExFor:Border.equals(Object)
        #ExFor:Border.equals(Border)
        #ExFor:Border.get_hash_code
        #ExFor:BorderCollection.count
        #ExFor:BorderCollection.equals(BorderCollection)
        #ExFor:BorderCollection.item(Int32)
        #ExSummary:Shows how border collections can share elements.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Paragraph 1.")
        builder.write("Paragraph 2.")

        # Since we used the same border configuration while creating
        # these paragraphs, their border collections share the same elements.
        firstParagraphBorders = doc.first_section.body.first_paragraph.paragraph_format.borders
        secondParagraphBorders = builder.current_paragraph.paragraph_format.borders
        self.assertEqual(6, firstParagraphBorders.count) #ExSkip

        for i in range(0, firstParagraphBorders.count) :
            
            self.assertTrue(firstParagraphBorders[i].equals(secondParagraphBorders[i]))
            self.assertEqual(firstParagraphBorders[i].get_hash_code(), secondParagraphBorders[i].get_hash_code())
            self.assertFalse(firstParagraphBorders[i].is_visible)
            

        for border in secondParagraphBorders :
            border.line_style = aw.LineStyle.DOT_DASH

        # After changing the line style of the borders in just the second paragraph,
        # the border collections no longer share the same elements.
        for i in range(0, firstParagraphBorders.count) :
            
            self.assertFalse(firstParagraphBorders[i].equals(secondParagraphBorders[i]))
            self.assertNotEqual(firstParagraphBorders[i].get_hash_code(), secondParagraphBorders[i].get_hash_code())

            # Changing the appearance of an empty border makes it visible.
            self.assertTrue(secondParagraphBorders[i].is_visible)
            

        doc.save(aeb.ArtifactsDir + "Border.shared_elements.docx")
        #ExEnd

        doc = aw.Document(aeb.ArtifactsDir + "Border.shared_elements.docx")
        paragraphs = doc.first_section.body.paragraphs

        for testBorder in paragraphs[0].paragraph_format.borders :
            self.assertEqual(aw.LineStyle.NONE, testBorder.line_style)

        for testBorder in paragraphs[1].paragraph_format.borders :
            self.assertEqual(aw.LineStyle.DOT_DASH, testBorder.line_style)
        

    def test_horizontal_borders(self) :
        
        #ExStart
        #ExFor:BorderCollection.horizontal
        #ExSummary:Shows how to apply settings to horizontal borders to a paragraph's format.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create a red horizontal border for the paragraph. Any paragraphs created afterwards will inherit these border settings.
        borders = doc.first_section.body.first_paragraph.paragraph_format.borders
        #borders.horizontal.color = Color.red
        borders.horizontal.line_style = aw.LineStyle.DASH_SMALL_GAP
        borders.horizontal.line_width = 3

        # Write text to the document without creating a new paragraph afterward.
        # Since there is no paragraph underneath, the horizontal border will not be visible.
        builder.write("Paragraph above horizontal border.")

        # Once we add a second paragraph, the border of the first paragraph will become visible.
        builder.insert_paragraph()
        builder.write("Paragraph below horizontal border.")

        doc.save(aeb.ArtifactsDir + "Border.horizontal_borders.docx")
        #ExEnd

        doc = aw.Document(aeb.ArtifactsDir + "Border.horizontal_borders.docx")
        paragraphs = doc.first_section.body.paragraphs

        self.assertEqual(aw.LineStyle.DASH_SMALL_GAP, paragraphs[0].paragraph_format.borders.horizontal.line_style)
        self.assertEqual(aw.LineStyle.DASH_SMALL_GAP, paragraphs[1].paragraph_format.borders.horizontal.line_style)
        

    def test_vertical_borders(self) :
        
        #ExStart
        #ExFor:BorderCollection.horizontal
        #ExFor:BorderCollection.vertical
        #ExFor:Cell.last_paragraph
        #ExSummary:Shows how to apply settings to vertical borders to a table row's format.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create a table with red and blue inner borders.
        table = builder.start_table()

        for i in range(0, 3) :
            
            builder.insert_cell()
            builder.write(f"Row {i + 1}, Column 1")
            builder.insert_cell()
            builder.write(f"Row {i + 1}, Column 2")

            row = builder.end_row()
            borders = row.row_format.borders

            # Adjust the appearance of borders that will appear between rows.
            #borders.horizontal.color = Color.red
            borders.horizontal.line_style = aw.LineStyle.DOT
            borders.horizontal.line_width = 2.0

            # Adjust the appearance of borders that will appear between cells.
            #borders.vertical.color = Color.blue
            borders.vertical.line_style = aw.LineStyle.DOT
            borders.vertical.line_width = 2.0
            

        # A row format, and a cell's inner paragraph use different border settings.
        border = table.first_row.first_cell.last_paragraph.paragraph_format.borders.vertical

        #self.assertEqual(Color.empty.to_argb(), border.color.to_argb())
        self.assertEqual(0.0, border.line_width)
        self.assertEqual(aw.LineStyle.NONE, border.line_style)

        doc.save(aeb.ArtifactsDir + "Border.vertical_borders.docx")
        #ExEnd

#        doc = aw.Document(aeb.ArtifactsDir + "Border.vertical_borders.docx")
#        table = doc.first_section.body.tables[0]
#
#        for row in table.get_child_nodes(aw.NodeType.ROW, True) :
#            
#            #self.assertEqual(Color.red.to_argb(), row.row_format.borders.horizontal.color.to_argb())
#            self.assertEqual(aw.LineStyle.DOT, row.row_format.borders.horizontal.line_style)
#            self.assertEqual(2.0, row.row_format.borders.horizontal.line_width)
#
#            #self.assertEqual(Color.blue.to_argb(), row.row_format.borders.vertical.color.to_argb())
#            self.assertEqual(aw.LineStyle.DOT, row.row_format.borders.vertical.line_style)
#            self.assertEqual(2.0, row.row_format.borders.vertical.line_width)
            
        
    
if __name__ == '__main__':
    unittest.main()    