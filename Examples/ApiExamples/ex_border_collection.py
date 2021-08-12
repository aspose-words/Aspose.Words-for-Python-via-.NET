import unittest

import api_example_base as aeb
from document_helper import DocumentHelper

import aspose.words as aw

class ExBorderCollection(aeb.ApiExampleBase):
    
    def test_get_borders_enumerator(self) :
        
        #ExStart
        #ExFor:BorderCollection.get_enumerator
        #ExSummary:Shows how to iterate over and edit all of the borders in a paragraph format object.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Configure the builder's paragraph format settings to create a green wave border on all sides.
        borders = builder.paragraph_format.borders

        for border in borders :
            #border.color = Color.green
            border.line_style = aw.LineStyle.WAVE
            border.line_width = 3
                
            

        # Insert a paragraph. Our border settings will determine the appearance of its border.
        builder.writeln("Hello world!")

        doc.save(aeb.artifacts_dir + "BorderCollection.get_borders_enumerator.docx")
        #ExEnd

        doc = aw.Document(aeb.artifacts_dir + "BorderCollection.get_borders_enumerator.docx")

        for border in doc.first_section.body.first_paragraph.paragraph_format.borders :
            #self.assertEqual(Color.green.to_argb(), border.color.to_argb())
            self.assertEqual(aw.LineStyle.WAVE, border.line_style)
            self.assertEqual(3.0, border.line_width)
            
        

    def test_remove_all_borders(self) :
        
        #ExStart
        #ExFor:BorderCollection.clear_formatting
        #ExSummary:Shows how to remove all borders from all paragraphs in a document.
        doc = aw.Document(aeb.my_dir + "Borders.docx")

        # The first paragraph of this document has visible borders with these settings.
        firstParagraphBorders = doc.first_section.body.first_paragraph.paragraph_format.borders

        #self.assertEqual(Color.red.to_argb(), firstParagraphBorders.color.to_argb())
        self.assertEqual(aw.LineStyle.SINGLE, firstParagraphBorders.line_style)
        self.assertEqual(3.0, firstParagraphBorders.line_width)

        # Use the "ClearFormatting" method on each paragraph to remove all borders.
        for i in range(0, doc.first_section.body.paragraphs.count) :
            
            paragraph = doc.first_section.body.paragraphs[i]

            paragraph.paragraph_format.borders.clear_formatting()

            for border in paragraph.paragraph_format.borders :
                
                #self.assertEqual(Color.empty.to_argb(), border.color.to_argb())
                self.assertEqual(aw.LineStyle.NONE, border.line_style)
                self.assertEqual(0.0, border.line_width)
                
            
            
        doc.save(aeb.artifacts_dir + "BorderCollection.remove_all_borders.docx")
        #ExEnd

        doc = aw.Document(aeb.artifacts_dir + "BorderCollection.remove_all_borders.docx")

        for border in doc.first_section.body.first_paragraph.paragraph_format.borders :
            
            #self.assertEqual(Color.empty.to_argb(), border.color.to_argb())
            self.assertEqual(aw.LineStyle.NONE, border.line_style)
            self.assertEqual(0.0, border.line_width)
            
        
    
if __name__ == '__main__':
    unittest.main()    