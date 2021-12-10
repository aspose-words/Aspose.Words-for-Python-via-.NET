import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class ExBorderCollection(ApiExampleBase):

    def test_get_borders_enumerator(self):

        #ExStart
        #ExFor:BorderCollection.GetEnumerator
        #ExSummary:Shows how to iterate over and edit all of the borders in a paragraph format object.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Configure the builder's paragraph format settings to create a green wave border on all sides.
        borders = builder.paragraph_format.borders

        for border in borders:
            border.color = drawing.Color.green
            border.line_style = aw.LineStyle.WAVE
            border.line_width = 3

        # Insert a paragraph. Our border settings will determine the appearance of its border.
        builder.writeln("Hello world!")

        doc.save(ARTIFACTS_DIR + "BorderCollection.get_borders_enumerator.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "BorderCollection.get_borders_enumerator.docx")

        for border in doc.first_section.body.first_paragraph.paragraph_format.borders:
            self.assertEqual(drawing.Color.green.to_argb(), border.color.to_argb())
            self.assertEqual(aw.LineStyle.WAVE, border.line_style)
            self.assertEqual(3.0, border.line_width)

    def test_remove_all_borders(self):

        #ExStart
        #ExFor:BorderCollection.ClearFormatting
        #ExSummary:Shows how to remove all borders from all paragraphs in a document.
        doc = aw.Document(MY_DIR + "Borders.docx")

        # The first paragraph of this document has visible borders with these settings.
        first_paragraph_borders = doc.first_section.body.first_paragraph.paragraph_format.borders

        self.assertEqual(drawing.Color.red.to_argb(), first_paragraph_borders.color.to_argb())
        self.assertEqual(aw.LineStyle.SINGLE, first_paragraph_borders.line_style)
        self.assertEqual(3.0, first_paragraph_borders.line_width)

        # Use the "clear_formatting" method on each paragraph to remove all borders.
        for paragraph in doc.first_section.body.paragraphs:
            paragraph = paragraph.as_paragraph()
            paragraph.paragraph_format.borders.clear_formatting()

            for border in paragraph.paragraph_format.borders:
                self.assertEqual(drawing.Color.empty().to_argb(), border.color.to_argb())
                self.assertEqual(aw.LineStyle.NONE, border.line_style)
                self.assertEqual(0.0, border.line_width)

        doc.save(ARTIFACTS_DIR + "BorderCollection.remove_all_borders.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "BorderCollection.remove_all_borders.docx")

        for border in doc.first_section.body.first_paragraph.paragraph_format.borders:
            self.assertEqual(drawing.Color.empty().to_argb(), border.color.to_argb())
            self.assertEqual(aw.LineStyle.NONE, border.line_style)
            self.assertEqual(0.0, border.line_width)
