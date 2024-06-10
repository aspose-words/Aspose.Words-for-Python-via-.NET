# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.pydrawing
import aspose.words as aw
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExBorderCollection(ApiExampleBase):

    def test_remove_all_borders(self):
        #ExStart
        #ExFor:BorderCollection.clear_formatting
        #ExSummary:Shows how to remove all borders from all paragraphs in a document.
        doc = aw.Document(file_name=MY_DIR + 'Borders.docx')
        # The first paragraph of this document has visible borders with these settings.
        first_paragraph_borders = doc.first_section.body.first_paragraph.paragraph_format.borders
        self.assertEqual(aspose.pydrawing.Color.red.to_argb(), first_paragraph_borders.color.to_argb())
        self.assertEqual(aw.LineStyle.SINGLE, first_paragraph_borders.line_style)
        self.assertEqual(3, first_paragraph_borders.line_width)
        # Use the "ClearFormatting" method on each paragraph to remove all borders.
        for paragraph in doc.first_section.body.paragraphs:
            paragraph = paragraph.as_paragraph()
            paragraph.paragraph_format.borders.clear_formatting()
            for border in paragraph.paragraph_format.borders:
                self.assertEqual(aspose.pydrawing.Color.empty().to_argb(), border.color.to_argb())
                self.assertEqual(aw.LineStyle.NONE, border.line_style)
                self.assertEqual(0, border.line_width)
        doc.save(file_name=ARTIFACTS_DIR + 'BorderCollection.RemoveAllBorders.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'BorderCollection.RemoveAllBorders.docx')
        for border in doc.first_section.body.first_paragraph.paragraph_format.borders:
            self.assertEqual(aspose.pydrawing.Color.empty().to_argb(), border.color.to_argb())
            self.assertEqual(aw.LineStyle.NONE, border.line_style)
            self.assertEqual(0, border.line_width)

    def test_get_borders_enumerator(self):
        #ExStart
        #ExFor:BorderCollection.__iter__
        #ExSummary:Shows how to iterate over and edit all of the borders in a paragraph format object.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Configure the builder's paragraph format settings to create a green wave border on all sides.
        borders = builder.paragraph_format.borders
        for border in borders:
            border.color = aspose.pydrawing.Color.green
            border.line_style = aw.LineStyle.WAVE
            border.line_width = 3
        # Insert a paragraph. Our border settings will determine the appearance of its border.
        builder.writeln('Hello world!')
        doc.save(ARTIFACTS_DIR + 'BorderCollection.get_borders_enumerator.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'BorderCollection.get_borders_enumerator.docx')
        for border in doc.first_section.body.first_paragraph.paragraph_format.borders:
            self.assertEqual(aspose.pydrawing.Color.green.to_argb(), border.color.to_argb())
            self.assertEqual(aw.LineStyle.WAVE, border.line_style)
            self.assertEqual(3.0, border.line_width)