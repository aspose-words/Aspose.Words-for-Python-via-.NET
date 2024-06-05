# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.pydrawing
import aspose.words as aw
import aspose.words.themes
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExBorder(ApiExampleBase):

    def test_font_border(self):
        #ExStart
        #ExFor:Border
        #ExFor:Border.color
        #ExFor:Border.line_width
        #ExFor:Border.line_style
        #ExFor:Font.border
        #ExFor:LineStyle
        #ExFor:Font
        #ExFor:DocumentBuilder.font
        #ExFor:DocumentBuilder.write(str)
        #ExSummary:Shows how to insert a string surrounded by a border into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.font.border.color = aspose.pydrawing.Color.green
        builder.font.border.line_width = 2.5
        builder.font.border.line_style = aw.LineStyle.DASH_DOT_STROKER
        builder.write('Text surrounded by green border.')
        doc.save(file_name=ARTIFACTS_DIR + 'Border.FontBorder.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Border.FontBorder.docx')
        border = doc.first_section.body.first_paragraph.runs[0].font.border
        self.assertEqual(aspose.pydrawing.Color.green.to_argb(), border.color.to_argb())
        self.assertEqual(2.5, border.line_width)
        self.assertEqual(aw.LineStyle.DASH_DOT_STROKER, border.line_style)

    def test_paragraph_top_border(self):
        #ExStart
        #ExFor:BorderCollection
        #ExFor:Border.theme_color
        #ExFor:Border.tint_and_shade
        #ExFor:Border
        #ExFor:BorderType
        #ExFor:ParagraphFormat.borders
        #ExSummary:Shows how to insert a paragraph with a top border.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        top_border = builder.paragraph_format.borders.top
        top_border.line_width = 4
        top_border.line_style = aw.LineStyle.DASH_SMALL_GAP
        # Set ThemeColor only when LineWidth or LineStyle setted.
        top_border.theme_color = aw.themes.ThemeColor.ACCENT1
        top_border.tint_and_shade = 0.25
        builder.writeln('Text with a top border.')
        doc.save(file_name=ARTIFACTS_DIR + 'Border.ParagraphTopBorder.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Border.ParagraphTopBorder.docx')
        border = doc.first_section.body.first_paragraph.paragraph_format.borders.top
        self.assertEqual(4, border.line_width)
        self.assertEqual(aw.LineStyle.DASH_SMALL_GAP, border.line_style)
        self.assertEqual(aw.themes.ThemeColor.ACCENT1, border.theme_color)
        self.assertAlmostEqual(0.25, border.tint_and_shade, delta=0.01)

    def test_clear_formatting(self):
        #ExStart
        #ExFor:Border.clear_formatting
        #ExFor:Border.is_visible
        #ExSummary:Shows how to remove borders from a paragraph.
        doc = aw.Document(file_name=MY_DIR + 'Borders.docx')
        # Each paragraph has an individual set of borders.
        # We can access the settings for the appearance of these borders via the paragraph format object.
        borders = doc.first_section.body.first_paragraph.paragraph_format.borders
        self.assertEqual(aspose.pydrawing.Color.red.to_argb(), borders[0].color.to_argb())
        self.assertEqual(3, borders[0].line_width)
        self.assertEqual(aw.LineStyle.SINGLE, borders[0].line_style)
        self.assertTrue(borders[0].is_visible)
        # We can remove a border at once by running the ClearFormatting method.
        # Running this method on every border of a paragraph will remove all its borders.
        for border in borders:
            border.clear_formatting()
        self.assertEqual(aspose.pydrawing.Color.empty().to_argb(), borders[0].color.to_argb())
        self.assertEqual(0, borders[0].line_width)
        self.assertEqual(aw.LineStyle.NONE, borders[0].line_style)
        self.assertFalse(borders[0].is_visible)
        doc.save(file_name=ARTIFACTS_DIR + 'Border.ClearFormatting.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Border.ClearFormatting.docx')
        for test_border in doc.first_section.body.first_paragraph.paragraph_format.borders:
            self.assertEqual(aspose.pydrawing.Color.empty().to_argb(), test_border.color.to_argb())
            self.assertEqual(0, test_border.line_width)
            self.assertEqual(aw.LineStyle.NONE, test_border.line_style)

    def test_shared_elements(self):
        #ExStart
        #ExFor:Border.__eq__(object)
        #ExFor:Border.__eq__(Border)
        #ExFor:Border.__hash__
        #ExFor:BorderCollection.count
        #ExFor:BorderCollection.__eq__(BorderCollection)
        #ExFor:BorderCollection.__getitem__(int)
        #ExSummary:Shows how border collections can share elements.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln('Paragraph 1.')
        builder.write('Paragraph 2.')
        # Since we used the same border configuration while creating
        # these paragraphs, their border collections share the same elements.
        first_paragraph_borders = doc.first_section.body.first_paragraph.paragraph_format.borders
        second_paragraph_borders = builder.current_paragraph.paragraph_format.borders
        self.assertEqual(6, first_paragraph_borders.count)  #ExSkip
        i = 0
        while i < first_paragraph_borders.count:
            self.assertTrue(first_paragraph_borders[i].equals(rhs=second_paragraph_borders[i]))
            self.assertEqual(hash(first_paragraph_borders[i]), hash(second_paragraph_borders[i]))
            self.assertFalse(first_paragraph_borders[i].is_visible)
            i += 1
        for border in second_paragraph_borders:
            border.line_style = aw.LineStyle.DOT_DASH
        # After changing the line style of the borders in just the second paragraph,
        # the border collections no longer share the same elements.
        i = 0
        while i < first_paragraph_borders.count:
            self.assertFalse(first_paragraph_borders[i].equals(rhs=second_paragraph_borders[i]))
            self.assertNotEqual(hash(first_paragraph_borders[i]), hash(second_paragraph_borders[i]))
            # Changing the appearance of an empty border makes it visible.
            self.assertTrue(second_paragraph_borders[i].is_visible)
            i += 1
        doc.save(file_name=ARTIFACTS_DIR + 'Border.SharedElements.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Border.SharedElements.docx')
        paragraphs = doc.first_section.body.paragraphs
        for test_border in paragraphs[0].paragraph_format.borders:
            self.assertEqual(aw.LineStyle.NONE, test_border.line_style)
        for test_border in paragraphs[1].paragraph_format.borders:
            self.assertEqual(aw.LineStyle.DOT_DASH, test_border.line_style)

    def test_horizontal_borders(self):
        #ExStart
        #ExFor:BorderCollection.horizontal
        #ExSummary:Shows how to apply settings to horizontal borders to a paragraph's format.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Create a red horizontal border for the paragraph. Any paragraphs created afterwards will inherit these border settings.
        borders = doc.first_section.body.first_paragraph.paragraph_format.borders
        borders.horizontal.color = aspose.pydrawing.Color.red
        borders.horizontal.line_style = aw.LineStyle.DASH_SMALL_GAP
        borders.horizontal.line_width = 3
        # Write text to the document without creating a new paragraph afterward.
        # Since there is no paragraph underneath, the horizontal border will not be visible.
        builder.write('Paragraph above horizontal border.')
        # Once we add a second paragraph, the border of the first paragraph will become visible.
        builder.insert_paragraph()
        builder.write('Paragraph below horizontal border.')
        doc.save(file_name=ARTIFACTS_DIR + 'Border.HorizontalBorders.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Border.HorizontalBorders.docx')
        paragraphs = doc.first_section.body.paragraphs
        self.assertEqual(aw.LineStyle.DASH_SMALL_GAP, paragraphs[0].paragraph_format.borders.get_by_border_type(aw.BorderType.HORIZONTAL).line_style)
        self.assertEqual(aw.LineStyle.DASH_SMALL_GAP, paragraphs[1].paragraph_format.borders.get_by_border_type(aw.BorderType.HORIZONTAL).line_style)

    def test_vertical_borders(self):
        #ExStart
        #ExFor:BorderCollection.horizontal
        #ExFor:BorderCollection.vertical
        #ExFor:Cell.last_paragraph
        #ExSummary:Shows how to apply settings to vertical borders to a table row's format.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Create a table with red and blue inner borders.
        table = builder.start_table()
        for i in range(3):
            builder.insert_cell()
            builder.write(f'Row {i + 1}, Column 1')
            builder.insert_cell()
            builder.write(f'Row {i + 1}, Column 2')
            row = builder.end_row()
            borders = row.row_format.borders
            # Adjust the appearance of borders that will appear between rows.
            borders.horizontal.color = aspose.pydrawing.Color.red
            borders.horizontal.line_style = aw.LineStyle.DOT
            borders.horizontal.line_width = 2.0
            # Adjust the appearance of borders that will appear between cells.
            borders.vertical.color = aspose.pydrawing.Color.blue
            borders.vertical.line_style = aw.LineStyle.DOT
            borders.vertical.line_width = 2.0
        # A row format, and a cell's inner paragraph use different border settings.
        border = table.first_row.first_cell.last_paragraph.paragraph_format.borders.vertical
        self.assertEqual(aspose.pydrawing.Color.empty().to_argb(), border.color.to_argb())
        self.assertEqual(0.0, border.line_width)
        self.assertEqual(aw.LineStyle.NONE, border.line_style)
        doc.save(ARTIFACTS_DIR + 'Border.vertical_borders.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Border.vertical_borders.docx')
        table = doc.first_section.body.tables[0]
        for row in table.get_child_nodes(aw.NodeType.ROW, True):
            row = row.as_row()
            self.assertEqual(aspose.pydrawing.Color.red.to_argb(), row.row_format.borders.horizontal.color.to_argb())
            self.assertEqual(aw.LineStyle.DOT, row.row_format.borders.horizontal.line_style)
            self.assertEqual(2.0, row.row_format.borders.horizontal.line_width)
            self.assertEqual(aspose.pydrawing.Color.blue.to_argb(), row.row_format.borders.vertical.color.to_argb())
            self.assertEqual(aw.LineStyle.DOT, row.row_format.borders.vertical.line_style)
            self.assertEqual(2.0, row.row_format.borders.vertical.line_width)