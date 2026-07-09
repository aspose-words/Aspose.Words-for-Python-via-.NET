import datetime
import glob
from typing import List
import sys
# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import aspose.words.drawing
import aspose.words.saving
import document_helper
import os
import pathlib
import system_helper
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, GOLDS_DIR, IMAGE_DIR, MY_DIR, FONTS_DIR

class ExMarkdownSaveOptions(ApiExampleBase):

    def test_markdown_document_table_content_alignment(self):
        for table_content_alignment in [aw.saving.TableContentAlignment.LEFT, aw.saving.TableContentAlignment.RIGHT, aw.saving.TableContentAlignment.CENTER, aw.saving.TableContentAlignment.AUTO]:
            #ExStart
            #ExFor:TableContentAlignment
            #ExFor:MarkdownSaveOptions.table_content_alignment
            #ExSummary:Shows how to align contents in tables.
            builder = aw.DocumentBuilder()
            builder.insert_cell()
            builder.paragraph_format.alignment = aw.ParagraphAlignment.RIGHT
            builder.write('Cell1')
            builder.insert_cell()
            builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
            builder.write('Cell2')
            save_options = aw.saving.MarkdownSaveOptions()
            save_options.table_content_alignment = table_content_alignment
            builder.document.save(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.MarkdownDocumentTableContentAlignment.md', save_options=save_options)
            doc = aw.Document(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.MarkdownDocumentTableContentAlignment.md')
            table = doc.first_section.body.tables[0]
            switch_condition = table_content_alignment
            if switch_condition == aw.saving.TableContentAlignment.AUTO:
                self.assertEqual(aw.ParagraphAlignment.RIGHT, table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                self.assertEqual(aw.ParagraphAlignment.CENTER, table.first_row.cells[1].first_paragraph.paragraph_format.alignment)
            elif switch_condition == aw.saving.TableContentAlignment.LEFT:
                self.assertEqual(aw.ParagraphAlignment.LEFT, table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                self.assertEqual(aw.ParagraphAlignment.LEFT, table.first_row.cells[1].first_paragraph.paragraph_format.alignment)
            elif switch_condition == aw.saving.TableContentAlignment.CENTER:
                self.assertEqual(aw.ParagraphAlignment.CENTER, table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                self.assertEqual(aw.ParagraphAlignment.CENTER, table.first_row.cells[1].first_paragraph.paragraph_format.alignment)
            elif switch_condition == aw.saving.TableContentAlignment.RIGHT:
                self.assertEqual(aw.ParagraphAlignment.RIGHT, table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                self.assertEqual(aw.ParagraphAlignment.RIGHT, table.first_row.cells[1].first_paragraph.paragraph_format.alignment)
        #ExEnd

    @unittest.skipIf(sys.platform.startswith('win'), 'Discrepancy in assertion between Python and .Net')
    def test_rename_images(self):
        #ExStart
        #ExFor:MarkdownSaveOptions
        #ExFor:MarkdownSaveOptions.__init__
        #ExFor:MarkdownSaveOptions.image_saving_callback
        #ExFor:MarkdownSaveOptions.save_format
        #ExFor:IImageSavingCallback
        #ExSummary:Shows how to rename the image name during saving into Markdown document.
        doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
        save_options = aw.saving.MarkdownSaveOptions()
        # If we convert a document that contains images into Markdown, we will end up with one Markdown file which links to several images.
        # Each image will be in the form of a file in the local file system.
        # There is also a callback that can customize the name and file system location of each image.
        save_options.image_saving_callback = self.SavedImageRename('MarkdownSaveOptions.HandleDocument.md')
        save_options.save_format = aw.SaveFormat.MARKDOWN
        # The ImageSaving() method of our callback will be run at this time.
        doc.save(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.HandleDocument.md', save_options=save_options)
        self.assertEqual(1, len(list(filter(lambda f: f.endswith('.jpeg'), list(filter(lambda s: s.startswith(ARTIFACTS_DIR + 'MarkdownSaveOptions.HandleDocument.md shape'), list(system_helper.io.Directory.get_files(ARTIFACTS_DIR))))))))
        self.assertEqual(8, len(list(filter(lambda f: f.endswith('.png'), list(filter(lambda s: s.startswith(ARTIFACTS_DIR + 'MarkdownSaveOptions.HandleDocument.md shape'), list(system_helper.io.Directory.get_files(ARTIFACTS_DIR))))))))
        #ExEnd

    def test_export_images_as_base64(self):
        for export_images_as_base64 in [True, False]:
            #ExStart
            #ExFor:MarkdownSaveOptions.export_images_as_base64
            #ExSummary:Shows how to save a .md document with images embedded inside it.
            doc = aw.Document(file_name=MY_DIR + 'Images.docx')
            save_options = aw.saving.MarkdownSaveOptions()
            save_options.export_images_as_base64 = export_images_as_base64
            doc.save(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.ExportImagesAsBase64.md', save_options=save_options)
            out_doc_contents = system_helper.io.File.read_all_text(ARTIFACTS_DIR + 'MarkdownSaveOptions.ExportImagesAsBase64.md')
            self.assertTrue('data:image/jpeg;base64' in out_doc_contents if export_images_as_base64 else 'MarkdownSaveOptions.ExportImagesAsBase64.001.jpeg' in out_doc_contents)
            #ExEnd

    def test_list_export_mode(self):
        for markdown_list_export_mode in [aw.saving.MarkdownListExportMode.PLAIN_TEXT, aw.saving.MarkdownListExportMode.MARKDOWN_SYNTAX]:
            #ExStart
            #ExFor:MarkdownSaveOptions.list_export_mode
            #ExFor:MarkdownListExportMode
            #ExSummary:Shows how to list items will be written to the markdown document.
            doc = aw.Document(file_name=MY_DIR + 'List item.docx')
            # Use MarkdownListExportMode.PlainText or MarkdownListExportMode.MarkdownSyntax to export list.
            options = aw.saving.MarkdownSaveOptions()
            options.list_export_mode = markdown_list_export_mode
            doc.save(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.ListExportMode.md', save_options=options)
        #ExEnd

    def test_images_folder(self):
        #ExStart
        #ExFor:MarkdownSaveOptions.images_folder
        #ExFor:MarkdownSaveOptions.images_folder_alias
        #ExSummary:Shows how to specifies the name of the folder used to construct image URIs.
        builder = aw.DocumentBuilder()
        builder.writeln('Some image below:')
        builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
        images_folder = os.path.join(ARTIFACTS_DIR, 'ImagesDir')
        save_options = aw.saving.MarkdownSaveOptions()
        # Use the "ImagesFolder" property to assign a folder in the local file system into which
        # Aspose.Words will save all the document's linked images.
        save_options.images_folder = images_folder
        # Use the "ImagesFolderAlias" property to use this folder
        # when constructing image URIs instead of the images folder's name.
        save_options.images_folder_alias = 'http://example.com/images'
        builder.document.save(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.ImagesFolder.md', save_options=save_options)
        #ExEnd
        dir_files = system_helper.io.Directory.get_files(images_folder, 'MarkdownSaveOptions.ImagesFolder.001.jpeg')
        self.assertEqual(1, len(dir_files))
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.ImagesFolder.md')
        'http://example.com/images/MarkdownSaveOptions.ImagesFolder.001.jpeg' in doc.get_text()

    def test_export_underline_formatting(self):
        #ExStart:ExportUnderlineFormatting
        #ExFor:MarkdownSaveOptions.export_underline_formatting
        #ExSummary:Shows how to export underline formatting as ++.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.underline = aw.Underline.SINGLE
        builder.write('Lorem ipsum. Dolor sit amet.')
        save_options = aw.saving.MarkdownSaveOptions()
        save_options.export_underline_formatting = True
        doc.save(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.ExportUnderlineFormatting.md', save_options=save_options)
        #ExEnd:ExportUnderlineFormatting

    def test_link_export_mode(self):
        #ExStart:LinkExportMode
        #ExFor:MarkdownSaveOptions.link_export_mode
        #ExFor:MarkdownLinkExportMode
        #ExSummary:Shows how to links will be written to the .md file.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.insert_shape(shape_type=aw.drawing.ShapeType.BALLOON, width=100, height=100)
        # Image will be written as reference:
        # ![ref1]
        #
        # [ref1]: aw_ref.001.png
        save_options = aw.saving.MarkdownSaveOptions()
        save_options.link_export_mode = aw.saving.MarkdownLinkExportMode.REFERENCE
        doc.save(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.LinkExportMode.Reference.md', save_options=save_options)
        # Image will be written as inline:
        # ![](aw_inline.001.png)
        save_options.link_export_mode = aw.saving.MarkdownLinkExportMode.INLINE
        doc.save(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.LinkExportMode.Inline.md', save_options=save_options)
        #ExEnd:LinkExportMode
        out_doc_contents = system_helper.io.File.read_all_text(ARTIFACTS_DIR + 'MarkdownSaveOptions.LinkExportMode.Inline.md')
        self.assertEqual('![](MarkdownSaveOptions.LinkExportMode.Inline.001.png)', out_doc_contents.strip())

    def test_export_table_as_html(self):
        #ExStart:ExportTableAsHtml
        #ExFor:MarkdownExportAsHtml
        #ExFor:MarkdownSaveOptions.export_as_html
        #ExSummary:Shows how to export a table to Markdown as raw HTML.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Sample table:')
        # Create table.
        builder.insert_cell()
        builder.paragraph_format.alignment = aw.ParagraphAlignment.RIGHT
        builder.write('Cell1')
        builder.insert_cell()
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        builder.write('Cell2')
        save_options = aw.saving.MarkdownSaveOptions()
        save_options.export_as_html = aw.saving.MarkdownExportAsHtml.TABLES
        doc.save(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.ExportTableAsHtml.md', save_options=save_options)
        #ExEnd:ExportTableAsHtml
        new_line = system_helper.environment.Environment.new_line()
        out_doc_contents = system_helper.io.File.read_all_text(ARTIFACTS_DIR + 'MarkdownSaveOptions.ExportTableAsHtml.md')
        self.assertEqual(f'Sample table:{new_line}<table cellspacing="0" cellpadding="0" style="width:100%; border:0.75pt solid #000000; border-collapse:collapse">' + '<tr><td style="border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top">' + '<p style="margin-top:0pt; margin-bottom:0pt; text-align:right; font-size:12pt"><span style="font-family:\'Times New Roman\'">Cell1</span></p>' + '</td><td style="border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top">' + '<p style="margin-top:0pt; margin-bottom:0pt; text-align:center; font-size:12pt"><span style="font-family:\'Times New Roman\'">Cell2</span></p>' + '</td></tr></table>', out_doc_contents.strip())

    def test_image_resolution(self):
        #ExStart:ImageResolution
        #ExFor:MarkdownSaveOptions.image_resolution
        #ExSummary:Shows how to set the output resolution for images.
        doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
        save_options = aw.saving.MarkdownSaveOptions()
        save_options.image_resolution = 300
        doc.save(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.ImageResolution.md', save_options=save_options)
        #ExEnd:ImageResolution

    def test_office_math_export_mode(self):
        #ExStart:OfficeMathExportMode
        #ExFor:MarkdownSaveOptions.office_math_export_mode
        #ExFor:MarkdownOfficeMathExportMode
        #ExSummary:Shows how OfficeMath will be written to the document.
        doc = aw.Document(file_name=MY_DIR + 'Office math.docx')
        save_options = aw.saving.MarkdownSaveOptions()
        save_options.office_math_export_mode = aw.saving.MarkdownOfficeMathExportMode.IMAGE
        doc.save(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.OfficeMathExportMode.md', save_options=save_options)
        #ExEnd:OfficeMathExportMode

    def test_empty_paragraph_export_mode(self):
        for export_mode in [aw.saving.MarkdownEmptyParagraphExportMode.NONE, aw.saving.MarkdownEmptyParagraphExportMode.EMPTY_LINE, aw.saving.MarkdownEmptyParagraphExportMode.MARKDOWN_HARD_LINE_BREAK]:
            #ExStart:EmptyParagraphExportMode
            #ExFor:MarkdownEmptyParagraphExportMode
            #ExFor:MarkdownSaveOptions.empty_paragraph_export_mode
            #ExSummary:Shows how to export empty paragraphs.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.writeln('First')
            builder.writeln('\r\n\r\n\r\n')
            builder.writeln('Last')
            save_options = aw.saving.MarkdownSaveOptions()
            save_options.empty_paragraph_export_mode = export_mode
            doc.save(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.EmptyParagraphExportMode.md', save_options=save_options)
            result = system_helper.io.File.read_all_text(ARTIFACTS_DIR + 'MarkdownSaveOptions.EmptyParagraphExportMode.md')
            switch_condition = export_mode
            if switch_condition == aw.saving.MarkdownEmptyParagraphExportMode.NONE:
                self.assertEqual('First\r\n\r\nLast\r\n', result)
            elif switch_condition == aw.saving.MarkdownEmptyParagraphExportMode.EMPTY_LINE:
                self.assertEqual('First\r\n\r\n\r\n\r\n\r\nLast\r\n\r\n', result)
            elif switch_condition == aw.saving.MarkdownEmptyParagraphExportMode.MARKDOWN_HARD_LINE_BREAK:
                self.assertEqual('First\r\n\\\r\n\\\r\n\\\r\n\\\r\n\\\r\nLast\r\n<br>\r\n', result)
        #ExEnd:EmptyParagraphExportMode

    def test_non_compatible_tables(self):
        #ExStart:NonCompatibleTables
        #ExFor:MarkdownExportAsHtml
        #ExSummary:Shows how to export tables that cannot be correctly represented in pure Markdown as raw HTML.
        output_path = ARTIFACTS_DIR + 'MarkdownSaveOptions.NonCompatibleTables.md'
        doc = aw.Document(file_name=MY_DIR + 'Non compatible table.docx')
        # With the "NonCompatibleTables" option, you can export tables that have a complex structure with merged cells
        # or nested tables to raw HTML and leave simple tables in Markdown format.
        save_options = aw.saving.MarkdownSaveOptions()
        save_options.export_as_html = aw.saving.MarkdownExportAsHtml.NON_COMPATIBLE_TABLES
        doc.save(file_name=output_path, save_options=save_options)
        #ExEnd:NonCompatibleTables
        document_helper.DocumentHelper.find_text_in_file(output_path, '<table><tr><th rowspan="2" valign="top">Heading 1</th>')
        document_helper.DocumentHelper.find_text_in_file(output_path, '|Heading 1|Heading 2|')

    def test_export_office_math_as_latex(self):
        #ExStart:ExportOfficeMathAsLatex
        #ExFor:MarkdownSaveOptions.office_math_export_mode
        #ExFor:MarkdownOfficeMathExportMode
        #ExSummary:Shows how to export OfficeMath object as Latex.
        doc = aw.Document(file_name=MY_DIR + 'Office math.docx')
        save_options = aw.saving.MarkdownSaveOptions()
        save_options.office_math_export_mode = aw.saving.MarkdownOfficeMathExportMode.LATEX
        doc.save(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.ExportOfficeMathAsLatex.md', save_options=save_options)
        #ExEnd:ExportOfficeMathAsLatex
        self.assertTrue(document_helper.DocumentHelper.compare_docs(ARTIFACTS_DIR + 'MarkdownSaveOptions.ExportOfficeMathAsLatex.md', GOLDS_DIR + 'MarkdownSaveOptions.ExportOfficeMathAsLatex.Gold.md'))

    def test_export_office_math_as_mark_it_down(self):
        #ExStart:ExportOfficeMathAsMarkItDown
        #ExFor:MarkdownSaveOptions.office_math_export_mode
        #ExFor:MarkdownOfficeMathExportMode
        #ExSummary:Shows how to export OfficeMath object as MarkItDown.
        doc = aw.Document(file_name=MY_DIR + 'Office math.docx')
        save_options = aw.saving.MarkdownSaveOptions()
        save_options.office_math_export_mode = aw.saving.MarkdownOfficeMathExportMode.MARK_IT_DOWN
        doc.save(file_name=ARTIFACTS_DIR + 'MarkdownSaveOptions.ExportOfficeMathAsMarkItDown.md', save_options=save_options)
        #ExEnd:ExportOfficeMathAsMarkItDown
        self.assertTrue(document_helper.DocumentHelper.compare_docs(ARTIFACTS_DIR + 'MarkdownSaveOptions.ExportOfficeMathAsMarkItDown.md', GOLDS_DIR + 'MarkdownSaveOptions.ExportOfficeMathAsMarkItDown.Gold.md'))
    #ExStart
    #ExFor:MarkdownSaveOptions
    #ExFor:MarkdownSaveOptions.__init__
    #ExFor:MarkdownSaveOptions.image_saving_callback
    #ExFor:MarkdownSaveOptions.save_format
    #ExFor:IImageSavingCallback
    #ExSummary:Shows how to rename the image name during saving into Markdown document (SavedImageRename).

    class SavedImageRename(aw.saving.IImageSavingCallback):

        def __init__(self, out_file_name):
            self.m_out_file_name = out_file_name
            self.m_count = 0

        def image_saving(self, args):
            from pathlib import Path
            self.m_count += 1
            image_file_name = f'{self.m_out_file_name} shape {self.m_count}, of type {args.current_shape.shape_type}{Path(args.image_file_name).suffix}'
            args.image_file_name = image_file_name
            args.image_stream = system_helper.io.FileStream(ARTIFACTS_DIR + image_file_name, system_helper.io.FileMode.CREATE)
            assert args.is_image_available
            assert not args.keep_image_stream_open
    #ExEnd