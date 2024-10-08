# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
import datetime
import glob
from typing import List
import sys
import aspose.words as aw
import aspose.words.drawing
import aspose.words.saving
import os
import system_helper
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, IMAGE_DIR, MY_DIR, FONTS_DIR

class ExMarkdownSaveOptions(ApiExampleBase):

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


    def test_markdown_document_table_content_alignment(self):
        parameters = [aw.saving.TableContentAlignment.LEFT, aw.saving.TableContentAlignment.RIGHT, aw.saving.TableContentAlignment.CENTER, aw.saving.TableContentAlignment.AUTO]
        for table_content_alignment in parameters:
            with self.subTest(table_content_alignment=table_content_alignment):
                builder = aw.DocumentBuilder()
                builder.insert_cell()
                builder.paragraph_format.alignment = aw.ParagraphAlignment.RIGHT
                builder.write('Cell1')
                builder.insert_cell()
                builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
                builder.write('Cell2')
                save_options = aw.saving.MarkdownSaveOptions()
                save_options.table_content_alignment = table_content_alignment
                builder.document.save(ARTIFACTS_DIR + 'MarkdownSaveOptions.markdown_document_table_content_alignment.md', save_options)
                doc = aw.Document(ARTIFACTS_DIR + 'MarkdownSaveOptions.markdown_document_table_content_alignment.md')
                table = doc.first_section.body.tables[0]
                if table_content_alignment == aw.saving.TableContentAlignment.AUTO:
                    self.assertEqual(aw.ParagraphAlignment.RIGHT, table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                    self.assertEqual(aw.ParagraphAlignment.CENTER, table.first_row.cells[1].first_paragraph.paragraph_format.alignment)
                elif table_content_alignment == aw.saving.TableContentAlignment.LEFT:
                    self.assertEqual(aw.ParagraphAlignment.LEFT, table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                    self.assertEqual(aw.ParagraphAlignment.LEFT, table.first_row.cells[1].first_paragraph.paragraph_format.alignment)
                elif table_content_alignment == aw.saving.TableContentAlignment.CENTER:
                    self.assertEqual(aw.ParagraphAlignment.CENTER, table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                    self.assertEqual(aw.ParagraphAlignment.CENTER, table.first_row.cells[1].first_paragraph.paragraph_format.alignment)
                elif table_content_alignment == aw.saving.TableContentAlignment.RIGHT:
                    self.assertEqual(aw.ParagraphAlignment.RIGHT, table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                    self.assertEqual(aw.ParagraphAlignment.RIGHT, table.first_row.cells[1].first_paragraph.paragraph_format.alignment)

    @unittest.skipUnless(sys.platform.startswith('win'), 'Windows encoding')
    def test_export_images_as_base64(self):
        for export_images_as_base64 in (True, False):
            with self.subTest(export_images_as_base64=export_images_as_base64):
                #ExStart
                #ExFor:MarkdownSaveOptions.export_images_as_base64
                #ExSummary:Shows how to save a .md document with images embedded inside it.
                doc = aw.Document(MY_DIR + 'Images.docx')
                save_options = aw.saving.MarkdownSaveOptions()
                save_options.export_images_as_base64 = export_images_as_base64
                doc.save(ARTIFACTS_DIR + 'MarkdownSaveOptions.ExportImagesAsBase64.md', save_options)
                with open(ARTIFACTS_DIR + 'MarkdownSaveOptions.ExportImagesAsBase64.md') as stream:
                    out_doc_contents = stream.read()
                if export_images_as_base64:
                    self.assertIn('data:image/jpeg;base64', out_doc_contents)
                else:
                    self.assertIn('MarkdownSaveOptions.ExportImagesAsBase64.001.jpeg', out_doc_contents)
                #ExEnd

    def test_list_export_mode(self):
        for markdownListExportMode in [aw.saving.MarkdownListExportMode.MARKDOWN_SYNTAX, aw.saving.MarkdownListExportMode.PLAIN_TEXT]:
            #ExStart
            #ExFor:MarkdownSaveOptions.list_export_mode
            #ExSummary:Shows how to list items will be written to the markdown document.
            doc = aw.Document(MY_DIR + 'List item.docx')
            # Use MarkdownListExportMode.PLAIN_TEXT or MarkdownListExportMode.MARKDOWN_SYNTAX to export list.
            options = aw.saving.MarkdownSaveOptions()
            options.list_export_mode = markdownListExportMode
            doc.save(ARTIFACTS_DIR + 'MarkdownSaveOptions.ListExportMode.md', options)
            #ExEnd

    def test_images_folder(self):
        #ExStart
        #ExFor:MarkdownSaveOptions.images_folder
        #ExFor:MarkdownSaveOptions.images_folder_alias
        #ExSummary: Shows how to specifies the name of the folder used to construct image URIs.
        builder = aw.DocumentBuilder()
        builder.writeln('Some image below:')
        builder.insert_image(IMAGE_DIR + 'Logo.jpg')
        saveOptions = aw.saving.MarkdownSaveOptions()
        # Use the "ImagesFolder" property to assign a folder in the local file system into which
        # Aspose.Words will save all the document's linked images.
        saveOptions.images_folder = ARTIFACTS_DIR + 'ImagesDir/'
        # Use the "ImagesFolderAlias" property to use this folder
        # when constructing image URIs instead of the images folder's name.
        saveOptions.images_folder_alias = 'http://example.com/images'
        builder.document.save(ARTIFACTS_DIR + 'MarkdownSaveOptions.ImagesFolder.md', saveOptions)
        #ExEnd
