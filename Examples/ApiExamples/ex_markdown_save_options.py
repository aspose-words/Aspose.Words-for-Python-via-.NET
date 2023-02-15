# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import datetime
import glob
import os
from typing import List

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, FONTS_DIR, IMAGE_DIR

class ExMarkdownSaveOptions(ApiExampleBase):

    def test_markdown_document_table_content_alignment(self):
        
        parameters = [
            aw.saving.TableContentAlignment.LEFT,
            aw.saving.TableContentAlignment.RIGHT,
            aw.saving.TableContentAlignment.CENTER,
            aw.saving.TableContentAlignment.AUTO,
            ]
        for table_content_alignment in parameters:
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

                builder.document.save(ARTIFACTS_DIR + "MarkdownSaveOptions.markdown_document_table_content_alignment.md", save_options)

                doc = aw.Document(ARTIFACTS_DIR + "MarkdownSaveOptions.markdown_document_table_content_alignment.md")
                table = doc.first_section.body.tables[0]

                if table_content_alignment == aw.saving.TableContentAlignment.AUTO:
                    self.assertEqual(
                        aw.ParagraphAlignment.RIGHT,
                        table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                    self.assertEqual(
                        aw.ParagraphAlignment.CENTER,
                        table.first_row.cells[1].first_paragraph.paragraph_format.alignment)

                elif table_content_alignment == aw.saving.TableContentAlignment.LEFT:
                    self.assertEqual(
                        aw.ParagraphAlignment.LEFT,
                        table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                    self.assertEqual(
                        aw.ParagraphAlignment.LEFT,
                        table.first_row.cells[1].first_paragraph.paragraph_format.alignment)

                elif table_content_alignment == aw.saving.TableContentAlignment.CENTER:
                    self.assertEqual(
                        aw.ParagraphAlignment.CENTER,
                        table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                    self.assertEqual(
                        aw.ParagraphAlignment.CENTER,
                        table.first_row.cells[1].first_paragraph.paragraph_format.alignment)

                elif table_content_alignment == aw.saving.TableContentAlignment.RIGHT:
                    self.assertEqual(
                        aw.ParagraphAlignment.RIGHT,
                        table.first_row.cells[0].first_paragraph.paragraph_format.alignment)
                    self.assertEqual(
                        aw.ParagraphAlignment.RIGHT,
                        table.first_row.cells[1].first_paragraph.paragraph_format.alignment)

    ##ExStart
    ##ExFor:MarkdownSaveOptions.image_saving_callback
    ##ExFor:IImageSavingCallback
    ##ExSummary:Shows how to rename the image name during saving into Markdown document.
    #def test_rename_images(self):
    #
    #    doc = aw.Document(MY_DIR + "Rendering.docx")
    #
    #    save_options = aw.saving.MarkdownSaveOptions()
    #
    #    # If we convert a document that contains images into Markdown, we will end up with one Markdown file which links to several images.
    #    # Each image will be in the form of a file in the local file system.
    #    # There is also a callback that can customize the name and file system location of each image.
    #    save_options.image_saving_callback = ExMarkdownSaveOptions.SavedImageRename("MarkdownSaveOptions.HandleDocument.md")
    #
    #    # The image_saving() method of our callback will be run at this time.
    #    doc.save(ARTIFACTS_DIR + "MarkdownSaveOptions.handle_document.md", save_options);
    #
    #    jpeg_file_names = glob.glob(ARTIFACTS_DIR + "MarkdownSaveOptions.handle_document.md shape*.jpeg")
    #    png_file_names = glob.glob(ARTIFACTS_DIR + "MarkdownSaveOptions.handle_document.md shape*.png")
    #
    #    self.assertEqual(1, len(jpeg_file_names))
    #    self.assertEqual(8, len(png_file_names))
    #
    #class SavedImageRename(aw.saving.IImageSavingCallback):
    #    """Renames saved images that are produced when an Markdown document is saved."""
    #
    #    def __init__(self, out_file_name: str):
    #        self.out_file_name = out_file_name
    #        self.count = 0
    #
    #    def image_saving(self, args: aw.saving.ImageSavingArgs):
    #        self.count += 1
    #
    #        file_ext = args.image_file_name.rpartition('.')[2]
    #        image_file_name = f"{self.out_file_name} shape {self.count}, of type {args.current_shape.shape_type}.{file_ext}"
    #
    #        args.image_file_name = image_file_name;
    #        args.image_stream = open(ARTIFACTS_DIR + image_file_name, "wb")
    #
    #        self.assertTrue(args.image_stream.can_write)
    #        self.assertTrue(args.is_image_available)
    #        self.assertFalse(args.keep_image_stream_open)
    #
    ##ExEnd

    def test_export_images_as_base64(self):
        for export_images_as_base64 in (True, False):
            with self.subTest(export_images_as_base64=export_images_as_base64):
                #ExStart
                #ExFor:MarkdownSaveOptions.export_images_as_base64
                #ExSummary:Shows how to save a .md document with images embedded inside it.
                doc = aw.Document(MY_DIR + "Images.docx")

                save_options = aw.saving.MarkdownSaveOptions()
                save_options.export_images_as_base64 = export_images_as_base64

                doc.save(ARTIFACTS_DIR + "MarkdownSaveOptions.ExportImagesAsBase64.md", save_options)

                with open(ARTIFACTS_DIR + "MarkdownSaveOptions.ExportImagesAsBase64.md") as stream:
                    out_doc_contents = stream.read()

                if export_images_as_base64:
                    self.assertIn("data:image/jpeg;base64", out_doc_contents)
                else:
                    self.assertIn("MarkdownSaveOptions.ExportImagesAsBase64.001.jpeg", out_doc_contents)
                #ExEnd

    def test_list_export_mode(self):
        for markdownListExportMode in [aw.saving.MarkdownListExportMode.MARKDOWN_SYNTAX, aw.saving.MarkdownListExportMode.PLAIN_TEXT]:
            #ExStart
            #ExFor:MarkdownSaveOptions.list_export_mode
            #ExSummary:Shows how to list items will be written to the markdown document.
            doc = aw.Document(MY_DIR + "List item.docx");

            # Use MarkdownListExportMode.PLAIN_TEXT or MarkdownListExportMode.MARKDOWN_SYNTAX to export list.
            options = aw.saving.MarkdownSaveOptions()
            options.list_export_mode = markdownListExportMode
            doc.save(ARTIFACTS_DIR + "MarkdownSaveOptions.ListExportMode.md", options)
            #ExEnd
