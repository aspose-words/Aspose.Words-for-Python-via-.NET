# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.pydrawing as drawing
import sys
import os
import io
import aspose.words.fonts
from aspose.words import Document
from aspose.words.saving import PdfTextCompression
from aspose.words.saving import PdfTextCompression
from datetime import timedelta, timezone
import aspose.words as aw
import aspose.words.digitalsignatures
import aspose.words.saving
import aspose.words.settings
import datetime
import system_helper
import test_util
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, IMAGE_DIR, MY_DIR, FONTS_DIR

class ExPdfSaveOptions(ApiExampleBase):

    def test_headings_outline_levels(self):
        #ExStart
        #ExFor:ParagraphFormat.is_heading
        #ExFor:PdfSaveOptions.outline_options
        #ExFor:PdfSaveOptions.save_format
        #ExSummary:Shows how to limit the headings' level that will appear in the outline of a saved PDF document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert headings that can serve as TOC entries of levels 1, 2, and then 3.
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
        self.assertTrue(builder.paragraph_format.is_heading)
        builder.writeln('Heading 1')
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING2
        builder.writeln('Heading 1.1')
        builder.writeln('Heading 1.2')
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING3
        builder.writeln('Heading 1.2.1')
        builder.writeln('Heading 1.2.2')
        # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        # to modify how that method converts the document to .PDF.
        save_options = aw.saving.PdfSaveOptions()
        save_options.save_format = aw.SaveFormat.PDF
        # The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
        # Clicking on an entry in this outline will take us to the location of its respective heading.
        # Set the "HeadingsOutlineLevels" property to "2" to exclude all headings whose levels are above 2 from the outline.
        # The last two headings we have inserted above will not appear.
        save_options.outline_options.headings_outline_levels = 2
        doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.HeadingsOutlineLevels.pdf', save_options=save_options)
        #ExEnd

    def test_create_missing_outline_levels(self):
        for create_missing_outline_levels in [False, True]:
            #ExStart
            #ExFor:OutlineOptions.create_missing_outline_levels
            #ExFor:PdfSaveOptions.outline_options
            #ExSummary:Shows how to work with outline levels that do not contain any corresponding headings when saving a PDF document.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            # Insert headings that can serve as TOC entries of levels 1 and 5.
            builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
            self.assertTrue(builder.paragraph_format.is_heading)
            builder.writeln('Heading 1')
            builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING5
            builder.writeln('Heading 1.1.1.1.1')
            builder.writeln('Heading 1.1.1.1.2')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            save_options = aw.saving.PdfSaveOptions()
            # The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
            # Clicking on an entry in this outline will take us to the location of its respective heading.
            # Set the "HeadingsOutlineLevels" property to "5" to include all headings of levels 5 and below in the outline.
            save_options.outline_options.headings_outline_levels = 5
            # This document contains headings of levels 1 and 5, and no headings with levels of 2, 3, and 4.
            # The output PDF document will treat outline levels 2, 3, and 4 as "missing".
            # Set the "CreateMissingOutlineLevels" property to "true" to include all missing levels in the outline,
            # leaving blank outline entries since there are no usable headings.
            # Set the "CreateMissingOutlineLevels" property to "false" to ignore missing outline levels,
            # and treat the outline level 5 headings as level 2.
            save_options.outline_options.create_missing_outline_levels = create_missing_outline_levels
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.CreateMissingOutlineLevels.pdf', save_options=save_options)
            #ExEnd

    def test_table_heading_outlines(self):
        for create_outlines_for_headings_in_tables in [False, True]:
            #ExStart
            #ExFor:OutlineOptions.create_outlines_for_headings_in_tables
            #ExSummary:Shows how to create PDF document outline entries for headings inside tables.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            # Create a table with three rows. The first row,
            # whose text we will format in a heading-type style, will serve as the column header.
            builder.start_table()
            builder.insert_cell()
            builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
            builder.write('Customers')
            builder.end_row()
            builder.insert_cell()
            builder.paragraph_format.style_identifier = aw.StyleIdentifier.NORMAL
            builder.write('John Doe')
            builder.end_row()
            builder.insert_cell()
            builder.write('Jane Doe')
            builder.end_table()
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            pdf_save_options = aw.saving.PdfSaveOptions()
            # The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
            # Clicking on an entry in this outline will take us to the location of its respective heading.
            # Set the "HeadingsOutlineLevels" property to "1" to get the outline
            # to only register headings with heading levels that are no larger than 1.
            pdf_save_options.outline_options.headings_outline_levels = 1
            # Set the "CreateOutlinesForHeadingsInTables" property to "false" to exclude all headings within tables,
            # such as the one we have created above from the outline.
            # Set the "CreateOutlinesForHeadingsInTables" property to "true" to include all headings within tables
            # in the outline, provided that they have a heading level that is no larger than the value of the "HeadingsOutlineLevels" property.
            pdf_save_options.outline_options.create_outlines_for_headings_in_tables = create_outlines_for_headings_in_tables
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.TableHeadingOutlines.pdf', save_options=pdf_save_options)
            #ExEnd

    def test_expanded_outline_levels(self):
        #ExStart
        #ExFor:Document.save(str,SaveOptions)
        #ExFor:PdfSaveOptions
        #ExFor:OutlineOptions.headings_outline_levels
        #ExFor:OutlineOptions.expanded_outline_levels
        #ExSummary:Shows how to convert a whole document to PDF with three levels in the document outline.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert headings of levels 1 to 5.
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
        self.assertTrue(builder.paragraph_format.is_heading)
        builder.writeln('Heading 1')
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING2
        builder.writeln('Heading 1.1')
        builder.writeln('Heading 1.2')
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING3
        builder.writeln('Heading 1.2.1')
        builder.writeln('Heading 1.2.2')
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING4
        builder.writeln('Heading 1.2.2.1')
        builder.writeln('Heading 1.2.2.2')
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING5
        builder.writeln('Heading 1.2.2.2.1')
        builder.writeln('Heading 1.2.2.2.2')
        # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        # to modify how that method converts the document to .PDF.
        options = aw.saving.PdfSaveOptions()
        # The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
        # Clicking on an entry in this outline will take us to the location of its respective heading.
        # Set the "HeadingsOutlineLevels" property to "4" to exclude all headings whose levels are above 4 from the outline.
        options.outline_options.headings_outline_levels = 4
        # If an outline entry has subsequent entries of a higher level inbetween itself and the next entry of the same or lower level,
        # an arrow will appear to the left of the entry. This entry is the "owner" of several such "sub-entries".
        # In our document, the outline entries from the 5th heading level are sub-entries of the second 4th level outline entry,
        # the 4th and 5th heading level entries are sub-entries of the second 3rd level entry, and so on.
        # In the outline, we can click on the arrow of the "owner" entry to collapse/expand all its sub-entries.
        # Set the "ExpandedOutlineLevels" property to "2" to automatically expand all heading level 2 and lower outline entries
        # and collapse all level and 3 and higher entries when we open the document.
        options.outline_options.expanded_outline_levels = 2
        doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.ExpandedOutlineLevels.pdf', save_options=options)
        #ExEnd

    def test_update_fields(self):
        for update_fields in [False, True]:
            #ExStart
            #ExFor:PdfSaveOptions.clone
            #ExFor:SaveOptions.update_fields
            #ExSummary:Shows how to update all the fields in a document immediately before saving it to PDF.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            # Insert text with PAGE and NUMPAGES fields. These fields do not display the correct value in real time.
            # We will need to manually update them using updating methods such as "Field.Update()", and "Document.UpdateFields()"
            # each time we need them to display accurate values.
            builder.write('Page ')
            builder.insert_field(field_code='PAGE', field_value='')
            builder.write(' of ')
            builder.insert_field(field_code='NUMPAGES', field_value='')
            builder.insert_break(aw.BreakType.PAGE_BREAK)
            builder.writeln('Hello World!')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            options = aw.saving.PdfSaveOptions()
            # Set the "UpdateFields" property to "false" to not update all the fields in a document right before a save operation.
            # This is the preferable option if we know that all our fields will be up to date before saving.
            # Set the "UpdateFields" property to "true" to iterate through all the document
            # fields and update them before we save it as a PDF. This will make sure that all the fields will display
            # the most accurate values in the PDF.
            options.update_fields = update_fields
            # We can clone PdfSaveOptions objects.
            self.assertNotEqual(options, options.clone())
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.UpdateFields.pdf', save_options=options)
            #ExEnd

    def test_preserve_form_fields(self):
        for preserve_form_fields in [False, True]:
            #ExStart
            #ExFor:PdfSaveOptions.preserve_form_fields
            #ExSummary:Shows how to save a document to the PDF format using the Save method and the PdfSaveOptions class.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.write('Please select a fruit: ')
            # Insert a combo box which will allow a user to choose an option from a collection of strings.
            builder.insert_combo_box('MyComboBox', ['Apple', 'Banana', 'Cherry'], 0)
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            pdf_options = aw.saving.PdfSaveOptions()
            # Set the "PreserveFormFields" property to "true" to save form fields as interactive objects in the output PDF.
            # Set the "PreserveFormFields" property to "false" to freeze all form fields in the document at
            # their current values and display them as plain text in the output PDF.
            pdf_options.preserve_form_fields = preserve_form_fields
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.PreserveFormFields.pdf', save_options=pdf_options)
            #ExEnd

    def test_compliance(self):
        for pdf_compliance in [aw.saving.PdfCompliance.PDF_A2U, aw.saving.PdfCompliance.PDF_A3A, aw.saving.PdfCompliance.PDF_A3U, aw.saving.PdfCompliance.PDF17, aw.saving.PdfCompliance.PDF_A2A, aw.saving.PdfCompliance.PDF_UA1, aw.saving.PdfCompliance.PDF20, aw.saving.PdfCompliance.PDF_A4, aw.saving.PdfCompliance.PDF_A4F, aw.saving.PdfCompliance.PDF_A4_UA_2, aw.saving.PdfCompliance.PDF_UA2]:
            #ExStart
            #ExFor:PdfSaveOptions.compliance
            #ExFor:PdfCompliance
            #ExSummary:Shows how to set the PDF standards compliance level of saved PDF documents.
            doc = aw.Document(file_name=MY_DIR + 'Images.docx')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            # Note that some PdfSaveOptions are prohibited when saving to one of the standards and automatically fixed.
            # Use IWarningCallback to know which options are automatically fixed.
            save_options = aw.saving.PdfSaveOptions()
            # Set the "Compliance" property to "PdfCompliance.PdfA1b" to comply with the "PDF/A-1b" standard,
            # which aims to preserve the visual appearance of the document as Aspose.Words convert it to PDF.
            # Set the "Compliance" property to "PdfCompliance.Pdf17" to comply with the "1.7" standard.
            # Set the "Compliance" property to "PdfCompliance.PdfA1a" to comply with the "PDF/A-1a" standard,
            # which complies with "PDF/A-1b" as well as preserving the document structure of the original document.
            # Set the "Compliance" property to "PdfCompliance.PdfUa1" to comply with the "PDF/UA-1" (ISO 14289-1) standard,
            # which aims to define represent electronic documents in PDF that allow the file to be accessible.
            # Set the "Compliance" property to "PdfCompliance.Pdf20" to comply with the "PDF 2.0" (ISO 32000-2) standard.
            # Set the "Compliance" property to "PdfCompliance.PdfA4" to comply with the "PDF/A-4" (ISO 19004:2020) standard,
            # which preserving document static visual appearance over time.
            # Set the "Compliance" property to "PdfCompliance.PdfA4Ua2" to comply with both PDF/A-4 (ISO 19005-4:2020)
            # and PDF/UA-2 (ISO 14289-2:2024) standards.
            # Set the "Compliance" property to "PdfCompliance.PdfUa2" to comply with the PDF/UA-2 (ISO 14289-2:2024) standard.
            # This helps with making documents searchable but may significantly increase the size of already large documents.
            save_options.compliance = pdf_compliance
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.Compliance.pdf', save_options=save_options)
        #ExEnd

    @unittest.skipUnless(sys.platform.startswith('win'), 'requires Windows')
    def test_text_compression(self):
        for pdf_text_compression in [aw.saving.PdfTextCompression.NONE, aw.saving.PdfTextCompression.FLATE]:
            #ExStart
            #ExFor:PdfSaveOptions
            #ExFor:PdfSaveOptions.text_compression
            #ExFor:PdfTextCompression
            #ExSummary:Shows how to apply text compression when saving a document to PDF.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            i = 0
            while i < 100:
                builder.writeln('Lorem ipsum dolor sit amet, consectetur adipiscing elit, ' + 'sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.')
                i += 1
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            options = aw.saving.PdfSaveOptions()
            # Set the "TextCompression" property to "PdfTextCompression.None" to not apply any
            # compression to text when we save the document to PDF.
            # Set the "TextCompression" property to "PdfTextCompression.Flate" to apply ZIP compression
            # to text when we save the document to PDF. The larger the document, the bigger the impact that this will have.
            options.text_compression = pdf_text_compression
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.TextCompression.pdf', save_options=options)
            #ExEnd
            file_path = ARTIFACTS_DIR + 'PdfSaveOptions.TextCompression.pdf'
            tested_file_length = system_helper.io.FileInfo(ARTIFACTS_DIR + 'PdfSaveOptions.TextCompression.pdf').length()
            switch_condition = pdf_text_compression
            if switch_condition == aw.saving.PdfTextCompression.NONE:
                self.assertTrue(tested_file_length < 69000)
                test_util.TestUtil.file_contains_string('<</Length 11 0 R>>stream', file_path)
            elif switch_condition == aw.saving.PdfTextCompression.FLATE:
                self.assertTrue(tested_file_length < 27000)
                test_util.TestUtil.file_contains_string('<</Length 11 0 R/Filter/FlateDecode>>stream', file_path)

    def test_image_compression(self):
        for pdf_image_compression in [aw.saving.PdfImageCompression.AUTO, aw.saving.PdfImageCompression.JPEG]:
            #ExStart
            #ExFor:PdfSaveOptions.image_compression
            #ExFor:PdfSaveOptions.jpeg_quality
            #ExFor:PdfImageCompression
            #ExSummary:Shows how to specify a compression type for all images in a document that we are converting to PDF.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.writeln('Jpeg image:')
            builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
            builder.insert_paragraph()
            builder.writeln('Png image:')
            builder.insert_image(file_name=IMAGE_DIR + 'Transparent background logo.png')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            pdf_save_options = aw.saving.PdfSaveOptions()
            # Set the "ImageCompression" property to "PdfImageCompression.Auto" to use the
            # "ImageCompression" property to control the quality of the Jpeg images that end up in the output PDF.
            # Set the "ImageCompression" property to "PdfImageCompression.Jpeg" to use the
            # "ImageCompression" property to control the quality of all images that end up in the output PDF.
            pdf_save_options.image_compression = pdf_image_compression
            # Set the "JpegQuality" property to "10" to strengthen compression at the cost of image quality.
            pdf_save_options.jpeg_quality = 10
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.ImageCompression.pdf', save_options=pdf_save_options)
        #ExEnd

    def test_image_color_space_export_mode(self):
        for pdf_image_color_space_export_mode in [aw.saving.PdfImageColorSpaceExportMode.AUTO, aw.saving.PdfImageColorSpaceExportMode.SIMPLE_CMYK]:
            #ExStart
            #ExFor:PdfImageColorSpaceExportMode
            #ExFor:PdfSaveOptions.image_color_space_export_mode
            #ExSummary:Shows how to set a different color space for images in a document as we export it to PDF.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.writeln('Jpeg image:')
            builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
            builder.insert_paragraph()
            builder.writeln('Png image:')
            builder.insert_image(file_name=IMAGE_DIR + 'Transparent background logo.png')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            pdf_save_options = aw.saving.PdfSaveOptions()
            # Set the "ImageColorSpaceExportMode" property to "PdfImageColorSpaceExportMode.Auto" to get Aspose.Words to
            # automatically select the color space for images in the document that it converts to PDF.
            # In most cases, the color space will be RGB.
            # Set the "ImageColorSpaceExportMode" property to "PdfImageColorSpaceExportMode.SimpleCmyk"
            # to use the CMYK color space for all images in the saved PDF.
            # Aspose.Words will also apply Flate compression to all images and ignore the "ImageCompression" property's value.
            pdf_save_options.image_color_space_export_mode = pdf_image_color_space_export_mode
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.ImageColorSpaceExportMode.pdf', save_options=pdf_save_options)
        #ExEnd

    def test_downsample_options(self):
        #ExStart
        #ExFor:DownsampleOptions
        #ExFor:DownsampleOptions.downsample_images
        #ExFor:DownsampleOptions.resolution
        #ExFor:DownsampleOptions.resolution_threshold
        #ExFor:PdfSaveOptions.downsample_options
        #ExSummary:Shows how to change the resolution of images in the PDF document.
        doc = aw.Document(file_name=MY_DIR + 'Images.docx')
        # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        # to modify how that method converts the document to .PDF.
        options = aw.saving.PdfSaveOptions()
        # By default, Aspose.Words downsample all images in a document that we save to PDF to 220 ppi.
        self.assertTrue(options.downsample_options.downsample_images)
        self.assertEqual(220, options.downsample_options.resolution)
        self.assertEqual(0, options.downsample_options.resolution_threshold)
        doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.DownsampleOptions.Default.pdf', save_options=options)
        # Set the "Resolution" property to "36" to downsample all images to 36 ppi.
        options.downsample_options.resolution = 36
        # Set the "ResolutionThreshold" property to only apply the downsampling to
        # images with a resolution that is above 128 ppi.
        options.downsample_options.resolution_threshold = 128
        # Only the first two images from the document will be downsampled at this stage.
        doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.DownsampleOptions.LowerResolution.pdf', save_options=options)
        #ExEnd

    def test_color_rendering(self):
        for color_mode in [aw.saving.ColorMode.GRAYSCALE, aw.saving.ColorMode.NORMAL]:
            #ExStart
            #ExFor:PdfSaveOptions
            #ExFor:ColorMode
            #ExFor:FixedPageSaveOptions.color_mode
            #ExSummary:Shows how to change image color with saving options property.
            doc = aw.Document(file_name=MY_DIR + 'Images.docx')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            # Set the "ColorMode" property to "Grayscale" to render all images from the document in black and white.
            # The size of the output document may be larger with this setting.
            # Set the "ColorMode" property to "Normal" to render all images in color.
            pdf_save_options = aw.saving.PdfSaveOptions()
            pdf_save_options.color_mode = color_mode
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.ColorRendering.pdf', save_options=pdf_save_options)
            #ExEnd

    def test_doc_title(self):
        for display_doc_title in [False, True]:
            #ExStart
            #ExFor:PdfSaveOptions.display_doc_title
            #ExSummary:Shows how to display the title of the document as the title bar.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.writeln('Hello world!')
            doc.built_in_document_properties.title = 'Windows bar pdf title'
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            # Set the "DisplayDocTitle" to "true" to get some PDF readers, such as Adobe Acrobat Pro,
            # to display the value of the document's "Title" built-in property in the tab that belongs to this document.
            # Set the "DisplayDocTitle" to "false" to get such readers to display the document's filename.
            pdf_save_options = aw.saving.PdfSaveOptions()
            pdf_save_options.display_doc_title = display_doc_title
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.DocTitle.pdf', save_options=pdf_save_options)
            #ExEnd

    def test_memory_optimization(self):
        for memory_optimization in [False, True]:
            #ExStart
            #ExFor:SaveOptions.create_save_options(SaveFormat)
            #ExFor:SaveOptions.memory_optimization
            #ExSummary:Shows an option to optimize memory consumption when rendering large documents to PDF.
            doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            save_options = aw.saving.SaveOptions.create_save_options(save_format=aw.SaveFormat.PDF)
            # Set the "MemoryOptimization" property to "true" to lower the memory footprint of large documents' saving operations
            # at the cost of increasing the duration of the operation.
            # Set the "MemoryOptimization" property to "false" to save the document as a PDF normally.
            save_options.memory_optimization = memory_optimization
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.MemoryOptimization.pdf', save_options=save_options)
            #ExEnd

    def test_escape_uri(self):
        for uri, result in [('https://www.google.com/search?q= aspose', 'https://www.google.com/search?q=%20aspose'), ('https://www.google.com/search?q=%20aspose', 'https://www.google.com/search?q=%20aspose')]:
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.insert_hyperlink('Testlink', uri, False)
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.EscapedUri.pdf')

    def test_header_footer_bookmarks_export_mode(self):
        for header_footer_bookmarks_export_mode in [aw.saving.HeaderFooterBookmarksExportMode.NONE, aw.saving.HeaderFooterBookmarksExportMode.FIRST, aw.saving.HeaderFooterBookmarksExportMode.ALL]:
            #ExStart
            #ExFor:HeaderFooterBookmarksExportMode
            #ExFor:OutlineOptions
            #ExFor:OutlineOptions.default_bookmarks_outline_level
            #ExFor:PdfSaveOptions.header_footer_bookmarks_export_mode
            #ExFor:PdfSaveOptions.page_mode
            #ExFor:PdfPageMode
            #ExSummary:Shows to process bookmarks in headers/footers in a document that we are rendering to PDF.
            doc = aw.Document(file_name=MY_DIR + 'Bookmarks in headers and footers.docx')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            save_options = aw.saving.PdfSaveOptions()
            # Set the "PageMode" property to "PdfPageMode.UseOutlines" to display the outline navigation pane in the output PDF.
            save_options.page_mode = aw.saving.PdfPageMode.USE_OUTLINES
            # Set the "DefaultBookmarksOutlineLevel" property to "1" to display all
            # bookmarks at the first level of the outline in the output PDF.
            save_options.outline_options.default_bookmarks_outline_level = 1
            # Set the "HeaderFooterBookmarksExportMode" property to "HeaderFooterBookmarksExportMode.None" to
            # not export any bookmarks that are inside headers/footers.
            # Set the "HeaderFooterBookmarksExportMode" property to "HeaderFooterBookmarksExportMode.First" to
            # only export bookmarks in the first section's header/footers.
            # Set the "HeaderFooterBookmarksExportMode" property to "HeaderFooterBookmarksExportMode.All" to
            # export bookmarks that are in all headers/footers.
            save_options.header_footer_bookmarks_export_mode = header_footer_bookmarks_export_mode
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf', save_options=save_options)
        #ExEnd

    def test_emulate_rendering_to_size_on_page(self):
        for render_to_size in [False, True]:
            #ExStart
            #ExFor:MetafileRenderingOptions.emulate_rendering_to_size_on_page
            #ExFor:MetafileRenderingOptions.emulate_rendering_to_size_on_page_resolution
            #ExSummary:Shows how to display of the metafile according to the size on page.
            doc = aw.Document(file_name=MY_DIR + 'WMF with text.docx')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            save_options = aw.saving.PdfSaveOptions()
            # Set the "EmulateRenderingToSizeOnPage" property to "true"
            # to emulate rendering according to the metafile size on page.
            # Set the "EmulateRenderingToSizeOnPage" property to "false"
            # to emulate metafile rendering to its default size in pixels.
            save_options.metafile_rendering_options.emulate_rendering_to_size_on_page = render_to_size
            save_options.metafile_rendering_options.emulate_rendering_to_size_on_page_resolution = 50
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.EmulateRenderingToSizeOnPage.pdf', save_options=save_options)
            #ExEnd

    @unittest.skipUnless(sys.platform.startswith('win'), 'requires Windows')
    def test_embed_windows_fonts(self):
        for pdf_font_embedding_mode in [aw.saving.PdfFontEmbeddingMode.EMBED_ALL, aw.saving.PdfFontEmbeddingMode.EMBED_NONE, aw.saving.PdfFontEmbeddingMode.EMBED_NONSTANDARD]:
            #ExStart
            #ExFor:PdfSaveOptions.font_embedding_mode
            #ExFor:PdfFontEmbeddingMode
            #ExSummary:Shows how to set Aspose.Words to skip embedding Arial and Times New Roman fonts into a PDF document.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            # "Arial" is a standard font, and "Courier New" is a nonstandard font.
            builder.font.name = 'Arial'
            builder.writeln('Hello world!')
            builder.font.name = 'Courier New'
            builder.writeln('The quick brown fox jumps over the lazy dog.')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            options = aw.saving.PdfSaveOptions()
            # Set the "EmbedFullFonts" property to "true" to embed every glyph of every embedded font in the output PDF.
            options.embed_full_fonts = True
            # Set the "FontEmbeddingMode" property to "EmbedAll" to embed all fonts in the output PDF.
            # Set the "FontEmbeddingMode" property to "EmbedNonstandard" to only allow nonstandard fonts' embedding in the output PDF.
            # Set the "FontEmbeddingMode" property to "EmbedNone" to not embed any fonts in the output PDF.
            options.font_embedding_mode = pdf_font_embedding_mode
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.EmbedWindowsFonts.pdf', save_options=options)
            #ExEnd
            tested_file_length = system_helper.io.FileInfo(ARTIFACTS_DIR + 'PdfSaveOptions.EmbedWindowsFonts.pdf').length()
            switch_condition = pdf_font_embedding_mode
            if switch_condition == aw.saving.PdfFontEmbeddingMode.EMBED_ALL:
                self.assertTrue(tested_file_length < 1040000)
            elif switch_condition == aw.saving.PdfFontEmbeddingMode.EMBED_NONSTANDARD:
                self.assertTrue(tested_file_length < 492000)
            elif switch_condition == aw.saving.PdfFontEmbeddingMode.EMBED_NONE:
                self.assertTrue(tested_file_length < 4300)

    @unittest.skipUnless(sys.platform.startswith('win'), 'requires Windows')
    def test_embed_core_fonts(self):
        for use_core_fonts in [False, True]:
            #ExStart
            #ExFor:PdfSaveOptions.use_core_fonts
            #ExSummary:Shows how enable/disable PDF Type 1 font substitution.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.font.name = 'Arial'
            builder.writeln('Hello world!')
            builder.font.name = 'Courier New'
            builder.writeln('The quick brown fox jumps over the lazy dog.')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            options = aw.saving.PdfSaveOptions()
            # Set the "UseCoreFonts" property to "true" to replace some fonts,
            # including the two fonts in our document, with their PDF Type 1 equivalents.
            # Set the "UseCoreFonts" property to "false" to not apply PDF Type 1 fonts.
            options.use_core_fonts = use_core_fonts
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.EmbedCoreFonts.pdf', save_options=options)
            #ExEnd
            tested_file_length = system_helper.io.FileInfo(ARTIFACTS_DIR + 'PdfSaveOptions.EmbedCoreFonts.pdf').length()
            if use_core_fonts:
                self.assertTrue(tested_file_length < 2000)
            else:
                self.assertTrue(tested_file_length < 33500)

    def test_additional_text_positioning(self):
        for apply_additional_text_positioning in [False, True]:
            #ExStart
            #ExFor:PdfSaveOptions.additional_text_positioning
            #ExSummary:Show how to write additional text positioning operators.
            doc = aw.Document(file_name=MY_DIR + 'Text positioning operators.docx')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            save_options = aw.saving.PdfSaveOptions()
            save_options.text_compression = aw.saving.PdfTextCompression.NONE
            # Set the "AdditionalTextPositioning" property to "true" to attempt to fix incorrect
            # element positioning in the output PDF, should there be any, at the cost of increased file size.
            # Set the "AdditionalTextPositioning" property to "false" to render the document as usual.
            save_options.additional_text_positioning = apply_additional_text_positioning
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.AdditionalTextPositioning.pdf', save_options=save_options)
            #ExEnd

    def test_save_as_pdf_book_fold(self):
        for render_text_as_bookfold in [False, True]:
            #ExStart
            #ExFor:PdfSaveOptions.use_book_fold_printing_settings
            #ExSummary:Shows how to save a document to the PDF format in the form of a book fold.
            doc = aw.Document(file_name=MY_DIR + 'Paragraphs.docx')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            options = aw.saving.PdfSaveOptions()
            # Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
            # in the output PDF in a way that helps us use it to make a booklet.
            # Set the "UseBookFoldPrintingSettings" property to "false" to render the PDF normally.
            options.use_book_fold_printing_settings = render_text_as_bookfold
            # If we are rendering the document as a booklet, we must set the "MultiplePages"
            # properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
            if render_text_as_bookfold:
                for s in doc.sections:
                    s = s.as_section()
                    s.page_setup.multiple_pages = aw.settings.MultiplePagesType.BOOK_FOLD_PRINTING
            # Once we print this document on both sides of the pages, we can fold all the pages down the middle at once,
            # and the contents will line up in a way that creates a booklet.
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.SaveAsPdfBookFold.pdf', save_options=options)
            #ExEnd

    def test_zoom_behaviour(self):
        #ExStart
        #ExFor:PdfSaveOptions.zoom_behavior
        #ExFor:PdfSaveOptions.zoom_factor
        #ExFor:PdfZoomBehavior
        #ExSummary:Shows how to set the default zooming that a reader applies when opening a rendered PDF document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Hello world!')
        # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        # to modify how that method converts the document to .PDF.
        # Set the "ZoomBehavior" property to "PdfZoomBehavior.ZoomFactor" to get a PDF reader to
        # apply a percentage-based zoom factor when we open the document with it.
        # Set the "ZoomFactor" property to "25" to give the zoom factor a value of 25%.
        options = aw.saving.PdfSaveOptions()
        options.zoom_behavior = aw.saving.PdfZoomBehavior.ZOOM_FACTOR
        options.zoom_factor = 25
        # When we open this document using a reader such as Adobe Acrobat, we will see the document scaled at 1/4 of its actual size.
        doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.ZoomBehaviour.pdf', save_options=options)
        #ExEnd

    def test_drawing_ml_effects(self):
        for effects_rendering_mode in [aw.saving.DmlEffectsRenderingMode.NONE, aw.saving.DmlEffectsRenderingMode.SIMPLIFIED, aw.saving.DmlEffectsRenderingMode.FINE]:
            #ExStart
            #ExFor:DmlRenderingMode
            #ExFor:DmlEffectsRenderingMode
            #ExFor:PdfSaveOptions.dml_effects_rendering_mode
            #ExFor:SaveOptions.dml_effects_rendering_mode
            #ExFor:SaveOptions.dml_rendering_mode
            #ExSummary:Shows how to configure the rendering quality of DrawingML effects in a document as we save it to PDF.
            doc = aw.Document(file_name=MY_DIR + 'DrawingML shape effects.docx')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            options = aw.saving.PdfSaveOptions()
            # Set the "DmlEffectsRenderingMode" property to "DmlEffectsRenderingMode.None" to discard all DrawingML effects.
            # Set the "DmlEffectsRenderingMode" property to "DmlEffectsRenderingMode.Simplified"
            # to render a simplified version of DrawingML effects.
            # Set the "DmlEffectsRenderingMode" property to "DmlEffectsRenderingMode.Fine" to
            # render DrawingML effects with more accuracy and also with more processing cost.
            options.dml_effects_rendering_mode = effects_rendering_mode
            self.assertEqual(aw.saving.DmlRenderingMode.DRAWING_ML, options.dml_rendering_mode)
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.DrawingMLEffects.pdf', save_options=options)
        #ExEnd

    def test_drawing_ml_fallback(self):
        for dml_rendering_mode in [aw.saving.DmlRenderingMode.FALLBACK, aw.saving.DmlRenderingMode.DRAWING_ML]:
            #ExStart
            #ExFor:DmlRenderingMode
            #ExFor:SaveOptions.dml_rendering_mode
            #ExSummary:Shows how to render fallback shapes when saving to PDF.
            doc = aw.Document(file_name=MY_DIR + 'DrawingML shape fallbacks.docx')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            options = aw.saving.PdfSaveOptions()
            # Set the "DmlRenderingMode" property to "DmlRenderingMode.Fallback"
            # to substitute DML shapes with their fallback shapes.
            # Set the "DmlRenderingMode" property to "DmlRenderingMode.DrawingML"
            # to render the DML shapes themselves.
            options.dml_rendering_mode = dml_rendering_mode
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.DrawingMLFallback.pdf', save_options=options)
            #ExEnd
            switch_condition = dml_rendering_mode
            if switch_condition == aw.saving.DmlRenderingMode.DRAWING_ML:
                test_util.TestUtil.file_contains_string('<</Type/Page/Parent 3 0 R/Contents 6 0 R/MediaBox[0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABC 12 0 R>>>>/Group<</Type/Group/S/Transparency/CS/DeviceRGB>>>>', ARTIFACTS_DIR + 'PdfSaveOptions.DrawingMLFallback.pdf')
            elif switch_condition == aw.saving.DmlRenderingMode.FALLBACK:
                test_util.TestUtil.file_contains_string('<</Type/Page/Parent 3 0 R/Contents 6 0 R/MediaBox[0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABE 14 0 R>>/ExtGState<</GS1 11 0 R/GS2 12 0 R/GS3 17 0 R>>>>/Group<</Type/Group/S/Transparency/CS/DeviceRGB>>>>', ARTIFACTS_DIR + 'PdfSaveOptions.DrawingMLFallback.pdf')

    def test_export_document_structure(self):
        for export_document_structure in [False, True]:
            #ExStart
            #ExFor:PdfSaveOptions.export_document_structure
            #ExSummary:Shows how to preserve document structure elements, which can assist in programmatically interpreting our document.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.paragraph_format.style = doc.styles.get_by_name('Heading 1')
            builder.writeln('Hello world!')
            builder.paragraph_format.style = doc.styles.get_by_name('Normal')
            builder.write('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            options = aw.saving.PdfSaveOptions()
            # Set the "ExportDocumentStructure" property to "true" to make the document structure, such tags, available via the
            # "Content" navigation pane of Adobe Acrobat at the cost of increased file size.
            # Set the "ExportDocumentStructure" property to "false" to not export the document structure.
            options.export_document_structure = export_document_structure
            # Suppose we export document structure while saving this document. In that case,
            # we can open it using Adobe Acrobat and find tags for elements such as the heading
            # and the next paragraph via "View" -> "Show/Hide" -> "Navigation panes" -> "Tags".
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.ExportDocumentStructure.pdf', save_options=options)
            #ExEnd
            if export_document_structure:
                test_util.TestUtil.file_contains_string('<</Type/Page/Parent 3 0 R/Contents 6 0 R/MediaBox[0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABD 13 0 R>>/ExtGState<</GS1 11 0 R/GS2 16 0 R>>>>/Group<</Type/Group/S/Transparency/CS/DeviceRGB>>/StructParents 0/Tabs/S>>', ARTIFACTS_DIR + 'PdfSaveOptions.ExportDocumentStructure.pdf')
            else:
                test_util.TestUtil.file_contains_string('<</Type/Page/Parent 3 0 R/Contents 6 0 R/MediaBox[0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABC 12 0 R>>>>/Group<</Type/Group/S/Transparency/CS/DeviceRGB>>>>', ARTIFACTS_DIR + 'PdfSaveOptions.ExportDocumentStructure.pdf')

    @unittest.skip("drawing.Image type isn't supported yet")
    def test_interpolate_images(self):
        for interpolate_images in [False, True]:
            #ExStart
            #ExFor:PdfSaveOptions.interpolate_images
            #ExSummary:Shows how to perform interpolation on images while saving a document to PDF.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.insert_image(file_name=IMAGE_DIR + 'Transparent background logo.png')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            save_options = aw.saving.PdfSaveOptions()
            # Set the "InterpolateImages" property to "true" to get the reader that opens this document to interpolate images.
            # Their resolution should be lower than that of the device that is displaying the document.
            # Set the "InterpolateImages" property to "false" to make it so that the reader does not apply any interpolation.
            save_options.interpolate_images = interpolate_images
            # When we open this document with a reader such as Adobe Acrobat, we will need to zoom in on the image
            # to see the interpolation effect if we saved the document with it enabled.
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.InterpolateImages.pdf', save_options=save_options)
            #ExEnd
            if interpolate_images:
                test_util.TestUtil.file_contains_string('<</Type/XObject/Subtype/Image/Width 400/Height 400/ColorSpace/DeviceRGB/BitsPerComponent 8/SMask 10 0 R/Interpolate true/Length 11 0 R/Filter/FlateDecode>>', ARTIFACTS_DIR + 'PdfSaveOptions.InterpolateImages.pdf')
            else:
                test_util.TestUtil.file_contains_string('<</Type/XObject/Subtype/Image/Width 400/Height 400/ColorSpace/DeviceRGB/BitsPerComponent 8/SMask 10 0 R/Length 11 0 R/Filter/FlateDecode>>', ARTIFACTS_DIR + 'PdfSaveOptions.InterpolateImages.pdf')

    def test_render_metafile(self):
        for rendering_mode in [aw.saving.EmfPlusDualRenderingMode.EMF, aw.saving.EmfPlusDualRenderingMode.EMF_PLUS, aw.saving.EmfPlusDualRenderingMode.EMF_PLUS_WITH_FALLBACK]:
            #ExStart
            #ExFor:EmfPlusDualRenderingMode
            #ExFor:MetafileRenderingOptions.emf_plus_dual_rendering_mode
            #ExFor:MetafileRenderingOptions.use_emf_embedded_to_wmf
            #ExSummary:Shows how to configure Enhanced Windows Metafile-related rendering options when saving to PDF.
            doc = aw.Document(file_name=MY_DIR + 'EMF.docx')
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            save_options = aw.saving.PdfSaveOptions()
            # Set the "EmfPlusDualRenderingMode" property to "EmfPlusDualRenderingMode.Emf"
            # to only render the EMF part of an EMF+ dual metafile.
            # Set the "EmfPlusDualRenderingMode" property to "EmfPlusDualRenderingMode.EmfPlus" to
            # to render the EMF+ part of an EMF+ dual metafile.
            # Set the "EmfPlusDualRenderingMode" property to "EmfPlusDualRenderingMode.EmfPlusWithFallback"
            # to render the EMF+ part of an EMF+ dual metafile if all of the EMF+ records are supported.
            # Otherwise, Aspose.Words will render the EMF part.
            save_options.metafile_rendering_options.emf_plus_dual_rendering_mode = rendering_mode
            # Set the "UseEmfEmbeddedToWmf" property to "true" to render embedded EMF data
            # for metafiles that we can render as vector graphics.
            save_options.metafile_rendering_options.use_emf_embedded_to_wmf = True
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.RenderMetafile.pdf', save_options=save_options)
        #ExEnd

    def test_encryption_permissions(self):
        #ExStart
        #ExFor:PdfEncryptionDetails.__init__(str,str,PdfPermissions)
        #ExFor:PdfSaveOptions.encryption_details
        #ExFor:PdfEncryptionDetails.permissions
        #ExFor:PdfEncryptionDetails.owner_password
        #ExFor:PdfEncryptionDetails.user_password
        #ExFor:PdfPermissions
        #ExFor:PdfEncryptionDetails
        #ExSummary:Shows how to set permissions on a saved PDF document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Hello world!')
        # Extend permissions to allow the editing of annotations.
        encryption_details = aw.saving.PdfEncryptionDetails(user_password='password', owner_password='', permissions=aw.saving.PdfPermissions.MODIFY_ANNOTATIONS | aw.saving.PdfPermissions.DOCUMENT_ASSEMBLY)
        # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        # to modify how that method converts the document to .PDF.
        save_options = aw.saving.PdfSaveOptions()
        # Enable encryption via the "EncryptionDetails" property.
        save_options.encryption_details = encryption_details
        # When we open this document, we will need to provide the password before accessing its contents.
        doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.EncryptionPermissions.pdf', save_options=save_options)
        #ExEnd

    def test_export_language_to_span_tag(self):
        #ExStart
        #ExFor:PdfSaveOptions.export_language_to_span_tag
        #ExSummary:Shows how to create a "Span" tag in the document structure to export the text language.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Hello world!')
        builder.writeln('Hola mundo!')
        save_options = aw.saving.PdfSaveOptions()
        # Note, when "ExportDocumentStructure" is false, "ExportLanguageToSpanTag" is ignored.
        save_options.export_document_structure = True
        save_options.export_language_to_span_tag = True
        doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.ExportLanguageToSpanTag.pdf', save_options=save_options)
        #ExEnd

    def test_attachments_embedding_mode(self):
        #ExStart:AttachmentsEmbeddingMode
        #ExFor:PdfSaveOptions.attachments_embedding_mode
        #ExSummary:Shows how to add embed attachments to the PDF document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.insert_ole_object(file_name=MY_DIR + 'Spreadsheet.xlsx', prog_id='Excel.Sheet', is_linked=False, as_icon=True, presentation=None)
        save_options = aw.saving.PdfSaveOptions()
        save_options.attachments_embedding_mode = aw.saving.PdfAttachmentsEmbeddingMode.ANNOTATIONS
        doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.AttachmentsEmbeddingMode.pdf', save_options=save_options)
        #ExEnd:AttachmentsEmbeddingMode

    def test_cache_background_graphics(self):
        #ExStart
        #ExFor:PdfSaveOptions.cache_background_graphics
        #ExSummary:Shows how to cache graphics placed in document's background.
        doc = aw.Document(file_name=MY_DIR + 'Background images.docx')
        save_options = aw.saving.PdfSaveOptions()
        save_options.cache_background_graphics = True
        doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.CacheBackgroundGraphics.pdf', save_options=save_options)
        aspose_to_pdf_size = system_helper.io.FileInfo(ARTIFACTS_DIR + 'PdfSaveOptions.CacheBackgroundGraphics.pdf').length()
        word_to_pdf_size = system_helper.io.FileInfo(MY_DIR + 'Background images (word to pdf).pdf').length()
        self.assertLess(aspose_to_pdf_size, word_to_pdf_size)
        #ExEnd

    def test_export_paragraph_graphics_to_artifact(self):
        #ExStart
        #ExFor:PdfSaveOptions.export_paragraph_graphics_to_artifact
        #ExSummary:Shows how to export paragraph graphics as artifact (underlines, text emphasis, etc.).
        doc = aw.Document(file_name=MY_DIR + 'PDF artifacts.docx')
        save_options = aw.saving.PdfSaveOptions()
        save_options.export_document_structure = True
        save_options.export_paragraph_graphics_to_artifact = True
        save_options.text_compression = aw.saving.PdfTextCompression.NONE
        doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.ExportParagraphGraphicsToArtifact.pdf', save_options=save_options)
        #ExEnd

    def test_page_layout(self):
        #ExStart:PageLayout
        #ExFor:PdfSaveOptions.page_layout
        #ExFor:PdfPageLayout
        #ExSummary:Shows how to display pages when opened in a PDF reader.
        doc = aw.Document(file_name=MY_DIR + 'Big document.docx')
        # Display the pages two at a time, with odd-numbered pages on the left.
        save_options = aw.saving.PdfSaveOptions()
        save_options.page_layout = aw.saving.PdfPageLayout.TWO_PAGE_LEFT
        doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.PageLayout.pdf', save_options=save_options)
        #ExEnd:PageLayout

    def test_sdt_tag_as_form_field_name(self):
        #ExStart:SdtTagAsFormFieldName
        #ExFor:PdfSaveOptions.use_sdt_tag_as_form_field_name
        #ExSummary:Shows how to use SDT control Tag or Id property as a name of form field in PDF.
        doc = aw.Document(file_name=MY_DIR + 'Form fields.docx')
        save_options = aw.saving.PdfSaveOptions()
        save_options.preserve_form_fields = True
        # When set to 'false', SDT control Id property is used as a name of form field in PDF.
        # When set to 'true', SDT control Tag property is used as a name of form field in PDF.
        save_options.use_sdt_tag_as_form_field_name = True
        doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.SdtTagAsFormFieldName.pdf', save_options=save_options)
        #ExEnd:SdtTagAsFormFieldName

    def test_render_choice_form_field_border(self):
        #ExStart:RenderChoiceFormFieldBorder
        #ExFor:PdfSaveOptions.render_choice_form_field_border
        #ExSummary:Shows how to render PDF choice form field border.
        doc = aw.Document(file_name=MY_DIR + 'Legacy drop-down.docx')
        save_options = aw.saving.PdfSaveOptions()
        save_options.preserve_form_fields = True
        save_options.render_choice_form_field_border = True
        doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.RenderChoiceFormFieldBorder.pdf', save_options=save_options)
        #ExEnd:RenderChoiceFormFieldBorder

    def test_one_page(self):
        #ExStart
        #ExFor:FixedPageSaveOptions.page_set
        #ExFor:Document.save(BytesIO,SaveOptions)
        #ExSummary:Shows how to convert only some of the pages in a document to PDF.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln('Page 1.')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln('Page 2.')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln('Page 3.')
        with open(ARTIFACTS_DIR + 'PdfSaveOptions.one_page.pdf', 'wb') as stream:
            # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
            # to modify how that method converts the document to .PDF.
            options = aw.saving.PdfSaveOptions()
            # Set the "page_index" to "1" to render a portion of the document starting from the second page.
            options.page_set = aw.saving.PageSet(1)
            # This document will contain one page starting from page two, which will only contain the second page.
            doc.save(stream, options)
        #ExEnd

    @unittest.skip("Discrepancy in assertion between Python and .Net")
    def test_open_hyperlinks_in_new_window(self):
        for open_hyperlinks_in_new_window in [False, True]:
            #ExStart
            #ExFor:PdfSaveOptions.open_hyperlinks_in_new_window
            #ExSummary:Shows how to save hyperlinks in a document we convert to PDF so that they open new pages when we click on them.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.insert_hyperlink('Testlink', 'https://www.google.com/search?q=%20aspose', False)
            # Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .PDF.
            options = aw.saving.PdfSaveOptions()
            # Set the "OpenHyperlinksInNewWindow" property to "true" to save all hyperlinks using Javascript code
            # that forces readers to open these links in new windows/browser tabs.
            # Set the "OpenHyperlinksInNewWindow" property to "false" to save all hyperlinks normally.
            options.open_hyperlinks_in_new_window = open_hyperlinks_in_new_window
            doc.save(file_name=ARTIFACTS_DIR + 'PdfSaveOptions.OpenHyperlinksInNewWindow.pdf', save_options=options)
            #ExEnd
            if open_hyperlinks_in_new_window:
                test_util.TestUtil.file_contains_string('<</Type/Annot/Subtype/Link/Rect[72 706.20098877 111.32800293 720]/BS' + '<</Type/Border/S/S/W 0>>/A<</Type/Action/S/JavaScript/JS(app.launchURL\\("https://www.google.com/search?q=%20aspose", true\\);)>>>>', ARTIFACTS_DIR + 'PdfSaveOptions.OpenHyperlinksInNewWindow.pdf')
            else:
                test_util.TestUtil.file_contains_string('<</Type/Annot/Subtype/Link/Rect[72 706.20098877 111.32800293 720]/BS' + '<</Type/Border/S/S/W 0>>/A<</Type/Action/S/URI/URI(https://www.google.com/search?q=%20aspose)>>>>', ARTIFACTS_DIR + 'PdfSaveOptions.OpenHyperlinksInNewWindow.pdf')

    @unittest.skip("Discrepancy in assertion between Python and .Net")
    def test_embed_full_fonts(self):
        for embed_full_fonts in (False, True):
            with self.subTest(embed_full_fonts=embed_full_fonts):
                #ExStart
                #ExFor:PdfSaveOptions.__init__
                #ExFor:PdfSaveOptions.embed_full_fonts
                #ExSummary:Shows how to enable or disable subsetting when embedding fonts while rendering a document to PDF.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.font.name = 'Arial'
                builder.writeln('Hello world!')
                builder.font.name = 'Arvo'
                builder.writeln('The quick brown fox jumps over the lazy dog.')
                # Configure our font sources to ensure that we have access to both the fonts in this document.
                original_fonts_sources = aw.fonts.FontSettings.default_instance.get_fonts_sources()
                folder_font_source = aw.fonts.FolderFontSource(FONTS_DIR, True)
                aw.fonts.FontSettings.default_instance.set_fonts_sources([original_fonts_sources[0], folder_font_source])
                font_sources = aw.fonts.FontSettings.default_instance.get_fonts_sources()
                self.assertTrue(any((font.full_font_name == 'Arial' for font in font_sources[0].get_available_fonts())))
                self.assertTrue(any((font.full_font_name == 'Arvo' for font in font_sources[1].get_available_fonts())))
                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()
                # Since our document contains a custom font, embedding in the output document may be desirable.
                # Set the "embed_full_fonts" property to "True" to embed every glyph of every embedded font in the output PDF.
                # The document's size may become very large, but we will have full use of all fonts if we edit the PDF.
                # Set the "embed_full_fonts" property to "False" to apply subsetting to fonts, saving only the glyphs
                # that the document is using. The file will be considerably smaller,
                # but we may need access to any custom fonts if we edit the document.
                options.embed_full_fonts = embed_full_fonts
                doc.save(ARTIFACTS_DIR + 'PdfSaveOptions.embed_full_fonts.pdf', options)
                if embed_full_fonts:
                    self.assertLess(500000, os.path.getsize(ARTIFACTS_DIR + 'PdfSaveOptions.embed_full_fonts.pdf'))
                else:
                    self.assertGreater(25000, os.path.getsize(ARTIFACTS_DIR + 'PdfSaveOptions.embed_full_fonts.pdf'))
                # Restore the original font sources.
                aw.fonts.FontSettings.default_instance.set_fonts_sources(original_fonts_sources)
                #ExEnd

    @unittest.skip("system.globalization.CultureInfo type isn't supported yet")
    def test_page_mode(self):
        for page_mode in (aw.saving.PdfPageMode.FULL_SCREEN, aw.saving.PdfPageMode.USE_THUMBS, aw.saving.PdfPageMode.USE_OC, aw.saving.PdfPageMode.USE_OUTLINES, aw.saving.PdfPageMode.USE_NONE, aw.saving.PdfPageMode.USE_ATTACHMENTS):
            with self.subTest(page_mode=page_mode):
                #ExStart
                #ExFor:PdfSaveOptions.page_mode
                #ExFor:PdfPageMode
                #ExSummary:Shows how to set instructions for some PDF readers to follow when opening an output document.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.writeln('Hello world!')
                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()
                # Set the "page_mode" property to "PdfPageMode.FULL_SCREEN" to get the PDF reader to open the saved
                # document in full-screen mode, which takes over the monitor's display and has no controls visible.
                # Set the "page_mode" property to "PdfPageMode.USE_THUMBS" to get the PDF reader to display a separate panel
                # with a thumbnail for each page in the document.
                # Set the "page_mode" property to "PdfPageMode.USE_OC" to get the PDF reader to display a separate panel
                # that allows us to work with any layers present in the document.
                # Set the "page_mode" property to "PdfPageMode.USE_OUTLINES" to get the PDF reader
                # also to display the outline, if possible.
                # Set the "page_mode" property to "PdfPageMode.USE_NONE" to get the PDF reader to display just the document itself.
                # Set the "page_mode" property to "PdfPageMode.USE_ATTACHMENTS" to make visible attachments panel.
                options.page_mode = page_mode
                doc.save(ARTIFACTS_DIR + 'PdfSaveOptions.page_mode.pdf', options)
                #ExEnd
                doc_locale_name = CultureInfo(doc.styles.default_font.locale_id).name
                with open(ARTIFACTS_DIR + 'PdfSaveOptions.page_mode.pdf', 'rb') as file:
                    content = file.read().decode('utf-8')
                    if page_mode == aw.saving.PdfPageMode.FULL_SCREEN:
                        self.assertIn('<</Type /Catalog/Pages 3 0 R/PageMode /FullScreen/Lang({})/Metadata 4 0 R>>\r\n'.format(doc_locale_name), content)
                    elif page_mode == aw.saving.PdfPageMode.USE_THUMBS:
                        self.assertIn('<</Type /Catalog/Pages 3 0 R/PageMode /UseThumbs/Lang({})/Metadata 4 0 R>>'.format(doc_locale_name), content)
                    elif page_mode == aw.saving.PdfPageMode.USE_OC:
                        self.assertIn('<</Type /Catalog/Pages 3 0 R/PageMode /UseOC/Lang({})/Metadata 4 0 R>>\r\n'.format(doc_locale_name), content)
                    elif page_mode in (aw.saving.PdfPageMode.USE_OUTLINES, aw.saving.PdfPageMode.USE_NONE):
                        self.assertIn('<</Type /Catalog/Pages 3 0 R/Lang({})/Metadata 4 0 R>>\r\n'.format(doc_locale_name), content)
                    elif page_mode == aw.saving.PdfPageMode.USE_ATTACHMENTS:
                        self.assertIn(f'<</Type /Catalog/Pages 3 0 R/PageMode /UseAttachments/Lang({doc_locale_name})/Metadata 4 0 R>>\r\n', content)

    def test_note_hyperlinks(self):
        for create_note_hyperlinks in (False, True):
            with self.subTest(create_note_hyperlinks=create_note_hyperlinks):
                #ExStart
                #ExFor:PdfSaveOptions.create_note_hyperlinks
                #ExSummary:Shows how to make footnotes and endnotes function as hyperlinks.
                doc = aw.Document(MY_DIR + 'Footnotes and endnotes.docx')
                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()
                # Set the "create_note_hyperlinks" property to "True" to turn all footnote/endnote symbols
                # in the text act as links that, upon clicking, take us to their respective footnotes/endnotes.
                # Set the "create_note_hyperlinks" property to "False" not to have footnote/endnote symbols link to anything.
                options.create_note_hyperlinks = create_note_hyperlinks
                doc.save(ARTIFACTS_DIR + 'PdfSaveOptions.note_hyperlinks.pdf', options)
                #ExEnd
                with open(ARTIFACTS_DIR + 'PdfSaveOptions.note_hyperlinks.pdf', 'rb') as file:
                    content = file.read()
                if create_note_hyperlinks:
                    self.assertIn(b'<</Type/Annot/Subtype/Link/Rect[157.80099487 720.90106201 159.35600281 733.55004883]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 677 0]>>', content)
                    self.assertIn(b'<</Type/Annot/Subtype/Link/Rect[202.16900635 720.90106201 206.06201172 733.55004883]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 79 0]>>', content)
                    self.assertIn(b'<</Type/Annot/Subtype/Link/Rect[212.23199463 699.2510376 215.34199524 711.90002441]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 654 0]>>', content)
                    self.assertIn(b'<</Type/Annot/Subtype/Link/Rect[258.15499878 699.2510376 262.04800415 711.90002441]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 68 0]>>', content)
                    self.assertIn(b'<</Type/Annot/Subtype/Link/Rect[85.05000305 68.19904327 88.66500092 79.69804382]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 202 733 0]>>', content)
                    self.assertIn(b'<</Type/Annot/Subtype/Link/Rect[85.05000305 56.70004272 88.66500092 68.19904327]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 258 711 0]>>', content)
                    self.assertIn(b'<</Type/Annot/Subtype/Link/Rect[85.05000305 666.10205078 86.4940033 677.60107422]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 157 733 0]>>', content)
                    self.assertIn(b'<</Type/Annot/Subtype/Link/Rect[85.05000305 643.10406494 87.93800354 654.60308838]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 212 711 0]>>', content)
                else:
                    self.assertNotIn(b'<</Type/Annot/Subtype/Link/Rect', content)

    def test_custom_properties_export(self):
        for pdf_custom_properties_export_mode in (aw.saving.PdfCustomPropertiesExport.NONE, aw.saving.PdfCustomPropertiesExport.STANDARD, aw.saving.PdfCustomPropertiesExport.METADATA):
            with self.subTest(pdf_custom_properties_export_mode=pdf_custom_properties_export_mode):
                #ExStart
                #ExFor:PdfCustomPropertiesExport
                #ExFor:PdfSaveOptions.custom_properties_export
                #ExSummary:Shows how to export custom properties while converting a document to PDF.
                doc = aw.Document()
                doc.custom_document_properties.add('Company', 'My value')
                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()
                # Set the "custom_properties_export" property to "PdfCustomPropertiesExport.NONE" to discard
                # custom document properties as we save the document to .PDF.
                # Set the "custom_properties_export" property to "PdfCustomPropertiesExport.STANDARD"
                # to preserve custom properties within the output PDF document.
                # Set the "custom_properties_export" property to "PdfCustomPropertiesExport.METADATA"
                # to preserve custom properties in an XMP packet.
                options.custom_properties_export = pdf_custom_properties_export_mode
                doc.save(ARTIFACTS_DIR + 'PdfSaveOptions.custom_properties_export.pdf', options)
                #ExEnd
                with open(ARTIFACTS_DIR + 'PdfSaveOptions.custom_properties_export.pdf', 'rb') as file:
                    content = file.read()
                if pdf_custom_properties_export_mode == aw.saving.PdfCustomPropertiesExport.NONE:
                    self.assertNotIn(doc.custom_document_properties[0].name.encode('ascii'), content)
                    self.assertNotIn(b'<</Type/Metadata/Subtype/XML/Length 8 0 R/Filter/FlateDecode>>', content)
                elif pdf_custom_properties_export_mode == aw.saving.PdfCustomPropertiesExport.STANDARD:
                    self.assertIn(b'<</Creator(\xfe\xff\x00A\x00s\x00p\x00o\x00s\x00e\x00.\x00W\x00o\x00r\x00d\x00s)/Producer(\xfe\xff\x00A\x00s\x00p\x00o\x00s\x00e\x00.\x00W\x00o\x00r\x00d\x00s\x00 \x00f\x00o\x00r\x00', content)
                    self.assertIn(b'/Company(\xfe\xff\x00M\x00y\x00 \x00v\x00a\x00l\x00u\x00e)>>', content)
                elif pdf_custom_properties_export_mode == aw.saving.PdfCustomPropertiesExport.METADATA:
                    self.assertIn(b'<</Type/Metadata/Subtype/XML/Length 8 0 R/Filter/FlateDecode>>', content)

    @unittest.skip("drawing.Image type isn't supported yet")
    def test_preblend_images(self):
        for preblend_images in (False, True):
            with self.subTest(preblend_images=preblend_images):
                #ExStart
                #ExFor:PdfSaveOptions.preblend_images
                #ExSummary:Shows how to preblend images with transparent backgrounds while saving a document to PDF.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                img = drawing.Image.from_file(IMAGE_DIR + 'Transparent background logo.png')
                builder.insert_image(img)
                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()
                # Set the "preblend_images" property to "True" to preblend transparent images
                # with a background, which may reduce artifacts.
                # Set the "preblend_images" property to "False" to render transparent images normally.
                options.preblend_images = preblend_images
                doc.save(ARTIFACTS_DIR + 'PdfSaveOptions.preblend_images.pdf', options)
                #ExEnd
                pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + 'PdfSaveOptions.preblend_images.pdf')
                image = pdf_document.pages[1].resources.images[1]
                with open(ARTIFACTS_DIR + 'PdfSaveOptions.preblend_images.pdf', 'rb') as file:
                    content = file.read()
                with io.BytesIO() as stream:
                    image.save(stream)
                    if preblend_images:
                        self.assertIn('11 0 obj\r\n20849 ', content)
                        self.assertEqual(17898, len(stream.getvalue()))
                    else:
                        self.assertIn('11 0 obj\r\n19289 ', content)
                        self.assertEqual(19216, len(stream.getvalue()))

    def test_pdf_digital_signature(self):
        #ExStart
        #ExFor:PdfDigitalSignatureDetails
        #ExFor:PdfDigitalSignatureDetails.__init__(CertificateHolder,str,str,datetime)
        #ExFor:PdfDigitalSignatureDetails.hash_algorithm
        #ExFor:PdfDigitalSignatureDetails.location
        #ExFor:PdfDigitalSignatureDetails.reason
        #ExFor:PdfDigitalSignatureDetails.signature_date
        #ExFor:PdfDigitalSignatureHashAlgorithm
        #ExFor:PdfSaveOptions.digital_signature_details
        #ExSummary:Shows how to sign a generated PDF document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln('Contents of signed PDF.')
        certificate_holder = aw.digitalsignatures.CertificateHolder.create(MY_DIR + 'morzal.pfx', 'aw')
        # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
        # to modify how that method converts the document to .PDF.
        options = aw.saving.PdfSaveOptions()
        # Configure the "digital_signature_details" object of the "SaveOptions" object to
        # digitally sign the document as we render it with the "save" method.
        signing_time = datetime.datetime.now()
        import aspose.words.saving as aws
        options.digital_signature_details = aw.saving.PdfDigitalSignatureDetails(certificate_holder, 'Test Signing', 'My Office', signing_time)
        options.digital_signature_details.hash_algorithm = aw.saving.PdfDigitalSignatureHashAlgorithm.RIPE_MD160
        self.assertEqual('Test Signing', options.digital_signature_details.reason)
        self.assertEqual('My Office', options.digital_signature_details.location)
        self.assertEqual(signing_time.astimezone(timezone.utc), options.digital_signature_details.signature_date)
        doc.save(ARTIFACTS_DIR + 'PdfSaveOptions.pdf_digital_signature.pdf', options)
        #ExEnd
        with open(ARTIFACTS_DIR + 'PdfSaveOptions.pdf_digital_signature.pdf', 'rb') as file:
            content = file.read()
        self.assertIn(b'7 0 obj\r\n' + b'<</Type/Annot/Subtype/Widget/Rect[0 0 0 0]/FT/Sig/T', content)
        self.assertFalse(aw.FileFormatUtil.detect_file_format(ARTIFACTS_DIR + 'PdfSaveOptions.pdf_digital_signature.pdf').has_digital_signature)

    def test_pdf_digital_signature_timestamp(self):
        #ExStart
        #ExFor:PdfDigitalSignatureDetails.timestamp_settings
        #ExFor:PdfDigitalSignatureTimestampSettings
        #ExFor:PdfDigitalSignatureTimestampSettings.__init__(str,str,str)
        #ExFor:PdfDigitalSignatureTimestampSettings.__init__(str,str,str,TimeSpan)
        #ExFor:PdfDigitalSignatureTimestampSettings.password
        #ExFor:PdfDigitalSignatureTimestampSettings.server_url
        #ExFor:PdfDigitalSignatureTimestampSettings.timeout
        #ExFor:PdfDigitalSignatureTimestampSettings.user_name
        #ExSummary:Shows how to sign a saved PDF document digitally and timestamp it.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln('Signed PDF contents.')
        # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
        # to modify how that method converts the document to .PDF.
        options = aw.saving.PdfSaveOptions()
        # Create a digital signature and assign it to our SaveOptions object to sign the document when we save it to PDF.
        certificate_holder = aw.digitalsignatures.CertificateHolder.create(MY_DIR + 'morzal.pfx', 'aw')
        options.digital_signature_details = aw.saving.PdfDigitalSignatureDetails(certificate_holder, 'Test Signing', 'Aspose Office', datetime.datetime.now())
        # Create a timestamp authority-verified timestamp.
        options.digital_signature_details.timestamp_settings = aw.saving.PdfDigitalSignatureTimestampSettings('https://freetsa.org/tsr', 'JohnDoe', 'MyPassword')
        # The default lifespan of the timestamp is 100 seconds.
        self.assertEqual(100.0, options.digital_signature_details.timestamp_settings.timeout.total_seconds())
        # We can set our timeout period via the constructor.
        options.digital_signature_details.timestamp_settings = aw.saving.PdfDigitalSignatureTimestampSettings('https://freetsa.org/tsr', 'JohnDoe', 'MyPassword', timedelta(minutes=30))
        self.assertEqual(1800.0, options.digital_signature_details.timestamp_settings.timeout.total_seconds())
        self.assertEqual('https://freetsa.org/tsr', options.digital_signature_details.timestamp_settings.server_url)
        self.assertEqual('JohnDoe', options.digital_signature_details.timestamp_settings.user_name)
        self.assertEqual('MyPassword', options.digital_signature_details.timestamp_settings.password)
        # The "save" method will apply our signature to the output document at this time.
        doc.save(ARTIFACTS_DIR + 'PdfSaveOptions.pdf_digital_signature_timestamp.pdf', options)
        #ExEnd
        self.assertFalse(aw.FileFormatUtil.detect_file_format(ARTIFACTS_DIR + 'PdfSaveOptions.pdf_digital_signature_timestamp.pdf').has_digital_signature)
        with open(ARTIFACTS_DIR + 'PdfSaveOptions.pdf_digital_signature_timestamp.pdf', 'rb') as file:
            content = file.read()
        self.assertIn(b'7 0 obj\r\n' + b'<</Type/Annot/Subtype/Widget/Rect[0 0 0 0]/FT/Sig/T', content)

    def test_set_numeral_format(self):
        for numeral_format in (aw.saving.NumeralFormat.ARABIC_INDIC, aw.saving.NumeralFormat.CONTEXT, aw.saving.NumeralFormat.EASTERN_ARABIC_INDIC, aw.saving.NumeralFormat.EUROPEAN, aw.saving.NumeralFormat.SYSTEM):
            with self.subTest(numeral_forma=numeral_format):
                #ExStart
                #ExFor:FixedPageSaveOptions.numeral_format
                #ExFor:NumeralFormat
                #ExSummary:Shows how to set the numeral format used when saving to PDF.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.font.locale_id = 4096  # CultureInfo("ar-AR").lcid
                builder.writeln('1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 50, 100')
                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()
                # Set the "numeral_format" property to "NumeralFormat.ARABIC_INDIC" to
                # use glyphs from the U+0660 to U+0669 range as numbers.
                # Set the "numeral_format" property to "NumeralFormat.CONTEXT" to
                # look up the locale to determine what number of glyphs to use.
                # Set the "numeral_format" property to "NumeralFormat.EASTERN_ARABIC_INDIC" to
                # use glyphs from the U+06F0 to U+06F9 range as numbers.
                # Set the "numeral_format" property to "NumeralFormat.EUROPEAN" to use european numerals.
                # Set the "numeral_format" property to "NumeralFormat.SYSTEM" to determine the symbol set from regional settings.
                options.numeral_format = numeral_format
                doc.save(ARTIFACTS_DIR + 'PdfSaveOptions.set_numeral_format.pdf', options)
                #ExEnd

    def test_export_page_set(self):
        #ExStart
        #ExFor:FixedPageSaveOptions.page_set
        #ExSummary:Shows how to export Odd pages from the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        for i in range(5):
            builder.writeln('Page ' + str(i + 1) + '(' + ('odd' if i % 2 == 0 else 'even') + ')')
            if i < 4:
                builder.insert_break(aw.BreakType.PAGE_BREAK)
        # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
        # to modify how that method converts the document to .PDF.
        options = aw.saving.PdfSaveOptions()
        # Below are three "page_set" properties that we can use to filter out a set of pages from
        # our document to save in an output PDF document based on the parity of their page numbers.
        # 1 -  Save only the even-numbered pages:
        options.page_set = aw.saving.PageSet.even
        doc.save(ARTIFACTS_DIR + 'PdfSaveOptions.export_page_set.even.pdf', options)
        # 2 -  Save only the odd-numbered pages:
        options.page_set = aw.saving.PageSet.odd
        doc.save(ARTIFACTS_DIR + 'PdfSaveOptions.export_page_set.odd.pdf', options)
        # 3 -  Save every page:
        options.page_set = aw.saving.PageSet.all
        doc.save(ARTIFACTS_DIR + 'PdfSaveOptions.export_page_set.all.pdf', options)
        #ExEnd