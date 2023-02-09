# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import io
import os
from datetime import datetime, timedelta, timezone

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, IMAGE_DIR, FONTS_DIR

class ExPdfSaveOptions(ApiExampleBase):

    def test_one_page(self):

        #ExStart
        #ExFor:FixedPageSaveOptions.page_set
        #ExFor:Document.save(BytesIO,SaveOptions)
        #ExSummary:Shows how to convert only some of the pages in a document to PDF.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Page 1.")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln("Page 2.")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln("Page 3.")

        with open(ARTIFACTS_DIR + "PdfSaveOptions.one_page.pdf", "wb") as stream:

            # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
            # to modify how that method converts the document to .PDF.
            options = aw.saving.PdfSaveOptions()

            # Set the "page_index" to "1" to render a portion of the document starting from the second page.
            options.page_set = aw.saving.PageSet(1)

            # This document will contain one page starting from page two, which will only contain the second page.
            doc.save(stream, options)

        #ExEnd

        #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.one_page.pdf")

        #self.assertEqual(1, pdf_document.pages.count)

        #text_fragment_absorber = aspose.pdf.text.TextFragmentAbsorber()
        #pdf_document.pages.accept(text_fragment_absorber)

        #self.assertEqual("Page 2.", text_fragment_absorber.text)

    def test_headings_outline_levels(self):

        #ExStart
        #ExFor:ParagraphFormat.is_heading
        #ExFor:PdfSaveOptions.outline_options
        #ExFor:PdfSaveOptions.save_format
        #ExSummary:Shows how to limit the headings' level that will appear in the outline of a saved PDF document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert headings that can serve as TOC entries of levels 1, 2, and then 3.
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1

        self.assertTrue(builder.paragraph_format.is_heading)

        builder.writeln("Heading 1")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING2

        builder.writeln("Heading 1.1")
        builder.writeln("Heading 1.2")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING3

        builder.writeln("Heading 1.2.1")
        builder.writeln("Heading 1.2.2")

        # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
        # to modify how that method converts the document to .PDF.
        save_options = aw.saving.PdfSaveOptions()
        save_options.save_format = aw.SaveFormat.PDF

        # The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
        # Clicking on an entry in this outline will take us to the location of its respective heading.
        # Set the "headings_outline_levels" property to "2" to exclude all headings whose levels are above 2 from the outline.
        # The last two headings we have inserted above will not appear.
        save_options.outline_options.headings_outline_levels = 2

        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.headings_outline_levels.pdf", save_options)
        #ExEnd

        #bookmark_editor = aspose.pdf.facades.PdfBookmarkEditor()
        #bookmark_editor.bind_pdf(ARTIFACTS_DIR + "PdfSaveOptions.headings_outline_levels.pdf")

        #bookmarks = bookmark_editor.extract_bookmarks()

        #self.assertEqual(3, bookmarks.count)

    def test_create_missing_outline_levels(self):

        for create_missing_outline_levels in (False, True):
            with self.subTest(create_missing_outline_levels=create_missing_outline_levels):
                #ExStart
                #ExFor:OutlineOptions.create_missing_outline_levels
                #ExFor:PdfSaveOptions.outline_options
                #ExSummary:Shows how to work with outline levels that do not contain any corresponding headings when saving a PDF document.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Insert headings that can serve as TOC entries of levels 1 and 5.
                builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1

                self.assertTrue(builder.paragraph_format.is_heading)

                builder.writeln("Heading 1")

                builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING5

                builder.writeln("Heading 1.1.1.1.1")
                builder.writeln("Heading 1.1.1.1.2")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                save_options = aw.saving.PdfSaveOptions()

                # The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
                # Clicking on an entry in this outline will take us to the location of its respective heading.
                # Set the "headings_outline_levels" property to "5" to include all headings of levels 5 and below in the outline.
                save_options.outline_options.headings_outline_levels = 5

                # This document contains headings of levels 1 and 5, and no headings with levels of 2, 3, and 4.
                # The output PDF document will treat outline levels 2, 3, and 4 as "missing".
                # Set the "create_missing_outline_levels" property to "True" to include all missing levels in the outline,
                # leaving blank outline entries since there are no usable headings.
                # Set the "create_missing_outline_levels" property to "False" to ignore missing outline levels,
                # and treat the outline level 5 headings as level 2.
                save_options.outline_options.create_missing_outline_levels = create_missing_outline_levels

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.create_missing_outline_levels.pdf", save_options)
                #ExEnd

                #bookmark_editor = aspose.pdf.facades.PdfBookmarkEditor()
                #bookmark_editor.bind_pdf(ARTIFACTS_DIR + "PdfSaveOptions.create_missing_outline_levels.pdf")

                #bookmarks = bookmark_editor.extract_bookmarks()

                #self.assertEqual(6 if create_missing_outline_levels else 3, bookmarks.count)
        #endif

    def test_table_heading_outlines(self):

        for create_outlines_for_headings_in_tables in (False, True):
            with self.subTest(create_outlines_for_headings_in_tables=create_outlines_for_headings_in_tables):
                #ExStart
                #ExFor:OutlineOptions.create_outlines_for_headings_in_tables
                #ExSummary:Shows how to create PDF document outline entries for headings inside tables.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Create a table with three rows. The first row,
                # whose text we will format in a heading-type style, will serve as the column header.
                builder.start_table()
                builder.insert_cell()
                builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
                builder.write("Customers")
                builder.end_row()
                builder.insert_cell()
                builder.paragraph_format.style_identifier = aw.StyleIdentifier.NORMAL
                builder.write("John Doe")
                builder.end_row()
                builder.insert_cell()
                builder.write("Jane Doe")
                builder.end_table()

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                pdf_save_options = aw.saving.PdfSaveOptions()

                # The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
                # Clicking on an entry in this outline will take us to the location of its respective heading.
                # Set the "headings_outline_levels" property to "1" to get the outline
                # to only register headings with heading levels that are no larger than 1.
                pdf_save_options.outline_options.headings_outline_levels = 1

                # Set the "create_outlines_for_headings_in_tables" property to "False" to exclude all headings within tables,
                # such as the one we have created above from the outline.
                # Set the "create_outlines_for_headings_in_tables" property to "True" to include all headings within tables
                # in the outline, provided that they have a heading level that is no larger than the value of the "headings_outline_levels" property.
                pdf_save_options.outline_options.create_outlines_for_headings_in_tables = create_outlines_for_headings_in_tables

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.table_heading_outlines.pdf", pdf_save_options)
                #ExEnd

                #pdf_doc = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.table_heading_outlines.pdf")

                #if create_outlines_for_headings_in_tables:
                #    self.assertEqual(1, pdf_doc.outlines.count)
                #    self.assertEqual("Customers", pdf_doc.outlines[1].title)
                #else:
                #    self.assertEqual(0, pdf_doc.outlines.count)

                #table_absorber = aspose.pdf.text.TableAbsorber()
                #table_absorber.visit(pdf_doc.pages[1])

                #self.assertEqual("Customers", table_absorber.table_list[0].row_list[0].cell_list[0].text_fragments[1].text)
                #self.assertEqual("John Doe", table_absorber.table_list[0].row_list[1].cell_list[0].text_fragments[1].text)
                #self.assertEqual("Jane Doe", table_absorber.table_list[0].row_list[2].cell_list[0].text_fragments[1].text)

    def test_expanded_outline_levels(self):

        #ExStart
        #ExFor:Document.save(str,SaveOptions)
        #ExFor:PdfSaveOptions
        #ExFor:OutlineOptions.headings_outline_levels
        #ExFor:OutlineOptions.expanded_outline_levels
        #ExSummary:Shows how to convert a whole document to PDF with three levels in the document outline.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert headings of levels 1 to 5.
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1

        self.assertTrue(builder.paragraph_format.is_heading)

        builder.writeln("Heading 1")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING2

        builder.writeln("Heading 1.1")
        builder.writeln("Heading 1.2")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING3

        builder.writeln("Heading 1.2.1")
        builder.writeln("Heading 1.2.2")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING4

        builder.writeln("Heading 1.2.2.1")
        builder.writeln("Heading 1.2.2.2")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING5

        builder.writeln("Heading 1.2.2.2.1")
        builder.writeln("Heading 1.2.2.2.2")

        # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
        # to modify how that method converts the document to .PDF.
        options = aw.saving.PdfSaveOptions()

        # The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
        # Clicking on an entry in this outline will take us to the location of its respective heading.
        # Set the "headings_outline_levels" property to "4" to exclude all headings whose levels are above 4 from the outline.
        options.outline_options.headings_outline_levels = 4

        # If an outline entry has subsequent entries of a higher level inbetween itself and the next entry of the same or lower level,
        # an arrow will appear to the left of the entry. This entry is the "owner" of several such "sub-entries".
        # In our document, the outline entries from the 5th heading level are sub-entries of the second 4th level outline entry,
        # the 4th and 5th heading level entries are sub-entries of the second 3rd level entry, and so on.
        # In the outline, we can click on the arrow of the "owner" entry to collapse/expand all its sub-entries.
        # Set the "expanded_outline_levels" property to "2" to automatically expand all heading level 2 and lower outline entries
        # and collapse all level and 3 and higher entries when we open the document.
        options.outline_options.expanded_outline_levels = 2

        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.expanded_outline_levels.pdf", options)
        #ExEnd

        #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.expanded_outline_levels.pdf")

        #self.assertEqual(1, pdf_document.outlines.count)
        #self.assertEqual(5, pdf_document.outlines.visible_count)

        #self.assertTrue(pdf_document.outlines[1].open)
        #self.assertEqual(1, pdf_document.outlines[1].level)

        #self.assertFalse(pdf_document.outlines[1][1].open)
        #self.assertEqual(2, pdf_document.outlines[1][1].level)

        #self.assertTrue(pdf_document.outlines[1][2].open)
        #self.assertEqual(2, pdf_document.outlines[1][2].level)

    def test_update_fields(self):

        for update_fields in (False, True):
            with self.subTest(update_fields=update_fields):
                #ExStart
                #ExFor:PdfSaveOptions.clone
                #ExFor:SaveOptions.update_fields
                #ExSummary:Shows how to update all the fields in a document immediately before saving it to PDF.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Insert text with PAGE and NUMPAGES fields. These fields do not display the correct value in real time.
                # We will need to manually update them using updating methods such as "Field.Update()", and "Document.UpdateFields()"
                # each time we need them to display accurate values.
                builder.write("Page ")
                builder.insert_field("PAGE", "")
                builder.write(" of ")
                builder.insert_field("NUMPAGES", "")
                builder.insert_break(aw.BreakType.PAGE_BREAK)
                builder.writeln("Hello World!")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()

                # Set the "update_fields" property to "False" to not update all the fields in a document right before a save operation.
                # This is the preferable option if we know that all our fields will be up to date before saving.
                # Set the "update_fields" property to "True" to iterate through all the document
                # fields and update them before we save it as a PDF. This will make sure that all the fields will display
                # the most accurate values in the PDF.
                options.update_fields = update_fields

                # We can clone PdfSaveOptions objects.
                options_copy = options.clone()

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.update_fields.pdf", options)
                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.update_fields.pdf")

                #text_fragment_absorber = aspose.pdf.text.TextFragmentAbsorber()
                #pdf_document.pages.accept(text_fragment_absorber)

                #self.assertEqual("Page 1 of 2" if update_fields else "Page  of ", text_fragment_absorber.text_fragments[1].text)

    def test_preserve_form_fields(self):

        for preserve_form_fields in (False, True):
            with self.subTest(preserve_form_fields=preserve_form_fields):
                #ExStart
                #ExFor:PdfSaveOptions.preserve_form_fields
                #ExSummary:Shows how to save a document to the PDF format using the Save method and the PdfSaveOptions class.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.write("Please select a fruit: ")

                # Insert a combo box which will allow a user to choose an option from a collection of strings.
                builder.insert_combo_box("MyComboBox", ["Apple", "Banana", "Cherry"], 0)

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                pdf_options = aw.saving.PdfSaveOptions()

                # Set the "preserve_form_fields" property to "True" to save form fields as interactive objects in the output PDF.
                # Set the "preserve_form_fields" property to "False" to freeze all form fields in the document at
                # their current values and display them as plain text in the output PDF.
                pdf_options.preserve_form_fields = preserve_form_fields

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.preserve_form_fields.pdf", pdf_options)
                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.preserve_form_fields.pdf")

                #self.assertEqual(1, pdf_document.pages.count)

                #text_fragment_absorber = aspose.pdf.text.TextFragmentAbsorber()
                #pdf_document.pages.accept(text_fragment_absorber)

                #with open(ARTIFACTS_DIR + "PdfSaveOptions.preserve_form_fields.pdf", 'rb') as file:
                #    content = file.read().decode('utf-8')

                #if preserve_form_fields:
                #    self.assertEqual("Please select a fruit: ", text_fragment_absorber.text)
                #    self.assertIn("11 0 obj\r\n" +
                #                  "<</Type /Annot/Subtype /Widget/P 5 0 R/FT /Ch/F 4/Rect [168.39199829 707.35101318 217.87442017 722.64007568]/Ff 131072/T(\xFE\xFF\0M\0y\0C\0o\0m\0b\0o\0B\0o\0x)/Opt " +
                #                  "[(\xFE\xFF\0A\0p\0p\0l\0e) (\xFE\xFF\0B\0a\0n\0a\0n\0a) (\xFE\xFF\0C\0h\0e\0r\0r\0y) ]/V(\xFE\xFF\0A\0p\0p\0l\0e)/DA(0 g /FAAABD 12 Tf )/AP<</N 12 0 R>>>>",
                #                  content)

                #    form = pdf_document.form
                #    self.assertEqual(1, pdf_document.form.count)

                #    field = form.fields[0].as_combo_box_field()

                #    self.assertEqual("MyComboBox", field.full_name)
                #    self.assertEqual(3, field.options.count)
                #    self.assertEqual("Apple", field.value)
                #else:
                #    self.assertEqual("Please select a fruit: Apple", text_fragment_absorber.text)
                #    self.assertNotIn("/Widget", content)
                #    self.assertEqual(0, pdf_document.form.count)

    def test_compliance(self):

        for pdf_compliance in (aw.saving.PdfCompliance.PDF_A2U,
                               aw.saving.PdfCompliance.PDF17,
                               aw.saving.PdfCompliance.PDF_A2A,
                               aw.saving.PdfCompliance.PDF_UA1,
                               aw.saving.PdfCompliance.PDF20,
                               aw.saving.PdfCompliance.PDF_A4):
            with self.subTest(pdf_compliance=pdf_compliance):
                #ExStart
                #ExFor:PdfSaveOptions.compliance
                #ExFor:PdfCompliance
                #ExSummary:Shows how to set the PDF standards compliance level of saved PDF documents.
                doc = aw.Document(MY_DIR + "Images.docx")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                # Note that some PdfSaveOptions are prohibited when saving to one of the standards and automatically fixed.
                # Use IWarningCallback to know which options are automatically fixed.
                save_options = aw.saving.PdfSaveOptions()

                # Set the "compliance" property to "PdfCompliance.PDF_A1B" to comply with the "PDF/A-1b" standard,
                # which aims to preserve the visual appearance of the document as Aspose.Words convert it to PDF.
                # Set the "compliance" property to "PdfCompliance.PDF17" to comply with the "1.7" standard.
                # Set the "compliance" property to "PdfCompliance.PDF_A1A" to comply with the "PDF/A-1a" standard,
                # which complies with "PDF/A-1b" as well as preserving the document structure of the original document.
                # Set the "compliance" property to "PdfCompliance.PDF_UA1" to comply with the "PDF/UA-1" (ISO 14289-1) standard,
                # which aims to define represent electronic documents in PDF that allow the file to be accessible.
                # Set the "Compliance" property to "PdfCompliance.Pdf20" to comply with the "PDF 2.0" (ISO 32000-2) standard.
                # Set the "Compliance" property to "PdfCompliance.PdfA4" to comply with the "PDF/A-4" (ISO 19004:2020) standard,
                # which preserving document static visual appearance over time.
                # This helps with making documents searchable but may significantly increase the size of already large documents.
                save_options.compliance = pdf_compliance

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.compliance.pdf", save_options)
                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.compliance.pdf")

                #if pdf_compliance == aw.saving.PdfCompliance.PDF17:
                #    self.assertEqual(aspose.pdf.PdfFormat.V_1_7, pdf_document.pdf_format)
                #    self.assertEqual("1.7", pdf_document.version)

                #elif pdf_compliance == aw.saving.PdfCompliance.PDF_A2A:
                #    self.assertEqual(aspose.pdf.PdfFormat.PDF_A_2A, pdf_document.pdf_format)
                #    self.assertEqual("1.7", pdf_document.version)

                #elif pdf_compliance == aw.saving.PdfCompliance.PDF_A2U:
                #    self.assertEqual(aspose.pdf.PdfFormat.PDF_A_2U, pdf_document.pdf_format)
                #    self.assertEqual("1.7", pdf_document.version)

                #elif pdf_compliance == aw.saving.PdfCompliance.PDF_UA1:
                #    self.assertEqual(aspose.pdf.PdfFormat.PDF_UA_1, pdf_document.pdf_format)
                #    self.assertEqual("1.7", pdf_document.version)

                #elif pdf_compliance == aw.saving.PdfCompliance.PDF_20:
                #    self.assertEqual(aspose.pdf.PdfFormat.PDF_V_2_0, pdf_document.pdf_format)
                #    self.assertEqual("2.0", pdf_document.version)

                #elif pdf_compliance == aw.saving.PdfCompliance.PDF_A4:
                #    self.assertEqual(aspose.pdf.PdfFormat.PDF_A_4, pdf_document.pdf_format)
                #    self.assertEqual("2.0", pdf_document.version)

    def test_text_compression(self):

        for pdf_text_compression in (aw.saving.PdfTextCompression.NONE,
                                     aw.saving.PdfTextCompression.FLATE):
            with self.subTest(pdf_text_compression=pdf_text_compression):
                #ExStart
                #ExFor:PdfSaveOptions
                #ExFor:PdfSaveOptions.text_compression
                #ExFor:PdfTextCompression
                #ExSummary:Shows how to apply text compression when saving a document to PDF.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                for i in range(100):
                    builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                                    "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()

                # Set the "text_compression" property to "PdfTextCompression.NONE" to not apply any
                # compression to text when we save the document to PDF.
                # Set the "text_compression" property to "PdfTextCompression.FLATE" to apply ZIP compression
                # to text when we save the document to PDF. The larger the document, the bigger the impact that this will have.
                options.text_compression = pdf_text_compression

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.text_compression.pdf", options)
                #ExEnd

                if pdf_text_compression == aw.saving.PdfTextCompression.NONE:
                    self.assertLess(60000, os.path.getsize(ARTIFACTS_DIR + "PdfSaveOptions.text_compression.pdf"))
                    with open(ARTIFACTS_DIR + "PdfSaveOptions.text_compression.pdf", "rb") as file:
                        self.assertIn(b"12 0 obj\r\n<</Length 13 0 R>>stream", file.read())

                elif pdf_text_compression == aw.saving.PdfTextCompression.FLATE:
                    self.assertGreater(30000, os.path.getsize(ARTIFACTS_DIR + "PdfSaveOptions.text_compression.pdf"))
                    with open(ARTIFACTS_DIR + "PdfSaveOptions.text_compression.pdf", "rb") as file:
                        self.assertIn(b"12 0 obj\r\n<</Length 13 0 R/Filter /FlateDecode>>stream", file.read())

    def test_image_compression(self):

        for pdf_image_compression in (aw.saving.PdfImageCompression.AUTO,
                                      aw.saving.PdfImageCompression.JPEG):
            with self.subTest(pdf_image_compression=pdf_image_compression):
                #ExStart
                #ExFor:PdfSaveOptions.image_compression
                #ExFor:PdfSaveOptions.jpeg_quality
                #ExFor:PdfImageCompression
                #ExSummary:Shows how to specify a compression type for all images in a document that we are converting to PDF.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.writeln("Jpeg image:")
                builder.insert_image(IMAGE_DIR + "Logo.jpg")
                builder.insert_paragraph()
                builder.writeln("Png image:")
                builder.insert_image(IMAGE_DIR + "Transparent background logo.png")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                pdf_save_options = aw.saving.PdfSaveOptions()

                # Set the "image_compression" property to "PdfImageCompression.AUTO" to use the
                # "image_compression" property to control the quality of the Jpeg images that end up in the output PDF.
                # Set the "image_compression" property to "PdfImageCompression.JPEG" to use the
                # "image_compression" property to control the quality of all images that end up in the output PDF.
                pdf_save_options.image_compression = pdf_image_compression

                # Set the "jpeg_quality" property to "10" to strengthen compression at the cost of image quality.
                pdf_save_options.jpeg_quality = 10

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.image_compression.pdf", pdf_save_options)
                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.image_compression.pdf")
                #with pdf_document.pages[1].resources.images[1].to_stream() as pdf_doc_image_stream:
                #    self.verify_image(400, 400, pdf_doc_image_stream)

                #with pdf_document.pages[1].resources.images[2].to_stream() as pdf_doc_image_stream:
                #    if pdf_image_compression == aw.saving.PdfImageCompression.AUTO:
                #        self.assertLess(50000, os.path.getsize(ARTIFACTS_DIR + "PdfSaveOptions.image_compression.pdf"))
                #        with self.assertRaises(Exception):
                #            self.verify_image(400, 400, pdf_doc_image_stream)

                #    elif pdf_image_compression == aw.saving.PdfImageCompression.JPEG:
                #        self.assertLess(42000, os.path.getsize(ARTIFACTS_DIR + "PdfSaveOptions.image_compression.pdf"))
                #        with self.assertRaises(Exception):
                #            self.verify_image(400, 400, pdf_doc_image_stream)

    def test_image_color_space_export_mode(self):

        for pdf_image_color_space_export_mode in (aw.saving.PdfImageColorSpaceExportMode.AUTO,
                                                  aw.saving.PdfImageColorSpaceExportMode.SIMPLE_CMYK):
            with self.subTest(pdf_image_color_space_export_mode=pdf_image_color_space_export_mode):
                #ExStart
                #ExFor:PdfImageColorSpaceExportMode
                #ExFor:PdfSaveOptions.image_color_space_export_mode
                #ExSummary:Shows how to set a different color space for images in a document as we export it to PDF.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.writeln("Jpeg image:")
                builder.insert_image(IMAGE_DIR + "Logo.jpg")
                builder.insert_paragraph()
                builder.writeln("Png image:")
                builder.insert_image(IMAGE_DIR + "Transparent background logo.png")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                pdf_save_options = aw.saving.PdfSaveOptions()

                # Set the "image_color_space_export_mode" property to "PdfImageColorSpaceExportMode.AUTO" to get Aspose.Words to
                # automatically select the color space for images in the document that it converts to PDF.
                # In most cases, the color space will be RGB.
                # Set the "image_color_space_export_mode" property to "PdfImageColorSpaceExportMode.SIMPLE_CMYK"
                # to use the CMYK color space for all images in the saved PDF.
                # Aspose.Words will also apply Flate compression to all images and ignore the "image_compression" property's value.
                pdf_save_options.image_color_space_export_mode = pdf_image_color_space_export_mode

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.image_color_space_export_mode.pdf", pdf_save_options)
                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.image_color_space_export_mode.pdf")
                #pdf_doc_image = pdf_document.pages[1].resources.images[1]

                #if pdf_image_color_space_export_mode == aw.saving.PdfImageColorSpaceExportMode.AUTO:
                #    self.assertLess(20000, pdf_doc_image.to_stream().length)

                #elif pdf_image_color_space_export_mode == aw.saving.PdfImageColorSpaceExportMode.SIMPLE_CMYK:
                #    self.assertLess(100000, pdf_doc_image.to_stream().length)

                #self.assertEqual(400, pdf_doc_image.width)
                #self.assertEqual(400, pdf_doc_image.height)
                #self.assertEqual(aspose.pdf.ColorType.RGB, pdf_doc_image.get_color_type())

                #pdf_doc_image = pdf_document.pages[1].resources.images[2]

                #if pdf_image_color_space_export_mode == aw.saving.PdfImageColorSpaceExportMode.AUTO:
                #    self.assertLess(25000, pdf_doc_image.to_stream().length)

                #elif pdf_image_color_space_export_mode == aw.saving.PdfImageColorSpaceExportMode.SIMPLE_CMYK:
                #    self.assertLess(18000, pdf_doc_image.to_stream().length)

                #self.assertEqual(400, pdf_doc_image.width)
                #self.assertEqual(400, pdf_doc_image.height)
                #self.assertEqual(aspose.pdf.ColorType.RGB, pdf_doc_image.get_color_type())

    def test_downsample_options(self):

        #ExStart
        #ExFor:DownsampleOptions
        #ExFor:DownsampleOptions.downsample_images
        #ExFor:DownsampleOptions.resolution
        #ExFor:DownsampleOptions.resolution_threshold
        #ExFor:PdfSaveOptions.downsample_options
        #ExSummary:Shows how to change the resolution of images in the PDF document.
        doc = aw.Document(MY_DIR + "Images.docx")

        # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
        # to modify how that method converts the document to .PDF.
        options = aw.saving.PdfSaveOptions()

        # By default, Aspose.Words downsample all images in a document that we save to PDF to 220 ppi.
        self.assertTrue(options.downsample_options.downsample_images)
        self.assertEqual(220, options.downsample_options.resolution)
        self.assertEqual(0, options.downsample_options.resolution_threshold)

        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.downsample_options.default.pdf", options)

        # Set the "resolution" property to "36" to downsample all images to 36 ppi.
        options.downsample_options.resolution = 36

        # Set the "resolution_threshold" property to only apply the downsampling to
        # images with a resolution that is above 128 ppi.
        options.downsample_options.resolution_threshold = 128

        # Only the first two images from the document will be downsampled at this stage.
        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.downsample_options.lower_resolution.pdf", options)
        #ExEnd

        #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.downsample_options.default.pdf")
        #pdf_doc_image = pdf_document.pages[1].resources.images[1]

        #self.assertLess(300000, pdf_doc_image.to_stream().length)
        #self.assertEqual(aspose.pdf.ColorType.RGB, pdf_doc_image.get_color_type())

    def test_color_rendering(self):

        for color_mode in (aw.saving.ColorMode.GRAYSCALE,
                           aw.saving.ColorMode.NORMAL):
            with self.subTest(color_mode=color_mode):
                #ExStart
                #ExFor:PdfSaveOptions
                #ExFor:ColorMode
                #ExFor:FixedPageSaveOptions.color_mode
                #ExSummary:Shows how to change image color with saving options property.
                doc = aw.Document(MY_DIR + "Images.docx")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                # Set the "color_mode" property to "GRAYSCALE" to render all images from the document in black and white.
                # The size of the output document may be larger with this setting.
                # Set the "color_mode" property to "NORMAL" to render all images in color.
                pdf_save_options = aw.saving.PdfSaveOptions()
                pdf_save_options.color_mode = color_mode

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.color_rendering.pdf", pdf_save_options)
                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.color_rendering.pdf")
                #pdf_doc_image = pdf_document.pages[1].resources.images[1]

                #if color_mode == aw.saving.ColorMode.NORMAL:
                #    self.assertLess(300000, pdf_doc_image.to_stream().length)
                #    self.assertEqual(aspose.pdf.ColorType.RGB, pdf_doc_image.get_color_type())

                #elif color_mode == aw.saving.ColorMode.GRAYSCALE:
                #    self.assertLess(1000000, pdf_doc_image.to_stream().length)
                #    self.assertEqual(aspose.pdf.ColorType.GRAYSCALE, pdf_doc_image.get_color_type())

    def test_doc_title(self):

        for display_doc_title in (False, True):
            with self.subTest(display_doc_title=display_doc_title):
                #ExStart
                #ExFor:PdfSaveOptions.display_doc_title
                #ExSummary:Shows how to display the title of the document as the title bar.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.writeln("Hello world!")

                doc.built_in_document_properties.title = "Windows bar pdf title"

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                # Set the "display_doc_title" to "True" to get some PDF readers, such as Adobe Acrobat Pro,
                # to display the value of the document's "title" built-in property in the tab that belongs to this document.
                # Set the "display_doc_title" to "False" to get such readers to display the document's filename.
                pdf_save_options = aw.saving.PdfSaveOptions()
                pdf_save_options.display_doc_title = display_doc_title

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.doc_title.pdf", pdf_save_options)
                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.doc_title.pdf")

                #self.assertEqual(display_doc_title, pdf_document.display_doc_title)
                #self.assertEqual("Windows bar pdf title", pdf_document.info.title)

    def test_memory_optimization(self):

        for memory_optimization in (False, True):
            with self.subTest(memory_optimization=memory_optimization):
                #ExStart
                #ExFor:SaveOptions.create_save_options(SaveFormat)
                #ExFor:SaveOptions.memory_optimization
                #ExSummary:Shows an option to optimize memory consumption when rendering large documents to PDF.
                doc = aw.Document(MY_DIR + "Rendering.docx")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                save_options = aw.saving.SaveOptions.create_save_options(aw.SaveFormat.PDF)

                # Set the "memory_optimization" property to "True" to lower the memory footprint of large documents' saving operations
                # at the cost of increasing the duration of the operation.
                # Set the "memory_optimization" property to "False" to save the document as a PDF normally.
                save_options.memory_optimization = memory_optimization

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.memory_optimization.pdf", save_options)
                #ExEnd

    def test_escape_uri(self):

        parameters = [
            (r"https://www.google.com/search?q= aspose", "https://www.google.com/search?q=%20aspose"),
            (r"https://www.google.com/search?q=%20aspose", "https://www.google.com/search?q=%20aspose"),
            ]

        for uri, result in parameters:
            with self.subTest(uri=uri, result=result):
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.insert_hyperlink("Testlink", uri, False)

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.escaped_uri.pdf")

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.escaped_uri.pdf")

                #page = pdf_document.pages[1]
                #link_annot = page.annotations[1].as_link_annotation()

                #action = link_Annot.action.as_go_to_uri_action()

                #self.assertEqual(result, action.uri)

    def test_open_hyperlinks_in_new_window(self):

        for open_hyperlinks_in_new_window in (False, True):
            with self.subTest(open_hyperlinks_in_new_window=open_hyperlinks_in_new_window):
                #ExStart
                #ExFor:PdfSaveOptions.open_hyperlinks_in_new_window
                #ExSummary:Shows how to save hyperlinks in a document we convert to PDF so that they open new pages when we click on them.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.insert_hyperlink("Testlink", "https://www.google.com/search?q=%20aspose", False)

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()

                # Set the "open_hyperlinks_in_new_window" property to "True" to save all hyperlinks using Javascript code
                # that forces readers to open these links in new windows/browser tabs.
                # Set the "open_hyperlinks_in_new_window" property to "False" to save all hyperlinks normally.
                options.open_hyperlinks_in_new_window = open_hyperlinks_in_new_window

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.open_hyperlinks_in_new_window.pdf", options)
                #ExEnd

                with open(ARTIFACTS_DIR + "PdfSaveOptions.open_hyperlinks_in_new_window.pdf", "rb") as file:
                    content = file.read()
                    if open_hyperlinks_in_new_window:
                        self.assertIn(
                            b"<</Type /Annot/Subtype /Link/Rect [72 706.20098877 111.32800293 720]/BS " +
                            b"<</Type/Border/S/S/W 0>>/A<</Type /Action/S /JavaScript/JS(app.launchURL\\(\"https://www.google.com/search?q=%20aspose\", true\\);)>>>>",
                            content)
                    else:
                        self.assertIn(
                            b"<</Type /Annot/Subtype /Link/Rect [72 706.20098877 111.32800293 720]/BS " +
                            b"<</Type/Border/S/S/W 0>>/A<</Type /Action/S /URI/URI(https://www.google.com/search?q=%20aspose)>>>>",
                            content)

                #pdf_document = aspose.pdf.document(ARTIFACTS_DIR + "PdfSaveOptions.open_hyperlinks_in_new_window.pdf")

                #page = pdf_document.pages[1]
                #link_annot = page.annotations[1].as_link_annotation()

                #self.assertEqual(type(JavascriptAction) if open_hyperlinks_in_new_window else type(GoToURIAction),
                #    link_annot.action.get_type())

    ##ExStart
    ##ExFor:MetafileRenderingMode
    ##ExFor:MetafileRenderingOptions
    ##ExFor:MetafileRenderingOptions.emulate_raster_operations
    ##ExFor:MetafileRenderingOptions.rendering_mode
    ##ExFor:IWarningCallback
    ##ExFor:FixedPageSaveOptions.metafile_rendering_options
    ##ExSummary:Shows added a fallback to bitmap rendering and changing type of warnings about unsupported metafile records.
    #def test_handle_binary_raster_warnings(self):

    #    doc = aw.Document(MY_DIR + "WMF with image.docx")

    #    metafile_rendering_options = aw.saving.MetafileRenderingOptions()

    #    # Set the "emulate_raster_operations" property to "False" to fall back to bitmap when
    #    # it encounters a metafile, which will require raster operations to render in the output PDF.
    #    metafile_rendering_options.emulate_raster_operations = False

    #    # Set the "rendering_mode" property to "VECTOR_WITH_FALLBACK" to try to render every metafile using vector graphics.
    #    metafile_rendering_options.rendering_mode = aw.saving.MetafileRenderingMode.VECTOR_WITH_FALLBACK

    #    # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
    #    # to modify how that method converts the document to .PDF and applies the configuration
    #    # in our MetafileRenderingOptions object to the saving operation.
    #    save_options = aw.saving.PdfSaveOptions()
    #    save_options.metafile_rendering_options = metafile_rendering_options

    #    callback = ExPdfSaveOptions.HandleDocumentWarnings()
    #    doc.warning_callback = callback

    #    doc.save(ARTIFACTS_DIR + "PdfSaveOptions.handle_binary_raster_warnings.pdf", save_options)

    #    self.assertEqual(1, callback.warnings.count)
    #    self.assertEqual("'R2_XORPEN' binary raster operation is partly supported.",
    #        callback.warnings[0].description)

    #class HandleDocumentWarnings(aw.IWarningCallback):
    #    """Prints and collects formatting loss-related warnings that occur upon saving a document."""

    #    def __init__(self):
    #        self.warnings = aw.WarningInfoCollection()

    #    def warning(self, info: aw.WarningInfo):

    #        if info.warning_type == aw.WarningType.MINOR_FORMATTING_LOSS:
    #            print("Unsupported operation: " + info.description)
    #            self.warnings.warning(info)

    ##ExEnd

    def test_header_footer_bookmarks_export_mode(self):

        for header_footer_bookmarks_export_mode in (aw.saving.HeaderFooterBookmarksExportMode.NONE,
                                                    aw.saving.HeaderFooterBookmarksExportMode.FIRST,
                                                    aw.saving.HeaderFooterBookmarksExportMode.ALL):
            with self.subTest(header_footer_bookmarks_export_mode=header_footer_bookmarks_export_mode):
                #ExStart
                #ExFor:HeaderFooterBookmarksExportMode
                #ExFor:OutlineOptions
                #ExFor:OutlineOptions.default_bookmarks_outline_level
                #ExFor:PdfSaveOptions.header_footer_bookmarks_export_mode
                #ExFor:PdfSaveOptions.page_mode
                #ExFor:PdfPageMode
                #ExSummary:Shows to process bookmarks in headers/footers in a document that we are rendering to PDF.
                doc = aw.Document(MY_DIR + "Bookmarks in headers and footers.docx")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                save_options = aw.saving.PdfSaveOptions()

                # Set the "page_mode" property to "PdfPageMode.USE_OUTLINES" to display the outline navigation pane in the output PDF.
                save_options.page_mode = aw.saving.PdfPageMode.USE_OUTLINES

                # Set the "default_bookmarks_outline_level" property to "1" to display all
                # bookmarks at the first level of the outline in the output PDF.
                save_options.outline_options.default_bookmarks_outline_level = 1

                # Set the "header_footer_bookmarks_export_mode" property to "HeaderFooterBookmarksExportMode.NONE" to
                # not export any bookmarks that are inside headers/footers.
                # Set the "header_footer_bookmarks_export_mode" property to "HeaderFooterBookmarksExportMode.FIRST" to
                # only export bookmarks in the first section's header/footers.
                # Set the "header_footer_bookmarks_export_mode" property to "HeaderFooterBookmarksExportMode.ALL" to
                # export bookmarks that are in all headers/footers.
                save_options.header_footer_bookmarks_export_mode = header_footer_bookmarks_export_mode

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.header_footer_bookmarks_export_mode.pdf", save_options)
                #ExEnd

                #pdf_doc = aspose.pdf.document(ARTIFACTS_DIR + "PdfSaveOptions.header_footer_bookmarks_export_mode.pdf")
                #input_doc_locale_name = CultureInfo(doc.styles.default_font.locale_id).name

                #text_fragment_absorber = aspose.pdf.text.TextFragmentAbsorber()
                #pdf_doc.pages.accept(text_fragment_absorber)

                #with open(ARTIFACTS_DIR + "PdfSaveOptions.header_footer_bookmarks_export_mode.pdf", "rb") as file:
                #    data = file.read().decode('utf-8')

                #if header_footer_bookmarks_export_mode == aw.saving.HeaderFooterBookmarksExportMode.NONE:
                #    self.assertIn(f"<</Type /Catalog/Pages 3 0 R/Lang({input_doc_locale_name})/Metadata 4 0 R>>\r\n", data)
                #    self.assertEqual(0, pdf_doc.outlines.count)

                #elif header_footer_bookmarks_export_mode in (aw.saving.HeaderFooterBookmarksExportMode.FIRST,
                #                                             aw.saving.HeaderFooterBookmarksExportMode.ALL):
                #    self.assertIn(f"<</Type /Catalog/Pages 3 0 R/Outlines 14 0 R/PageMode /UseOutlines/Lang({inputDocLocaleName})/Metadata 4 0 R>>", data)

                #    outline_item_collection = pdf_doc.outlines

                #    self.assertEqual(4, outline_item_collection.count)
                #    self.assertEqual("Bookmark_1", outline_item_collection[1].title)
                #    self.assertEqual("1 XYZ 233 806 0", outline_item_collection[1].destination.to_string())

                #    self.assertEqual("Bookmark_2", outline_item_collection[2].title)
                #    self.assertEqual("1 XYZ 84 47 0", outline_item_collection[2].destination.to_string())

                #    self.assertEqual("Bookmark_3", outline_item_collection[3].title)
                #    self.assertEqual("2 XYZ 85 806 0", outline_item_collection[3].destination.to_string())

                #    self.assertEqual("Bookmark_4", outline_item_collection[4].title)
                #    self.assertEqual("2 XYZ 85 48 0", outline_item_collection[4].destination.to_string())

    #def test_unsupported_image_format_warning(self):

    #    doc = aw.Document(MY_DIR + "Corrupted image.docx")

    #    save_warning_callback = ExpPdfSaveOptions.SaveWarningCallback()
    #    doc.warning_callback = save_warning_callback

    #    doc.save(ARTIFACTS_DIR + "PdfSaveOption.unsupported_image_format_warning.pdf", aw.SaveFormat.PDF)

    #    self.assertEqual(
    #        save_warning_callback.save_warnings[0].description,
    #        "Image can not be processed. Possibly unsupported image format.")

    #class SaveWarningCallback(aw.IWarningCallback):

    #    def __init__(self):
    #        self.save_warnings = aw.WarningInfoCollection()

    #    def warning(self, info: aw.WarningInfo):

    #        if info.WarningType == aw.WarningType.MINOR_FORMATTING_LOSS:
    #            print(f"{info.warning_type}: {info.description}.")
    #            self.save_warnings.warning(info)

    def test_fonts_scaled_to_metafile_size(self):

        for scale_wmf_fonts in (False, True):
            with self.subTest(scale_wmf_fonts=scale_wmf_fonts):
                #ExStart
                #ExFor:MetafileRenderingOptions.scale_wmf_fonts_to_metafile_size
                #ExSummary:Shows how to WMF fonts scaling according to metafile size on the page.
                doc = aw.Document(MY_DIR + "WMF with text.docx")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                save_options = aw.saving.PdfSaveOptions()

                # Set the "scale_wmf_fonts_to_metafile_size" property to "True" to scale fonts
                # that format text within WMF images according to the size of the metafile on the page.
                # Set the "scale_wmf_fonts_to_metafile_size" property to "False" to
                # preserve the default scale of these fonts.
                save_options.metafile_rendering_options.scale_wmf_fonts_to_metafile_size = scale_wmf_fonts

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.fonts_scaled_to_metafile_size.pdf", save_options)
                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.fonts_scaled_to_metafile_size.pdf")
                #text_absorber = aspose.pdf.text.TextFragmentAbsorber()

                #pdf_document.pages[1].accept(text_absorber)
                #text_fragment_rectangle = text_absorber.text_fragments[3].rectangle

                #self.assertAlmostEqual(1.589 if scale_wmf_fonts else 5.045, text_fragment_rectangle.width, delta=0.001)

    def test_embed_full_fonts(self):

        for embed_full_fonts in (False, True):
            with self.subTest(embed_full_fonts=embed_full_fonts):
                #ExStart
                #ExFor:PdfSaveOptions.__init__
                #ExFor:PdfSaveOptions.embed_full_fonts
                #ExSummary:Shows how to enable or disable subsetting when embedding fonts while rendering a document to PDF.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.font.name = "Arial"
                builder.writeln("Hello world!")
                builder.font.name = "Arvo"
                builder.writeln("The quick brown fox jumps over the lazy dog.")

                # Configure our font sources to ensure that we have access to both the fonts in this document.
                original_fonts_sources = aw.fonts.FontSettings.default_instance.get_fonts_sources()
                folder_font_source = aw.fonts.FolderFontSource(FONTS_DIR, True)
                aw.fonts.FontSettings.default_instance.set_fonts_sources([original_fonts_sources[0], folder_font_source])

                font_sources = aw.fonts.FontSettings.default_instance.get_fonts_sources()
                self.assertTrue(any(font.full_font_name == "Arial" for font in font_sources[0].get_available_fonts()))
                self.assertTrue(any(font.full_font_name == "Arvo" for font in font_sources[1].get_available_fonts()))

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

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.embed_full_fonts.pdf", options)

                if embed_full_fonts:
                    self.assertLess(500000, os.path.getsize(ARTIFACTS_DIR + "PdfSaveOptions.embed_full_fonts.pdf"))
                else:
                    self.assertGreater(25000, os.path.getsize(ARTIFACTS_DIR + "PdfSaveOptions.embed_full_fonts.pdf"))

                # Restore the original font sources.
                aw.fonts.FontSettings.default_instance.set_fonts_sources(original_fonts_sources)
                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.embed_full_fonts.pdf")

                #pdf_doc_fonts = pdf_document.font_utilities.get_all_fonts()

                #self.assertEqual("ArialMT", pdf_doc_fonts[0].font_name)
                #self.assertNotEqual(embed_full_fonts, pdf_doc_fonts[0].is_subset)

                #self.assertEqual("Arvo", pdf_doc_fonts[1].font_name)
                #self.assertNotEqual(embed_full_fonts, pdf_doc_fonts[1].is_subset)

    def test_embed_windows_fonts(self):

        for pdf_font_embedding_mode in (aw.saving.PdfFontEmbeddingMode.EMBED_ALL,
                                        aw.saving.PdfFontEmbeddingMode.EMBED_NONE,
                                        aw.saving.PdfFontEmbeddingMode.EMBED_NONSTANDARD):
            with self.subTest(pdf_font_embedding_mode=pdf_font_embedding_mode):
                #ExStart
                #ExFor:PdfSaveOptions.font_embedding_mode
                #ExFor:PdfFontEmbeddingMode
                #ExSummary:Shows how to set Aspose.Words to skip embedding Arial and Times New Roman fonts into a PDF document.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # "Arial" is a standard font, and "Courier New" is a nonstandard font.
                builder.font.name = "Arial"
                builder.writeln("Hello world!")
                builder.font.name = "Courier New"
                builder.writeln("The quick brown fox jumps over the lazy dog.")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()

                # Set the "embed_full_fonts" property to "True" to embed every glyph of every embedded font in the output PDF.
                options.embed_full_fonts = True

                # Set the "font_embedding_mode" property to "EMBED_ALL" to embed all fonts in the output PDF.
                # Set the "font_embedding_mode" property to "EMBED_NONSTANDARD" to only allow nonstandard fonts' embedding in the output PDF.
                # Set the "font_embedding_mode" property to "EMBED_NONE" to not embed any fonts in the output PDF.
                options.font_embedding_mode = pdf_font_embedding_mode

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.embed_windows_fonts.pdf", options)

                if pdf_font_embedding_mode == aw.saving.PdfFontEmbeddingMode.EMBED_ALL:
                    self.assertLess(1000000, os.path.getsize(ARTIFACTS_DIR + "PdfSaveOptions.embed_windows_fonts.pdf"))

                elif pdf_font_embedding_mode == aw.saving.PdfFontEmbeddingMode.EMBED_NONSTANDARD:
                    self.assertLess(480000, os.path.getsize(ARTIFACTS_DIR + "PdfSaveOptions.embed_windows_fonts.pdf"))

                elif pdf_font_embedding_mode == aw.saving.PdfFontEmbeddingMode.EMBED_NONE:
                    self.assertGreater(4243, os.path.getsize(ARTIFACTS_DIR + "PdfSaveOptions.embed_windows_fonts.pdf"))

                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.embed_windows_fonts.pdf")

                #pdf_doc_fonts = pdf_document.font_utilities.get_all_fonts()

                #self.assertEqual("ArialMT", pdf_doc_fonts[0].font_name)
                #self.assertEqual(
                #    pdf_font_embedding_mode == aw.saving.PdfFontEmbeddingMode.EMBED_ALL,
                #    pdf_doc_fonts[0].is_embedded)

                #self.assertEqual("CourierNewPSMT", pdf_doc_fonts[1].font_name)
                #self.assertEqual(
                #    pdf_font_embedding_mode in (aw.saving.PdfFontEmbeddingMode.EMBED_ALL, aw.PdfFontEmbeddingMode.EMBED_NONSTANDARD),
                #    pdf_doc_fonts[1].is_embedded)

    def test_embed_core_fonts(self):

        for use_core_fonts in (False, True):
            with self.subTest(use_core_fonts=use_core_fonts):
                #ExStart
                #ExFor:PdfSaveOptions.use_core_fonts
                #ExSummary:Shows how enable/disable PDF Type 1 font substitution.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.font.name = "Arial"
                builder.writeln("Hello world!")
                builder.font.name = "Courier New"
                builder.writeln("The quick brown fox jumps over the lazy dog.")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()

                # Set the "use_core_fonts" property to "True" to replace some fonts,
                # including the two fonts in our document, with their PDF Type 1 equivalents.
                # Set the "use_core_fonts" property to "False" to not apply PDF Type 1 fonts.
                options.use_core_fonts = use_core_fonts

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.embed_core_fonts.pdf", options)

                if use_core_fonts:
                    self.assertGreater(3000, os.path.getsize(ARTIFACTS_DIR + "PdfSaveOptions.embed_core_fonts.pdf"))
                else:
                    self.assertLess(30000, os.path.getsize(ARTIFACTS_DIR + "PdfSaveOptions.embed_core_fonts.pdf"))
                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.embed_core_fonts.pdf")

                #pdf_doc_fonts = pdf_document.font_utilities.get_all_fonts()

                #if use_core_fonts:
                #    self.assertEqual("Helvetica", pdf_doc_fonts[0].font_name)
                #    self.assertEqual("Courier", pdf_doc_fonts[1].font_name)
                #else:
                #    self.assertEqual("ArialMT", pdf_doc_fonts[0].font_name)
                #    self.assertEqual("CourierNewPSMT", pdf_doc_fonts[1].font_name)

                #self.assertNotEqual(use_core_fonts, pdf_doc_fonts[0].is_embedded)
                #self.assertNotEqual(use_core_fonts, pdf_doc_fonts[1].is_embedded)

    def test_additional_text_positioning(self):

        for apply_additional_text_positioning in (False, True):
            with self.subTest(apply_additional_text_positioning=apply_additional_text_positioning):
                #ExStart
                #ExFor:PdfSaveOptions.additional_text_positioning
                #ExSummary:Show how to write additional text positioning operators.
                doc = aw.Document(MY_DIR + "Text positioning operators.docx")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                save_options = aw.saving.PdfSaveOptions()
                save_options.text_compression = aw.saving.PdfTextCompression.NONE

                # Set the "additional_text_positioning" property to "True" to attempt to fix incorrect
                # element positioning in the output PDF, should there be any, at the cost of increased file size.
                # Set the "additional_text_positioning" property to "False" to render the document as usual.
                save_options.additional_text_positioning = apply_additional_text_positioning

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.additional_text_positioning.pdf", save_options)
                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.additional_text_positioning.pdf")
                #text_absorber = aspose.pdf.text.TextFragmentAbsorber()

                #pdf_document.pages[1].accept(text_absorber)

                #tj_operator = text_absorber.text_fragments[1].page.contents[85].as_set_glyphs_position_show_text()

                #if apply_additional_text_positioning:
                #    self.assertLess(100000, os.path.getsize(ARTIFACTS_DIR + "PdfSaveOptions.additional_text_positioning.pdf"))
                #    self.assertEqual(
                #        "[0 (S) 0 (a) 0 (m) 0 (s) 0 (t) 0 (a) -1 (g) 1 (,) 0 ( ) 0 (1) 0 (0) 0 (.) 0 ( ) 0 (N) 0 (o) 0 (v) 0 (e) 0 (m) 0 (b) 0 (e) 0 (r) -1 ( ) 1 (2) -1 (0) 0 (1) 0 (8)] TJ",
                #        tj_operator.to_string())
                #else:
                #    self.assertLess(97000, os.path.getsize(ARTIFACTS_DIR + "PdfSaveOptions.additional_text_positioning.pdf"))
                #    self.assertEqual(
                #        "[(Samsta) -1 (g) 1 (, 10. November) -1 ( ) 1 (2) -1 (018)] TJ",
                #        tj_operator.to_string())

    def test_save_as_pdf_book_fold(self):

        for render_text_as_bookfold in (False, True):
            with self.subTest(render_text_as_bookfold=render_text_as_bookfold):
                #ExStart
                #ExFor:PdfSaveOptions.use_book_fold_printing_settings
                #ExSummary:Shows how to save a document to the PDF format in the form of a book fold.
                doc = aw.Document(MY_DIR + "Paragraphs.docx")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()

                # Set the "use_book_fold_printing_settings" property to "True" to arrange the contents
                # in the output PDF in a way that helps us use it to make a booklet.
                # Set the "use_book_fold_printing_settings" property to "False" to render the PDF normally.
                options.use_book_fold_printing_settings = render_text_as_bookfold

                # If we are rendering the document as a booklet, we must set the "multiple_pages"
                # properties of the page setup objects of all sections to "MultiplePagesType.BOOK-FOLD_PRINTING".
                if render_text_as_bookfold:
                    for section in doc.sections:
                        section = section.as_section()
                        section.page_setup.multiple_pages = aw.settings.MultiplePagesType.BOOK_FOLD_PRINTING

                # Once we print this document on both sides of the pages, we can fold all the pages down the middle at once,
                # and the contents will line up in a way that creates a booklet.
                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.save_as_pdf_book_fold.pdf", options)
                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.save_as_pdf_book_fold.pdf")
                #text_absorber = TextAbsorber()

                #pdf_document.pages.accept(text_absorber)

                #if render_text_as_bookfold:
                #    self.assertTrue(text_absorber.text.index_of("Heading #1", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #2", StringComparison.ORDINAL))
                #    self.assertTrue(text_absorber.text.index_of("Heading #2", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #3", StringComparison.ORDINAL))
                #    self.assertTrue(text_absorber.text.index_of("Heading #3", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #4", StringComparison.ORDINAL))
                #    self.assertTrue(text_absorber.text.index_of("Heading #4", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #5", StringComparison.ORDINAL))
                #    self.assertTrue(text_absorber.text.index_of("Heading #5", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #6", StringComparison.ORDINAL))
                #    self.assertTrue(text_absorber.text.index_of("Heading #6", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #7", StringComparison.ORDINAL))
                #    self.assertFalse(text_absorber.text.index_of("Heading #7", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #8", StringComparison.ORDINAL))
                #    self.assertTrue(text_absorber.text.index_of("Heading #8", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #9", StringComparison.ORDINAL))
                #    self.assertFalse(text_absorber.text.index_of("Heading #9", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #10", StringComparison.ORDINAL))
                #else:
                #    self.assertTrue(text_absorber.text.index_of("Heading #1", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #2", StringComparison.ORDINAL))
                #    self.assertTrue(text_absorber.text.index_of("Heading #2", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #3", StringComparison.ORDINAL))
                #    self.assertTrue(text_absorber.text.index_of("Heading #3", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #4", StringComparison.ORDINAL))
                #    self.assertTrue(text_absorber.text.index_of("Heading #4", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #5", StringComparison.ORDINAL))
                #    self.assertTrue(text_absorber.text.index_of("Heading #5", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #6", StringComparison.ORDINAL))
                #    self.assertTrue(text_absorber.text.index_of("Heading #6", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #7", StringComparison.ORDINAL))
                #    self.assertTrue(text_absorber.text.index_of("Heading #7", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #8", StringComparison.ORDINAL))
                #    self.assertTrue(text_absorber.text.index_of("Heading #8", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #9", StringComparison.ORDINAL))
                #    self.assertTrue(text_absorber.text.index_of("Heading #9", StringComparison.ORDINAL) < text_absorber.text.index_of("Heading #10", StringComparison.ORDINAL))

    def test_zoom_behaviour(self):

        #ExStart
        #ExFor:PdfSaveOptions.zoom_behavior
        #ExFor:PdfSaveOptions.zoom_factor
        #ExFor:PdfZoomBehavior
        #ExSummary:Shows how to set the default zooming that a reader applies when opening a rendered PDF document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world!")

        # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
        # to modify how that method converts the document to .PDF.
        # Set the "zoom_behavior" property to "PdfZoomBehavior.ZOOM_FACTOR" to get a PDF reader to
        # apply a percentage-based zoom factor when we open the document with it.
        # Set the "zoom_factor" property to "25" to give the zoom factor a value of 25%.
        options = aw.saving.PdfSaveOptions()
        options.zoom_behavior = aw.saving.PdfZoomBehavior.ZOOM_FACTOR
        options.zoom_factor = 25

        # When we open this document using a reader such as Adobe Acrobat, we will see the document scaled at 1/4 of its actual size.
        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.zoom_behaviour.pdf", options)
        #ExEnd

        #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.zoom_behaviour.pdf")
        #action = pdf_document.open_action.as_go_to_action()

        #self.assertEqual(0.25, action.destination.as_xyz_explicit_destination().zoom)

    def test_page_mode(self):

        for page_mode in (aw.saving.PdfPageMode.FULL_SCREEN,
                          aw.saving.PdfPageMode.USE_THUMBS,
                          aw.saving.PdfPageMode.USE_OC,
                          aw.saving.PdfPageMode.USE_OUTLINES,
                          aw.saving.PdfPageMode.USE_NONE,
                          aw.saving.PdfPageMode.USE_ATTACHMENTS):
            with self.subTest(page_mode=page_mode):
                #ExStart
                #ExFor:PdfSaveOptions.page_mode
                #ExFor:PdfPageMode
                #ExSummary:Shows how to set instructions for some PDF readers to follow when opening an output document.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.writeln("Hello world!")

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

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.page_mode.pdf", options)
                #ExEnd

                doc_locale_name = CultureInfo(doc.styles.default_font.locale_id).name

                with open(ARTIFACTS_DIR + "PdfSaveOptions.page_mode.pdf", "rb") as file:
                    content = file.read().decode('utf-8')

                    if page_mode == aw.saving.PdfPageMode.FULL_SCREEN:
                        self.assertIn(
                            "<</Type /Catalog/Pages 3 0 R/PageMode /FullScreen/Lang({})/Metadata 4 0 R>>\r\n".format(doc_locale_name),
                            content)

                    elif page_mode == aw.saving.PdfPageMode.USE_THUMBS:
                        self.assertIn(
                            "<</Type /Catalog/Pages 3 0 R/PageMode /UseThumbs/Lang({})/Metadata 4 0 R>>".format(doc_locale_name),
                            content)

                    elif page_mode == aw.saving.PdfPageMode.USE_OC:
                        self.assertIn(
                            "<</Type /Catalog/Pages 3 0 R/PageMode /UseOC/Lang({})/Metadata 4 0 R>>\r\n".format(doc_locale_name),
                            content)

                    elif page_mode in (aw.saving.PdfPageMode.USE_OUTLINES, aw.saving.PdfPageMode.USE_NONE):
                        self.assertIn(
                            "<</Type /Catalog/Pages 3 0 R/Lang({})/Metadata 4 0 R>>\r\n".format(doc_locale_name),
                            content)
                     
                    elif page_mode == aw.saving.PdfPageMode.USE_ATTACHMENTS:
                        self.assertIn(
                            f"<</Type /Catalog/Pages 3 0 R/PageMode /UseAttachments/Lang({doc_locale_name})/Metadata 4 0 R>>\r\n",
                            content)

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.page_mode.pdf")

                #if page_mode in (aw.saving.PdfPageMode.USE_NONE, aw.saving.PdfPageMode.USE_OUTLINES):
                #    self.assertEqual(aspose.pdf.PageMode.USE_NONE, pdf_document.page_mode)

                #elif page_mode == aw.saving.PdfPageMode.USE_THUMBS:
                #    self.assertEqual(aspose.pdf.PageMode.USE_THUMBS, pdf_document.page_mode)

                #elif page_mode == aw.saving.PdfPageMode.FULL_SCREEN:
                #    self.assertEqual(aspose.pdf.PageMode.FULL_SCREEN, pdf_document.page_mode)

                #elif page_mode == aw.saving.PdfPageMode.USE_OC:
                #    self.assertEqual(aspose.pdf.PageMode.USE_OC, pdf_document.page_mode)

                #elif page_mode == aw.saving.PdfPageMode.USE_ATTACHMENTS:
                #    self.assertEqual(aspose.pdf.PageMode.USE_ATTACHMENTS, pdf_document.page_mode)

    def test_note_hyperlinks(self):

        for create_note_hyperlinks in (False, True):
            with self.subTest(create_note_hyperlinks=create_note_hyperlinks):
                #ExStart
                #ExFor:PdfSaveOptions.create_note_hyperlinks
                #ExSummary:Shows how to make footnotes and endnotes function as hyperlinks.
                doc = aw.Document(MY_DIR + "Footnotes and endnotes.docx")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()

                # Set the "create_note_hyperlinks" property to "True" to turn all footnote/endnote symbols
                # in the text act as links that, upon clicking, take us to their respective footnotes/endnotes.
                # Set the "create_note_hyperlinks" property to "False" not to have footnote/endnote symbols link to anything.
                options.create_note_hyperlinks = create_note_hyperlinks

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.note_hyperlinks.pdf", options)
                #ExEnd

                with open(ARTIFACTS_DIR + "PdfSaveOptions.note_hyperlinks.pdf", "rb") as file:
                    content = file.read()

                if create_note_hyperlinks:
                    self.assertIn(
                        b"<</Type /Annot/Subtype /Link/Rect [157.80099487 720.90106201 159.35600281 733.55004883]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 677 0]>>",
                        content)
                    self.assertIn(
                        b"<</Type /Annot/Subtype /Link/Rect [202.16900635 720.90106201 206.06201172 733.55004883]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 79 0]>>",
                        content)
                    self.assertIn(
                        b"<</Type /Annot/Subtype /Link/Rect [212.23199463 699.2510376 215.34199524 711.90002441]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 654 0]>>",
                        content)
                    self.assertIn(
                        b"<</Type /Annot/Subtype /Link/Rect [258.15499878 699.2510376 262.04800415 711.90002441]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 68 0]>>",
                        content)
                    self.assertIn(
                        b"<</Type /Annot/Subtype /Link/Rect [85.05000305 68.19904327 88.66500092 79.69804382]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 202 733 0]>>",
                        content)
                    self.assertIn(
                        b"<</Type /Annot/Subtype /Link/Rect [85.05000305 56.70004272 88.66500092 68.19904327]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 258 711 0]>>",
                        content)
                    self.assertIn(
                        b"<</Type /Annot/Subtype /Link/Rect [85.05000305 666.10205078 86.4940033 677.60107422]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 157 733 0]>>",
                        content)
                    self.assertIn(
                        b"<</Type /Annot/Subtype /Link/Rect [85.05000305 643.10406494 87.93800354 654.60308838]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 212 711 0]>>",
                        content)
                else:
                    self.assertNotIn(
                        b"<</Type /Annot/Subtype /Link/Rect",
                        content)

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.note_hyperlinks.pdf")
                #page = pdf_document.pages[1]
                #annotation_selector = aspose.pdf.AnnotationSelector(aspose.pdf.LinkAnnotation(page, aspose.pdf.Rectangle.TRIVIAL))

                #page.accept(annotation_selector)

                #link_annotations = [x.as_link_annotation() for x in annotation_selector.selected]

                #if create_note_hyperlinks:
                #    self.assertEqual(8, len([a for a in link_annotations if a.annotation_type == aspose.pdf.annotations.AnnotationType.LINK]))

                #    self.assertEqual("1 XYZ 85 677 0", link_annotations[0].destination.to_string())
                #    self.assertEqual("1 XYZ 85 79 0", link_annotations[1].destination.to_string())
                #    self.assertEqual("1 XYZ 85 654 0", link_annotations[2].destination.to_string())
                #    self.assertEqual("1 XYZ 85 68 0", link_annotations[3].destination.to_string())
                #    self.assertEqual("1 XYZ 202 733 0", link_annotations[4].destination.to_string())
                #    self.assertEqual("1 XYZ 258 711 0", link_annotations[5].destination.to_string())
                #    self.assertEqual("1 XYZ 157 733 0", link_annotations[6].destination.to_string())
                #    self.assertEqual("1 XYZ 212 711 0", link_annotations[7].destination.to_string())
                #else:
                #    self.assertEqual(0, annotation_selector.selected.count)

    def test_custom_properties_export(self):

        for pdf_custom_properties_export_mode in (aw.saving.PdfCustomPropertiesExport.NONE,
                                                  aw.saving.PdfCustomPropertiesExport.STANDARD,
                                                  aw.saving.PdfCustomPropertiesExport.METADATA):
            with self.subTest(pdf_custom_properties_export_mode=pdf_custom_properties_export_mode):
                #ExStart
                #ExFor:PdfCustomPropertiesExport
                #ExFor:PdfSaveOptions.custom_properties_export
                #ExSummary:Shows how to export custom properties while converting a document to PDF.
                doc = aw.Document()

                doc.custom_document_properties.add("Company", "My value")

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

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.custom_properties_export.pdf", options)
                #ExEnd

                with open(ARTIFACTS_DIR + "PdfSaveOptions.custom_properties_export.pdf", "rb") as file:
                    content = file.read()

                if pdf_custom_properties_export_mode == aw.saving.PdfCustomPropertiesExport.NONE:
                    self.assertNotIn(doc.custom_document_properties[0].name.encode('ascii'), content)
                    self.assertNotIn(
                        b"<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>",
                        content)

                elif pdf_custom_properties_export_mode == aw.saving.PdfCustomPropertiesExport.STANDARD:
                    self.assertIn(
                        b"<</Creator(\xFE\xFF\0A\0s\0p\0o\0s\0e\0.\0W\0o\0r\0d\0s)/Producer(\xFE\xFF\0A\0s\0p\0o\0s\0e\0.\0W\0o\0r\0d\0s\0 \0f\0o\0r\0",
                        content)
                    self.assertIn(
                        b"/Company (\xFE\xFF\0M\0y\0 \0v\0a\0l\0u\0e)>>",
                        content)

                elif pdf_custom_properties_export_mode == aw.saving.PdfCustomPropertiesExport.METADATA:
                    self.assertIn(
                        b"<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>",
                        content)

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.custom_properties_export.pdf")

                #self.assertEqual("Aspose.Words", pdf_document.info.creator)
                #self.assertTrue(pdf_document.info.producer.startswith("Aspose.Words"))

                #if pdf_custom_properties_export_mode == aw.saving.PdfCustomPropertiesExport.NONE:
                #    self.assertEqual(2, pdf_document.info.count)
                #    self.assertEqual(3, pdf_document.metadata.count)

                #elif pdf_custom_properties_export_mode == aw.saving.PdfCustomPropertiesExport.METADATA:
                #    self.assertEqual(2, pdf_document.info.count)
                #    self.assertEqual(4, pdf_document.metadata.count)

                #    self.assertEqual("Aspose.Words", pdf_document.metadata["xmp:CreatorTool"].to_string())
                #    self.assertEqual("Company", pdf_document.metadata["custprops:Property1"].to_string())

                #elif pdf_custom_properties_export_mode == aw.saving.PdfCustomPropertiesExport.STANDARD:
                #    self.assertEqual(3, pdf_document.info.count)
                #    self.assertEqual(3, pdf_document.metadata.count)

                #    self.assertEqual("My value", pdf_document.info["Company"])

    def test_drawing_ml_effects(self):

        for effects_rendering_mode in (aw.saving.DmlEffectsRenderingMode.NONE,
                                       aw.saving.DmlEffectsRenderingMode.SIMPLIFIED,
                                       aw.saving.DmlEffectsRenderingMode.FINE):
            with self.subTest(effects_rendering_mode=effects_rendering_mode):
                #ExStart
                #ExFor:DmlRenderingMode
                #ExFor:DmlEffectsRenderingMode
                #ExFor:PdfSaveOptions.dml_effects_rendering_mode
                #ExFor:SaveOptions.dml_effects_rendering_mode
                #ExFor:SaveOptions.dml_rendering_mode
                #ExSummary:Shows how to configure the rendering quality of DrawingML effects in a document as we save it to PDF.
                doc = aw.Document(MY_DIR + "DrawingML shape effects.docx")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()

                # Set the "dml_effects_rendering_mode" property to "DmlEffectsRenderingMode.NONE" to discard all DrawingML effects.
                # Set the "dml_effects_rendering_mode" property to "DmlEffectsRenderingMode.SIMPLIFIED"
                # to render a simplified version of DrawingML effects.
                # Set the "dml_effects_rendering_mode" property to "DmlEffectsRenderingMode.FINE" to
                # render DrawingML effects with more accuracy and also with more processing cost.
                options.dml_effects_rendering_mode = effects_rendering_mode

                self.assertEqual(aw.saving.DmlRenderingMode.DRAWING_ML, options.dml_rendering_mode)

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.drawing_ml_effects.pdf", options)
                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.drawing_ml_effects.pdf")

                #image_placement_absorber = aspose.pdf.ImagePlacementAbsorber()
                #image_placement_absorber.visit(pdf_document.pages[1])

                #table_absorber = aspose.pdf.text.TableAbsorber()
                #table_absorber.visit(pdf_document.pages[1])

                #with open(ARTIFACTS_DIR + "PdfSaveOptions.drawing_m_l_effects.pdf", "rb") as file:
                #    content = file.read()

                #if effects_rendering_mode in (aw.saving.DmlEffectsRenderingMode.NONE,
                #                              aw.saving.DmlEffectsRenderingMode.SIMPLIFIED):
                #    self.assertIn(
                #        b"5 0 obj\r\n" +
                #        b"<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                #        content)
                #    self.assertEqual(0, image_placement_absorber.image_placements.count)
                #    self.assertEqual(28, table_absorber.table_list.count)

                #elif effects_rendering_mode == aw.saving.DmlEffectsRenderingMode.FINE:
                #    self.assertIn(
                #        b"5 0 obj\r\n<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R>>/XObject<</X1 10 0 R/X2 11 0 R/X3 12 0 R/X4 13 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                #        content)
                #    self.assertEqual(21, image_placement_absorber.image_placements.count)
                #    self.assertEqual(4, table_absorber.table_list.count)

    def test_drawing_ml_fallback(self):

        for dml_rendering_mode in (aw.saving.DmlRenderingMode.FALLBACK,
                                   aw.saving.DmlRenderingMode.DRAWING_ML):
            with self.subTest(dml_rendering_mode=dml_rendering_mode):
                #ExStart
                #ExFor:DmlRenderingMode
                #ExFor:SaveOptions.dml_rendering_mode
                #ExSummary:Shows how to render fallback shapes when saving to PDF.
                doc = aw.Document(MY_DIR + "DrawingML shape fallbacks.docx")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()

                # Set the "dml_rendering_mode" property to "DmlRenderingMode.FALLBACK"
                # to substitute DML shapes with their fallback shapes.
                # Set the "dml_rendering_mode" property to "DmlRenderingMode.DRAWING_ML"
                # to render the DML shapes themselves.
                options.dml_rendering_mode = dml_rendering_mode

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.drawing_ml_fallback.pdf", options)
                #ExEnd

                with open(ARTIFACTS_DIR + "PdfSaveOptions.drawing_ml_fallback.pdf", "rb") as file:
                    content = file.read()

                if dml_rendering_mode == aw.saving.DmlRenderingMode.DRAWING_ML:
                    self.assertIn(
                        b"<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABB 11 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        content)

                elif dml_rendering_mode ==  aw.saving.DmlRenderingMode.FALLBACK:
                    self.assertIn(
                        b"5 0 obj\r\n<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABD 13 0 R>>/ExtGState<</GS1 10 0 R/GS2 11 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        content)

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.drawing_ml_fallback.pdf")

                #image_placement_absorber = aspose.pdf.ImagePlacementAbsorber()
                #image_placement_absorber.visit(pdf_document.pages[1])

                #table_absorber = aspose.pdf.text.TableAbsorber()
                #table_absorber.visit(pdf_document.pages[1])

                #if dml_rendering_mode == aw.saving.DmlRenderingMode.DRAWING_ML:
                #    self.assertEqual(6, table_absorber.table_list.count)

                #elif dml_rendering_mode == aw.saving.DmlRenderingMode.FALLBACK:
                #    self.assertEqual(15, table_absorber.table_list.count)

    def test_export_document_structure(self):

        for export_document_structure in (False, True):
            with self.subTest(export_document_structure=export_document_structure):
                #ExStart
                #ExFor:PdfSaveOptions.export_document_structure
                #ExSummary:Shows how to preserve document structure elements, which can assist in programmatically interpreting our document.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.paragraph_format.style = doc.styles.get_by_name("Heading 1")
                builder.writeln("Hello world!")
                builder.paragraph_format.style = doc.styles.get_by_name("Normal")
                builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()

                # Set the "export_document_structure" property to "True" to make the document structure, such tags, available via the
                # "Content" navigation pane of Adobe Acrobat at the cost of increased file size.
                # Set the "export_document_structure" property to "False" to not export the document structure.
                options.export_document_structure = export_document_structure

                # Suppose we export document structure while saving this document. In that case,
                # we can open it using Adobe Acrobat and find tags for elements such as the heading
                # and the next paragraph via "View" -> "Show/Hide" -> "Navigation panes" -> "Tags".
                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.export_document_structure.pdf", options)
                #ExEnd

                with open(ARTIFACTS_DIR + "PdfSaveOptions.export_document_structure.pdf", "rb") as file:
                    content = file.read()

                if export_document_structure:
                    self.assertIn(
                        b"5 0 obj\r\n" +
                        b"<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABC 12 0 R>>/ExtGState<</GS1 10 0 R/GS2 14 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>/StructParents 0/Tabs /S>>",
                        content)
                else:
                    self.assertIn(
                        b"5 0 obj\r\n" +
                        b"<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABB 11 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        content)

    def test_preblend_images(self):

        for preblend_images in (False, True):
            with self.subTest(preblend_images=preblend_images):
                #ExStart
                #ExFor:PdfSaveOptions.preblend_images
                #ExSummary:Shows how to preblend images with transparent backgrounds while saving a document to PDF.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                img = drawing.Image.from_file(IMAGE_DIR + "Transparent background logo.png")
                builder.insert_image(img)

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                options = aw.saving.PdfSaveOptions()

                # Set the "preblend_images" property to "True" to preblend transparent images
                # with a background, which may reduce artifacts.
                # Set the "preblend_images" property to "False" to render transparent images normally.
                options.preblend_images = preblend_images

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.preblend_images.pdf", options)
                #ExEnd

                pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.preblend_images.pdf")
                image = pdf_document.pages[1].resources.images[1]

                with open(ARTIFACTS_DIR + "PdfSaveOptions.preblend_images.pdf", "rb") as file:
                    content = file.read()

                with io.BytesIO() as stream:
                    image.save(stream)

                    if preblend_images:
                        self.assertIn("11 0 obj\r\n20849 ", content)
                        self.assertEqual(17898, len(stream.getvalue()))
                    else:
                        self.assertIn("11 0 obj\r\n19289 ", content)
                        self.assertEqual(19216, len(stream.getvalue()))

    def test_interpolate_images(self):

        for interpolate_images in (False, True):
            with self.subTest(interpolate_images=interpolate_images):
                #ExStart
                #ExFor:PdfSaveOptions.interpolate_images
                #ExSummary:Shows how to perform interpolation on images while saving a document to PDF.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                img = drawing.Image.from_file(IMAGE_DIR + "Transparent background logo.png")
                builder.insert_image(img)

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                save_options = aw.saving.PdfSaveOptions()

                # Set the "interpolate_images" property to "True" to get the reader that opens this document to interpolate images.
                # Their resolution should be lower than that of the device that is displaying the document.
                # Set the "interpolate_images" property to "False" to make it so that the reader does not apply any interpolation.
                save_options.interpolate_images = interpolate_images

                # When we open this document with a reader such as Adobe Acrobat, we will need to zoom in on the image
                # to see the interpolation effect if we saved the document with it enabled.
                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.interpolate_images.pdf", save_options)
                #ExEnd

                with open(ARTIFACTS_DIR + "PdfSaveOptions.interpolate_images.pdf", "rb") as file:
                    content = file.read()

                if interpolate_images:
                    self.assertIn(
                        b"7 0 obj\r\n" +
                        b"<</Type /XObject/Subtype /Image/Width 400/Height 400/ColorSpace /DeviceRGB/BitsPerComponent 8/SMask 10 0 R/Interpolate True/Length 11 0 R/Filter /FlateDecode>>",
                        content)
                else:
                    self.assertIn(
                        b"7 0 obj\r\n" +
                        b"<</Type /XObject/Subtype /Image/Width 400/Height 400/ColorSpace /DeviceRGB/BitsPerComponent 8/SMask 10 0 R/Length 11 0 R/Filter /FlateDecode>>",
                        content)

    #def test_dml3d_effects_rendering_mode_test(self):

    #    doc = aw.Document(MY_DIR + "DrawingML shape 3D effects.docx")

    #    warning_callback = ExPdfSaveOptions.RenderCallback()
    #    doc.warning_callback = warning_callback

    #    save_options = aw.saving.PdfSaveOptions()
    #    save_options.dml3_d_effects_rendering_mode = aw.saving.Dml3DEffectsRenderingMode.ADVANCED

    #    doc.save(ARTIFACTS_DIR + "PdfSaveOptions.dml3_d_effects_rendering_mode_test.pdf", save_options)

    #    self.assertEqual(38, warning_callback.count)

    #class RenderCallback(aw.IWarningCallback):

    #    def __init__(self):
    #        self.warnings: List[aw.WarningInfo] = []

    #    def warning(info: aw.WarningInfo):

    #        print(f"{info.warning_type}: {info.description}.")
    #        self.warnings.add(info)

    #    def __getitem__(self, i) -> aw.WarningInfo:
    #        return self.warnings[i]

    #    def clear(self):
    #        """Clears warning collection."""
    #        self.warnings.clear()

    #    @property
    #    def count(self):
    #        return len(self.warnings)

    #    def contains(self, source: aw.WarningSource, type: aw.WarningType, description: str) -> bool:
    #        """Returns True if a warning with the specified properties has been generated."""

    #        return any(warning for warning in self.warnings
    #                   if warning.source == source and warning.warning_type == type and warning.description == description)

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
        builder.writeln("Contents of signed PDF.")

        certificate_holder = aw.digitalsignatures.CertificateHolder.create(MY_DIR + "morzal.pfx", "aw")

        # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
        # to modify how that method converts the document to .PDF.
        options = aw.saving.PdfSaveOptions()

        # Configure the "digital_signature_details" object of the "SaveOptions" object to
        # digitally sign the document as we render it with the "save" method.
        signing_time = datetime.now()
        import aspose.words.saving as aws
        options.digital_signature_details = aw.saving.PdfDigitalSignatureDetails(certificate_holder, "Test Signing", "My Office", signing_time)
        options.digital_signature_details.hash_algorithm = aw.saving.PdfDigitalSignatureHashAlgorithm.RIPE_MD160

        self.assertEqual("Test Signing", options.digital_signature_details.reason)
        self.assertEqual("My Office", options.digital_signature_details.location)
        self.assertEqual(signing_time.astimezone(timezone.utc), options.digital_signature_details.signature_date)

        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.pdf_digital_signature.pdf", options)
        #ExEnd

        with open(ARTIFACTS_DIR + "PdfSaveOptions.pdf_digital_signature.pdf", "rb") as file:
            content = file.read()

        self.assertIn(
            b"7 0 obj\r\n" +
            b"<</Type /Annot/Subtype /Widget/Rect [0 0 0 0]/FT /Sig/T",
            content)

        self.assertFalse(aw.FileFormatUtil.detect_file_format(ARTIFACTS_DIR + "PdfSaveOptions.pdf_digital_signature.pdf").has_digital_signature)

        #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.pdf_digital_signature.pdf")

        #self.assertTrue(pdf_document.form.signatures_exist)

        #signature_field = pdf_document.form[1].as_signature_field()

        #self.assertEqual("AsposeDigitalSignature", signature_field.full_name)
        #self.assertEqual("AsposeDigitalSignature", signature_field.partial_name)
        #self.assertEqual(type(aspose.pdf.forms.PKCS7_DETACHED), signature_field.signature.get_type())
        #self.assertEqual(date.today(), signature_field.signature.date.date())
        #self.assertEqual("\xFE\xFF\0M\0o\0r\0z\0a\0l\0.\0M\0e", signature_field.signature.authority)
        #self.assertEqual("\xFE\xFF\0M\0y\0 \0O\0f\0f\0i\0c\0e", signature_field.signature.location)
        #self.assertEqual("\xFE\xFF\0T\0e\0s\0t\0 \0S\0i\0g\0n\0i\0n\0g", signature_field.signature.reason)

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
        builder.writeln("Signed PDF contents.")

        # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
        # to modify how that method converts the document to .PDF.
        options = aw.saving.PdfSaveOptions()

        # Create a digital signature and assign it to our SaveOptions object to sign the document when we save it to PDF.
        certificate_holder = aw.digitalsignatures.CertificateHolder.create(MY_DIR + "morzal.pfx", "aw")
        options.digital_signature_details = aw.saving.PdfDigitalSignatureDetails(certificate_holder, "Test Signing", "Aspose Office", datetime.now())

        # Create a timestamp authority-verified timestamp.
        options.digital_signature_details.timestamp_settings = aw.saving.PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword")

        # The default lifespan of the timestamp is 100 seconds.
        self.assertEqual(100.0, options.digital_signature_details.timestamp_settings.timeout.total_seconds())

        # We can set our timeout period via the constructor.
        options.digital_signature_details.timestamp_settings = aw.saving.PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword", timedelta(minutes=30))

        self.assertEqual(1800.0, options.digital_signature_details.timestamp_settings.timeout.total_seconds())
        self.assertEqual("https://freetsa.org/tsr", options.digital_signature_details.timestamp_settings.server_url)
        self.assertEqual("JohnDoe", options.digital_signature_details.timestamp_settings.user_name)
        self.assertEqual("MyPassword", options.digital_signature_details.timestamp_settings.password)

        # The "save" method will apply our signature to the output document at this time.
        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.pdf_digital_signature_timestamp.pdf", options)
        #ExEnd

        self.assertFalse(aw.FileFormatUtil.detect_file_format(ARTIFACTS_DIR + "PdfSaveOptions.pdf_digital_signature_timestamp.pdf").has_digital_signature)

        with open(ARTIFACTS_DIR + "PdfSaveOptions.pdf_digital_signature_timestamp.pdf", "rb") as file:
            content = file.read()

        self.assertIn(
            b"7 0 obj\r\n" +
            b"<</Type /Annot/Subtype /Widget/Rect [0 0 0 0]/FT /Sig/T",
            content)

        #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.pdf_digital_signature_timestamp.pdf")

        #self.assertTrue(pdf_document.form.signatures_exist)

        #signature_field = pdf_document.form[1].as_signature_field()

        #self.assertEqual("AsposeDigitalSignature", signature_field.full_name)
        #self.assertEqual("AsposeDigitalSignature", signature_field.partial_name)
        #self.assertEqual(type(aspose.pdf.forms.PKCS7_DETACHED), signature_field.signature.get_type())
        #self.assertEqual(datetime(1, 1, 1, 0, 0, 0), signature_field.signature.date)
        #self.assertEqual("\xFE\xFF\0M\0o\0r\0z\0a\0l\0.\0M\0e", signature_field.signature.authority)
        #self.assertEqual("\xFE\xFF\0A\0s\0p\0o\0s\0e\0 \0O\0f\0f\0i\0c\0e", signature_field.signature.location)
        #self.assertEqual("\xFE\xFF\0T\0e\0s\0t\0 \0S\0i\0g\0n\0i\0n\0g", signature_field.signature.reason)
        #self.assertIsNone(signature_field.signature.timestamp_settings)

    def test_render_metafile(self):

        for rendering_mode in (aw.saving.EmfPlusDualRenderingMode.EMF,
                               aw.saving.EmfPlusDualRenderingMode.EMF_PLUS,
                               aw.saving.EmfPlusDualRenderingMode.EMF_PLUS_WITH_FALLBACK):
            with self.subTest(rendering_mode=rendering_mode):
                #ExStart
                #ExFor:EmfPlusDualRenderingMode
                #ExFor:MetafileRenderingOptions.emf_plus_dual_rendering_mode
                #ExFor:MetafileRenderingOptions.use_emf_embedded_to_wmf
                #ExSummary:Shows how to configure Enhanced Windows Metafile-related rendering options when saving to PDF.
                doc = aw.Document(MY_DIR + "EMF.docx")

                # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to .PDF.
                save_options = aw.saving.PdfSaveOptions()

                # Set the "emf_plus_dual_rendering_mode" property to "EmfPlusDualRenderingMode.EMF"
                # to only render the EMF part of an EMF+ dual metafile.
                # Set the "emf_plus_dual_rendering_mode" property to "EmfPlusDualRenderingMode.EMF_PLUS" to
                # to render the EMF+ part of an EMF+ dual metafile.
                # Set the "emf_plus_dual_rendering_mode" property to "EmfPlusDualRenderingMode.EMF_PLUS_WITH_FALLBACK"
                # to render the EMF+ part of an EMF+ dual metafile if all of the EMF+ records are supported.
                # Otherwise, Aspose.Words will render the EMF part.
                save_options.metafile_rendering_options.emf_plus_dual_rendering_mode = rendering_mode

                # Set the "use_emf_embedded_to_wmf" property to "True" to render embedded EMF data
                # for metafiles that we can render as vector graphics.
                save_options.metafile_rendering_options.use_emf_embedded_to_wmf = True

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.render_metafile.pdf", save_options)
                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.render_metafile.pdf")

                #with open(ARTIFACTS_DIR + "PdfSaveOptions.render_metafile.pdf", "rb") as file:
                #    content = file.read()

                #if rendering_mode in (aw.saving.EmfPlusDualRenderingMode.EMF,
                #                      aw.saving.EmfPlusDualRenderingMode.EMF_PLUS_WITH_FALLBACK):
                #    self.assertEqual(0, pdf_document.pages[1].resources.images.count)
                #    self.assertIn(
                #        b"5 0 obj\r\n" +
                #        b"<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 595.29998779 841.90002441]/Resources<</Font<</FAAAAI 8 0 R/FAAABB 11 0 R/FAAABE 14 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                #        content)
                #    break

                #elif rendering_mode == aw.saving.EmfPlusDualRenderingMode.EMF_PLUS:
                #    self.assertEqual(1, pdf_document.pages[1].resources.images.count)
                #    self.assertIn(
                #        b"5 0 obj\r\n" +
                #        b"<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 595.29998779 841.90002441]/Resources<</Font<</FAAAAI 8 0 R/FAAABC 12 0 R/FAAABF 15 0 R>>/XObject<</X1 10 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                #        content)

    def test_encryption_permissions(self):

        #ExStart
        #ExFor:PdfEncryptionDetails.__init__
        #ExFor:PdfSaveOptions.encryption_details
        #ExFor:PdfEncryptionDetails.permissions
        #ExFor:PdfEncryptionDetails.owner_password
        #ExFor:PdfEncryptionDetails.user_password
        #ExFor:PdfPermissions
        #ExFor:PdfEncryptionDetails
        #ExSummary:Shows how to set permissions on a saved PDF document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello world!")

        encryption_details = aw.saving.PdfEncryptionDetails("password", "")

        # Start by disallowing all permissions.
        encryption_details.permissions = aw.saving.PdfPermissions.DISALLOW_ALL

        # Extend permissions to allow the editing of annotations.
        encryption_details.permissions = aw.saving.PdfPermissions.MODIFY_ANNOTATIONS | aw.saving.PdfPermissions.DOCUMENT_ASSEMBLY

        # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
        # to modify how that method converts the document to .PDF.
        save_options = aw.saving.PdfSaveOptions()

        # Enable encryption via the "encryption_details" property.
        save_options.encryption_details = encryption_details

        # When we open this document, we will need to provide the password before accessing its contents.
        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.encryption_permissions.pdf", save_options)
        #ExEnd

        #with self.assertRaises(Exception):
        #    aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.encryption_permissions.pdf")

        #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.encryption_permissions.pdf", "password")
        #text_absorber = aspose.pdf.text.TextFragmentAbsorber()

        #pdf_document.pages[1].accept(text_absorber)

        #self.assertEqual("Hello world!", text_absorber.text)

    def test_set_numeral_format(self):

        for numeral_format in (aw.saving.NumeralFormat.ARABIC_INDIC,
                               aw.saving.NumeralFormat.CONTEXT,
                               aw.saving.NumeralFormat.EASTERN_ARABIC_INDIC,
                               aw.saving.NumeralFormat.EUROPEAN,
                               aw.saving.NumeralFormat.SYSTEM):
            with self.subTest(numeral_forma=numeral_format):
                #ExStart
                #ExFor:FixedPageSaveOptions.numeral_format
                #ExFor:NumeralFormat
                #ExSummary:Shows how to set the numeral format used when saving to PDF.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.font.locale_id = 4096 # CultureInfo("ar-AR").lcid
                builder.writeln("1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 50, 100")

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

                doc.save(ARTIFACTS_DIR + "PdfSaveOptions.set_numeral_format.pdf", options)
                #ExEnd

                #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.set_numeral_format.pdf")
                #text_absorber = aspose.pdf.text.TextFragmentAbsorber()

                #pdf_document.pages[1].accept(text_absorber)

                #if numeral_format == aw.saving.NumeralFormat.EUROPEAN:
                #    self.assertEqual("1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 50, 100", text_absorber.text)

                #elif numeral_format == aw.saving.NumeralFormat.ARABIC_INDIC:
                #    self.assertEqual(", ٢, ٣, ٤, ٥, ٦, ٧, ٨, ٩, ١٠, ٥٠, ١١٠٠", text_absorber.text)

                #elif numeral_format == aw.saving.NumeralFormat.EASTERN_ARABIC_INDIC:
                #    self.assertEqual("۱۰۰ ,۵۰ ,۱۰ ,۹ ,۸ ,۷ ,۶ ,۵ ,۴ ,۳ ,۲ ,۱", text_absorber.text)

    def test_export_page_set(self):

        #ExStart
        #ExFor:FixedPageSaveOptions.page_set
        #ExSummary:Shows how to export Odd pages from the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        for i in range(5):
            builder.writeln("Page " + str(i + 1) + "(" + ("odd" if i % 2 == 0 else "even") + ")")
            if i < 4:
                builder.insert_break(aw.BreakType.PAGE_BREAK)

        # Create a "PdfSaveOptions" object that we can pass to the document's "save" method
        # to modify how that method converts the document to .PDF.
        options = aw.saving.PdfSaveOptions()

        # Below are three "page_set" properties that we can use to filter out a set of pages from
        # our document to save in an output PDF document based on the parity of their page numbers.
        # 1 -  Save only the even-numbered pages:
        options.page_set = aw.saving.PageSet.even

        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.export_page_set.even.pdf", options)

        # 2 -  Save only the odd-numbered pages:
        options.page_set = aw.saving.PageSet.odd

        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.export_page_set.odd.pdf", options)

        # 3 -  Save every page:
        options.page_set = aw.saving.PageSet.all

        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.export_page_set.all.pdf", options)
        #ExEnd

        #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.export_page_set.even.pdf")
        #text_absorber = aspose.pdf.text.TextAbsorber()
        #pdf_document.pages.accept(text_absorber)

        #self.assertEqual("Page 2 (even)\r\n" +
        #                 "Page 4 (even)", text_absorber.text)

        #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.export_page_set.odd.pdf")
        #text_absorber = aspose.pdf.text.TextAbsorber()
        #pdf_document.pages.accept(text_absorber)

        #self.assertEqual("Page 1 (odd)\r\n" +
        #                 "Page 3 (odd)\r\n" +
        #                 "Page 5 (odd)", text_absorber.text)

        #pdf_document = aspose.pdf.Document(ARTIFACTS_DIR + "PdfSaveOptions.export_page_set.all.pdf")
        #text_absorber = aspose.pdf.text.TextAbsorber()
        #pdf_document.pages.accept(text_absorber)

        #self.assertEqual("Page 1 (odd)\r\n" +
        #                 "Page 2 (even)\r\n" +
        #                 "Page 3 (odd)\r\n" +
        #                 "Page 4 (even)\r\n" +
        #                 "Page 5 (odd)", text_absorber.text)

    def test_export_language_to_span_tag(self):

        #ExStart
        #ExFor:PdfSaveOptions.export_language_to_span_tag
        #ExSummary:Shows how to create a "Span" tag in the document structure to export the text language.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello world!")
        builder.writeln("Hola mundo!")

        save_options = aw.saving.PdfSaveOptions()

        # Note, when "export_document_structure" is "False", "export_language_to_span_tag" is ignored.
        save_options.export_document_structure = True
        save_options.export_language_to_span_tag = True

        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.export_language_to_span_tag.pdf", save_options)
        #ExEnd

    def test_pdf_embed_attachments(self):
        #ExStart
        #ExFor:PdfSaveOptions.embed_attachments
        #ExSummary:Shows how to add embed attachments to the PDF document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc);

        builder.insert_ole_object(MY_DIR + "Spreadsheet.xlsx", "Excel.Sheet", False, True, None)

        options = aw.saving.PdfSaveOptions()
        options.embed_attachments = True

        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.PdfEmbedAttachments.pdf", options)
        #ExEnd

    def test_cache_background_graphics(self):
        #ExStart
        #ExFor:PdfSaveOptions.cache_background_graphics
        #ExSummary:Shows how to cache graphics placed in document's background.
        doc = aw.Document(MY_DIR + "Background images.docx")

        save_options = aw.saving.PdfSaveOptions()
        save_options.cache_background_graphics = True

        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.CacheBackgroundGraphics.pdf", save_options)

        aspose_to_pdf_size = os.stat(ARTIFACTS_DIR + "PdfSaveOptions.CacheBackgroundGraphics.pdf").st_size
        word_to_pdf_size = os.stat(MY_DIR + "Background images (word to pdf).pdf").st_size

        self.assertLess(aspose_to_pdf_size, word_to_pdf_size)
        #ExEnd