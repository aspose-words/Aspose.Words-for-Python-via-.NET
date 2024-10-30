# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import os
import sys
import aspose.words as aw
import aspose.words.digitalsignatures
import aspose.words.saving
import aspose.words.settings
import datetime
import system_helper
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExXpsSaveOptions(ApiExampleBase):

    def test_outline_levels(self):
        #ExStart
        #ExFor:XpsSaveOptions
        #ExFor:XpsSaveOptions.__init__
        #ExFor:XpsSaveOptions.outline_options
        #ExFor:XpsSaveOptions.save_format
        #ExSummary:Shows how to limit the headings' level that will appear in the outline of a saved XPS document.
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
        # Create an "XpsSaveOptions" object that we can pass to the document's "Save" method
        # to modify how that method converts the document to .XPS.
        save_options = aw.saving.XpsSaveOptions()
        self.assertEqual(aw.SaveFormat.XPS, save_options.save_format)
        # The output XPS document will contain an outline, a table of contents that lists headings in the document body.
        # Clicking on an entry in this outline will take us to the location of its respective heading.
        # Set the "HeadingsOutlineLevels" property to "2" to exclude all headings whose levels are above 2 from the outline.
        # The last two headings we have inserted above will not appear.
        save_options.outline_options.headings_outline_levels = 2
        doc.save(file_name=ARTIFACTS_DIR + 'XpsSaveOptions.OutlineLevels.xps', save_options=save_options)
        #ExEnd

    def test_book_fold(self):
        for render_text_as_book_fold in [False, True]:
            #ExStart
            #ExFor:XpsSaveOptions.__init__(SaveFormat)
            #ExFor:XpsSaveOptions.use_book_fold_printing_settings
            #ExSummary:Shows how to save a document to the XPS format in the form of a book fold.
            doc = aw.Document(file_name=MY_DIR + 'Paragraphs.docx')
            # Create an "XpsSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to .XPS.
            xps_options = aw.saving.XpsSaveOptions(aw.SaveFormat.XPS)
            # Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
            # in the output XPS in a way that helps us use it to make a booklet.
            # Set the "UseBookFoldPrintingSettings" property to "false" to render the XPS normally.
            xps_options.use_book_fold_printing_settings = render_text_as_book_fold
            # If we are rendering the document as a booklet, we must set the "MultiplePages"
            # properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
            if render_text_as_book_fold:
                for s in doc.sections:
                    s = s.as_section()
                    s.page_setup.multiple_pages = aw.settings.MultiplePagesType.BOOK_FOLD_PRINTING
            # Once we print this document, we can turn it into a booklet by stacking the pages
            # to come out of the printer and folding down the middle.
            doc.save(file_name=ARTIFACTS_DIR + 'XpsSaveOptions.BookFold.xps', save_options=xps_options)
            #ExEnd

    def test_export_exact_pages(self):
        #ExStart
        #ExFor:FixedPageSaveOptions.page_set
        #ExFor:PageSet.__init__(List[int])
        #ExSummary:Shows how to extract pages based on exact page indices.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Add five pages to the document.
        i = 1
        while i < 6:
            builder.write('Page ' + str(i))
            builder.insert_break(aw.BreakType.PAGE_BREAK)
            i += 1
        # Create an "XpsSaveOptions" object, which we can pass to the document's "Save" method
        # to modify how that method converts the document to .XPS.
        xps_options = aw.saving.XpsSaveOptions()
        # Use the "PageSet" property to select a set of the document's pages to save to output XPS.
        # In this case, we will choose, via a zero-based index, only three pages: page 1, page 2, and page 4.
        xps_options.page_set = aw.saving.PageSet(pages=[0, 1, 3])
        doc.save(file_name=ARTIFACTS_DIR + 'XpsSaveOptions.ExportExactPages.xps', save_options=xps_options)
        #ExEnd

    def test_xps_digital_signature(self):
        #ExStart:XpsDigitalSignature
        #ExFor:XpsSaveOptions.digital_signature_details
        #ExSummary:Shows how to sign XPS document.
        doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        certificate_holder = aw.digitalsignatures.CertificateHolder.create(file_name=MY_DIR + 'morzal.pfx', password='aw')
        options = aw.digitalsignatures.SignOptions()
        options.sign_time = datetime.datetime.now()
        options.comments = 'Some comments'
        digital_signature_details = aw.saving.DigitalSignatureDetails(certificate_holder, options)
        save_options = aw.saving.XpsSaveOptions()
        save_options.digital_signature_details = digital_signature_details
        self.assertEqual(certificate_holder, digital_signature_details.certificate_holder)
        self.assertEqual('Some comments', digital_signature_details.sign_options.comments)
        doc.save(file_name=ARTIFACTS_DIR + 'XpsSaveOptions.XpsDigitalSignature.docx', save_options=save_options)
        #ExEnd:XpsDigitalSignature

    @unittest.skipUnless(sys.platform.startswith('win'), 'different calculation on Linux')
    def test_optimize_output(self):
        for optimize_output in (False, True):
            with self.subTest(optimize_output=optimize_output):
                #ExStart
                #ExFor:FixedPageSaveOptions.optimize_output
                #ExSummary:Shows how to optimize document objects while saving to xps.
                doc = aw.Document(MY_DIR + 'Unoptimized document.docx')
                # Create an "XpsSaveOptions" object to pass to the document's "save" method
                # to modify how that method converts the document to .XPS.
                save_options = aw.saving.XpsSaveOptions()
                # Set the "optimize_output" property to "True" to take measures such as removing nested or empty canvases
                # and concatenating adjacent runs with identical formatting to optimize the output document's content.
                # This may affect the appearance of the document.
                # Set the "optimize_output" property to "False" to save the document normally.
                save_options.optimize_output = optimize_output
                doc.save(ARTIFACTS_DIR + 'XpsSaveOptions.optimize_output.xps', save_options)
                #ExEnd
                out_file_size = os.path.getsize(ARTIFACTS_DIR + 'XpsSaveOptions.optimize_output.xps')
                if optimize_output:
                    self.assertLess(out_file_size, 50000)
                else:
                    self.assertGreater(out_file_size, 60000)