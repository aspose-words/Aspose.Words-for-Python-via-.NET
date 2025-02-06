# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import aspose.words.saving
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExPclSaveOptions(ApiExampleBase):

    def test_rasterize_elements(self):
        #ExStart
        #ExFor:PclSaveOptions
        #ExFor:PclSaveOptions.save_format
        #ExFor:PclSaveOptions.rasterize_transformed_elements
        #ExSummary:Shows how to rasterize complex elements while saving a document to PCL.
        doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
        save_options = aw.saving.PclSaveOptions()
        save_options.save_format = aw.SaveFormat.PCL
        save_options.rasterize_transformed_elements = True
        doc.save(file_name=ARTIFACTS_DIR + 'PclSaveOptions.RasterizeElements.pcl', save_options=save_options)
        #ExEnd

    def test_fallback_font_name(self):
        #ExStart
        #ExFor:PclSaveOptions.fallback_font_name
        #ExSummary:Shows how to declare a font that a printer will apply to printed text as a substitute should its original font be unavailable.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.name = 'Non-existent font'
        builder.write('Hello world!')
        save_options = aw.saving.PclSaveOptions()
        save_options.fallback_font_name = 'Times New Roman'
        # This document will instruct the printer to apply "Times New Roman" to the text with the missing font.
        # Should "Times New Roman" also be unavailable, the printer will default to the "Arial" font.
        doc.save(file_name=ARTIFACTS_DIR + 'PclSaveOptions.SetPrinterFont.pcl', save_options=save_options)
        #ExEnd

    def test_add_printer_font(self):
        #ExStart
        #ExFor:PclSaveOptions.add_printer_font(str,str)
        #ExSummary:Shows how to get a printer to substitute all instances of a specific font with a different font.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.name = 'Courier'
        builder.write('Hello world!')
        save_options = aw.saving.PclSaveOptions()
        save_options.add_printer_font('Courier New', 'Courier')
        # When printing this document, the printer will use the "Courier New" font
        # to access places where our document used the "Courier" font.
        doc.save(file_name=ARTIFACTS_DIR + 'PclSaveOptions.AddPrinterFont.pcl', save_options=save_options)
        #ExEnd

    def test_get_preserved_paper_tray_information(self):
        doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
        # Paper tray information is now preserved when saving document to PCL format.
        # Following information is transferred from document's model to PCL file.
        for section in doc.sections:
            section = section.as_section()
            section.page_setup.first_page_tray = 15
            section.page_setup.other_pages_tray = 12
        doc.save(file_name=ARTIFACTS_DIR + 'PclSaveOptions.GetPreservedPaperTrayInformation.pcl')