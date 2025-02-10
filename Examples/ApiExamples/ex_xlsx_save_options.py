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

class ExXlsxSaveOptions(ApiExampleBase):

    def test_compress_xlsx(self):
        #ExStart
        #ExFor:XlsxSaveOptions
        #ExFor:XlsxSaveOptions.compression_level
        #ExFor:XlsxSaveOptions.save_format
        #ExSummary:Shows how to compress XLSX document.
        doc = aw.Document(file_name=MY_DIR + 'Shape with linked chart.docx')
        xlsx_save_options = aw.saving.XlsxSaveOptions()
        xlsx_save_options.compression_level = aw.saving.CompressionLevel.MAXIMUM
        xlsx_save_options.save_format = aw.SaveFormat.XLSX
        doc.save(file_name=ARTIFACTS_DIR + 'XlsxSaveOptions.CompressXlsx.xlsx', save_options=xlsx_save_options)
        #ExEnd

    def test_selection_mode(self):
        #ExStart:SelectionMode
        #ExFor:XlsxSaveOptions.section_mode
        #ExFor:XlsxSectionMode
        #ExSummary:Shows how to save document as a separate worksheets.
        doc = aw.Document(file_name=MY_DIR + 'Big document.docx')
        # Each section of a document will be created as a separate worksheet.
        # Use 'SingleWorksheet' to display all document on one worksheet.
        xlsx_save_options = aw.saving.XlsxSaveOptions()
        xlsx_save_options.section_mode = aw.saving.XlsxSectionMode.MULTIPLE_WORKSHEETS
        doc.save(file_name=ARTIFACTS_DIR + 'XlsxSaveOptions.SelectionMode.xlsx', save_options=xlsx_save_options)
        #ExEnd:SelectionMode

    def test_date_time_parsing_mode(self):
        #ExStart:DateTimeParsingMode
        #ExFor:XlsxSaveOptions.date_time_parsing_mode
        #ExFor:XlsxDateTimeParsingMode
        #ExSummary:Shows how to specify autodetection of the date time format.
        doc = aw.Document(file_name=MY_DIR + 'Xlsx DateTime.docx')
        save_options = aw.saving.XlsxSaveOptions()
        # Specify using datetime format autodetection.
        save_options.date_time_parsing_mode = aw.saving.XlsxDateTimeParsingMode.AUTO
        doc.save(file_name=ARTIFACTS_DIR + 'XlsxSaveOptions.DateTimeParsingMode.xlsx', save_options=save_options)
        #ExEnd:DateTimeParsingMode