# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
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
        #ExFor:XlsxSaveOptions.compression_level
        #ExSummary:Shows how to compress XLSX document.
        doc = aw.Document(file_name=MY_DIR + 'Shape with linked chart.docx')
        xlsx_save_options = aw.saving.XlsxSaveOptions()
        xlsx_save_options.compression_level = aw.saving.CompressionLevel.MAXIMUM
        doc.save(file_name=ARTIFACTS_DIR + 'XlsxSaveOptions.CompressXlsx.xlsx', save_options=xlsx_save_options)
        #ExEnd

    def test_selection_mode(self):
        #ExStart:SelectionMode
        #ExFor:XlsxSaveOptions.section_mode
        #ExSummary:Shows how to save document as a separate worksheets.
        doc = aw.Document(file_name=MY_DIR + 'Big document.docx')
        # Each section of a document will be created as a separate worksheet.
        # Use 'SingleWorksheet' to display all document on one worksheet.
        xlsx_save_options = aw.saving.XlsxSaveOptions()
        xlsx_save_options.section_mode = aw.saving.XlsxSectionMode.MULTIPLE_WORKSHEETS
        doc.save(file_name=ARTIFACTS_DIR + 'XlsxSaveOptions.SelectionMode.xlsx', save_options=xlsx_save_options)
        #ExEnd:SelectionMode