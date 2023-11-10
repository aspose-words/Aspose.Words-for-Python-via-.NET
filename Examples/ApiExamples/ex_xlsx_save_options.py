# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

from aspose.words import Document
from aspose.words.saving import XlsxSaveOptions, CompressionLevel, XlsxSectionMode

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExXlsxSaveOptions(ApiExampleBase):

    def test_compress_xlsx(self):
        #ExStart
        #ExFor:XlsxSaveOptions.compression_level
        #ExSummary:Shows how to compress XLSX document.

        doc = Document(MY_DIR + "Shape with linked chart.docx")

        xlsx_save_options = XlsxSaveOptions()
        xlsx_save_options.compression_level = CompressionLevel.MAXIMUM

        doc.save(ARTIFACTS_DIR + "XlsxSaveOptions.CompressXlsx.xlsx", xlsx_save_options)
        #ExEnd

    def test_selection_mode(self):
        #ExStart
        #ExFor: XlsxSaveOptions.section_mode
        #ExSummary:Shows how to save document as a separate worksheets.
        doc = Document(MY_DIR + "Big document.docx")

        # Each section of a document will be created as a separate worksheet.
        # Use 'SingleWorksheet' to display all document on one worksheet.

        xlsx_save_options = XlsxSaveOptions()
        xlsx_save_options.section_mode = XlsxSectionMode.MULTIPLE_WORKSHEETS

        doc.save(ARTIFACTS_DIR + "XlsxSaveOptions.SelectionMode.xlsx", xlsx_save_options)
        #ExEnd
