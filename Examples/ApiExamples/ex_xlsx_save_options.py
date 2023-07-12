# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

from aspose.words import Document
from aspose.words.saving import XlsxSaveOptions, CompressionLevel

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



