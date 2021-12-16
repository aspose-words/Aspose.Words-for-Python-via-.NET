# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw


class ExMossRtf2Docx:

    @staticmethod
    def convert_rtf_to_docx(in_file_name: str, out_file_name: str):

        # Load an RTF file into Aspose.Words.
        doc = aw.Document(in_file_name)

        # Save the document in the OOXML format.
        doc.save(out_file_name, aw.SaveFormat.DOCX)
