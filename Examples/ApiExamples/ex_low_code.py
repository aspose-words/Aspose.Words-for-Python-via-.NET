# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import os
import pathlib

from aspose.words import SaveFormat, DocumentBuilder
from aspose.words.lowcode import Merger, MergeFormatMode
from aspose.words.saving import OoxmlSaveOptions
from aspose.pydrawing import Color
import io

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExLowCode(ApiExampleBase):

    def test_merge_document(self):
        #ExStart
        #ExFor:Merger.merge(output_file: str, input_files: list[str])
        #ExFor:Merger.merge(input_files: list[str], merge_format_mode: aspose.words.lowcode.MergeFormatMode)
        #ExFor:Merger.merge(output_file: str, input_files: list[str], save_options: aspose.words.saving.SaveOptions, merge_format_mode: aspose.words.lowcode.MergeFormatMode)
        #ExFor:Merger.merge(output_file: str, input_files: list[str], save_format: aspose.words.SaveFormat, merge_format_mode: aspose.words.lowcode.MergeFormatMode)
        #ExSummary:Shows how to merge documents into a single output document.

        # There is a several ways to merge documents:
        Merger.merge(ARTIFACTS_DIR + "LowCode.MergeDocument.SimpleMerge.docx",
                     [MY_DIR + "Big document.docx", MY_DIR + "Tables.docx"])

        save_options = OoxmlSaveOptions()
        save_options.password = "Aspose.Words"
        Merger.merge(ARTIFACTS_DIR + "LowCode.MergeDocument.SaveOptions.docx",
                     [MY_DIR + "Big document.docx", MY_DIR + "Tables.docx"],
                     save_options, MergeFormatMode.KEEP_SOURCE_FORMATTING)

        Merger.merge(ARTIFACTS_DIR + "LowCode.MergeDocument.SaveFormat.pdf",
                     [MY_DIR + "Big document.docx", MY_DIR + "Tables.docx"],
                     SaveFormat.PDF, MergeFormatMode.KEEP_SOURCE_LAYOUT)

        doc = Merger.merge([MY_DIR + "Big document.docx", MY_DIR + "Tables.docx"], MergeFormatMode.MERGE_FORMATTING)
        doc.save(ARTIFACTS_DIR + "LowCode.MergeDocument.DocumentInstance.docx")
        #ExEnd

    def test_merge_stream_document(self):
        #ExStart
        #ExFor:Merger.merge_stream(input_streams: list[io.BytesIO], merge_format_mode: aspose.words.lowcode.MergeFormatMode)
        #ExFor:Merger.merge_stream(output_stream: io.BytesIO, input_streams: list[io.BytesIO], save_options: aspose.words.saving.SaveOptions, merge_format_mode: aspose.words.lowcode.MergeFormatMode)
        #ExFor:Merger.merge_stream(output_stream: io.BytesIO, input_streams: list[io.BytesIO], save_format: aspose.words.SaveFormat)
        #ExSummary:Shows how to merge documents from stream into a single output document.

        # There is a several ways to merge documents from stream:

        first_file_in = open(MY_DIR + "Big document.docx", mode="rb")
        first_stream_in = io.BytesIO(first_file_in.read())

        second_file_in = open(MY_DIR + "Tables.docx", mode="rb")
        second_stream_in = io.BytesIO(second_file_in.read())

        out = io.BytesIO()

        save_options = OoxmlSaveOptions()
        save_options.password = "Aspose.Words"
        Merger.merge_stream(out, [first_stream_in, second_stream_in], save_options,
                     MergeFormatMode.KEEP_SOURCE_FORMATTING)
        out.flush()
        pathlib.Path(ARTIFACTS_DIR + "LowCode.MergeStreamDocument.SaveOptions.docx").write_bytes(out.getvalue())

        out.seek(0)

        Merger.merge_stream(out, [first_stream_in, second_stream_in], SaveFormat.DOCX)
        out.flush()

        pathlib.Path(ARTIFACTS_DIR + "LowCode.MergeStreamDocument.SaveFormat.docx").write_bytes(out.getvalue())
        out.close()

        doc = Merger.merge_stream([first_stream_in, second_stream_in], MergeFormatMode.MERGE_FORMATTING)
        doc.save(ARTIFACTS_DIR + "LowCode.MergeStreamDocument.DocumentInstance.docx")
        first_file_in.close()
        second_file_in.close()
        #ExEnd

    def test_merge_document_instances(self):
        #ExStart
        #ExFor: Merger.merge_docs
        #ExSummary:Shows how to merge input documents to a single document instance.

        first_doc = DocumentBuilder()
        first_doc.font.size = 16
        first_doc.font.color = Color.blue
        first_doc.write("Hello first word!")

        second_doc = DocumentBuilder()
        second_doc.write("Hello second word!")

        merged_doc = Merger.merge_docs([first_doc.document, second_doc.document], MergeFormatMode.KEEP_SOURCE_LAYOUT)

        self.assertEqual("Hello first word!\fHello second word!\f", merged_doc.get_text())
        #ExEnd
