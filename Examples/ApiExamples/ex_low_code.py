# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from aspose.words.saving import OoxmlSaveOptions
from aspose.words.lowcode import Merger, MergeFormatMode
from aspose.words import SaveFormat, DocumentBuilder
from aspose.pydrawing import Color
import io
import pathlib
import os
from aspose.pydrawing import Color
import unittest
import aspose.words as aw
import aspose.words.lowcode
import aspose.words.saving
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExLowCode(ApiExampleBase):

    def test_merge_documents(self):
        #ExStart
        #ExFor:Merger.merge(str,List[str])
        #ExFor:Merger.merge(List[str],MergeFormatMode)
        #ExFor:Merger.merge(str,List[str],SaveOptions,MergeFormatMode)
        #ExFor:Merger.merge(str,List[str],SaveFormat,MergeFormatMode)
        #ExSummary:Shows how to merge documents into a single output document.
        #There is a several ways to merge documents:
        aw.lowcode.Merger.merge(output_file=ARTIFACTS_DIR + 'LowCode.MergeDocument.SimpleMerge.docx', input_files=[MY_DIR + 'Big document.docx', MY_DIR + 'Tables.docx'])
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.password = 'Aspose.Words'
        aw.lowcode.Merger.merge(output_file=ARTIFACTS_DIR + 'LowCode.MergeDocument.SaveOptions.docx', input_files=[MY_DIR + 'Big document.docx', MY_DIR + 'Tables.docx'], save_options=save_options, merge_format_mode=aw.lowcode.MergeFormatMode.KEEP_SOURCE_FORMATTING)
        aw.lowcode.Merger.merge(output_file=ARTIFACTS_DIR + 'LowCode.MergeDocument.SaveFormat.pdf', input_files=[MY_DIR + 'Big document.docx', MY_DIR + 'Tables.docx'], save_format=aw.SaveFormat.PDF, merge_format_mode=aw.lowcode.MergeFormatMode.KEEP_SOURCE_LAYOUT)
        doc = aw.lowcode.Merger.merge(input_files=[MY_DIR + 'Big document.docx', MY_DIR + 'Tables.docx'], merge_format_mode=aw.lowcode.MergeFormatMode.MERGE_FORMATTING)
        doc.save(file_name=ARTIFACTS_DIR + 'LowCode.MergeDocument.DocumentInstance.docx')
        #ExEnd

    def test_merge_stream_document(self):
        #ExStart
        #ExFor:Merger.merge_stream(input_streams: list[io.BytesIO], merge_format_mode: aspose.words.lowcode.MergeFormatMode)
        #ExFor:Merger.merge_stream(output_stream: io.BytesIO, input_streams: list[io.BytesIO], save_options: aspose.words.saving.SaveOptions, merge_format_mode: aspose.words.lowcode.MergeFormatMode)
        #ExFor:Merger.merge_stream(output_stream: io.BytesIO, input_streams: list[io.BytesIO], save_format: aspose.words.SaveFormat)
        #ExSummary:Shows how to merge documents from stream into a single output document.
        # There is a several ways to merge documents from stream:
        first_file_in = open(MY_DIR + 'Big document.docx', mode='rb')
        first_stream_in = io.BytesIO(first_file_in.read())
        second_file_in = open(MY_DIR + 'Tables.docx', mode='rb')
        second_stream_in = io.BytesIO(second_file_in.read())
        out = io.BytesIO()
        save_options = OoxmlSaveOptions()
        save_options.password = 'Aspose.Words'
        Merger.merge_stream(out, [first_stream_in, second_stream_in], save_options, MergeFormatMode.KEEP_SOURCE_FORMATTING)
        out.flush()
        pathlib.Path(ARTIFACTS_DIR + 'LowCode.MergeStreamDocument.SaveOptions.docx').write_bytes(out.getvalue())
        out.seek(0)
        Merger.merge_stream(out, [first_stream_in, second_stream_in], SaveFormat.DOCX)
        out.flush()
        pathlib.Path(ARTIFACTS_DIR + 'LowCode.MergeStreamDocument.SaveFormat.docx').write_bytes(out.getvalue())
        out.close()
        doc = Merger.merge_stream([first_stream_in, second_stream_in], MergeFormatMode.MERGE_FORMATTING)
        doc.save(ARTIFACTS_DIR + 'LowCode.MergeStreamDocument.DocumentInstance.docx')
        first_file_in.close()
        second_file_in.close()
        #ExEnd

    def test_merge_document_instances(self):
        #ExStart:MergeDocumentInstances
        #ExFor:Merger.merge(List[Document],MergeFormatMode)
        #ExSummary:Shows how to merge input documents to a single document instance.
        first_doc = DocumentBuilder()
        first_doc.font.size = 16
        first_doc.font.color = Color.blue
        first_doc.write('Hello first word!')
        second_doc = DocumentBuilder()
        second_doc.write('Hello second word!')
        merged_doc = Merger.merge_docs([first_doc.document, second_doc.document], MergeFormatMode.KEEP_SOURCE_LAYOUT)
        self.assertEqual('Hello first word!\x0cHello second word!\x0c', merged_doc.get_text())
        #ExEnd:MergeDocumentInstances