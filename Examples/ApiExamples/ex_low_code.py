# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from aspose.words import SaveFormat, DocumentBuilder
from aspose.words.lowcode import Merger, MergeFormatMode
from aspose.words.saving import OoxmlSaveOptions
import unittest
from aspose.pydrawing import Color
import os
import pathlib
import io
from aspose.pydrawing import Color
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
        #ExFor:LowCode.merge_format_mode
        #ExFor:LowCode.merger
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

    def test_convert(self):
        #ExStart:Convert
        #ExFor:Converter.convert(str,str)
        #ExFor:Converter.convert(str,str,SaveFormat)
        #ExFor:Converter.convert(str,str,SaveOptions)
        #ExSummary:Shows how to convert documents with a single line of code.
        aw.lowcode.Converter.convert(input_file=MY_DIR + 'Document.docx', output_file=ARTIFACTS_DIR + 'LowCode.Convert.pdf')
        aw.lowcode.Converter.convert(input_file=MY_DIR + 'Document.docx', output_file=ARTIFACTS_DIR + 'LowCode.Convert.rtf', save_format=aw.SaveFormat.RTF)
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.password = 'Aspose.Words'
        aw.lowcode.Converter.convert(input_file=MY_DIR + 'Document.doc', output_file=ARTIFACTS_DIR + 'LowCode.Convert.docx', save_options=save_options)
        #ExEnd:Convert

    def test_convert_to_images(self):
        #ExStart:ConvertToImages
        #ExFor:Converter.convert_to_images(str,str)
        #ExFor:Converter.convert_to_images(str,str,SaveFormat)
        #ExFor:Converter.convert_to_images(str,str,ImageSaveOptions)
        #ExSummary:Shows how to convert document to images.
        aw.lowcode.Converter.convert_to_images(input_file=MY_DIR + 'Big document.docx', output_file=ARTIFACTS_DIR + 'LowCode.ConvertToImages.png')
        aw.lowcode.Converter.convert_to_images(input_file=MY_DIR + 'Big document.docx', output_file=ARTIFACTS_DIR + 'LowCode.ConvertToImages.jpeg', save_format=aw.SaveFormat.JPEG)
        image_save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
        image_save_options.page_set = aw.saving.PageSet(page=1)
        aw.lowcode.Converter.convert_to_images(input_file=MY_DIR + 'Big document.docx', output_file=ARTIFACTS_DIR + 'LowCode.ConvertToImages.png', save_options=image_save_options)
        #ExEnd:ConvertToImages

    def test_convert_to_images_stream(self):
        #ExStart:ConvertToImagesStream
        #ExFor:Converter.convert_to_images(str,SaveFormat)
        #ExFor:Converter.convert_to_images(str,ImageSaveOptions)
        #ExFor:Converter.convert_to_images(Document,SaveFormat)
        #ExFor:Converter.convert_to_images(Document,ImageSaveOptions)
        #ExSummary:Shows how to convert document to images stream.
        streams = aw.lowcode.Converter.convert_to_images(input_file=MY_DIR + 'Big document.docx', save_format=aw.SaveFormat.PNG)
        image_save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
        image_save_options.page_set = aw.saving.PageSet(page=1)
        streams = aw.lowcode.Converter.convert_to_images(input_file=MY_DIR + 'Big document.docx', save_options=image_save_options)
        streams = aw.lowcode.Converter.convert_to_images(doc=aw.Document(file_name=MY_DIR + 'Big document.docx'), save_format=aw.SaveFormat.PNG)
        streams = aw.lowcode.Converter.convert_to_images(doc=aw.Document(file_name=MY_DIR + 'Big document.docx'), save_options=image_save_options)
        #ExEnd:ConvertToImagesStream

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

    def test_convert_stream(self):
        #ExStart:ConvertStream
        #ExFor:Converter.convert(io.BytesIO, io.BytesIO, SaveFormat)
        #ExFor:Converter.convert(io.BytesIO, io.BytesIO, SaveOptions)
        #ExSummary: Shows how to convert documents with a single line of code (Stream).
        first_file_in = open(MY_DIR + 'Big document.docx', mode='rb')
        first_stream_in = io.BytesIO(first_file_in.read())
        out = io.BytesIO()
        aw.lowcode.Converter.convert(first_stream_in, out, aw.SaveFormat.DOCX)
        out.flush()
        pathlib.Path(ARTIFACTS_DIR + 'LowCode.ConvertStream.SaveFormat.docx').write_bytes(out.getvalue())
        out.seek(0)
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.password = 'Aspose.Words'
        aw.lowcode.Converter.convert(first_stream_in, out, save_options)
        out.flush()
        pathlib.Path(ARTIFACTS_DIR + 'LowCode.ConvertStream.SaveOptions.docx').write_bytes(out.getvalue())
        out.close()
        first_stream_in.close()
        first_file_in.close()
        #ExEnd:ConvertStream

    def test_convert_to_images_from_stream(self):
        #ExStart:ConvertToImagesFromStream
        #ExFor:Converter.convert_to_images(io.BytesIO, SaveFormat)
        #ExFor:Converter.convert_to_images(io.BytesIO, ImageSaveOptions)
        #ExSummary:Shows how to convert document to images from stream.
        first_file_in = open(MY_DIR + 'Big document.docx', mode='rb')
        first_stream_in = io.BytesIO(first_file_in.read())
        streams = aw.lowcode.Converter.convert_to_images(first_stream_in, aw.SaveFormat.JPEG)
        save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
        save_options.page_set = aw.saving.PageSet(1)
        streams = aw.lowcode.Converter.convert_to_images(first_stream_in, save_options)
        #ExEnd:ConvertToImagesFromStream