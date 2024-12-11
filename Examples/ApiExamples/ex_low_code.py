# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from aspose.pydrawing import Color
from aspose.pydrawing import Color
from aspose.words.saving import OoxmlSaveOptions
from aspose.words.lowcode import Merger, MergeFormatMode
from aspose.words import SaveFormat, DocumentBuilder
import io
import pathlib
import os
import unittest
import aspose.pydrawing
import aspose.words as aw
import aspose.words.comparing
import aspose.words.loading
import aspose.words.lowcode
import aspose.words.lowcode.mailmerging
import aspose.words.lowcode.splitting
import aspose.words.replacing
import aspose.words.saving
import datetime
import system_helper
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, IMAGE_DIR, MY_DIR

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

    def test_merge_stream_document(self):
        #ExStart
        #ExFor:Merger.merge(List[BytesIO],MergeFormatMode)
        #ExFor:Merger.merge(BytesIO,List[BytesIO],SaveOptions,MergeFormatMode)
        #ExFor:Merger.merge(BytesIO,List[BytesIO],SaveFormat)
        #ExSummary:Shows how to merge documents from stream into a single output document.
        #There is a several ways to merge documents from stream:
        with system_helper.io.FileStream(MY_DIR + 'Big document.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as first_stream_in:
            with system_helper.io.FileStream(MY_DIR + 'Tables.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as second_stream_in:
                save_options = aw.saving.OoxmlSaveOptions()
                save_options.password = 'Aspose.Words'
                with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.MergeStreamDocument.SaveOptions.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                    aw.lowcode.Merger.merge_stream(output_stream=stream_out, input_streams=[first_stream_in, second_stream_in], save_options=save_options, merge_format_mode=aw.lowcode.MergeFormatMode.KEEP_SOURCE_FORMATTING)
                with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.MergeStreamDocument.SaveFormat.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                    aw.lowcode.Merger.merge_stream(output_stream=stream_out, input_streams=[first_stream_in, second_stream_in], save_format=aw.SaveFormat.DOCX)
                doc = aw.lowcode.Merger.merge_stream(input_streams=[first_stream_in, second_stream_in], merge_format_mode=aw.lowcode.MergeFormatMode.MERGE_FORMATTING)
                doc.save(file_name=ARTIFACTS_DIR + 'LowCode.MergeStreamDocument.DocumentInstance.docx')
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

    def test_convert_stream(self):
        #ExStart:ConvertStream
        #ExFor:Converter.convert(BytesIO,BytesIO,SaveFormat)
        #ExFor:Converter.convert(BytesIO,BytesIO,SaveOptions)
        #ExSummary:Shows how to convert documents with a single line of code (Stream).
        with system_helper.io.FileStream(MY_DIR + 'Big document.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.ConvertStream.SaveFormat.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                aw.lowcode.Converter.convert(input_stream=stream_in, output_stream=stream_out, save_format=aw.SaveFormat.DOCX)
            save_options = aw.saving.OoxmlSaveOptions()
            save_options.password = 'Aspose.Words'
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.ConvertStream.SaveOptions.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                aw.lowcode.Converter.convert(input_stream=stream_in, output_stream=stream_out, save_options=save_options)
        #ExEnd:ConvertStream

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

    def test_convert_to_images_from_stream(self):
        #ExStart:ConvertToImagesFromStream
        #ExFor:Converter.convert_to_images(BytesIO,SaveFormat)
        #ExFor:Converter.convert_to_images(BytesIO,ImageSaveOptions)
        #ExSummary:Shows how to convert document to images from stream.
        with system_helper.io.FileStream(MY_DIR + 'Big document.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            streams = aw.lowcode.Converter.convert_to_images(input_stream=stream_in, save_format=aw.SaveFormat.JPEG)
            image_save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
            image_save_options.page_set = aw.saving.PageSet(page=1)
            streams = aw.lowcode.Converter.convert_to_images(input_stream=stream_in, save_options=image_save_options)
        #ExEnd:ConvertToImagesFromStream

    def test_pdf_renderer(self):
        raise NotImplementedError('Unsupported call of method named SaveTo')

    def test_compare_documents(self):
        #ExStart:CompareDocuments
        #ExFor:Comparer.compare(str,str,str,str,datetime)
        #ExFor:Comparer.compare(str,str,str,SaveFormat,str,datetime)
        #ExFor:Comparer.compare(str,str,str,str,datetime,CompareOptions)
        #ExFor:Comparer.compare(str,str,str,SaveFormat,str,datetime,CompareOptions)
        #ExSummary:Shows how to simple compare documents.
        # There is a several ways to compare documents:
        first_doc = MY_DIR + 'Table column bookmarks.docx'
        second_doc = MY_DIR + 'Table column bookmarks.doc'
        aw.lowcode.Comparer.compare(v1=first_doc, v2=second_doc, output_file_name=ARTIFACTS_DIR + 'LowCode.CompareDocuments.1.docx', author='Author', date_time=datetime.datetime(1, 1, 1))
        aw.lowcode.Comparer.compare(v1=first_doc, v2=second_doc, output_file_name=ARTIFACTS_DIR + 'LowCode.CompareDocuments.2.docx', save_format=aw.SaveFormat.DOCX, author='Author', date_time=datetime.datetime(1, 1, 1))
        compare_options = aw.comparing.CompareOptions()
        compare_options.ignore_case_changes = True
        aw.lowcode.Comparer.compare(v1=first_doc, v2=second_doc, output_file_name=ARTIFACTS_DIR + 'LowCode.CompareDocuments.3.docx', author='Author', date_time=datetime.datetime(1, 1, 1), compare_options=compare_options)
        aw.lowcode.Comparer.compare(v1=first_doc, v2=second_doc, output_file_name=ARTIFACTS_DIR + 'LowCode.CompareDocuments.4.docx', save_format=aw.SaveFormat.DOCX, author='Author', date_time=datetime.datetime(1, 1, 1), compare_options=compare_options)
        #ExEnd:CompareDocuments

    def test_compare_stream_documents(self):
        #ExStart:CompareStreamDocuments
        #ExFor:Comparer.compare(BytesIO,BytesIO,BytesIO,SaveFormat,str,datetime)
        #ExFor:Comparer.compare(BytesIO,BytesIO,BytesIO,SaveFormat,str,datetime,CompareOptions)
        #ExSummary:Shows how to compare documents from the stream.
        # There is a several ways to compare documents from the stream:
        with system_helper.io.FileStream(MY_DIR + 'Table column bookmarks.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as first_stream_in:
            with system_helper.io.FileStream(MY_DIR + 'Table column bookmarks.doc', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as second_stream_in:
                with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.CompareStreamDocuments.1.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                    aw.lowcode.Comparer.compare(v1=first_stream_in, v2=second_stream_in, output_stream=stream_out, save_format=aw.SaveFormat.DOCX, author='Author', date_time=datetime.datetime(1, 1, 1))
                with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.CompareStreamDocuments.2.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                    compare_options = aw.comparing.CompareOptions()
                    compare_options.ignore_case_changes = True
                    aw.lowcode.Comparer.compare(v1=first_stream_in, v2=second_stream_in, output_stream=stream_out, save_format=aw.SaveFormat.DOCX, author='Author', date_time=datetime.datetime(1, 1, 1), compare_options=compare_options)
        #ExEnd:CompareStreamDocuments

    def test_mail_merge(self):
        #ExStart:MailMerge
        #ExFor:MailMerger.execute(str,str,List[str],List[object])
        #ExFor:MailMerger.execute(str,str,SaveFormat,List[str],List[object])
        #ExFor:MailMerger.execute(str,str,SaveFormat,MailMergeOptions,List[str],List[object])
        #ExSummary:Shows how to do mail merge operation for a single record.
        # There is a several ways to do mail merge operation:
        doc = MY_DIR + 'Mail merge.doc'
        field_names = ['FirstName', 'Location', 'SpecialCharsInName()']
        field_values = ['James Bond', 'London', 'Classified']
        aw.lowcode.MailMerger.execute(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.MailMerge.1.docx', field_names=field_names, field_values=field_values)
        aw.lowcode.MailMerger.execute(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.MailMerge.2.docx', save_format=aw.SaveFormat.DOCX, field_names=field_names, field_values=field_values)
        mail_merge_options = aw.lowcode.mailmerging.MailMergeOptions()
        mail_merge_options.trim_whitespaces = True
        aw.lowcode.MailMerger.execute(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.MailMerge.3.docx', save_format=aw.SaveFormat.DOCX, mail_merge_options=mail_merge_options, field_names=field_names, field_values=field_values)
        #ExEnd:MailMerge

    def test_mail_merge_stream(self):
        #ExStart:MailMergeStream
        #ExFor:MailMerger.execute(BytesIO,BytesIO,SaveFormat,List[str],List[object])
        #ExFor:MailMerger.execute(BytesIO,BytesIO,SaveFormat,MailMergeOptions,List[str],List[object])
        #ExSummary:Shows how to do mail merge operation for a single record from the stream.
        # There is a several ways to do mail merge operation using documents from the stream:
        field_names = ['FirstName', 'Location', 'SpecialCharsInName()']
        field_values = ['James Bond', 'London', 'Classified']
        with system_helper.io.FileStream(MY_DIR + 'Mail merge.doc', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.MailMergeStream.1.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                aw.lowcode.MailMerger.execute(input_stream=stream_in, output_stream=stream_out, save_format=aw.SaveFormat.DOCX, field_names=field_names, field_values=field_values)
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.MailMergeStream.2.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                mail_merge_options = aw.lowcode.mailmerging.MailMergeOptions()
                mail_merge_options.trim_whitespaces = True
                aw.lowcode.MailMerger.execute(input_stream=stream_in, output_stream=stream_out, save_format=aw.SaveFormat.DOCX, mail_merge_options=mail_merge_options, field_names=field_names, field_values=field_values)
        #ExEnd:MailMergeStream

    def test_replace(self):
        #ExStart:Replace
        #ExFor:Replacer.replace(str,str,str,str)
        #ExFor:Replacer.replace(str,str,SaveFormat,str,str)
        #ExFor:Replacer.replace(str,str,SaveFormat,str,str,FindReplaceOptions)
        #ExSummary:Shows how to replace string in the document.
        # There is a several ways to replace string in the document:
        doc = MY_DIR + 'Footer.docx'
        pattern = '(C)2006 Aspose Pty Ltd.'
        replacement = 'Copyright (C) 2024 by Aspose Pty Ltd.'
        options = aw.replacing.FindReplaceOptions()
        options.find_whole_words_only = False
        aw.lowcode.Replacer.replace(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.Replace.1.docx', pattern=pattern, replacement=replacement)
        aw.lowcode.Replacer.replace(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.Replace.2.docx', save_format=aw.SaveFormat.DOCX, pattern=pattern, replacement=replacement)
        aw.lowcode.Replacer.replace(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.Replace.3.docx', save_format=aw.SaveFormat.DOCX, pattern=pattern, replacement=replacement, options=options)
        #ExEnd:Replace

    def test_replace_stream(self):
        #ExStart:ReplaceStream
        #ExFor:Replacer.replace(BytesIO,BytesIO,SaveFormat,str,str)
        #ExFor:Replacer.replace(BytesIO,BytesIO,SaveFormat,str,str,FindReplaceOptions)
        #ExSummary:Shows how to replace string in the document using documents from the stream.
        # There is a several ways to replace string in the document using documents from the stream:
        pattern = '(C)2006 Aspose Pty Ltd.'
        replacement = 'Copyright (C) 2024 by Aspose Pty Ltd.'
        with system_helper.io.FileStream(MY_DIR + 'Footer.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.ReplaceStream.1.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                aw.lowcode.Replacer.replace(input_stream=stream_in, output_stream=stream_out, save_format=aw.SaveFormat.DOCX, pattern=pattern, replacement=replacement)
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.ReplaceStream.2.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                options = aw.replacing.FindReplaceOptions()
                options.find_whole_words_only = False
                aw.lowcode.Replacer.replace(input_stream=stream_in, output_stream=stream_out, save_format=aw.SaveFormat.DOCX, pattern=pattern, replacement=replacement, options=options)
        #ExEnd:ReplaceStream

    def test_remove_blank_pages(self):
        #ExStart:RemoveBlankPages
        #ExFor:Splitter.remove_blank_pages(str,str)
        #ExFor:Splitter.remove_blank_pages(str,str,SaveFormat)
        #ExSummary:Shows how to remove empty pages from the document.
        # There is a several ways to remove empty pages from the document:
        doc = MY_DIR + 'Blank pages.docx'
        aw.lowcode.Splitter.remove_blank_pages(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.RemoveBlankPages.1.docx')
        aw.lowcode.Splitter.remove_blank_pages(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.RemoveBlankPages.2.docx', save_format=aw.SaveFormat.DOCX)
        #ExEnd:RemoveBlankPages

    def test_remove_blank_pages_stream(self):
        #ExStart:RemoveBlankPagesStream
        #ExFor:Splitter.remove_blank_pages(BytesIO,BytesIO,SaveFormat)
        #ExSummary:Shows how to remove empty pages from the document from the stream.
        with system_helper.io.FileStream(MY_DIR + 'Blank pages.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.RemoveBlankPagesStream.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                aw.lowcode.Splitter.remove_blank_pages(input_stream=stream_in, output_stream=stream_out, save_format=aw.SaveFormat.DOCX)
        #ExEnd:RemoveBlankPagesStream

    def test_extract_pages(self):
        #ExStart:ExtractPages
        #ExFor:Splitter.extract_pages(str,str,int,int)
        #ExFor:Splitter.extract_pages(str,str,SaveFormat,int,int)
        #ExSummary:Shows how to extract pages from the document.
        # There is a several ways to extract pages from the document:
        doc = MY_DIR + 'Big document.docx'
        aw.lowcode.Splitter.extract_pages(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.ExtractPages.1.docx', start_page_index=0, page_count=2)
        aw.lowcode.Splitter.extract_pages(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.ExtractPages.2.docx', save_format=aw.SaveFormat.DOCX, start_page_index=0, page_count=2)
        #ExEnd:ExtractPages

    def test_extract_pages_stream(self):
        #ExStart:ExtractPagesStream
        #ExFor:Splitter.extract_pages(BytesIO,BytesIO,SaveFormat,int,int)
        #ExSummary:Shows how to extract pages from the document from the stream.
        with system_helper.io.FileStream(MY_DIR + 'Big document.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.ExtractPagesStream.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                aw.lowcode.Splitter.extract_pages(input_stream=stream_in, output_stream=stream_out, save_format=aw.SaveFormat.DOCX, start_page_index=0, page_count=2)
        #ExEnd:ExtractPagesStream

    def test_split_document(self):
        #ExStart:SplitDocument
        #ExFor:Splitter.split(str,str,SplitOptions)
        #ExFor:Splitter.split(str,str,SaveFormat,SplitOptions)
        #ExSummary:Shows how to split document by pages.
        doc = MY_DIR + 'Big document.docx'
        options = aw.lowcode.splitting.SplitOptions()
        options.split_criteria = aw.lowcode.splitting.SplitCriteria.PAGE
        aw.lowcode.Splitter.split(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.SplitDocument.1.docx', options=options)
        aw.lowcode.Splitter.split(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.SplitDocument.2.docx', save_format=aw.SaveFormat.DOCX, options=options)
        #ExEnd:SplitDocument

    def test_split_document_stream(self):
        #ExStart:SplitDocumentStream
        #ExFor:Splitter.split(BytesIO,SaveFormat,SplitOptions)
        #ExSummary:Shows how to split document from the stream by pages.
        with system_helper.io.FileStream(MY_DIR + 'Big document.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            options = aw.lowcode.splitting.SplitOptions()
            options.split_criteria = aw.lowcode.splitting.SplitCriteria.PAGE
            stream = aw.lowcode.Splitter.split(input_stream=stream_in, save_format=aw.SaveFormat.DOCX, options=options)
        #ExEnd:SplitDocumentStream

    def test_watermark_text(self):
        #ExStart:WatermarkText
        #ExFor:Watermarker.set_text(str,str,str)
        #ExFor:Watermarker.set_text(str,str,SaveFormat,str)
        #ExFor:Watermarker.set_text(str,str,str,TextWatermarkOptions)
        #ExFor:Watermarker.set_text(str,str,SaveFormat,str,TextWatermarkOptions)
        #ExSummary:Shows how to insert watermark text to the document.
        doc = MY_DIR + 'Big document.docx'
        watermark_text = 'This is a watermark'
        aw.lowcode.Watermarker.set_text(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.WatermarkText.1.docx', watermark_text=watermark_text)
        aw.lowcode.Watermarker.set_text(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.WatermarkText.2.docx', save_format=aw.SaveFormat.DOCX, watermark_text=watermark_text)
        watermark_options = aw.TextWatermarkOptions()
        watermark_options.color = aspose.pydrawing.Color.red
        aw.lowcode.Watermarker.set_text(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.WatermarkText.3.docx', watermark_text=watermark_text, options=watermark_options)
        aw.lowcode.Watermarker.set_text(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.WatermarkText.4.docx', save_format=aw.SaveFormat.DOCX, watermark_text=watermark_text, options=watermark_options)
        #ExEnd:WatermarkText

    def test_watermark_text_stream(self):
        #ExStart:WatermarkTextStream
        #ExFor:Watermarker.set_text(BytesIO,BytesIO,SaveFormat,str)
        #ExFor:Watermarker.set_text(BytesIO,BytesIO,SaveFormat,str,TextWatermarkOptions)
        #ExSummary:Shows how to insert watermark text to the document from the stream.
        watermark_text = 'This is a watermark'
        with system_helper.io.FileStream(MY_DIR + 'Document.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.WatermarkTextStream.1.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                aw.lowcode.Watermarker.set_text(input_stream=stream_in, output_stream=stream_out, save_format=aw.SaveFormat.DOCX, watermark_text=watermark_text)
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.WatermarkTextStream.2.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                options = aw.TextWatermarkOptions()
                options.color = aspose.pydrawing.Color.red
                aw.lowcode.Watermarker.set_text(input_stream=stream_in, output_stream=stream_out, save_format=aw.SaveFormat.DOCX, watermark_text=watermark_text, options=options)
        #ExEnd:WatermarkTextStream

    def test_watermark_image(self):
        #ExStart:WatermarkImage
        #ExFor:Watermarker.set_image(str,str,str)
        #ExFor:Watermarker.set_image(str,str,SaveFormat,str)
        #ExFor:Watermarker.set_image(str,str,str,ImageWatermarkOptions)
        #ExFor:Watermarker.set_image(str,str,SaveFormat,str,ImageWatermarkOptions)
        #ExSummary:Shows how to insert watermark image to the document.
        doc = MY_DIR + 'Document.docx'
        watermark_image = IMAGE_DIR + 'Logo.jpg'
        aw.lowcode.Watermarker.set_image(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.SetWatermarkImage.1.docx', watermark_image_file_name=watermark_image)
        aw.lowcode.Watermarker.set_image(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.SetWatermarkText.2.docx', save_format=aw.SaveFormat.DOCX, watermark_image_file_name=watermark_image)
        options = aw.ImageWatermarkOptions()
        options.scale = 50
        aw.lowcode.Watermarker.set_image(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.SetWatermarkText.3.docx', watermark_image_file_name=watermark_image, options=options)
        aw.lowcode.Watermarker.set_image(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.SetWatermarkText.4.docx', save_format=aw.SaveFormat.DOCX, watermark_image_file_name=watermark_image, options=options)
        #ExEnd:WatermarkImage

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
