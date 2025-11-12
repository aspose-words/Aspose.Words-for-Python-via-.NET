# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.pydrawing
import aspose.words as aw
import aspose.words.comparing
import aspose.words.loading
import aspose.words.lowcode
import aspose.words.replacing
import aspose.words.saving
import datetime
import system_helper
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, IMAGE_DIR, MY_DIR

class ExLowCode(ApiExampleBase):

    def test_merge_documents(self):
        #ExStart
        #ExFor:Merger.merge(str,List[str])
        #ExFor:Merger.merge(List[str],MergeFormatMode)
        #ExFor:Merger.merge(List[str],List[LoadOptions],MergeFormatMode)
        #ExFor:Merger.merge(str,List[str],SaveOptions,MergeFormatMode)
        #ExFor:Merger.merge(str,List[str],SaveFormat,MergeFormatMode)
        #ExFor:Merger.merge(str,List[str],List[LoadOptions],SaveOptions,MergeFormatMode)
        #ExFor:LowCode.merge_format_mode
        #ExFor:LowCode.merger
        #ExSummary:Shows how to merge documents into a single output document.
        #There is a several ways to merge documents:
        input_doc1 = MY_DIR + 'Big document.docx'
        input_doc2 = MY_DIR + 'Tables.docx'
        aw.lowcode.Merger.merge(output_file=ARTIFACTS_DIR + 'LowCode.MergeDocument.1.docx', input_files=[input_doc1, input_doc2])
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.password = 'Aspose.Words'
        aw.lowcode.Merger.merge(output_file=ARTIFACTS_DIR + 'LowCode.MergeDocument.2.docx', input_files=[input_doc1, input_doc2], save_options=save_options, merge_format_mode=aw.lowcode.MergeFormatMode.KEEP_SOURCE_FORMATTING)
        aw.lowcode.Merger.merge(output_file=ARTIFACTS_DIR + 'LowCode.MergeDocument.3.pdf', input_files=[input_doc1, input_doc2], save_format=aw.SaveFormat.PDF, merge_format_mode=aw.lowcode.MergeFormatMode.KEEP_SOURCE_LAYOUT)
        first_load_options = aw.loading.LoadOptions()
        first_load_options.ignore_ole_data = True
        second_load_options = aw.loading.LoadOptions()
        second_load_options.ignore_ole_data = False
        aw.lowcode.Merger.merge(output_file=ARTIFACTS_DIR + 'LowCode.MergeDocument.4.docx', input_files=[input_doc1, input_doc2], load_options=[first_load_options, second_load_options], save_options=save_options, merge_format_mode=aw.lowcode.MergeFormatMode.KEEP_SOURCE_FORMATTING)
        doc = aw.lowcode.Merger.merge(input_files=[input_doc1, input_doc2], merge_format_mode=aw.lowcode.MergeFormatMode.MERGE_FORMATTING)
        doc.save(file_name=ARTIFACTS_DIR + 'LowCode.MergeDocument.5.docx')
        doc = aw.lowcode.Merger.merge(input_files=[input_doc1, input_doc2], load_options=[first_load_options, second_load_options], merge_format_mode=aw.lowcode.MergeFormatMode.MERGE_FORMATTING)
        doc.save(file_name=ARTIFACTS_DIR + 'LowCode.MergeDocument.6.docx')
        #ExEnd

    def test_merge_context_documents(self):
        #ExStart:MergeContextDocuments
        #ExFor:Processor
        #ExFor:Processor.from_file(str,LoadOptions)
        #ExFor:Processor.to_file(str,SaveOptions)
        #ExFor:Processor.to_file(str,SaveFormat)
        #ExFor:Processor.execute
        #ExFor:Merger.create(MergerContext)
        #ExFor:MergerContext
        #ExSummary:Shows how to merge documents into a single output document using context.
        #There is a several ways to merge documents:
        input_doc1 = MY_DIR + 'Big document.docx'
        input_doc2 = MY_DIR + 'Tables.docx'
        context = aw.lowcode.MergerContext()
        context.merge_format_mode = aw.lowcode.MergeFormatMode.KEEP_SOURCE_FORMATTING
        aw.lowcode.Merger.create(context).from_file(input=input_doc1).from_file(input=input_doc2).to_file(output=ARTIFACTS_DIR + 'LowCode.MergeContextDocuments.1.docx').execute()
        first_load_options = aw.loading.LoadOptions()
        first_load_options.ignore_ole_data = True
        second_load_options = aw.loading.LoadOptions()
        second_load_options.ignore_ole_data = False
        context2 = aw.lowcode.MergerContext()
        context2.merge_format_mode = aw.lowcode.MergeFormatMode.KEEP_SOURCE_FORMATTING
        aw.lowcode.Merger.create(context2).from_file(input=input_doc1, load_options=first_load_options).from_file(input=input_doc2, load_options=second_load_options).to_file(output=ARTIFACTS_DIR + 'LowCode.MergeContextDocuments.2.docx', save_format=aw.SaveFormat.DOCX).execute()
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.password = 'Aspose.Words'
        context3 = aw.lowcode.MergerContext()
        context3.merge_format_mode = aw.lowcode.MergeFormatMode.KEEP_SOURCE_FORMATTING
        aw.lowcode.Merger.create(context3).from_file(input=input_doc1).from_file(input=input_doc2).to_file(output=ARTIFACTS_DIR + 'LowCode.MergeContextDocuments.3.docx', save_options=save_options).execute()
        #ExEnd:MergeContextDocuments

    def test_merge_stream_document(self):
        #ExStart
        #ExFor:Merger.merge_streams(List[BytesIO],MergeFormatMode)
        #ExFor:Merger.merge_streams(List[BytesIO],List[LoadOptions],MergeFormatMode)
        #ExFor:Merger.merge_streams(BytesIO,List[BytesIO],SaveOptions,MergeFormatMode)
        #ExFor:Merger.merge_streams(BytesIO,List[BytesIO],List[LoadOptions],SaveOptions,MergeFormatMode)
        #ExFor:Merger.merge_streams(BytesIO,List[BytesIO],SaveFormat)
        #ExSummary:Shows how to merge documents from stream into a single output document.
        #There is a several ways to merge documents from stream:
        with system_helper.io.FileStream(MY_DIR + 'Big document.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as first_stream_in:
            with system_helper.io.FileStream(MY_DIR + 'Tables.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as second_stream_in:
                save_options = aw.saving.OoxmlSaveOptions()
                save_options.password = 'Aspose.Words'
                with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.MergeStreamDocument.1.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                    aw.lowcode.Merger.merge_streams(output_stream=stream_out, input_streams=[first_stream_in, second_stream_in], save_options=save_options, merge_format_mode=aw.lowcode.MergeFormatMode.KEEP_SOURCE_FORMATTING)
                with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.MergeStreamDocument.2.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                    aw.lowcode.Merger.merge_streams(output_stream=stream_out, input_streams=[first_stream_in, second_stream_in], save_format=aw.SaveFormat.DOCX)
                first_load_options = aw.loading.LoadOptions()
                first_load_options.ignore_ole_data = True
                second_load_options = aw.loading.LoadOptions()
                second_load_options.ignore_ole_data = False
                with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.MergeStreamDocument.3.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                    aw.lowcode.Merger.merge_streams(output_stream=stream_out, input_streams=[first_stream_in, second_stream_in], load_options=[first_load_options, second_load_options], save_options=save_options, merge_format_mode=aw.lowcode.MergeFormatMode.KEEP_SOURCE_FORMATTING)
                first_doc = aw.lowcode.Merger.merge_streams(input_streams=[first_stream_in, second_stream_in], merge_format_mode=aw.lowcode.MergeFormatMode.MERGE_FORMATTING)
                first_doc.save(file_name=ARTIFACTS_DIR + 'LowCode.MergeStreamDocument.4.docx')
                second_doc = aw.lowcode.Merger.merge_streams(input_streams=[first_stream_in, second_stream_in], load_options=[first_load_options, second_load_options], merge_format_mode=aw.lowcode.MergeFormatMode.MERGE_FORMATTING)
                second_doc.save(file_name=ARTIFACTS_DIR + 'LowCode.MergeStreamDocument.5.docx')
        #ExEnd

    def test_merge_stream_context_documents(self):
        #ExStart:MergeStreamContextDocuments
        #ExFor:Processor
        #ExFor:Processor.from_stream(BytesIO,LoadOptions)
        #ExFor:Processor.to_stream(BytesIO,SaveFormat)
        #ExFor:Processor.to_stream(BytesIO,SaveOptions)
        #ExFor:Processor.execute
        #ExFor:Merger.create(MergerContext)
        #ExFor:MergerContext
        #ExSummary:Shows how to merge documents from stream into a single output document using context.
        #There is a several ways to merge documents:
        input_doc1 = MY_DIR + 'Big document.docx'
        input_doc2 = MY_DIR + 'Tables.docx'
        with system_helper.io.FileStream(MY_DIR + 'Big document.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as first_stream_in:
            with system_helper.io.FileStream(MY_DIR + 'Tables.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as second_stream_in:
                save_options = aw.saving.OoxmlSaveOptions()
                save_options.password = 'Aspose.Words'
                with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.MergeStreamContextDocuments.1.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                    context = aw.lowcode.MergerContext()
                    context.merge_format_mode = aw.lowcode.MergeFormatMode.KEEP_SOURCE_FORMATTING
                    aw.lowcode.Merger.create(context).from_stream(input=first_stream_in).from_stream(input=second_stream_in).to_stream(output=stream_out, save_options=save_options).execute()
                first_load_options = aw.loading.LoadOptions()
                first_load_options.ignore_ole_data = True
                second_load_options = aw.loading.LoadOptions()
                second_load_options.ignore_ole_data = False
                with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.MergeStreamContextDocuments.2.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                    context2 = aw.lowcode.MergerContext()
                    context2.merge_format_mode = aw.lowcode.MergeFormatMode.KEEP_SOURCE_FORMATTING
                    aw.lowcode.Merger.create(context2).from_stream(input=first_stream_in, load_options=first_load_options).from_stream(input=second_stream_in, load_options=second_load_options).to_stream(output=stream_out, save_format=aw.SaveFormat.DOCX).execute()
        #ExEnd:MergeStreamContextDocuments

    def test_merge_document_instances(self):
        #ExStart:MergeDocumentInstances
        #ExFor:Merger.merge_docs(List[Document],MergeFormatMode)
        #ExSummary:Shows how to merge input documents to a single document instance.
        first_doc = aw.DocumentBuilder()
        first_doc.font.size = 16
        first_doc.font.color = aspose.pydrawing.Color.blue
        first_doc.write('Hello first word!')
        second_doc = aw.DocumentBuilder()
        second_doc.write('Hello second word!')
        merged_doc = aw.lowcode.Merger.merge_docs(input_documents=[first_doc.document, second_doc.document], merge_format_mode=aw.lowcode.MergeFormatMode.KEEP_SOURCE_LAYOUT)
        self.assertEqual('Hello first word!\x0cHello second word!\x0c', merged_doc.get_text())
        #ExEnd:MergeDocumentInstances

    def test_convert(self):
        #ExStart:Convert
        #ExFor:Converter.convert(str,str)
        #ExFor:Converter.convert(str,str,SaveFormat)
        #ExFor:Converter.convert(str,str,SaveOptions)
        #ExFor:Converter.convert(str,LoadOptions,str,SaveOptions)
        #ExSummary:Shows how to convert documents with a single line of code.
        doc = MY_DIR + 'Document.docx'
        aw.lowcode.Converter.convert(input_file=doc, output_file=ARTIFACTS_DIR + 'LowCode.Convert.pdf')
        aw.lowcode.Converter.convert(input_file=doc, output_file=ARTIFACTS_DIR + 'LowCode.Convert.SaveFormat.rtf', save_format=aw.SaveFormat.RTF)
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.password = 'Aspose.Words'
        load_options = aw.loading.LoadOptions()
        load_options.ignore_ole_data = True
        aw.lowcode.Converter.convert(input_file=doc, load_options=load_options, output_file=ARTIFACTS_DIR + 'LowCode.Convert.LoadOptions.docx', save_options=save_options)
        aw.lowcode.Converter.convert(input_file=doc, output_file=ARTIFACTS_DIR + 'LowCode.Convert.SaveOptions.docx', save_options=save_options)
        #ExEnd:Convert

    def test_convert_context(self):
        #ExStart:ConvertContext
        #ExFor:Processor
        #ExFor:Processor.from_file(str,LoadOptions)
        #ExFor:Processor.to_file(str,SaveOptions)
        #ExFor:Processor.execute
        #ExFor:Converter.create(ConverterContext)
        #ExFor:ConverterContext
        #ExSummary:Shows how to convert documents with a single line of code using context.
        doc = MY_DIR + 'Big document.docx'
        aw.lowcode.Converter.create(aw.lowcode.ConverterContext()).from_file(input=doc).to_file(output=ARTIFACTS_DIR + 'LowCode.ConvertContext.1.pdf').execute()
        aw.lowcode.Converter.create(aw.lowcode.ConverterContext()).from_file(input=doc).to_file(output=ARTIFACTS_DIR + 'LowCode.ConvertContext.2.pdf', save_format=aw.SaveFormat.RTF).execute()
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.password = 'Aspose.Words'
        load_options = aw.loading.LoadOptions()
        load_options.ignore_ole_data = True
        aw.lowcode.Converter.create(aw.lowcode.ConverterContext()).from_file(input=doc, load_options=load_options).to_file(output=ARTIFACTS_DIR + 'LowCode.ConvertContext.3.docx', save_options=save_options).execute()
        aw.lowcode.Converter.create(aw.lowcode.ConverterContext()).from_file(input=doc).to_file(output=ARTIFACTS_DIR + 'LowCode.ConvertContext.4.png', save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)).execute()
        #ExEnd:ConvertContext

    def test_convert_stream(self):
        #ExStart:ConvertStream
        #ExFor:Converter.convert(BytesIO,BytesIO,SaveFormat)
        #ExFor:Converter.convert(BytesIO,BytesIO,SaveOptions)
        #ExFor:Converter.convert(BytesIO,LoadOptions,BytesIO,SaveOptions)
        #ExSummary:Shows how to convert documents with a single line of code (Stream).
        with system_helper.io.FileStream(MY_DIR + 'Big document.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.ConvertStream.1.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                aw.lowcode.Converter.convert(input_stream=stream_in, output_stream=stream_out, save_format=aw.SaveFormat.DOCX)
            save_options = aw.saving.OoxmlSaveOptions()
            save_options.password = 'Aspose.Words'
            load_options = aw.loading.LoadOptions()
            load_options.ignore_ole_data = True
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.ConvertStream.2.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                aw.lowcode.Converter.convert(input_stream=stream_in, load_options=load_options, output_stream=stream_out, save_options=save_options)
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.ConvertStream.3.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                aw.lowcode.Converter.convert(input_stream=stream_in, output_stream=stream_out, save_options=save_options)
        #ExEnd:ConvertStream

    def test_convert_to_images(self):
        #ExStart:ConvertToImages
        #ExFor:Converter.convert_to_images(str,str)
        #ExFor:Converter.convert_to_images(str,str,SaveFormat)
        #ExFor:Converter.convert_to_images(str,str,ImageSaveOptions)
        #ExFor:Converter.convert_to_images(str,LoadOptions,str,ImageSaveOptions)
        #ExSummary:Shows how to convert document to images.
        doc = MY_DIR + 'Big document.docx'
        aw.lowcode.Converter.convert(input_file=doc, output_file=ARTIFACTS_DIR + 'LowCode.ConvertToImages.1.png')
        aw.lowcode.Converter.convert(input_file=doc, output_file=ARTIFACTS_DIR + 'LowCode.ConvertToImages.2.jpeg', save_format=aw.SaveFormat.JPEG)
        load_options = aw.loading.LoadOptions()
        load_options.ignore_ole_data = False
        image_save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
        image_save_options.page_set = aw.saving.PageSet(page=1)
        aw.lowcode.Converter.convert(input_file=doc, load_options=load_options, output_file=ARTIFACTS_DIR + 'LowCode.ConvertToImages.3.png', save_options=image_save_options)
        aw.lowcode.Converter.convert(input_file=doc, output_file=ARTIFACTS_DIR + 'LowCode.ConvertToImages.4.png', save_options=image_save_options)
        #ExEnd:ConvertToImages

    def test_convert_to_images_stream(self):
        #ExStart:ConvertToImagesStream
        #ExFor:Converter.convert_to_images(str,SaveFormat)
        #ExFor:Converter.convert_to_images(str,ImageSaveOptions)
        #ExFor:Converter.convert_to_images(Document,SaveFormat)
        #ExFor:Converter.convert_to_images(Document,ImageSaveOptions)
        #ExSummary:Shows how to convert document to images stream.
        doc = MY_DIR + 'Big document.docx'
        streams = aw.lowcode.Converter.convert_to_images(input_file=doc, save_format=aw.SaveFormat.PNG)
        image_save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
        image_save_options.page_set = aw.saving.PageSet(page=1)
        streams = aw.lowcode.Converter.convert_to_images(input_file=doc, save_options=image_save_options)
        streams = aw.lowcode.Converter.convert_to_images(doc=aw.Document(file_name=doc), save_format=aw.SaveFormat.PNG)
        streams = aw.lowcode.Converter.convert_to_images(doc=aw.Document(file_name=doc), save_options=image_save_options)
        #ExEnd:ConvertToImagesStream

    def test_convert_to_images_from_stream(self):
        #ExStart:ConvertToImagesFromStream
        #ExFor:Converter.convert_to_images(BytesIO,SaveFormat)
        #ExFor:Converter.convert_to_images(BytesIO,ImageSaveOptions)
        #ExFor:Converter.convert_to_images(BytesIO,LoadOptions,ImageSaveOptions)
        #ExSummary:Shows how to convert document to images from stream.
        with system_helper.io.FileStream(MY_DIR + 'Big document.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            streams = aw.lowcode.Converter.convert_to_images(input_stream=stream_in, save_format=aw.SaveFormat.JPEG)
            image_save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
            image_save_options.page_set = aw.saving.PageSet(page=1)
            streams = aw.lowcode.Converter.convert_to_images(input_stream=stream_in, save_options=image_save_options)
            load_options = aw.loading.LoadOptions()
            load_options.ignore_ole_data = False
            aw.lowcode.Converter.convert_to_images(input_stream=stream_in, load_options=load_options, save_options=image_save_options)
        #ExEnd:ConvertToImagesFromStream

    def test_compare_documents(self):
        #ExStart:CompareDocuments
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

    def test_compare_context_documents(self):
        #ExStart:CompareContextDocuments
        #ExFor:Comparer.create(ComparerContext)
        #ExFor:ComparerContext
        #ExFor:ComparerContext.compare_options
        #ExSummary:Shows how to simple compare documents using context.
        # There is a several ways to compare documents:
        first_doc = MY_DIR + 'Table column bookmarks.docx'
        second_doc = MY_DIR + 'Table column bookmarks.doc'
        comparer_context = aw.lowcode.ComparerContext()
        comparer_context.compare_options.ignore_case_changes = True
        comparer_context.author = 'Author'
        comparer_context.date_time = datetime.datetime(1, 1, 1)
        aw.lowcode.Comparer.create(comparer_context).from_file(input=first_doc).from_file(input=second_doc).to_file(output=ARTIFACTS_DIR + 'LowCode.CompareContextDocuments.docx').execute()
        #ExEnd:CompareContextDocuments

    def test_compare_stream_documents(self):
        #ExStart:CompareStreamDocuments
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

    def test_compare_context_stream_documents(self):
        #ExStart:CompareContextStreamDocuments
        #ExFor:Comparer.create(ComparerContext)
        #ExFor:ComparerContext
        #ExFor:ComparerContext.compare_options
        #ExSummary:Shows how to compare documents from the stream using context.
        # There is a several ways to compare documents from the stream:
        with system_helper.io.FileStream(MY_DIR + 'Table column bookmarks.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as first_stream_in:
            with system_helper.io.FileStream(MY_DIR + 'Table column bookmarks.doc', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as second_stream_in:
                comparer_context = aw.lowcode.ComparerContext()
                comparer_context.compare_options.ignore_case_changes = True
                comparer_context.author = 'Author'
                comparer_context.date_time = datetime.datetime(1, 1, 1)
                with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.CompareContextStreamDocuments.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                    aw.lowcode.Comparer.create(comparer_context).from_stream(input=first_stream_in).from_stream(input=second_stream_in).to_stream(output=stream_out, save_format=aw.SaveFormat.DOCX).execute()
        #ExEnd:CompareContextStreamDocuments

    def test_compare_documents_toimages(self):
        #ExStart:CompareDocumentsToimages
        #ExFor:Comparer.compare_to_images(BytesIO,BytesIO,ImageSaveOptions,str,datetime,CompareOptions)
        #ExSummary:Shows how to compare documents and save results as images.
        # There is a several ways to compare documents:
        first_doc = MY_DIR + 'Table column bookmarks.docx'
        second_doc = MY_DIR + 'Table column bookmarks.doc'
        pages = aw.lowcode.Comparer.compare_to_images(v1=first_doc, v2=second_doc, image_save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), author='Author', date_time=datetime.datetime(1, 1, 1))
        with system_helper.io.FileStream(first_doc, system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as first_stream_in:
            with system_helper.io.FileStream(second_doc, system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as second_stream_in:
                compare_options = aw.comparing.CompareOptions()
                compare_options.ignore_case_changes = True
                pages = aw.lowcode.Comparer.compare_to_images(v1=first_stream_in, v2=second_stream_in, image_save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), author='Author', date_time=datetime.datetime(1, 1, 1), compare_options=compare_options)
        #ExEnd:CompareDocumentsToimages

    def test_mail_merge(self):
        #ExStart:MailMerge
        #ExFor:MailMergeOptions
        #ExFor:MailMergeOptions.trim_whitespaces
        #ExFor:MailMerger.execute(str,str,List[str],List[object])
        #ExFor:MailMerger.execute(str,str,SaveFormat,List[str],List[object],MailMergeOptions)
        #ExSummary:Shows how to do mail merge operation for a single record.
        # There is a several ways to do mail merge operation:
        doc = MY_DIR + 'Mail merge.doc'
        field_names = ['FirstName', 'Location', 'SpecialCharsInName()']
        field_values = ['James Bond', 'London', 'Classified']
        aw.lowcode.MailMerger.execute(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.MailMerge.1.docx', field_names=field_names, field_values=field_values)
        aw.lowcode.MailMerger.execute(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.MailMerge.2.docx', save_format=aw.SaveFormat.DOCX, field_names=field_names, field_values=field_values)
        mail_merge_options = aw.lowcode.MailMergeOptions()
        mail_merge_options.trim_whitespaces = True
        aw.lowcode.MailMerger.execute(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.MailMerge.3.docx', save_format=aw.SaveFormat.DOCX, field_names=field_names, field_values=field_values, mail_merge_options=mail_merge_options)
        #ExEnd:MailMerge

    def test_mail_merge_context(self):
        #ExStart:MailMergeContext
        #ExFor:MailMerger.create(MailMergerContext)
        #ExFor:MailMergerContext
        #ExFor:MailMergerContext.set_simple_data_source(List[str],List[object])
        #ExFor:MailMergerContext.mail_merge_options
        #ExSummary:Shows how to do mail merge operation for a single record using context.
        # There is a several ways to do mail merge operation:
        doc = MY_DIR + 'Mail merge.doc'
        field_names = ['FirstName', 'Location', 'SpecialCharsInName()']
        field_values = ['James Bond', 'London', 'Classified']
        mail_merger_context = aw.lowcode.MailMergerContext()
        mail_merger_context.set_simple_data_source(field_names=field_names, field_values=field_values)
        mail_merger_context.mail_merge_options.trim_whitespaces = True
        aw.lowcode.MailMerger.create(mail_merger_context).from_file(input=doc).to_file(output=ARTIFACTS_DIR + 'LowCode.MailMergeContext.docx').execute()
        #ExEnd:MailMergeContext

    def test_mail_merge_to_images(self):
        #ExStart:MailMergeToImages
        #ExFor:MailMerger.execute_to_images(str,ImageSaveOptions,List[str],List[object],MailMergeOptions)
        #ExSummary:Shows how to do mail merge operation for a single record and save result to images.
        # There is a several ways to do mail merge operation:
        doc = MY_DIR + 'Mail merge.doc'
        field_names = ['FirstName', 'Location', 'SpecialCharsInName()']
        field_values = ['James Bond', 'London', 'Classified']
        images = aw.lowcode.MailMerger.execute_to_images(input_file_name=doc, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), field_names=field_names, field_values=field_values)
        mail_merge_options = aw.lowcode.MailMergeOptions()
        mail_merge_options.trim_whitespaces = True
        images = aw.lowcode.MailMerger.execute_to_images(input_file_name=doc, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), field_names=field_names, field_values=field_values, mail_merge_options=mail_merge_options)
        #ExEnd:MailMergeToImages

    def test_mail_merge_stream(self):
        #ExStart:MailMergeStream
        #ExFor:MailMerger.execute(BytesIO,BytesIO,SaveFormat,List[str],List[object],MailMergeOptions)
        #ExSummary:Shows how to do mail merge operation for a single record from the stream.
        # There is a several ways to do mail merge operation using documents from the stream:
        field_names = ['FirstName', 'Location', 'SpecialCharsInName()']
        field_values = ['James Bond', 'London', 'Classified']
        with system_helper.io.FileStream(MY_DIR + 'Mail merge.doc', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.MailMergeStream.1.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                aw.lowcode.MailMerger.execute(input_stream=stream_in, output_stream=stream_out, save_format=aw.SaveFormat.DOCX, field_names=field_names, field_values=field_values)
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.MailMergeStream.2.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                mail_merge_options = aw.lowcode.MailMergeOptions()
                mail_merge_options.trim_whitespaces = True
                aw.lowcode.MailMerger.execute(input_stream=stream_in, output_stream=stream_out, save_format=aw.SaveFormat.DOCX, field_names=field_names, field_values=field_values, mail_merge_options=mail_merge_options)
        #ExEnd:MailMergeStream

    def test_mail_merge_context_stream(self):
        #ExStart:MailMergeContextStream
        #ExFor:MailMerger.create(MailMergerContext)
        #ExFor:MailMergerContext
        #ExFor:MailMergerContext.set_simple_data_source(List[str],List[object])
        #ExFor:MailMergerContext.mail_merge_options
        #ExSummary:Shows how to do mail merge operation for a single record from the stream using context.
        # There is a several ways to do mail merge operation using documents from the stream:
        field_names = ['FirstName', 'Location', 'SpecialCharsInName()']
        field_values = ['James Bond', 'London', 'Classified']
        with system_helper.io.FileStream(MY_DIR + 'Mail merge.doc', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            mail_merger_context = aw.lowcode.MailMergerContext()
            mail_merger_context.set_simple_data_source(field_names=field_names, field_values=field_values)
            mail_merger_context.mail_merge_options.trim_whitespaces = True
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.MailMergeContextStream.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                aw.lowcode.MailMerger.create(mail_merger_context).from_stream(input=stream_in).to_stream(output=stream_out, save_format=aw.SaveFormat.DOCX).execute()
        #ExEnd:MailMergeContextStream

    def test_mail_merge_stream_to_images(self):
        #ExStart:MailMergeStreamToImages
        #ExFor:MailMerger.execute_to_images(BytesIO,ImageSaveOptions,List[str],List[object],MailMergeOptions)
        #ExSummary:Shows how to do mail merge operation for a single record from the stream and save result to images.
        # There is a several ways to do mail merge operation using documents from the stream:
        field_names = ['FirstName', 'Location', 'SpecialCharsInName()']
        field_values = ['James Bond', 'London', 'Classified']
        with system_helper.io.FileStream(MY_DIR + 'Mail merge.doc', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            images = aw.lowcode.MailMerger.execute_to_images(input_stream=stream_in, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), field_names=field_names, field_values=field_values)
            mail_merge_options = aw.lowcode.MailMergeOptions()
            mail_merge_options.trim_whitespaces = True
            images = aw.lowcode.MailMerger.execute_to_images(input_stream=stream_in, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), field_names=field_names, field_values=field_values, mail_merge_options=mail_merge_options)
        #ExEnd:MailMergeStreamToImages

    def test_replace(self):
        #ExStart:Replace
        #ExFor:Replacer.replace(str,str,str,str)
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

    def test_replace_context(self):
        #ExStart:ReplaceContext
        #ExFor:Replacer.create(ReplacerContext)
        #ExFor:ReplacerContext
        #ExFor:ReplacerContext.set_replacement(str,str)
        #ExFor:ReplacerContext.find_replace_options
        #ExSummary:Shows how to replace string in the document using context.
        # There is a several ways to replace string in the document:
        doc = MY_DIR + 'Footer.docx'
        pattern = '(C)2006 Aspose Pty Ltd.'
        replacement = 'Copyright (C) 2024 by Aspose Pty Ltd.'
        replacer_context = aw.lowcode.ReplacerContext()
        replacer_context.set_replacement(pattern=pattern, replacement=replacement)
        replacer_context.find_replace_options.find_whole_words_only = False
        aw.lowcode.Replacer.create(replacer_context).from_file(input=doc).to_file(output=ARTIFACTS_DIR + 'LowCode.ReplaceContext.docx').execute()
        #ExEnd:ReplaceContext

    def test_replace_to_images(self):
        #ExStart:ReplaceToImages
        #ExFor:Replacer.replace_to_images(str,ImageSaveOptions,str,str,FindReplaceOptions)
        #ExSummary:Shows how to replace string in the document and save result to images.
        # There is a several ways to replace string in the document:
        doc = MY_DIR + 'Footer.docx'
        pattern = '(C)2006 Aspose Pty Ltd.'
        replacement = 'Copyright (C) 2024 by Aspose Pty Ltd.'
        images = aw.lowcode.Replacer.replace_to_images(input_file_name=doc, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), pattern=pattern, replacement=replacement)
        options = aw.replacing.FindReplaceOptions()
        options.find_whole_words_only = False
        images = aw.lowcode.Replacer.replace_to_images(input_file_name=doc, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), pattern=pattern, replacement=replacement, options=options)
        #ExEnd:ReplaceToImages

    def test_replace_stream(self):
        #ExStart:ReplaceStream
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

    def test_replace_context_stream(self):
        #ExStart:ReplaceContextStream
        #ExFor:Replacer.create(ReplacerContext)
        #ExFor:ReplacerContext
        #ExFor:ReplacerContext.set_replacement(str,str)
        #ExFor:ReplacerContext.find_replace_options
        #ExSummary:Shows how to replace string in the document using documents from the stream using context.
        # There is a several ways to replace string in the document using documents from the stream:
        pattern = '(C)2006 Aspose Pty Ltd.'
        replacement = 'Copyright (C) 2024 by Aspose Pty Ltd.'
        with system_helper.io.FileStream(MY_DIR + 'Footer.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            replacer_context = aw.lowcode.ReplacerContext()
            replacer_context.set_replacement(pattern=pattern, replacement=replacement)
            replacer_context.find_replace_options.find_whole_words_only = False
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.ReplaceContextStream.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                aw.lowcode.Replacer.create(replacer_context).from_stream(input=stream_in).to_stream(output=stream_out, save_format=aw.SaveFormat.DOCX).execute()
        #ExEnd:ReplaceContextStream

    def test_replace_to_images_stream(self):
        #ExStart:ReplaceToImagesStream
        #ExFor:Replacer.replace_to_images(BytesIO,ImageSaveOptions,str,str,FindReplaceOptions)
        #ExSummary:Shows how to replace string in the document using documents from the stream and save result to images.
        # There is a several ways to replace string in the document using documents from the stream:
        pattern = '(C)2006 Aspose Pty Ltd.'
        replacement = 'Copyright (C) 2024 by Aspose Pty Ltd.'
        with system_helper.io.FileStream(MY_DIR + 'Footer.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            images = aw.lowcode.Replacer.replace_to_images(input_stream=stream_in, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), pattern=pattern, replacement=replacement)
            options = aw.replacing.FindReplaceOptions()
            options.find_whole_words_only = False
            images = aw.lowcode.Replacer.replace_to_images(input_stream=stream_in, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), pattern=pattern, replacement=replacement, options=options)
        #ExEnd:ReplaceToImagesStream

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
        #ExFor:SplitCriteria
        #ExFor:SplitOptions.split_criteria
        #ExFor:Splitter.split(str,str,SplitOptions)
        #ExFor:Splitter.split(str,str,SaveFormat,SplitOptions)
        #ExSummary:Shows how to split document by pages.
        doc = MY_DIR + 'Big document.docx'
        options = aw.lowcode.SplitOptions()
        options.split_criteria = aw.lowcode.SplitCriteria.PAGE
        aw.lowcode.Splitter.split(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.SplitDocument.1.docx', options=options)
        aw.lowcode.Splitter.split(input_file_name=doc, output_file_name=ARTIFACTS_DIR + 'LowCode.SplitDocument.2.docx', save_format=aw.SaveFormat.DOCX, options=options)
        #ExEnd:SplitDocument

    def test_split_context_document(self):
        #ExStart:SplitContextDocument
        #ExFor:Splitter.create(SplitterContext)
        #ExFor:SplitterContext
        #ExFor:SplitterContext.split_options
        #ExSummary:Shows how to split document by pages using context.
        doc = MY_DIR + 'Big document.docx'
        splitter_context = aw.lowcode.SplitterContext()
        splitter_context.split_options.split_criteria = aw.lowcode.SplitCriteria.PAGE
        aw.lowcode.Splitter.create(splitter_context).from_file(input=doc).to_file(output=ARTIFACTS_DIR + 'LowCode.SplitContextDocument.docx').execute()
        #ExEnd:SplitContextDocument

    def test_split_document_stream(self):
        #ExStart:SplitDocumentStream
        #ExFor:Splitter.split(BytesIO,SaveFormat,SplitOptions)
        #ExSummary:Shows how to split document from the stream by pages.
        with system_helper.io.FileStream(MY_DIR + 'Big document.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            options = aw.lowcode.SplitOptions()
            options.split_criteria = aw.lowcode.SplitCriteria.PAGE
            stream = aw.lowcode.Splitter.split(input_stream=stream_in, save_format=aw.SaveFormat.DOCX, options=options)
        #ExEnd:SplitDocumentStream

    def test_watermark_text(self):
        #ExStart:WatermarkText
        #ExFor:Watermarker.set_text(str,str,str)
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

    def test_watermark_context_text(self):
        #ExStart:WatermarkContextText
        #ExFor:Watermarker.create(WatermarkerContext)
        #ExFor:WatermarkerContext
        #ExFor:WatermarkerContext.text_watermark
        #ExFor:WatermarkerContext.text_watermark_options
        #ExSummary:Shows how to insert watermark text to the document using context.
        doc = MY_DIR + 'Big document.docx'
        watermark_text = 'This is a watermark'
        watermarker_context = aw.lowcode.WatermarkerContext()
        watermarker_context.text_watermark = watermark_text
        watermarker_context.text_watermark_options.color = aspose.pydrawing.Color.red
        aw.lowcode.Watermarker.create(watermarker_context).from_file(input=doc).to_file(output=ARTIFACTS_DIR + 'LowCode.WatermarkContextText.docx').execute()
        #ExEnd:WatermarkContextText

    def test_watermark_text_stream(self):
        #ExStart:WatermarkTextStream
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

    def test_watermark_context_text_stream(self):
        #ExStart:WatermarkContextTextStream
        #ExFor:Watermarker.create(WatermarkerContext)
        #ExFor:WatermarkerContext
        #ExFor:WatermarkerContext.text_watermark
        #ExFor:WatermarkerContext.text_watermark_options
        #ExSummary:Shows how to insert watermark text to the document from the stream using context.
        watermark_text = 'This is a watermark'
        with system_helper.io.FileStream(MY_DIR + 'Document.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            watermarker_context = aw.lowcode.WatermarkerContext()
            watermarker_context.text_watermark = watermark_text
            watermarker_context.text_watermark_options.color = aspose.pydrawing.Color.red
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.WatermarkContextTextStream.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                aw.lowcode.Watermarker.create(watermarker_context).from_stream(input=stream_in).to_stream(output=stream_out, save_format=aw.SaveFormat.DOCX).execute()
        #ExEnd:WatermarkContextTextStream

    def test_watermark_image(self):
        #ExStart:WatermarkImage
        #ExFor:Watermarker.set_image(str,str,str)
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

    def test_watermark_context_image(self):
        #ExStart:WatermarkContextImage
        #ExFor:Watermarker.create(WatermarkerContext)
        #ExFor:WatermarkerContext
        #ExFor:WatermarkerContext.image_watermark
        #ExFor:WatermarkerContext.image_watermark_options
        #ExSummary:Shows how to insert watermark image to the document using context.
        doc = MY_DIR + 'Document.docx'
        watermark_image = IMAGE_DIR + 'Logo.jpg'
        watermarker_context = aw.lowcode.WatermarkerContext()
        watermarker_context.image_watermark = system_helper.io.File.read_all_bytes(watermark_image)
        watermarker_context.image_watermark_options.scale = 50
        aw.lowcode.Watermarker.create(watermarker_context).from_file(input=doc).to_file(output=ARTIFACTS_DIR + 'LowCode.WatermarkContextImage.docx').execute()
        #ExEnd:WatermarkContextImage

    def test_watermark_context_image_stream(self):
        #ExStart:WatermarkContextImageStream
        #ExFor:Watermarker.create(WatermarkerContext)
        #ExFor:WatermarkerContext
        #ExFor:WatermarkerContext.image_watermark
        #ExFor:WatermarkerContext.image_watermark_options
        #ExSummary:Shows how to insert watermark image to the document from a stream using context.
        watermark_image = IMAGE_DIR + 'Logo.jpg'
        with system_helper.io.FileStream(MY_DIR + 'Document.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            watermarker_context = aw.lowcode.WatermarkerContext()
            watermarker_context.image_watermark = system_helper.io.File.read_all_bytes(watermark_image)
            watermarker_context.image_watermark_options.scale = 50
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'LowCode.WatermarkContextImageStream.docx', system_helper.io.FileMode.CREATE, system_helper.io.FileAccess.READ_WRITE) as stream_out:
                aw.lowcode.Watermarker.create(watermarker_context).from_stream(input=stream_in).to_stream(output=stream_out, save_format=aw.SaveFormat.DOCX).execute()
        #ExEnd:WatermarkContextImageStream

    def test_watermark_text_to_images(self):
        #ExStart:WatermarkTextToImages
        #ExFor:Watermarker.set_watermark_to_images(str,ImageSaveOptions,str,TextWatermarkOptions)
        #ExSummary:Shows how to insert watermark text to the document and save result to images.
        doc = MY_DIR + 'Big document.docx'
        watermark_text = 'This is a watermark'
        images = aw.lowcode.Watermarker.set_watermark_to_images(input_file_name=doc, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), watermark_text=watermark_text)
        watermark_options = aw.TextWatermarkOptions()
        watermark_options.color = aspose.pydrawing.Color.red
        images = aw.lowcode.Watermarker.set_watermark_to_images(input_file_name=doc, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), watermark_text=watermark_text, options=watermark_options)
        #ExEnd:WatermarkTextToImages

    def test_watermark_text_to_images_stream(self):
        #ExStart:WatermarkTextToImagesStream
        #ExFor:Watermarker.set_watermark_to_images(BytesIO,ImageSaveOptions,str,TextWatermarkOptions)
        #ExSummary:Shows how to insert watermark text to the document from the stream and save result to images.
        watermark_text = 'This is a watermark'
        with system_helper.io.FileStream(MY_DIR + 'Document.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            images = aw.lowcode.Watermarker.set_watermark_to_images(input_stream=stream_in, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), watermark_text=watermark_text)
            watermark_options = aw.TextWatermarkOptions()
            watermark_options.color = aspose.pydrawing.Color.red
            images = aw.lowcode.Watermarker.set_watermark_to_images(input_stream=stream_in, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), watermark_text=watermark_text, options=watermark_options)
        #ExEnd:WatermarkTextToImagesStream

    def test_watermark_image_to_images(self):
        #ExStart:WatermarkImageToImages
        #ExFor:Watermarker.set_watermark_to_images(str,ImageSaveOptions,bytes,ImageWatermarkOptions)
        #ExSummary:Shows how to insert watermark image to the document and save result to images.
        doc = MY_DIR + 'Document.docx'
        watermark_image = IMAGE_DIR + 'Logo.jpg'
        aw.lowcode.Watermarker.set_watermark_to_images(input_file_name=doc, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), watermark_image_bytes=system_helper.io.File.read_all_bytes(watermark_image))
        options = aw.ImageWatermarkOptions()
        options.scale = 50
        aw.lowcode.Watermarker.set_watermark_to_images(input_file_name=doc, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), watermark_image_bytes=system_helper.io.File.read_all_bytes(watermark_image), options=options)
        #ExEnd:WatermarkImageToImages

    def test_watermark_image_to_images_stream(self):
        #ExStart:WatermarkImageToImagesStream
        #ExFor:Watermarker.set_watermark_to_images(BytesIO,ImageSaveOptions,BytesIO,ImageWatermarkOptions)
        #ExSummary:Shows how to insert watermark image to the document from a stream and save result to images.
        watermark_image = IMAGE_DIR + 'Logo.jpg'
        with system_helper.io.FileStream(MY_DIR + 'Document.docx', system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as stream_in:
            with system_helper.io.FileStream(watermark_image, system_helper.io.FileMode.OPEN, system_helper.io.FileAccess.READ) as image_stream:
                aw.lowcode.Watermarker.set_watermark_to_images(input_stream=stream_in, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), watermark_image_stream=image_stream)
                options = aw.ImageWatermarkOptions()
                options.scale = 50
                aw.lowcode.Watermarker.set_watermark_to_images(input_stream=stream_in, save_options=aw.saving.ImageSaveOptions(aw.SaveFormat.PNG), watermark_image_stream=image_stream, options=options)
        #ExEnd:WatermarkImageToImagesStream