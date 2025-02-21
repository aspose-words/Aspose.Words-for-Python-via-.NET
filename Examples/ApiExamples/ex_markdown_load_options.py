# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from datetime import date, datetime
import aspose.words as aw
import os
import io
import system_helper
from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExMarkdownLoadOptions(ApiExampleBase):

    def test_preserve_empty_lines(self):
        #ExStart: PreserveEmptyLines
        #ExFor: MarkdownLoadOptions
        #ExFor:MarkdownLoadOptions.PreserveEmptyLines
        #ExSummary: Shows how to preserve empty line while load a document.
        md_text = f'{os.linesep}Line1{os.linesep}{os.linesep}Line2{os.linesep}{os.linesep}'
        stream = io.BytesIO(str.encode(md_text))
        stream.seek(0)
        load_options = aw.loading.MarkdownLoadOptions()
        load_options.preserve_empty_lines = True
        doc = aw.Document(stream, load_options)
        self.assertEqual('\rLine1\r\rLine2\r\x0c', doc.get_text())
        #ExEnd: PreserveEmptyLines

    def test_import_underline_formatting(self):
        #ExStart:ImportUnderlineFormatting
        #ExFor:MarkdownLoadOptions.ImportUnderlineFormatting
        #ExSummary:Shows how to recognize plus characters "++" as underline text formatting.
        bytes_ = '++12 and B++'.encode(encoding='ascii')
        with io.BytesIO(bytes_) as byte_stream:
            load_options = aw.loading.MarkdownLoadOptions()
            load_options.import_underline_formatting = True
            doc = aw.Document(byte_stream, load_options)
            para = doc.get_child(aw.NodeType.PARAGRAPH, 0, True).as_paragraph()
            self.assertEqual(aw.Underline.SINGLE, para.runs[0].font.underline)
            load_options = aw.loading.MarkdownLoadOptions()
            load_options.import_underline_formatting = False
            doc = aw.Document(byte_stream, load_options)
            para = doc.get_child(aw.NodeType.PARAGRAPH, 0, True).as_paragraph()
            self.assertEqual(aw.Underline.NONE, para.runs[0].font.underline)
        #ExEnd:ImportUnderlineFormatting