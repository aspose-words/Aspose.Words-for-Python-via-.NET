# -*- coding: utf-8 -*-
import io
import os
import aspose.words as aw
from datetime import date, datetime
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