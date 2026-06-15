# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import aspose.words.loading
import unittest
from api_example_base import ApiExampleBase, MY_DIR

class ExRtfLoadOptions(ApiExampleBase):

    def test_recognize_utf_8_text(self):
        for recognize_utf_8_text in [False, True]:
            #ExStart
            #ExFor:RtfLoadOptions
            #ExFor:RtfLoadOptions.__init__
            #ExFor:RtfLoadOptions.recognize_utf8_text
            #ExSummary:Shows how to detect UTF-8 characters while loading an RTF document.
            # Create an "RtfLoadOptions" object to modify how we load an RTF document.
            load_options = aw.loading.RtfLoadOptions()
            # Set the "RecognizeUtf8Text" property to "false" to assume that the document uses the ISO 8859-1 charset
            # and loads every character in the document.
            # Set the "RecognizeUtf8Text" property to "true" to parse any variable-length characters that may occur in the text.
            load_options.recognize_utf8_text = recognize_utf_8_text
            doc = aw.Document(file_name=MY_DIR + 'UTF-8 characters.rtf', load_options=load_options)
            self.assertEqual('“John Doe´s list of currency symbols”™\r' + '€, ¢, £, ¥, ¤' if recognize_utf_8_text else 'â€œJohn DoeÂ´s list of currency symbolsâ€\x9dâ„¢\r' + 'â‚¬, Â¢, Â£, Â¥, Â¤', doc.first_section.body.get_text().strip())
            #ExEnd