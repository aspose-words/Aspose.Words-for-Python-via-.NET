# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import aspose.words.loading
import io
import system_helper
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExChmLoadOptions(ApiExampleBase):

    def test_original_file_name(self):
        #ExStart
        #ExFor:ChmLoadOptions
        #ExFor:ChmLoadOptions.__init__
        #ExFor:ChmLoadOptions.original_file_name
        #ExSummary:Shows how to resolve URLs like "ms-its:myfile.chm::/index.htm".
        # Our document contains URLs like "ms-its:amhelp.chm::....htm", but it has a different name,
        # so file links don't work after saving it to HTML.
        # We need to define the original filename in 'ChmLoadOptions' to avoid this behavior.
        load_options = aw.loading.ChmLoadOptions()
        load_options.original_file_name = 'amhelp.chm'
        doc = aw.Document(stream=io.BytesIO(system_helper.io.File.read_all_bytes(MY_DIR + 'Document with ms-its links.chm')), load_options=load_options)
        doc.save(file_name=ARTIFACTS_DIR + 'ExChmLoadOptions.OriginalFileName.html')
        #ExEnd