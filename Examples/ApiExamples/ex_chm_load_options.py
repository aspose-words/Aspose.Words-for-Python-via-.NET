# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExChmLoadOptions(ApiExampleBase):

    def test_original_file_name(self):

        #ExStart
        #ExFor:ChmLoadOptions.original_file_name
        #ExSummary:Shows how to resolve URLs like "ms-its:myfile.chm::/index.htm".
        
        # Our document contains URLs like "ms-its:amhelp.chm::....htm", but it has a different name,
        # so file links don't work after saving it to HTML.
        # We need to define the original filename in 'ChmLoadOptions' to avoid this behavior.
        load_options = aw.loading.ChmLoadOptions()
        load_options.original_file_name = "amhelp.chm"

        with open(MY_DIR + "Document with ms-its links.chm", "rb") as stream:
            doc = aw.Document(stream, load_options)

        doc.save(ARTIFACTS_DIR + "ExChmLoadOptions.OriginalFileName.html")
        #ExEnd
