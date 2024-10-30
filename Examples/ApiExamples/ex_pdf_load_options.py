# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import aspose.words.loading
import unittest
from api_example_base import ApiExampleBase, MY_DIR

class ExPdfLoadOptions(ApiExampleBase):

    def test_skip_pdf_images(self):
        for is_skip_pdf_images in [True, False]:
            #ExStart
            #ExFor:PdfLoadOptions
            #ExFor:PdfLoadOptions.skip_pdf_images
            #ExFor:PdfLoadOptions.page_index
            #ExFor:PdfLoadOptions.page_count
            #ExSummary:Shows how to skip images during loading PDF files.
            options = aw.loading.PdfLoadOptions()
            options.skip_pdf_images = is_skip_pdf_images
            options.page_index = 0
            options.page_count = 1
            doc = aw.Document(file_name=MY_DIR + 'Images.pdf', load_options=options)
            shape_collection = doc.get_child_nodes(aw.NodeType.SHAPE, True)
            if is_skip_pdf_images:
                self.assertEqual(shape_collection.count, 0)
            else:
                self.assertNotEqual(shape_collection.count, 0)
            #ExEnd