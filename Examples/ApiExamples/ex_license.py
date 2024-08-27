# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import shutil
import aspose.words as aw
import os
from api_example_base import ApiExampleBase, LICENSE_PATH

class ExLicense(ApiExampleBase):

    def test_license_from_file_no_path(self):
        #ExStart
        #ExFor:License
        #ExFor:License.__init__
        #ExFor:License.set_license(str)
        #ExSummary:Shows how initialize a license for Aspose.Words using a license file in the local file system.
        # Set the license for our Aspose.Words product by passing the local file system filename of a valid license file.
        license = aw.License()
        license.set_license(LICENSE_PATH)
        #ExEnd
        license.set_license('')

    def test_license_from_stream(self):
        #ExStart
        #ExFor:License.set_license(BytesIO)
        #ExSummary:Shows how to initialize a license for Aspose.Words from a stream.
        # Set the license for our Aspose.Words product by passing a stream for a valid license file in our local file system.
        with open(LICENSE_PATH, 'rb') as my_stream:
            license = aw.License()
            license.set_license(my_stream)