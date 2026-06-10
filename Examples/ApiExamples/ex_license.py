import system_helper
import shutil
import unittest
# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import os
from api_example_base import ApiExampleBase, LICENSE_PATH

class ExLicense(ApiExampleBase):

    @unittest.skip('This test should be run manually')
    def test_license_from_file_no_path(self):
        #ExStart
        #ExFor:License
        #ExFor:License.__init__
        #ExFor:License.set_license(str)
        #ExSummary:Shows how to initialize a license for Aspose.Words using a license file in the local file system.
        import os
        import shutil
        test_license_file_name = 'Aspose.Total.NET.lic'
        # Set the license for our Aspose.Words product by passing the local file system filename of a valid license file.
        license_file_name = os.path.join(LICENSE_PATH, test_license_file_name)
        license = aw.License()
        license.set_license(license_name=license_file_name)
        # Create a copy of our license file in the binaries folder of our application.
        license_copy_file_name = os.path.join(AssemblyDir, test_license_file_name)
        shutil.copy2(license_file_name, license_copy_file_name)
        # If we pass a file's name without a path,
        # the SetLicense will search several local file system locations for this file.
        # One of those locations will be the "bin" folder, which contains a copy of our license file.
        license.set_license(license_name=test_license_file_name)
        #ExEnd
        license.set_license(license_name='')
        os.remove(license_copy_file_name)

    def test_license_from_stream(self):
        #ExStart
        #ExFor:License.set_license(BytesIO)
        #ExSummary:Shows how to initialize a license for Aspose.Words from a stream.
        # Set the license for our Aspose.Words product by passing a stream for a valid license file in our local file system.
        with system_helper.io.File.open_read(LICENSE_PATH) as my_stream:
            license = aw.License()
            license.set_license(stream=my_stream)
        #ExEnd