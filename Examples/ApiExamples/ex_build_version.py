# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw

from api_example_base import ApiExampleBase

class ExBuildVersion(ApiExampleBase):

    def test_print_build_version_info(self):

        #ExStart
        #ExFor:BuildVersionInfo
        #ExFor:BuildVersionInfo.product
        #ExFor:BuildVersionInfo.version
        #ExSummary:Shows how to display information about your installed version of Aspose.Words.
        print(f"I am currently using {aw.BuildVersionInfo.product}, version number {aw.BuildVersionInfo.version}!")
        #ExEnd

        self.assertEqual("Aspose.Words for Python via .NET", aw.BuildVersionInfo.product)
        self.assertRegex(aw.BuildVersionInfo.version, "[0-9]{2}.[0-9]{1,2}.[0-9]")
