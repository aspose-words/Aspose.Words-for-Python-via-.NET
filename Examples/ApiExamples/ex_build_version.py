import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

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
