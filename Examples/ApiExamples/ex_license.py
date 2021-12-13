import unittest
import io
import os
import shutil

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir, license_dir, assembly_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir
LICENSE_DIR = license_dir
ASSEMBLY_DIR = assembly_dir

class ExLicense(ApiExampleBase):

    def test_license_from_file_no_path(self):

        #ExStart
        #ExFor:License
        #ExFor:License.#ctor
        #ExFor:License.SetLicense(String)
        #ExSummary:Shows how initialize a license for Aspose.Words using a license file in the local file system.
        # Set the license for our Aspose.Words product by passing the local file system filename of a valid license file.
        license_file_name = os.path.join(LICENSE_DIR, "Aspose.Words.Python.lic")

        license = aw.License()
        license.set_license(license_file_name)

        # Create a copy of our license file in the binaries folder of our application.
        license_copy_file_name = os.path.join(ASSEMBLY_DIR, "Aspose.Words.Python.lic")
        shutil.copyfile(license_file_name, license_copy_file_name)

        # If we pass a file's name without a path,
        # the SetLicense will search several local file system locations for this file.
        # One of those locations will be the "bin" folder, which contains a copy of our license file.
        license.set_license("Aspose.Words.Python.lic")
        #ExEnd

        license.set_license("")
        os.unlink(license_copy_file_name)

    def test_license_from_stream(self):

        #ExStart
        #ExFor:License.SetLicense(Stream)
        #ExSummary:Shows how to initialize a license for Aspose.Words from a stream.
        # Set the license for our Aspose.Words product by passing a stream for a valid license file in our local file system.
        with open(os.path.join(LICENSE_DIR, "Aspose.words.n_e_t.lic", "rb")) as my_stream:
            license = aw.License()
            license.set_license(my_stream)

        #ExEnd
