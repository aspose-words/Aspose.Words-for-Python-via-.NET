import io

import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, LICENSE_PATH

class WorkingWithLicense(DocsExamplesBase):

    def test_apply_license_from_file(self):

        #ExStart:ApplyLicenseFromFile
        lic = aw.License()

        # Try to set license from the folder with the python script.
        try:
            lic.set_license("Aspose.Words.Python.NET.lic")
            print("License set successfully.")
        except RuntimeError as err:
            # We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license.
            print("\nThere was an error setting the license:", err)
        #ExEnd:ApplyLicenseFromFile

    def test_apply_license_from_stream(self):

        #ExStart:ApplyLicenseFromStream
        lic = aw.License()

        # Try to set license from the stream.
        try:
            with io.FileIO(LICENSE_PATH) as stream:
                lic.set_license(stream)
            print("License set successfully.")
        except RuntimeError as err:
            # We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license.
            print("\nThere was an error setting the license:", err)
        #ExEnd:ApplyLicenseFromStream

    def test_apply_metered_license(self):

        try:
            #ExStart:ApplyMeteredLicense
            # set metered public and private keys
            metered = aw.Metered()
            # Access the setMeteredKey property and pass public and private keys as parameters
            metered.set_metered_key("*****", "*****")

            # Load the document from disk.
            doc = aw.Document(MY_DIR + "Document.docx")
            #Get the page count of document
            print(doc.page_count)
            #ExEnd:ApplyMeteredLicense
        except RuntimeError as err:
            # We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license.
            print("\nThere was an error setting the license:", err)