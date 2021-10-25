import unittest
import aspose.words as aw
import os
import platform

# root_dir =  "X:/" if platform.system() == "Windows" else "/X/"
root_dir = os.path.abspath(os.curdir) + "/"
root_dir = root_dir[:root_dir.find("Aspose.Words-for-Python-via-.NET")]
api_examples_root = root_dir + "Aspose.Words-for-Python-via-.NET/Examples/"
license_path = root_dir + "Aspose.Words-for-Python-via-.NET/Temp/Aspose.Words.Python.NET.lic"
MY_DIR = api_examples_root + "Data/"
ARTIFACTS_DIR = MY_DIR + "Artifacts/"
temp_dir = MY_DIR + "Temp/"
IMAGES_DIR = MY_DIR + "Images/"
fonts_dir = MY_DIR + "MyFonts/"
database_dir = MY_DIR + "Database/"
JSON_DIR = MY_DIR + "JSON/"
aspose_logo_url = "https://www.aspose.cloud/templates/aspose/App_Themes/V3/images/words/header/aspose_words-for-net.png"


class DocsExamplesBase(unittest.TestCase):

    def setUp(self):
        if os.path.exists(ARTIFACTS_DIR):
            l = aw.License()
            l.set_license(license_path)
        if not os.path.exists(ARTIFACTS_DIR):
            os.makedirs(ARTIFACTS_DIR)
