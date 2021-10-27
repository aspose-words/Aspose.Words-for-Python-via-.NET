import unittest
import aspose.words as aw
import os

# root_dir =  "X:/" if platform.system() == "Windows" else "/X/"
ROOT_DIR = os.path.abspath(os.curdir) + "/"
ROOT_DIR = ROOT_DIR[:ROOT_DIR.find("Aspose.Words-for-Python-via-.NET")]
API_EXAMPLES_ROOT = ROOT_DIR + "Aspose.Words-for-Python-via-.NET/Examples/"
LECIENSE_PATH = ROOT_DIR + "Aspose.Words-for-Python-via-.NET/Temp/Aspose.Words.Python.NET.lic"
MY_DIR = API_EXAMPLES_ROOT + "Data/"
ARTIFACTS_DIR = MY_DIR + "Artifacts/"
TEMP_DIR = MY_DIR + "Temp/"
IMAGES_DIR = MY_DIR + "Images/"
FONTS_DIR = MY_DIR + "MyFonts/"
DATABASE_DIR = MY_DIR + "Database/"
JSON_DIR = MY_DIR + "JSON/"
ASPOSE_LOGO_URL = "https://www.aspose.cloud/templates/aspose/App_Themes/V3/images/words/header/aspose_words-for-net.png"


class DocsExamplesBase(unittest.TestCase):

    def setUp(self):
        if os.path.exists(LECIENSE_PATH):
            lic = aw.License()
            lic.set_license(LECIENSE_PATH)
        if not os.path.exists(ARTIFACTS_DIR):
            os.makedirs(ARTIFACTS_DIR)
