import unittest
import aspose.words as aw
import os
import platform

# root_dir =  "X:/" if platform.system() == "Windows" else "/X/"
root_dir = os.path.abspath(os.curdir) + "/"
root_dir = root_dir[:root_dir.find("Aspose.Words-for-Python-via-.NET")]
api_examples_root = root_dir + "Aspose.Words-for-Python-via-.NET/Examples/"
license_path = root_dir + "Aspose.Words-for-Python-via-.NET/Temp/Aspose.Words.Python.NET.lic"
my_dir = api_examples_root + "Data/"
artifacts_dir = my_dir + "Artifacts/"
golds_dir = my_dir + "Golds/"
temp_dir = my_dir + "Temp/"
image_dir = my_dir + "Images/"
fonts_dir = my_dir + "MyFonts/"
aspose_logo_url = "https://www.aspose.cloud/templates/aspose/App_Themes/V3/images/words/header/aspose_words-for-net.png"


class ApiExampleBase(unittest.TestCase):

    def setUp(self):
        if os.path.exists(artifacts_dir):
            l = aw.License()
            l.set_license(license_path)
        if not os.path.exists(artifacts_dir):
            os.makedirs(artifacts_dir)