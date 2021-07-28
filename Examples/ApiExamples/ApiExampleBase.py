import unittest
import aspose.words as aw
import os
import platform

root_dir =  "X:/" if platform.system() == "Windows" else "/X/"
api_examples_root = root_dir + "Aspose.Words-for-Python-via-.NET/Examples/"
license_path = root_dir + "awnet/TestData/Licenses/Aspose.Words.NET.lic"
MyDir = api_examples_root + "Data/"
ArtifactsDir = MyDir + "Artifacts/"
GoldsDir = MyDir + "Golds/"
TempDir = MyDir + "Temp/"
ImageDir = MyDir + "Images/"
FontsDir = MyDir + "MyFonts/"

class ApiExampleBase(unittest.TestCase):

    def setUp(self):
        if os.path.exists(ArtifactsDir):
            l = aw.License()
            l.set_license(license_path)
        if not os.path.exists(ArtifactsDir):
            os.makedirs(ArtifactsDir)