# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import unittest
import aspose.words as aw
import os
import platform

ROOT_DIR = os.path.abspath(os.curdir) + "/"
ROOT_DIR = ROOT_DIR[:ROOT_DIR.find("Aspose.Words-for-Python-via-.NET")]
API_EXAMPLES_ROOT = ROOT_DIR + "Aspose.Words-for-Python-via-.NET/Examples/"
LICENSE_PATH = os.getenv("ASPOSE_WORDS_PYTHON_LICENSE", "Aspose.Words.Python.NET.lic")
MY_DIR = API_EXAMPLES_ROOT + "Data/"
ARTIFACTS_DIR = MY_DIR + "Artifacts/"
GOLDS_DIR = MY_DIR + "Golds/"
TEMP_DIR = MY_DIR + "Temp/"
IMAGE_DIR = MY_DIR + "Images/"
FONTS_DIR = MY_DIR + "MyFonts/"
ASPOSE_LOGO_URL = "https://www.aspose.cloud/templates/aspose/App_Themes/V3/images/words/header/aspose_words-for-net.png"


class ApiExampleBase(unittest.TestCase):

    def setUp(self):
        if os.path.exists(ARTIFACTS_DIR):
            l = aw.License()
            l.set_license(LICENSE_PATH)
        if not os.path.exists(ARTIFACTS_DIR):
            os.makedirs(ARTIFACTS_DIR)
