import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class ExComHelper(ApiExampleBase):

    def test_com_helper(self):

        #ExStart
        #ExFor:ComHelper
        #ExFor:ComHelper.__init__
        #ExFor:ComHelper.open(BytesIO)
        #ExFor:ComHelper.open(str)
        #ExSummary:Shows how to open documents using the ComHelper class.
        # The ComHelper class allows us to load documents from within COM clients.
        com_helper = aw.ComHelper()

        # 1 -  Using a local system filename:
        doc = com_helper.open(MY_DIR + "Document.docx")

        self.assertEqual("Hello World!\r\rHello Word!\r\r\rHello World!", doc.get_text().strip())

        # 2 -  From a stream:
        with open(MY_DIR + "Document.docx", "rb") as stream:
            doc = com_helper.open(stream)
            self.assertEqual("Hello World!\r\rHello Word!\r\r\rHello World!", doc.get_text().strip())

        #ExEnd
