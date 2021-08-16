import unittest

import aspose.words as aw

import api_example_base as aeb


class ExComHelper(aeb.ApiExampleBase):

    def test_com_helper(self):

        #ExStart
        #ExFor:ComHelper
        #ExFor:ComHelper.#ctor
        #ExFor:ComHelper.open(Stream)
        #ExFor:ComHelper.open(String)
        #ExSummary:Shows how to open documents using the ComHelper class.
        # The ComHelper class allows us to load documents from within COM clients.
        com_helper = aw.ComHelper()

        # 1 -  Using a local system filename:
        doc = com_helper.open(aeb.my_dir + "Document.docx")

        self.assertEqual("Hello World!\r\rHello Word!\r\r\rHello World!", doc.get_text().strip())

        ### Streams are not allowed in current version of aw for python
        # 2 -  From a stream:
        # using (FileStream stream = new FileStream(MyDir + "Document.docx", FileMode.open))
        #
        #     doc = comHelper.open(stream)
        #
        #     self.assertEqual("Hello World!\r\rHello Word!\r\r\rHello World!", doc.get_text().strip())

        #ExEnd


if __name__ == '__main__':
    unittest.main()
