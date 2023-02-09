# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw
import time

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExMetered(ApiExampleBase):

    def test_test_metered_usage(self):

        with self.assertRaises(Exception):
            ExMetered.usage()

    @staticmethod
    def usage():

        #ExStart
        #ExFor:Metered
        #ExFor:Metered.__init__
        #ExFor:Metered.get_consumption_credit
        #ExFor:Metered.get_consumption_quantity
        #ExFor:Metered.set_metered_key(str,str)
        #ExSummary:Shows how to activate a Metered license and track credit/consumption.
        # Create a new Metered license, and then print its usage statistics.
        metered = aw.Metered()
        metered.set_metered_key("MyPublicKey", "MyPrivateKey")

        print("Credit before operation:", metered.get_consumption_credit())
        print("Consumption quantity before operation:", metered.get_consumption_quantity())

        # Operate using Aspose.Words, and then print our metered stats again to see how much we spent.
        doc = aw.Document(MY_DIR + "Document.docx")
        doc.save(ARTIFACTS_DIR + "Metered.usage.pdf")

        # Aspose Metered Licensing mechanism does not send the usage data to purchase server every time,
        # you need to use waiting.
        time.sleep(10)

        print("Credit after operation:", metered.get_consumption_credit())
        print("Consumption quantity after operation:", metered.get_consumption_quantity())
        #ExEnd
