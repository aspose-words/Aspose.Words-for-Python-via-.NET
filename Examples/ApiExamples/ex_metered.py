import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class ExMetered(ApiExampleBase):

    def test_test_metered_usage(self):

        with self.assertRaises(Exception):
            ExMetered.usage()

    @staticmethod
    def usage():

        #ExStart
        #ExFor:Metered
        #ExFor:Metered.#ctor
        #ExFor:Metered.GetConsumptionCredit
        #ExFor:Metered.GetConsumptionQuantity
        #ExFor:Metered.SetMeteredKey(String, String)
        #ExSummary:Shows how to activate a Metered license and track credit/consumption.
        # Create a new Metered license, and then print its usage statistics.
        metered = aw.Metered()
        metered.set_metered_key("MyPublicKey", "MyPrivateKey")

        print(f"Credit before operation: {metered.get_consumption_credit()}")
        print(f"Consumption quantity before operation: {metered.get_consumption_quantity()}")

        # Operate using Aspose.Words, and then print our metered stats again to see how much we spent.
        doc = aw.Document(MY_DIR + "Document.docx")
        doc.save(ARTIFACTS_DIR + "Metered.usage.pdf")

        print(f"Credit after operation: {metered.get_consumption_credit()}")
        print(f"Consumption quantity after operation: {metered.get_consumption_quantity()}")
        #ExEnd
