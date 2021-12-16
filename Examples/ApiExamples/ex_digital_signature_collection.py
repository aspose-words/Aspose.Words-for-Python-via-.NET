# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExDigitalSignatureCollection(ApiExampleBase):

    def test_get_enumerator(self):

        #ExStart
        #ExFor:DigitalSignatureCollection.__iter__
        #ExSummary:Shows how to print all the digital signatures of a signed document.
        digital_signatures = aw.digitalsignatures.DigitalSignatureUtil.load_signatures(MY_DIR + "Digitally signed.docx")

        for ds in digital_signatures:
            print(ds)

        #ExEnd

        self.assertEqual(1, digital_signatures.count)

        signature = digital_signatures[0]

        self.assertTrue(signature.is_valid)
        self.assertEqual(aw.digitalsignatures.DigitalSignatureType.XML_DSIG, signature.signature_type)
        self.assertEqual("12/23/2010 02:14:40 AM", signature.sign_time.strftime("%m/%d/%Y %H:%M:%S %p"))
        self.assertEqual("Test Sign", signature.comments)

        # System.Security.Cryptography.X509Certificates.X509Certificate2 is not supported. That is why the following information is not accesible.
        #self.assertEqual(signature.issuer_name, signature.certificate_holder.certificate.issuer_name.name)
        #self.assertEqual(signature.subject_name, signature.certificate_holder.certificate.subject_name.name)

        self.assertEqual("CN=VeriSign Class 3 Code Signing 2009-2 CA, " +
            "OU=Terms of use at https://www.verisign.com/rpa (c)09, " +
            "OU=VeriSign Trust Network, " +
            "O=\"VeriSign, Inc.\", " +
            "C=US", signature.issuer_name)

        self.assertEqual("CN=Aspose Pty Ltd, " +
            "OU=Digital ID Class 3 - Microsoft Software Validation v2, " +
            "O=Aspose Pty Ltd, " +
            "L=Lane Cove, " +
            "S=New South Wales, " +
            "C=AU", signature.subject_name)
