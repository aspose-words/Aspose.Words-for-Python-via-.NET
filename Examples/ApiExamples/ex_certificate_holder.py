# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import unittest
from api_example_base import ApiExampleBase

class ExCertificateHolder(ApiExampleBase):

    @unittest.skip('Unknown types: NetworkCredential and Pkcs12Store')
    def test_create(self):
        #ExStart
        #ExFor:CertificateHolder.create(bytes,SecureString)
        #ExFor:CertificateHolder.create(bytes,str)
        #ExFor:CertificateHolder.create(str,str,str)
        #ExSummary:Shows how to create CertificateHolder objects.
        # Below are four ways of creating CertificateHolder objects.
        # 1 -  Load a PKCS #12 file into a byte array and apply its password:
        with open(MY_DIR + 'morzal.pfx', 'rb') as file:
            cert_bytes = file.read()
        aw.digitalsignatures.CertificateHolder.create(cert_bytes, 'aw')
        # 2 -  Load a PKCS #12 file into a byte array, and apply a secure password:
        password = NetworkCredential('', 'aw').secure_password
        aw.digitalsignatures.CertificateHolder.create(cert_bytes, password)
        # If the certificate has private keys corresponding to aliases,
        # we can use the aliases to fetch their respective keys. First, we will check for valid aliases.
        with open(MY_DIR + 'morzal.pfx', 'rb') as cert_stream:
            pkcs12_store = Pkcs12Store(cert_stream, 'aw').build()
            pkcs12_store.load(cert_stream, 'aw')
            for alias in pkcs12_store.aliases:
                if pkcs12_store.is_key_entry(alias) and pkcs12_store.get_key(alias).key.is_private:
                    print('Valid alias found:', alias)
        # 3 -  Use a valid alias:
        aw.digitalsignatures.CertificateHolder.create(MY_DIR + 'morzal.pfx', 'aw', 'c20be521-11ea-4976-81ed-865fbbfc9f24')
        # 4 -  Pass "null" as the alias in order to use the first available alias that returns a private key:
        aw.digitalsignatures.CertificateHolder.create(MY_DIR + 'morzal.pfx', 'aw', None)
        #ExEnd