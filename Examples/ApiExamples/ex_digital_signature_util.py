# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from aspose.words.digitalsignatures import DigitalSignatureUtil
import aspose.words as aw
import aspose.words.digitalsignatures
import aspose.words.loading
import datetime
import system_helper
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExDigitalSignatureUtil(ApiExampleBase):

    def test_load(self):
        #ExStart
        #ExFor:DigitalSignatureUtil
        #ExFor:DigitalSignatureUtil.load_signatures(str)
        #ExFor:DigitalSignatureUtil.load_signatures(BytesIO)
        #ExSummary:Shows how to load signatures from a digitally signed document.
        # There are two ways of loading a signed document's collection of digital signatures using the DigitalSignatureUtil class.
        # 1 -  Load from a document from a local file system filename:
        digital_signatures = aw.digitalsignatures.DigitalSignatureUtil.load_signatures(file_name=MY_DIR + 'Digitally signed.docx')
        # If this collection is nonempty, then we can verify that the document is digitally signed.
        self.assertEqual(1, digital_signatures.count)
        # 2 -  Load from a document from a FileStream:
        with system_helper.io.FileStream(MY_DIR + 'Digitally signed.docx', system_helper.io.FileMode.OPEN) as stream:
            digital_signatures = aw.digitalsignatures.DigitalSignatureUtil.load_signatures(stream=stream)
            self.assertEqual(1, digital_signatures.count)
        #ExEnd

    def test_remove(self):
        #ExStart
        #ExFor:DigitalSignatureUtil
        #ExFor:DigitalSignatureUtil.load_signatures(str)
        #ExFor:DigitalSignatureUtil.remove_all_signatures(BytesIO,BytesIO)
        #ExFor:DigitalSignatureUtil.remove_all_signatures(str,str)
        #ExSummary:Shows how to remove digital signatures from a digitally signed document.
        # There are two ways of using the DigitalSignatureUtil class to remove digital signatures
        # from a signed document by saving an unsigned copy of it somewhere else in the local file system.
        # 1 - Determine the locations of both the signed document and the unsigned copy by filename strings:
        aw.digitalsignatures.DigitalSignatureUtil.remove_all_signatures(src_file_name=MY_DIR + 'Digitally signed.docx', dst_file_name=ARTIFACTS_DIR + 'DigitalSignatureUtil.LoadAndRemove.FromString.docx')
        # 2 - Determine the locations of both the signed document and the unsigned copy by file streams:
        with system_helper.io.FileStream(MY_DIR + 'Digitally signed.docx', system_helper.io.FileMode.OPEN) as stream_in:
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'DigitalSignatureUtil.LoadAndRemove.FromStream.docx', system_helper.io.FileMode.CREATE) as stream_out:
                aw.digitalsignatures.DigitalSignatureUtil.remove_all_signatures(src_stream=stream_in, dst_stream=stream_out)
        # Verify that both our output documents have no digital signatures.
        self.assertEqual(0, aw.digitalsignatures.DigitalSignatureUtil.load_signatures(file_name=ARTIFACTS_DIR + 'DigitalSignatureUtil.LoadAndRemove.FromString.docx').count)
        self.assertEqual(0, aw.digitalsignatures.DigitalSignatureUtil.load_signatures(file_name=ARTIFACTS_DIR + 'DigitalSignatureUtil.LoadAndRemove.FromStream.docx').count)
        #ExEnd

    def test_remove_signatures(self):
        aw.digitalsignatures.DigitalSignatureUtil.remove_all_signatures(src_file_name=MY_DIR + 'Digitally signed.odt', dst_file_name=ARTIFACTS_DIR + 'DigitalSignatureUtil.RemoveSignatures.odt')
        self.assertEqual(0, aw.digitalsignatures.DigitalSignatureUtil.load_signatures(file_name=ARTIFACTS_DIR + 'DigitalSignatureUtil.RemoveSignatures.odt').count)

    @unittest.skip('Discrepancy in assertion between Python and .Net')
    def test_sign_document(self):
        #ExStart
        #ExFor:CertificateHolder
        #ExFor:CertificateHolder.create(str,str)
        #ExFor:DigitalSignatureUtil.sign(BytesIO,BytesIO,CertificateHolder,SignOptions)
        #ExFor:DigitalSignatures.sign_options
        #ExFor:SignOptions.comments
        #ExFor:SignOptions.sign_time
        #ExSummary:Shows how to digitally sign documents.
        # Create an X.509 certificate from a PKCS#12 store, which should contain a private key.
        certificate_holder = aw.digitalsignatures.CertificateHolder.create(file_name=MY_DIR + 'morzal.pfx', password='aw')
        # Create a comment and date which will be applied with our new digital signature.
        sign_options = aw.digitalsignatures.SignOptions()
        sign_options.comments = 'My comment'
        sign_options.sign_time = datetime.datetime.now()
        # Take an unsigned document from the local file system via a file stream,
        # then create a signed copy of it determined by the filename of the output file stream.
        with system_helper.io.FileStream(MY_DIR + 'Document.docx', system_helper.io.FileMode.OPEN) as stream_in:
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'DigitalSignatureUtil.SignDocument.docx', system_helper.io.FileMode.OPEN_OR_CREATE) as stream_out:
                aw.digitalsignatures.DigitalSignatureUtil.sign(src_stream=stream_in, dst_stream=stream_out, cert_holder=certificate_holder, sign_options=sign_options)
        #ExEnd
        with system_helper.io.FileStream(ARTIFACTS_DIR + 'DigitalSignatureUtil.SignDocument.docx', system_helper.io.FileMode.OPEN) as stream:
            digital_signatures = aw.digitalsignatures.DigitalSignatureUtil.load_signatures(stream=stream)
            self.assertEqual(1, digital_signatures.count)
            signature = digital_signatures[0]
            self.assertTrue(signature.is_valid)
            self.assertEqual(aw.digitalsignatures.DigitalSignatureType.XML_DSIG, signature.signature_type)
            self.assertEqual(str(sign_options.sign_time), str(signature.sign_time))
            self.assertEqual('My comment', signature.comments)

    def test_decryption_password(self):
        #ExStart
        #ExFor:CertificateHolder
        #ExFor:SignOptions.decryption_password
        #ExFor:LoadOptions.password
        #ExSummary:Shows how to sign encrypted document file.
        # Create an X.509 certificate from a PKCS#12 store, which should contain a private key.
        certificate_holder = aw.digitalsignatures.CertificateHolder.create(file_name=MY_DIR + 'morzal.pfx', password='aw')
        # Create a comment, date, and decryption password which will be applied with our new digital signature.
        sign_options = aw.digitalsignatures.SignOptions()
        sign_options.comments = 'Comment'
        sign_options.sign_time = datetime.datetime.now()
        sign_options.decryption_password = 'docPassword'
        # Set a local system filename for the unsigned input document, and an output filename for its new digitally signed copy.
        input_file_name = MY_DIR + 'Encrypted.docx'
        output_file_name = ARTIFACTS_DIR + 'DigitalSignatureUtil.DecryptionPassword.docx'
        aw.digitalsignatures.DigitalSignatureUtil.sign(src_file_name=input_file_name, dst_file_name=output_file_name, cert_holder=certificate_holder, sign_options=sign_options)
        #ExEnd
        # Open encrypted document from a file.
        load_options = aw.loading.LoadOptions(password='docPassword')
        self.assertEqual(sign_options.decryption_password, load_options.password)
        # Check that encrypted document was successfully signed.
        signed_doc = aw.Document(file_name=output_file_name, load_options=load_options)
        signatures = signed_doc.digital_signatures
        self.assertEqual(1, signatures.count)
        self.assertTrue(signatures.is_valid)

    def test_sign_document_obfuscation_bug(self):
        ch = aw.digitalsignatures.CertificateHolder.create(file_name=MY_DIR + 'morzal.pfx', password='aw')
        doc = aw.Document(file_name=MY_DIR + 'Structured document tags.docx')
        output_file_name = ARTIFACTS_DIR + 'DigitalSignatureUtil.SignDocumentObfuscationBug.doc'
        sign_options = aw.digitalsignatures.SignOptions()
        sign_options.comments = 'Comment'
        sign_options.sign_time = datetime.datetime.now()
        aw.digitalsignatures.DigitalSignatureUtil.sign(src_file_name=doc.original_file_name, dst_file_name=output_file_name, cert_holder=ch, sign_options=sign_options)

    def test_incorrect_decryption_password(self):
        certificate_holder = aw.digitalsignatures.CertificateHolder.create(file_name=MY_DIR + 'morzal.pfx', password='aw')
        doc = aw.Document(file_name=MY_DIR + 'Encrypted.docx', load_options=aw.loading.LoadOptions(password='docPassword'))
        output_file_name = ARTIFACTS_DIR + 'DigitalSignatureUtil.IncorrectDecryptionPassword.docx'
        sign_options = aw.digitalsignatures.SignOptions()
        sign_options.comments = 'Comment'
        sign_options.sign_time = datetime.datetime.now()
        sign_options.decryption_password = 'docPassword1'
        self.assertRaises(Exception, lambda: aw.digitalsignatures.DigitalSignatureUtil.sign(src_file_name=doc.original_file_name, dst_file_name=output_file_name, cert_holder=certificate_holder, sign_options=sign_options))

    def test_no_arguments_for_sing(self):
        sign_options = aw.digitalsignatures.SignOptions()
        sign_options.comments = ''
        sign_options.sign_time = datetime.datetime.now()
        sign_options.decryption_password = ''
        self.assertRaises(Exception, lambda: aw.digitalsignatures.DigitalSignatureUtil.sign(src_file_name='', dst_file_name='', cert_holder=None, sign_options=sign_options))

    def test_no_certificate_for_sign(self):
        doc = aw.Document(file_name=MY_DIR + 'Digitally signed.docx')
        output_file_name = ARTIFACTS_DIR + 'DigitalSignatureUtil.NoCertificateForSign.docx'
        sign_options = aw.digitalsignatures.SignOptions()
        sign_options.comments = 'Comment'
        sign_options.sign_time = datetime.datetime.now()
        sign_options.decryption_password = 'docPassword'
        self.assertRaises(Exception, lambda: aw.digitalsignatures.DigitalSignatureUtil.sign(src_file_name=doc.original_file_name, dst_file_name=output_file_name, cert_holder=None, sign_options=sign_options))

    def test_xml_dsig(self):
        #ExStart:XmlDsig
        #ExFor:SignOptions.xml_dsig_level
        #ExFor:XmlDsigLevel
        #ExSummary:Shows how to sign document based on XML-DSig standard.
        certificate_holder = aw.digitalsignatures.CertificateHolder.create(file_name=MY_DIR + 'morzal.pfx', password='aw')
        sign_options = aw.digitalsignatures.SignOptions()
        sign_options.xml_dsig_level = aw.digitalsignatures.XmlDsigLevel.X_AD_ES_EPES
        input_file_name = MY_DIR + 'Document.docx'
        output_file_name = ARTIFACTS_DIR + 'DigitalSignatureUtil.XmlDsig.docx'
        aw.digitalsignatures.DigitalSignatureUtil.sign(src_file_name=input_file_name, dst_file_name=output_file_name, cert_holder=certificate_holder, sign_options=sign_options)
        #ExEnd:XmlDsig

    def test_sign_document_with_options(self):
        #ExStart:SignDocumentWithOptions
        #ExFor:SignOptions.windows_version
        #ExFor:SignOptions.application_version
        #ExFor:SignOptions.office_version
        #ExFor:SignOptions.horizontal_resolution
        #ExFor:SignOptions.vertical_resolution
        #ExFor:SignOptions.color_depth
        #ExFor:DigitalSignatureUtil.sign(str,str,CertificateHolder,SignOptions)
        #ExSummary:Shows how to sign a document with additional signing options.
        sign_options = aw.digitalsignatures.SignOptions()
        sign_options.windows_version = '10.0'
        sign_options.application_version = '16.0.19127'
        sign_options.office_version = '16.0.19127/27'
        sign_options.horizontal_resolution = 1024
        sign_options.vertical_resolution = 768
        sign_options.color_depth = 24
        cert_bytes = system_helper.io.File.read_all_bytes(MY_DIR + 'morzal.pfx')
        cert = aw.digitalsignatures.CertificateHolder.create(cert_bytes=cert_bytes, password='aw')
        aw.digitalsignatures.DigitalSignatureUtil.sign(src_file_name=MY_DIR + 'Digitally signed.docx', dst_file_name=ARTIFACTS_DIR + 'DigitalSignatureUtil.docx', cert_holder=cert, sign_options=sign_options)
        signed_doc = aw.Document(file_name=ARTIFACTS_DIR + 'DigitalSignatureUtil.docx')
        signature = signed_doc.digital_signatures[0]
        self.assertEqual(1, signed_doc.digital_signatures.count)
        self.assertTrue(signature.is_valid)
        self.assertEqual('10.0', signature.windows_version)
        self.assertEqual('16.0.19127', signature.application_version)
        self.assertEqual('16.0.19127/27', signature.office_version)
        self.assertEqual(1024, signature.horizontal_resolution)
        self.assertEqual(768, signature.vertical_resolution)
        self.assertEqual(24, signature.color_depth)
        #ExEnd:SignDocumentWithOptions