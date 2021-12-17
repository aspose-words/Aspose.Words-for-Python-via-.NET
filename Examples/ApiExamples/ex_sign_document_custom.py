# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import unittest
import io
import uuid
from typing import List

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, IMAGE_DIR

class ExSignDocumentCustom(ApiExampleBase):

    #ExStart
    #ExFor:CertificateHolder
    #ExFor:SignatureLineOptions.signer
    #ExFor:SignatureLineOptions.signer_title
    #ExFor:SignatureLine.id
    #ExFor:SignOptions.signature_line_id
    #ExFor:SignOptions.signature_line_image
    #ExFor:DigitalSignatureUtil.sign(str,str,CertificateHolder,SignOptions)
    #ExSummary:Shows how to add a signature line to a document, and then sign it using a digital certificate.
    def test_sign(self):

        signee_name = "Ron Williams"
        src_document_path = MY_DIR + "Document.docx"
        dst_document_path = ARTIFACTS_DIR + "SignDocumentCustom.sign.docx"
        certificate_path = MY_DIR + "morzal.pfx"
        certificate_password = "aw"

        for signee_info in self._create_signees():
            if signee_info.name == signee_name:
                self._sign_document(src_document_path, dst_document_path, signee_info, certificate_path, certificate_password)
                break
        else:
            raise Exception("Signee does not exist.")

    def _sign_document(self, src_document_path: str, dst_document_path: str,
        signee_info, certificate_path: str, certificate_password: str):
        """Creates a copy of a source document signed using provided signee information and X509 certificate."""

        document = aw.Document(src_document_path)
        builder = aw.DocumentBuilder(document)

        # Configure and insert a signature line, an object in the document that will display a signature that we sign it with.
        signature_line_options = aw.SignatureLineOptions()
        signature_line_options.signer = signee_info.name
        signature_line_options.signer_title = signee_info.position

        signature_line = builder.insert_signature_line(signature_line_options).signature_line
        signature_line.id = signee_info.person_id

        # First, we will save an unsigned version of our document.
        builder.document.save(dst_document_path)

        certificate_holder = aw.digitalsignatures.CertificateHolder.create(certificate_path, certificate_password)

        sign_options = aw.digitalsignatures.SignOptions()
        sign_options.signature_line_id = signee_info.person_id
        sign_options.signature_line_image = signee_info.image

        # Overwrite the unsigned document we saved above with a version signed using the certificate.
        aw.digitalsignatures.DigitalSignatureUtil.sign(dst_document_path, dst_document_path, certificate_holder, sign_options)

    def _image_to_byte_array(self, image_in: drawing.Image) -> bytes:
        """Converts an image to a byte array."""

        with io.BytesIO() as stream:
            image_in.save(stream, drawing.imaging.ImageFormat.png)
            return bytes(stream.getbuffer())

    class Signee:

        def __init__(self, guid: uuid.UUID, name: str, position: str, image: bytes):
            self.person_id = guid
            self.name = name
            self.position = position
            self.image = image

    def _create_signees(self):

        return [
            ExSignDocumentCustom.Signee(uuid.uuid4(), "Ron Williams", "Chief Executive Officer",
                self._image_to_byte_array(drawing.Image.from_file(IMAGE_DIR + "Logo.jpg"))),
            ExSignDocumentCustom.Signee(uuid.uuid4(), "Stephen Morse", "Head of Compliance",
                self._image_to_byte_array(drawing.Image.from_file(IMAGE_DIR + "Logo.jpg")))
        ]
    #ExEnd
