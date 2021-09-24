import unittest
import os
import sys
import array
import uuid
from datetime import date, datetime

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithDigitalSinatures(docs_base.DocsExamplesBase):
    
    def test_sign_document(self) :
        
        #ExStart:SingDocument
        certHolder = aw.digitalsignatures.CertificateHolder.create(docs_base.my_dir + "morzal.pfx", "aw")
            
        aw.digitalsignatures.DigitalSignatureUtil.sign(docs_base.my_dir + "Digitally signed.docx", docs_base.artifacts_dir + "Document.signed.docx", certHolder)
        #ExEnd:SingDocument
        

    def test_signing_encrypted_document(self) :
        
        #ExStart:SigningEncryptedDocument
        signOptions = aw.digitalsignatures.SignOptions()
        signOptions.decryption_password = "decryptionPassword" 

        certHolder = aw.digitalsignatures.CertificateHolder.create(docs_base.my_dir + "morzal.pfx", "aw")
            
        aw.digitalsignatures.DigitalSignatureUtil.sign(docs_base.my_dir + "Digitally signed.docx", docs_base.artifacts_dir + "Document.encrypted_document.docx",
            certHolder, signOptions)
        #ExEnd:SigningEncryptedDocument
        

    def test_creating_and_signing_new_signature_line(self) :
        
        #ExStart:CreatingAndSigningNewSignatureLine
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        signatureLine = builder.insert_signature_line(aw.SignatureLineOptions()).signature_line
            
        doc.save(docs_base.artifacts_dir + "SignDocuments.signature_line.docx")

        signOptions = aw.digitalsignatures.SignOptions()
            
        signOptions.signature_line_id = signatureLine.id
        with open(docs_base.images_dir + "Enhanced Windows MetaFile.emf", "rb") as image_file:
            signOptions.signature_line_image = image_file.read()
            

        certHolder = aw.digitalsignatures.CertificateHolder.create(docs_base.my_dir + "morzal.pfx", "aw")
            
        aw.digitalsignatures.DigitalSignatureUtil.sign(docs_base.artifacts_dir + "SignDocuments.signature_line.docx",
            docs_base.artifacts_dir + "SignDocuments.new_signature_line.docx", certHolder, signOptions)
        #ExEnd:CreatingAndSigningNewSignatureLine
        

    def test_signing_existing_signature_line(self) :
        
        #ExStart:SigningExistingSignatureLine
        doc = aw.Document(docs_base.my_dir + "Signature line.docx")
            
        signatureLine = doc.first_section.body.get_child(aw.NodeType.SHAPE, 0, True).as_shape().signature_line

        signOptions = aw.digitalsignatures.SignOptions()
            
        signOptions.signature_line_id = signatureLine.id

        imagefile = open(docs_base.images_dir + "Enhanced Windows MetaFile.emf", "rb")

        with open(docs_base.images_dir + "Enhanced Windows MetaFile.emf", "rb") as image_file:
            signOptions.signature_line_image = image_file.read()
            

        certHolder = aw.digitalsignatures.CertificateHolder.create(docs_base.my_dir + "morzal.pfx", "aw")
            
        aw.digitalsignatures.DigitalSignatureUtil.sign(docs_base.my_dir + "Digitally signed.docx",
            docs_base.artifacts_dir + "SignDocuments.signing_existing_signature_line.docx", certHolder, signOptions)
        #ExEnd:SigningExistingSignatureLine
        

    def test_set_signature_provider_id(self) :
        
        #ExStart:SetSignatureProviderID
        doc = aw.Document(docs_base.my_dir + "Signature line.docx")

        signatureLine = doc.first_section.body.get_child(aw.NodeType.SHAPE, 0, True).as_shape().signature_line

        signOptions = aw.digitalsignatures.SignOptions()
            
        signOptions.provider_id = signatureLine.provider_id
        signOptions.signature_line_id = signatureLine.id
            

        certHolder = aw.digitalsignatures.CertificateHolder.create(docs_base.my_dir + "morzal.pfx", "aw")

        aw.digitalsignatures.DigitalSignatureUtil.sign(docs_base.my_dir + "Digitally signed.docx",
            docs_base.artifacts_dir + "SignDocuments.set_signature_provider_id.docx", certHolder, signOptions)
        #ExEnd:SetSignatureProviderID
        

    def test_create_new_signature_line_and_set_provider_id(self) :
        
        #ExStart:CreateNewSignatureLineAndSetProviderID
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        signatureLineOptions = aw.SignatureLineOptions()
            
        signatureLineOptions.signer = "vderyushev"
        signatureLineOptions.signer_title = "QA"
        signatureLineOptions.email = "vderyushev@aspose.com"
        signatureLineOptions.show_date = True
        signatureLineOptions.default_instructions = False
        signatureLineOptions.instructions = "Please sign here."
        signatureLineOptions.allow_comments = True
            

        signatureLine = builder.insert_signature_line(signatureLineOptions).signature_line
        signatureLine.provider_id = uuid.UUID('{CF5A7BB4-8F3C-4756-9DF6-BEF7F13259A2}')
            
        doc.save(docs_base.artifacts_dir + "SignDocuments.signature_line_provider_id.docx")

        signOptions = aw.digitalsignatures.SignOptions()
            
        signOptions.signature_line_id = signatureLine.id
        signOptions.provider_id = signatureLine.provider_id
        signOptions.comments = "Document was signed by vderyushev"
        signOptions.sign_time = datetime.today()
            

        certHolder = aw.digitalsignatures.CertificateHolder.create(docs_base.my_dir + "morzal.pfx", "aw")

        aw.digitalsignatures.DigitalSignatureUtil.sign(docs_base.artifacts_dir + "SignDocuments.signature_line_provider_id.docx", 
            docs_base.artifacts_dir + "SignDocuments.create_new_signature_line_and_set_provider_id.docx", certHolder, signOptions)
        #ExEnd:CreateNewSignatureLineAndSetProviderID
        

    def test_access_and_verify_signature(self) :
        
        #ExStart:AccessAndVerifySignature
        doc = aw.Document(docs_base.my_dir + "Digitally signed.docx")

        for signature in doc.digital_signatures :
            
            print("*** Signature Found ***")
            print("Is valid: " + str(signature.is_valid))
            # This property is available in MS Word documents only.
            print("Reason for signing: " + signature.comments) 
            print("Time of signing: " + str(signature.sign_time))
            #print("Subject name: " + signature.certificate_holder.certificate.subject_name.name)
            #print("Issuer name: " + signature.certificate_holder.certificate.issuer_name.name)
            print()
            
        #ExEnd:AccessAndVerifySignature
        
    

if __name__ == '__main__':
    unittest.main()