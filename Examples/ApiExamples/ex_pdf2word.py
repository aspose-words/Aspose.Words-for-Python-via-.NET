# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw

from api_example_base import ApiExampleBase, ARTIFACTS_DIR

class ExPdf2Word(ApiExampleBase):

    def test_load_pdf(self):

        #ExStart
        #ExFor:Document.__init__(str)
        #ExSummary:Shows how to load a PDF.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Hello world!")

        doc.save(ARTIFACTS_DIR + "PDF2Word.load_pdf.pdf")

        # Below are two ways of loading PDF documents using Aspose products.
        # 1 -  Load as an Aspose.Words document:
        aspose_words_doc = aw.Document(ARTIFACTS_DIR + "PDF2Word.load_pdf.pdf")

        self.assertEqual("Hello world!", aspose_words_doc.get_text().strip())

        # 2 -  Load as an Aspose.Pdf document:
        aspose_pdf_doc = aw.Document(ARTIFACTS_DIR + "PDF2Word.load_pdf.pdf")

        text_fragment_absorber = aspose.pdf.text.TextFragmentAbsorber()
        aspose_pdf_doc.pages.accept(text_fragment_absorber)

        self.assertEqual("Hello world!", text_fragment_absorber.text.strip())
        #ExEnd

    @staticmethod
    def convert_pdf_to_docx():

        #ExStart
        #ExFor:Document.__init__(str)
        #ExFor:Document.save(str)
        #ExSummary:Shows how to convert a PDF to a .docx.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Hello world!")

        doc.save(ARTIFACTS_DIR + "PDF2Word.convert_pdf_to_docx.pdf")

        # Load the PDF document that we just saved, and convert it to .docx.
        pdf_doc = aw.Document(ARTIFACTS_DIR + "PDF2Word.convert_pdf_to_docx.pdf")

        pdf_doc.save(ARTIFACTS_DIR + "PDF2Word.convert_pdf_to_docx.docx")
        #ExEnd

    @staticmethod
    def convert_pdf_to_docx_custom():

        #ExStart
        #ExFor:Document.save(str,SaveOptions)
        #ExSummary:Shows how to convert a PDF to a .docx and customize the saving process with a SaveOptions object.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello world!")

        doc.save(ARTIFACTS_DIR + "PDF2Word.convert_pdf_to_docx_custom.pdf")

        # Load the PDF document that we just saved, and convert it to .docx.
        pdf_doc = aw.Document(ARTIFACTS_DIR + "PDF2Word.convert_pdf_to_docx_custom.pdf")

        save_options = aw.saving.OoxmlSaveOptions(aw.SaveFormat.DOCX)

        # Set the "password" property to encrypt the saved document with a password.
        save_options.password = "MyPassword"

        pdf_doc.save(ARTIFACTS_DIR + "PDF2Word.convert_pdf_to_docx_custom.docx", save_options)
        #ExEnd

    def load_pdf_using_plugin(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Hello world!")

        doc.save(ARTIFACTS_DIR + "PDF2Word.load_pdf_using_plugin.pdf")

        # Use the Pdf2Word plugin to open load a PDF document as an Aspose.Words document.
        pdf_doc = aw.Document()

        pdf2word = aw.pdf2word.PdfDocumentReaderPlugin()
        with open(ARTIFACTS_DIR + "PDF2Word.load_pdf_using_plugin.pdf", "rb") as stream:
            pdf2word.read(stream, aw.LoadOptions(), pdf_doc)

        builder = aw.DocumentBuilder(pdf_doc)

        builder.move_to_document_end()
        builder.writeln(" We are editing a PDF document that was loaded into Aspose.Words!")

        self.assertEqual("Hello world! We are editing a PDF document that was loaded into Aspose.Words!", pdf_doc.get_text().strip())

    def load_encrypted_pdf_using_plugin(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello world! This is an encrypted PDF document.")

        # Configure a SaveOptions object to encrypt this PDF document while saving it to the local file system.
        encryption_details = aw.saving.PdfEncryptionDetails("MyPassword", "", aw.saving.PdfEncryptionAlgorithm.RC4_128)

        self.assertEqual(aw.saving.PdfPermissions.DISALLOW_ALL, encryption_details.permissions)

        save_options = aw.saving.PdfSaveOptions()
        save_options.encryption_details = encryption_details

        doc.save(ARTIFACTS_DIR + "PDF2Word.load_encrypted_pdf_using_plugin.pdf", save_options)

        pdf_doc = aw.Document()

        # To load a password encrypted document, we need to pass a LoadOptions object
        # with the correct password stored in its "password" property.
        load_options = aw.loading.LoadOptions()
        load_options.password = "MyPassword"

        pdf2word = aw.pdf2word.PdfDocumentReaderPlugin()
        with open(ARTIFACTS_DIR + "PDF2Word.load_encrypted_pdf_using_plugin.pdf", "rb") as stream:
            # Pass the LoadOptions object into the Pdf2Word plugin's "read" method
            # the same way we would pass it into a document's "load" method.
            pdf2word.read(stream, aw.loading.LoadOptions("MyPassword"), pdf_doc)

        self.assertEqual("Hello world! This is an encrypted PDF document.", pdf_doc.get_text().strip())
