# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw

from api_example_base import ApiExampleBase, ARTIFACTS_DIR

class ExPlainTextDocument(ApiExampleBase):

    def test_load(self):

        #ExStart
        #ExFor:PlainTextDocument
        #ExFor:PlainTextDocument.__init__(str)
        #ExFor:PlainTextDocument.text
        #ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world!")

        doc.save(ARTIFACTS_DIR + "PlainTextDocument.load.docx")

        plaintext = aw.PlainTextDocument(ARTIFACTS_DIR + "PlainTextDocument.load.docx")

        self.assertEqual("Hello world!", plaintext.text.strip())
        #ExEnd

    def test_load_from_stream(self):

        #ExStart
        #ExFor:PlainTextDocument.__init__(BytesIO)
        #ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext using stream.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello world!")
        doc.save(ARTIFACTS_DIR + "PlainTextDocument.load_from_stream.docx")

        with open(ARTIFACTS_DIR + "PlainTextDocument.load_from_stream.docx", "rb") as stream:

            plaintext = aw.PlainTextDocument(stream)

            self.assertEqual("Hello world!", plaintext.text.strip())

        #ExEnd

    def test_load_encrypted(self):

        #ExStart
        #ExFor:PlainTextDocument.__init__(str,LoadOptions)
        #ExSummary:Shows how to load the contents of an encrypted Microsoft Word document in plaintext.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world!")

        save_options = aw.saving.OoxmlSaveOptions()
        save_options.password = "MyPassword"

        doc.save(ARTIFACTS_DIR + "PlainTextDocument.load_encrypted.docx", save_options)

        load_options = aw.loading.LoadOptions()
        load_options.password = "MyPassword"

        plaintext = aw.PlainTextDocument(ARTIFACTS_DIR + "PlainTextDocument.load_encrypted.docx", load_options)

        self.assertEqual("Hello world!", plaintext.text.strip())
        #ExEnd

    def test_load_encrypted_using_stream(self):

        #ExStart
        #ExFor:PlainTextDocument.__init__(BytesIO,LoadOptions)
        #ExSummary:Shows how to load the contents of an encrypted Microsoft Word document in plaintext using stream.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello world!")

        save_options = aw.saving.OoxmlSaveOptions()
        save_options.password = "MyPassword"

        doc.save(ARTIFACTS_DIR + "PlainTextDocument.load_encrypted_using_stream.docx", save_options)

        load_options = aw.loading.LoadOptions()
        load_options.password = "MyPassword"

        with open(ARTIFACTS_DIR + "PlainTextDocument.load_encrypted_using_stream.docx", "rb") as stream:

            plaintext = aw.PlainTextDocument(stream, load_options)

            self.assertEqual("Hello world!", plaintext.text.strip())

        #ExEnd

    def test_built_in_properties(self):

        #ExStart
        #ExFor:PlainTextDocument.built_in_document_properties
        #ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext and then access the original document's built-in properties.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello world!")
        doc.built_in_document_properties.author = "John Doe"

        doc.save(ARTIFACTS_DIR + "PlainTextDocument.built_in_properties.docx")

        plaintext = aw.PlainTextDocument(ARTIFACTS_DIR + "PlainTextDocument.built_in_properties.docx")

        self.assertEqual("Hello world!", plaintext.text.strip())
        self.assertEqual("John Doe", plaintext.built_in_document_properties.author)
        #ExEnd

    def test_custom_document_properties(self):

        #ExStart
        #ExFor:PlainTextDocument.custom_document_properties
        #ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext and then access the original document's custom properties.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello world!")
        doc.custom_document_properties.add("Location of writing", "123 Main St, London, UK")

        doc.save(ARTIFACTS_DIR + "PlainTextDocument.custom_document_properties.docx")

        plaintext = aw.PlainTextDocument(ARTIFACTS_DIR + "PlainTextDocument.custom_document_properties.docx")

        self.assertEqual("Hello world!", plaintext.text.strip())
        self.assertEqual("123 Main St, London, UK", plaintext.custom_document_properties.get_by_name("Location of writing").value)
        #ExEnd
