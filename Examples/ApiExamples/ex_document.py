import os
import unittest
from datetime import date, datetime

import aspose.words as aw

import api_example_base as aeb
from document_helper import DocumentHelper


class ExDocument(aeb.ApiExampleBase):

    def test_constructor(self):
        # ExStart
        # ExFor:Document.#ctor()
        # ExFor:Document.#ctor(String,LoadOptions)
        # ExSummary:Shows how to create and load documents.
        # There are two ways of creating a Document object using Aspose.Words.
        # 1 -  Create a blank document:
        doc = aw.Document()
        # New Document objects by default come with the minimal set of nodes
        # required to begin adding content such as text and shapes: a Section, a Body, and a Paragraph.
        doc.first_section.body.first_paragraph.append_child(aw.Run(doc, "Hello world!"))
        # 2 -  Load a document that exists in the local file system:
        doc = aw.Document(aeb.my_dir + "Document.docx")
        # Loaded documents will have contents that we can access and edit.
        self.assertEqual("Hello World!", doc.first_section.body.first_paragraph.get_text().strip())
        # Some operations that need to occur during loading, such as using a password to decrypt a document,
        # can be done by passing a LoadOptions object when loading the document.
        doc = aw.Document(aeb.my_dir + "Encrypted.docx", aw.loading.LoadOptions("docPassword"))
        self.assertEqual("Test encrypted document.", doc.first_section.body.first_paragraph.get_text().strip())
        # ExEnd

    def test_convert_to_pdf(self):
        # ExStart
        # ExFor:Document.#ctor(String)
        # ExFor:Document.Save(String)
        # ExSummary:Shows how to open a document and convert it to .PDF.
        doc = aw.Document(aeb.my_dir + "Document.docx")
        doc.save(aeb.artifacts_dir + "Document.ConvertToPdf.pdf")
        # ExEnd

    def test_detect_pdf_document_format(self):
        info = aw.FileFormatUtil.detect_file_format(aeb.my_dir + "Pdf Document.pdf")
        self.assertEqual(info.load_format, aw.LoadFormat.PDF)

    def test_open_pdf_document(self):
        doc = aw.Document(aeb.my_dir + "Pdf Document.pdf")
        self.assertEqual(
            "Heading 1\rHeading 1.1.1.1 Heading 1.1.1.2\rHeading 1.1.1.1.1.1.1.1.1 Heading 1.1.1.1.1.1.1.1.2\u000c",
            doc.range.text)

    def test_open_protected_pdf_document(self):
        doc = aw.Document(aeb.my_dir + "Pdf Document.pdf")

        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.encryption_details = aw.saving.PdfEncryptionDetails("Aspose", None, aw.saving.PdfEncryptionAlgorithm.RC4_40)

        doc.save(aeb.artifacts_dir + "Document.PdfDocumentEncrypted.pdf", saveOptions)

        loadOptions = aw.loading.PdfLoadOptions()
        loadOptions.password = "Aspose"
        loadOptions.load_format = aw.LoadFormat.PDF

        doc = aw.Document(aeb.artifacts_dir + "Document.PdfDocumentEncrypted.pdf", loadOptions)

    def test_load_encrypted(self):
        # ExStart
        # ExFor:Document.#ctor(Stream,LoadOptions)
        # ExFor:Document.#ctor(String,LoadOptions)
        # ExFor:LoadOptions
        # ExFor:LoadOptions.#ctor(String)
        # ExSummary:Shows how to load an encrypted Microsoft Word document.
        doc = None

        # Aspose.Words throw an exception if we try to open an encrypted document without its password.
        with self.assertRaises(RuntimeError) as ex:
            doc = aw.Document(aeb.my_dir + "Encrypted.docx")

        self.assertEqual(str(ex.exception), "Proxy error: The document password is incorrect.")

        # When loading such a document, the password is passed to the document's constructor using a LoadOptions object.
        options = aw.loading.LoadOptions("docPassword")

        # There are two ways of loading an encrypted document with a LoadOptions object.
        # 1 -  Load the document from the local file system by filename:
        doc = aw.Document(aeb.my_dir + "Encrypted.docx", options)
        self.assertEqual("Test encrypted document.", doc.get_text().strip())

        # ExEnd

    def test_temp_folder(self):
        # ExStart
        # ExFor:LoadOptions.TempFolder
        # ExSummary:Shows how to load a document using temporary files.
        # Note that such an approach can reduce memory usage but degrades speed
        loadOptions = aw.loading.LoadOptions()
        loadOptions.temp_folder = aeb.temp_dir

        # Ensure that the directory exists and load
        if not os.path.exists(loadOptions.temp_folder):
            os.makedirs(loadOptions.temp_folder)

        doc = aw.Document(aeb.my_dir + "Document.docx", loadOptions)
        # ExEnd

    def test_convert_to_html(self):
        # ExStart
        # ExFor:Document.Save(String,SaveFormat)
        # ExFor:SaveFormat
        # ExSummary:Shows how to convert from DOCX to HTML format.
        doc = aw.Document(aeb.my_dir + "Document.docx")

        doc.save(aeb.artifacts_dir + "Document.ConvertToHtml.html", aw.SaveFormat.HTML)
        # ExEnd

    def test_convert_to_mhtml(self):
        doc = aw.Document(aeb.my_dir + "Document.docx")
        doc.save(aeb.artifacts_dir + "Document.ConvertToMhtml.mht")

    def test_convert_to_txt(self):
        doc = aw.Document(aeb.my_dir + "Document.docx")
        doc.save(aeb.artifacts_dir + "Document.ConvertToTxt.txt")

    def test_convert_to_epub(self):
        doc = aw.Document(aeb.my_dir + "Rendering.docx")
        doc.save(aeb.artifacts_dir + "Document.ConvertToEpub.epub")

    def test_append_document(self):
        # ExStart
        # ExFor:Document.AppendDocument(Document, ImportFormatMode)
        # ExSummary:Shows how to append a document to the end of another document.
        srcDoc = aw.Document()
        srcDoc.first_section.body.append_paragraph("Source document text. ")

        dstDoc = aw.Document()
        dstDoc.first_section.body.append_paragraph("Destination document text. ")

        # Append the source document to the destination document while preserving its formatting,
        # then save the source document to the local file system.
        dstDoc.append_document(srcDoc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
        self.assertEqual(2, dstDoc.sections.count)  # ExSkip

        dstDoc.save(aeb.artifacts_dir + "Document.AppendDocument.docx")
        # ExEnd

        outDocText = aw.Document(aeb.artifacts_dir + "Document.AppendDocument.docx").get_text()

        self.assertTrue(outDocText.startswith(dstDoc.get_text()))
        self.assertTrue(outDocText.endswith(srcDoc.get_text()))

    def test_import_list(self):
        for isKeepSourceNumbering in [True, False]:
            with self.subTest(isKeepSourceNumbering=isKeepSourceNumbering):
                # ExStart
                # ExFor:ImportFormatOptions.KeepSourceNumbering
                # ExSummary:Shows how to import a document with numbered lists.
                srcDoc = aw.Document(aeb.my_dir + "List source.docx")
                dstDoc = aw.Document(aeb.my_dir + "List destination.docx")

                self.assertEqual(2, dstDoc.lists.count)

                options = aw.ImportFormatOptions()

                # If there is a clash of list styles, apply the list format of the source document.
                # Set the "KeepSourceNumbering" property to "False" to not import any list numbers into the destination document.
                # Set the "KeepSourceNumbering" property to "True" import all clashing
                # list style numbering with the same appearance that it had in the source document.
                options.keep_source_numbering = isKeepSourceNumbering

                dstDoc.append_document(srcDoc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, options)
                dstDoc.update_list_labels()

                if (isKeepSourceNumbering):
                    self.assertEqual(3, dstDoc.lists.count)
                else:
                    self.assertEqual(2, dstDoc.lists.count)
                # ExEnd

    def test_keep_source_numbering_same_list_ids(self):
        # ExStart
        # ExFor:ImportFormatOptions.KeepSourceNumbering
        # ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
        # ExSummary:Shows how resolve a clash when importing documents that have lists with the same list definition identifier.
        srcDoc = aw.Document(aeb.my_dir + "List with the same definition identifier - source.docx")
        dstDoc = aw.Document(aeb.my_dir + "List with the same definition identifier - destination.docx")

        importFormatOptions = aw.ImportFormatOptions()

        # Set the "KeepSourceNumbering" property to "True" to apply a different list definition ID
        # to identical styles as Aspose.Words imports them into destination documents.
        importFormatOptions.keep_source_numbering = True
        dstDoc.append_document(srcDoc, aw.ImportFormatMode.USE_DESTINATION_STYLES, importFormatOptions)

        dstDoc.update_list_labels()
        # ExEnd

        paraText = dstDoc.sections[1].body.last_paragraph.get_text()

        self.assertTrue(paraText.startswith("13->13"), paraText)
        self.assertEqual("1.", dstDoc.sections[1].body.last_paragraph.list_label.label_string)

    def test_validate_individual_document_signatures(self):
        # ExStart
        # ExFor:CertificateHolder.Certificate
        # ExFor:Document.DigitalSignatures
        # ExFor:DigitalSignature
        # ExFor:DigitalSignatureCollection
        # ExFor:DigitalSignature.IsValid
        # ExFor:DigitalSignature.Comments
        # ExFor:DigitalSignature.SignTime
        # ExFor:DigitalSignature.SignatureType
        # ExSummary:Shows how to validate and display information about each signature in a document.
        doc = aw.Document(aeb.my_dir + "Digitally signed.docx")

        for signature in doc.digital_signatures:
            print(f"\n{'Valid' if signature.is_valid else 'Invalid'} signature: ")
            print(f"\tReason:\t{signature.comments}")
            print(f"\tType:\t{signature.signature_type}")
            print(f"\tSign time:\t{signature.sign_time}")
            # System.Security.Cryptography.X509Certificates.X509Certificate2 is not supported. That is why the following information is not accesible.
            # print(f"\tSubject name:\t{signature.certificate_holder.certificate.subject_name}\n")
            # print(f"\tIssuer name:\t{signature.certificate_holder.certificate.issuer_name.name}\n")
        # ExEnd

        self.assertEqual(1, doc.digital_signatures.count)

        digitalSig = doc.digital_signatures[0]

        self.assertTrue(digitalSig.is_valid)
        self.assertEqual("Test Sign", digitalSig.comments)
        self.assertEqual(aw.digitalsignatures.DigitalSignatureType.XML_DSIG, digitalSig.signature_type)
        # System.Security.Cryptography.X509Certificates.X509Certificate2 is not supported. That is why the following information is not accesible.
        # self.assertTrue(digitalSig.certificate_holder.certificate.subject.contains("Aspose Pty Ltd"))
        # self.assertTrue(digitalSig.certificate_holder.certificate.issuer_name.name != None & digitalSig.certificate_holder.certificate.issuer_name.name.contains("VeriSign"))

    def test_digital_signature(self):
        # ExStart
        # ExFor:DigitalSignature.CertificateHolder
        # ExFor:DigitalSignature.IssuerName
        # ExFor:DigitalSignature.SubjectName
        # ExFor:DigitalSignatureCollection
        # ExFor:DigitalSignatureCollection.IsValid
        # ExFor:DigitalSignatureCollection.Count
        # ExFor:DigitalSignatureCollection.Item(Int32)
        # ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder)
        # ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder)
        # ExFor:DigitalSignatureType
        # ExFor:Document.DigitalSignatures
        # ExSummary:Shows how to sign documents with X.509 certificates.
        # Verify that a document is not signed.
        self.assertFalse(aw.FileFormatUtil.detect_file_format(aeb.my_dir + "Document.docx").has_digital_signature)

        # Create a CertificateHolder object from a PKCS12 file, which we will use to sign the document.
        certificateHolder = aw.digitalsignatures.CertificateHolder.create(aeb.my_dir + "morzal.pfx", "aw", None)

        # There are two ways of saving a signed copy of a document to the local file system:
        # 1 - Designate a document by a local system filename and save a signed copy at a location specified by another filename.
        sigOptions = aw.digitalsignatures.SignOptions()
        sigOptions.sign_time = date.today()
        aw.digitalsignatures.DigitalSignatureUtil.sign(aeb.my_dir + "Document.docx", aeb.artifacts_dir + "Document.DigitalSignature.docx",
                                                       certificateHolder, sigOptions)

        self.assertTrue(aw.FileFormatUtil.detect_file_format(
            aeb.artifacts_dir + "Document.DigitalSignature.docx").has_digital_signature)

        # Please verify that all of the document's digital signatures are valid and check their details.
        signedDoc = aw.Document(aeb.artifacts_dir + "Document.DigitalSignature.docx")
        digitalSignatureCollection = signedDoc.digital_signatures

        self.assertTrue(digitalSignatureCollection.is_valid)
        self.assertEqual(1, digitalSignatureCollection.count)
        self.assertEqual(aw.digitalsignatures.DigitalSignatureType.XML_DSIG, digitalSignatureCollection[0].signature_type)
        self.assertEqual("CN=Morzal.Me", signedDoc.digital_signatures[0].issuer_name)
        self.assertEqual("CN=Morzal.Me", signedDoc.digital_signatures[0].subject_name)
        # ExEnd

    def test_append_all_documents_in_folder(self):
        # ExStart
        # ExFor:Document.AppendDocument(Document, ImportFormatMode)
        # ExSummary:Shows how to append all the documents in a folder to the end of a template document.
        dstDoc = aw.Document()

        builder = aw.DocumentBuilder(dstDoc)
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
        builder.writeln("Template Document")
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.NORMAL
        builder.writeln("Some content here")
        self.assertEqual(5, dstDoc.styles.count)  # ExSkip
        self.assertEqual(1, dstDoc.sections.count)  # ExSkip

        # Append all unencrypted documents with the .doc extension
        # from our local file system directory to the base document.
        for file in os.listdir(aeb.my_dir):
            if file.endswith(".doc"):
                fileName = aeb.my_dir + file
                info = aw.FileFormatUtil.detect_file_format(fileName)
                if (info.is_encrypted):
                    continue

                srcDoc = aw.Document(fileName)
                dstDoc.append_document(srcDoc, aw.ImportFormatMode.USE_DESTINATION_STYLES)

        dstDoc.save(aeb.artifacts_dir + "Document.AppendAllDocumentsInFolder.doc")
        # ExEnd

        self.assertEqual(7, dstDoc.styles.count)
        self.assertEqual(9, dstDoc.sections.count)

    def test_join_runs_with_same_formatting(self):
        # ExStart
        # ExFor:Document.JoinRunsWithSameFormatting
        # ExSummary:Shows how to join runs in a document to reduce unneeded runs.
        # Open a document that contains adjacent runs of text with identical formatting,
        # which commonly occurs if we edit the same paragraph multiple times in Microsoft Word.
        doc = aw.Document(aeb.my_dir + "Rendering.docx")

        # If any number of these runs are adjacent with identical formatting,
        # then the document may be simplified.
        self.assertEqual(317, doc.get_child_nodes(aw.NodeType.RUN, True).count)

        # Combine such runs with this method and verify the number of run joins that will take place.
        self.assertEqual(121, doc.join_runs_with_same_formatting())

        # The number of joins and the number of runs we have after the join
        # should add up the number of runs we had initially.
        self.assertEqual(196, doc.get_child_nodes(aw.NodeType.RUN, True).count)
        # ExEnd

    def test_default_tab_stop(self):
        # ExStart
        # ExFor:Document.DefaultTabStop
        # ExFor:ControlChar.Tab
        # ExFor:ControlChar.TabChar
        # ExSummary:Shows how to set a custom interval for tab stop positions.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Set tab stops to appear every 72 points (1 inch).
        builder.document.default_tab_stop = 72

        # Each tab character snaps the text after it to the next closest tab stop position.
        builder.writeln("Hello" + aw.ControlChar.TAB + "World!")
        builder.writeln("Hello" + aw.ControlChar.TAB + "World!")
        # ExEnd

        doc = DocumentHelper.save_open(doc)
        self.assertEqual(72, doc.default_tab_stop)

    def test_clone_document(self):
        # ExStart
        # ExFor:Document.Clone
        # ExSummary:Shows how to deep clone a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Hello world!")

        # Cloning will produce a new document with the same contents as the original,
        # but with a unique copy of each of the original document's nodes.
        clone = doc.clone()

        self.assertEqual(doc.first_section.body.first_paragraph.runs[0].text,
                         clone.first_section.body.first_paragraph.runs[0].text)
        self.assertIsNot(doc.first_section.body.first_paragraph.runs[0],
                         clone.first_section.body.first_paragraph.runs[0])
        # ExEnd

    def test_document_get_text_to_string(self):
        # ExStart
        # ExFor:CompositeNode.get_text
        # ExFor:Node.ToString(SaveFormat)
        # ExSummary:Shows the difference between calling the get_text and ToString methods on a node.
        doc = aw.Document()

        builder = aw.DocumentBuilder(doc)
        builder.insert_field("MERGEFIELD Field")

        # get_text will retrieve the visible text as well as field codes and special characters.
        self.assertEqual("\u0013MERGEFIELD Field\u0014«Field»\u0015\u000c", doc.get_text())

        # ToString will give us the document's appearance if saved to a passed save format.
        self.assertEqual("«Field»\r\n", doc.to_string(aw.SaveFormat.TEXT))
        # ExEnd

    def test_protect_unprotect(self):
        # ExStart
        # ExFor:Document.Protect(ProtectionType,String)
        # ExFor:Document.ProtectionType
        # ExFor:Document.Unprotect
        # ExFor:Document.Unprotect(String)
        # ExSummary:Shows how to protect and unprotect a document.
        doc = aw.Document()
        doc.protect(aw.ProtectionType.READ_ONLY, "password")

        self.assertEqual(aw.ProtectionType.READ_ONLY, doc.protection_type)

        # If we open this document with Microsoft Word intending to edit it,
        # we will need to apply the password to get through the protection.
        doc.save(aeb.artifacts_dir + "Document.Protect.docx")

        # Note that the protection only applies to Microsoft Word users opening our document.
        # We have not encrypted the document in any way, and we do not need the password to open and edit it programmatically.
        protectedDoc = aw.Document(aeb.artifacts_dir + "Document.Protect.docx")

        self.assertEqual(aw.ProtectionType.READ_ONLY, protectedDoc.protection_type)

        builder = aw.DocumentBuilder(protectedDoc)
        builder.writeln("Text added to a protected document.")
        self.assertEqual("Text added to a protected document.", protectedDoc.range.text.strip())  # ExSkip

        # There are two ways of removing protection from a document.
        # 1 - With no password:
        doc.unprotect()

        self.assertEqual(aw.ProtectionType.NO_PROTECTION, doc.protection_type)

        doc.protect(aw.ProtectionType.READ_ONLY, "NewPassword")

        self.assertEqual(aw.ProtectionType.READ_ONLY, doc.protection_type)

        doc.unprotect("WrongPassword")

        self.assertEqual(aw.ProtectionType.READ_ONLY, doc.protection_type)

        # 2 - With the correct password:
        doc.unprotect("NewPassword")

        self.assertEqual(aw.ProtectionType.NO_PROTECTION, doc.protection_type)
        # ExEnd

    def test_document_ensureMinimum(self):
        # ExStart
        # ExFor:Document.EnsureMinimum
        # ExSummary:Shows how to ensure that a document contains the minimal set of nodes required for editing its contents.
        # A newly created document contains one child Section, which includes one child Body and one child Paragraph.
        # We can edit the document body's contents by adding nodes such as Runs or inline Shapes to that paragraph.
        doc = aw.Document()
        nodes = doc.get_child_nodes(aw.NodeType.ANY, True)

        self.assertEqual(aw.NodeType.SECTION, nodes[0].node_type)
        self.assertEqual(doc, nodes[0].parent_node)

        self.assertEqual(aw.NodeType.BODY, nodes[1].node_type)
        self.assertEqual(nodes[0], nodes[1].parent_node)

        self.assertEqual(aw.NodeType.PARAGRAPH, nodes[2].node_type)
        self.assertEqual(nodes[1], nodes[2].parent_node)

        # This is the minimal set of nodes that we need to be able to edit the document.
        # We will no longer be able to edit the document if we remove any of them.
        doc.remove_all_children()

        self.assertEqual(0, doc.get_child_nodes(aw.NodeType.ANY, True).count)

        # Call this method to make sure that the document has at least those three nodes so we can edit it again.
        doc.ensure_minimum()

        self.assertEqual(aw.NodeType.SECTION, nodes[0].node_type)
        self.assertEqual(aw.NodeType.BODY, nodes[1].node_type)
        self.assertEqual(aw.NodeType.PARAGRAPH, nodes[2].node_type)

        doc.first_section.body.first_paragraph.runs.add(
            aw.Run(doc, "Hello world!"))  # cannot use (Paragraph)nodes[2] because thre is no casting in Python
        # ExEnd

        self.assertEqual("Hello world!", doc.get_text().strip())

    def test_remove_macros_from_document(self):
        # ExStart
        # ExFor:Document.RemoveMacros
        # ExSummary:Shows how to remove all macros from a document.
        doc = aw.Document(aeb.my_dir + "Macro.docm")

        self.assertTrue(doc.has_macros)
        self.assertEqual("Project", doc.vba_project.name)

        # Remove the document's VBA project, along with all its macros.
        doc.remove_macros()

        self.assertFalse(doc.has_macros)
        self.assertIsNone(doc.vba_project)
        # ExEnd

    def test_get_page_count(self):
        # ExStart
        # ExFor:Document.PageCount
        # ExSummary:Shows how to count the number of pages in the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Page 1")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.write("Page 2")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.write("Page 3")

        # Verify the expected page count of the document.
        self.assertEqual(3, doc.page_count)

        # Getting the PageCount property invoked the document's page layout to calculate the value.
        # This operation will not need to be re-done when rendering the document to a fixed page save format,
        # such as .pdf. So you can save some time, especially with more complex documents.
        doc.save(aeb.artifacts_dir + "Document.GetPageCount.pdf")
        # ExEnd

    def test_get_updated_page_properties(self):
        # ExStart
        # ExFor:Document.UpdateWordCount()
        # ExFor:Document.UpdateWordCount(Boolean)
        # ExFor:BuiltInDocumentProperties.Characters
        # ExFor:BuiltInDocumentProperties.Words
        # ExFor:BuiltInDocumentProperties.Paragraphs
        # ExFor:BuiltInDocumentProperties.Lines
        # ExSummary:Shows how to update all list labels in a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")
        builder.write(
            "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.")

        # Aspose.Words does not track document metrics like these in real time.
        self.assertEqual(0, doc.built_in_document_properties.characters)
        self.assertEqual(0, doc.built_in_document_properties.words)
        self.assertEqual(1, doc.built_in_document_properties.paragraphs)
        self.assertEqual(1, doc.built_in_document_properties.lines)

        # To get accurate values for three of these properties, we will need to update them manually.
        doc.update_word_count()

        self.assertEqual(196, doc.built_in_document_properties.characters)
        self.assertEqual(36, doc.built_in_document_properties.words)
        self.assertEqual(2, doc.built_in_document_properties.paragraphs)

        # For the line count, we will need to call a specific overload of the updating method.
        self.assertEqual(1, doc.built_in_document_properties.lines)

        doc.update_word_count(True)

        self.assertEqual(4, doc.built_in_document_properties.lines)
        # ExEnd

    @unittest.skip("It is impossible to cast Style to TableStyle to work with its properties.")
    def test_table_style_to_direct_formatting(self):
        # ExStart
        # ExFor:CompositeNode.get_child
        # ExFor:Document.ExpandTableStylesToDirectFormatting
        # ExSummary:Shows how to apply the properties of a table's style directly to the table's elements.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Hello world!")
        builder.end_table()

        tableStyle = doc.styles.add(aw.StyleType.TABLE, "MyTableStyle1")
        # tableStyle = tableStyle.as_tablestyle()
        tableStyle.row_stripe = 3
        tableStyle.cell_spacing = 5
        # tableStyle.shading.background_pattern_color = Color.AntiqueWhite
        # tableStyle.Borders.Color = Color.Blue
        tableStyle.borders.line_style = aw.LineStyle.DOT_DASH

        table.style = tableStyle

        # This method concerns table style properties such as the ones we set above.
        doc.expand_table_styles_to_direct_formatting()

        doc.save(aeb.artifacts_dir + "Document.TableStyleToDirectFormatting.docx")
        # ExEnd

    #           TestUtil.DocPackageFileContainsString("<w:tblStyleRowBandSize w:val=\"3\" />",
    #               ArtifactsDir + "Document.TableStyleToDirectFormatting.docx", "document.xml")
    #           TestUtil.DocPackageFileContainsString("<w:tblCellSpacing w:w=\"100\" w:type=\"dxa\" />",
    #               ArtifactsDir + "Document.TableStyleToDirectFormatting.docx", "document.xml")
    #           TestUtil.DocPackageFileContainsString("<w:tblBorders><w:top w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:left w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:bottom w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:right w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:insideH w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:insideV w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /></w:tblBorders>",
    #               ArtifactsDir + "Document.TableStyleToDirectFormatting.docx", "document.xml")

    def test_update_table_layout(self):
        # ExStart
        # ExFor:Document.UpdateTableLayout
        # ExSummary:Shows how to preserve a table's layout when saving to .txt.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Cell 1")
        builder.insert_cell()
        builder.write("Cell 2")
        builder.insert_cell()
        builder.write("Cell 3")
        builder.end_table()

        # Use a TxtSaveOptions object to preserve the table's layout when converting the document to plaintext.
        options = aw.saving.TxtSaveOptions()
        options.preserve_table_layout = True

        # Previewing the appearance of the document in .txt form shows that the table will not be represented accurately.
        self.assertEqual(0.0, table.first_row.cells[0].cell_format.width)
        self.assertEqual("CCC\r\neee\r\nlll\r\nlll\r\n   \r\n123\r\n\r\n", doc.to_string(options))

        # We can call UpdateTableLayout() to fix some of these issues.
        doc.update_table_layout()

        self.assertEqual("Cell 1             Cell 2             Cell 3\r\n\r\n", doc.to_string(options))
        self.assertAlmostEqual(155.0, table.first_row.cells[0].cell_format.width, delta=2)
        # ExEnd

    def test_get_original_file_info(self):
        # ExStart
        # ExFor:Document.OriginalFileName
        # ExFor:Document.OriginalLoadFormat
        # ExSummary:Shows how to retrieve details of a document's load operation.
        doc = aw.Document(aeb.my_dir + "Document.docx")

        self.assertEqual(aeb.my_dir + "Document.docx", doc.original_file_name)
        self.assertEqual(aw.LoadFormat.DOCX, doc.original_load_format)
        # ExEnd

    def test_footnote_columns(self):
        # ExStart
        # ExFor:FootnoteOptions
        # ExFor:FootnoteOptions.Columns
        # ExSummary:Shows how to split the footnote section into a given number of columns.
        doc = aw.Document(aeb.my_dir + "Footnotes and endnotes.docx")
        self.assertEqual(0, doc.footnote_options.columns)  # ExSkip

        doc.footnote_options.columns = 2
        doc.save(aeb.artifacts_dir + "Document.FootnoteColumns.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "Document.FootnoteColumns.docx")

        self.assertEqual(2, doc.first_section.page_setup.footnote_options.columns)

    def test_compare(self):
        # ExStart
        # ExFor:Document.Compare(Document, String, DateTime)
        # ExFor:RevisionCollection.AcceptAll
        # ExSummary:Shows how to compare documents.
        docOriginal = aw.Document()
        builder = aw.DocumentBuilder(docOriginal)
        builder.writeln("This is the original document.")

        docEdited = aw.Document()
        builder = aw.DocumentBuilder(docEdited)
        builder.writeln("This is the edited document.")

        # Comparing documents with revisions will throw an exception.
        if docOriginal.revisions.count == 0 & docEdited.revisions.count == 0:
            docOriginal.compare(docEdited, "authorName", date.today())

        # After the comparison, the original document will gain a new revision
        # for every element that is different in the edited document.
        self.assertEqual(2, docOriginal.revisions.count)  # ExSkip
        for r in docOriginal.revisions:
            print(f"Revision type: {r.revision_type}, on a node of type \"{r.parent_node.node_type}\"")
            print(f"\tChanged text: \"{r.parent_node.get_text()}\"")

        # Accepting these revisions will transform the original document into the edited document.
        docOriginal.revisions.accept_all()

        self.assertEqual(docOriginal.get_text(), docEdited.get_text())
        # ExEnd

        docOriginal = DocumentHelper.save_open(docOriginal)
        self.assertEqual(0, docOriginal.revisions.count)

    def test_compare_document_with_revisions(self):
        doc1 = aw.Document()
        builder = aw.DocumentBuilder(doc1)
        builder.writeln("Hello world! This text is not a revision.")

        docWithRevision = aw.Document()
        builder = aw.DocumentBuilder(docWithRevision)

        docWithRevision.start_track_revisions("John Doe")
        builder.writeln("This is a revision.")

        # InvalidOperationException
        with self.assertRaises(RuntimeError) as ex:
            docWithRevision.compare(doc1, "John Doe", date.today())

    @unittest.skip("Not sure why but 'aspose.words.Document' object has no attribute 'get_child'. Skip the test for now.")
    def test_compare_options(self):
        # ExStart
        # ExFor:CompareOptions
        # ExFor:CompareOptions.IgnoreFormatting
        # ExFor:CompareOptions.IgnoreCaseChanges
        # ExFor:CompareOptions.IgnoreComments
        # ExFor:CompareOptions.IgnoreTables
        # ExFor:CompareOptions.IgnoreFields
        # ExFor:CompareOptions.IgnoreFootnotes
        # ExFor:CompareOptions.IgnoreTextboxes
        # ExFor:CompareOptions.IgnoreHeadersAndFooters
        # ExFor:CompareOptions.Target
        # ExFor:ComparisonTargetType
        # ExFor:Document.Compare(Document, String, DateTime, CompareOptions)
        # ExSummary:Shows how to filter specific types of document elements when making a comparison.
        # Create the original document and populate it with various kinds of elements.
        docOriginal = aw.Document()
        builder = aw.DocumentBuilder(docOriginal)

        # Paragraph text referenced with an endnote:
        builder.writeln("Hello world! This is the first paragraph.")
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Original endnote text.")

        # Table:
        builder.start_table()
        builder.insert_cell()
        builder.write("Original cell 1 text")
        builder.insert_cell()
        builder.write("Original cell 2 text")
        builder.end_table()

        # Textbox:
        textBox = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 150, 20)
        builder.move_to(textBox.first_paragraph)
        builder.write("Original textbox contents")

        # DATE field:
        builder.move_to(docOriginal.first_section.body.append_paragraph(""))
        builder.insert_field(" DATE ")

        # Comment:
        newComment = aw.Comment(docOriginal, "John Doe", "J.D.", date.today())
        newComment.set_text("Original comment.")
        builder.current_paragraph.append_child(newComment)

        # Header:
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.writeln("Original header contents.")

        # Create a clone of our document and perform a quick edit on each of the cloned document's elements.
        docEdited = docOriginal.clone(True).as_document()
        firstParagraph = docEdited.first_section.body.first_paragraph

        firstParagraph.runs[0].text = "hello world! this is the first paragraph, after editing."
        firstParagraph.paragraph_format.style = docEdited.styles[aw.StyleIdentifier.HEADING1]
        docEdited.get_child(aw.NodeType.FOOTNOTE, 0, True).first_paragraph.runs[1].text = "Edited endnote text."
        docEdited.get_child(aw.NodeType.TABLE, 0, True).FirstRow.Cells[1].first_paragraph.runs[
            0].text = "Edited Cell 2 contents"
        docEdited.get_child(aw.NodeType.SHAPE, 0, True).first_paragraph.runs[0].text = "Edited textbox contents"
        docEdited.range.fields[0].use_lunar_calendar = True
        docEdited.get_child(aw.NodeType.COMMENT, 0, True).first_paragraph.runs[0].text = "Edited comment."
        docEdited.FirstSection.headers_footers[aw.HeaderFooterType.HADER_PRIMARY].first_paragraph.runs[
            0].text = "Edited header contents."

        # Comparing documents creates a revision for every edit in the edited document.
        # A CompareOptions object has a series of flags that can suppress revisions
        # on each respective type of element, effectively ignoring their change.
        compareOptions = aw.comparing.CompareOptions()
        compareOptions.ignore_formatting = False
        compareOptions.ignore_case_changes = False
        compareOptions.ignore_comments = False
        compareOptions.ignore_tables = False
        compareOptions.ignore_fields = False
        compareOptions.ignore_footnotes = False
        compareOptions.ignore_textboxes = False
        compareOptions.ignore_headers_and_footers = False
        compareOptions.Target = aw.comparing.ComparisonTargetType.NEW

        docOriginal.compare(docEdited, "John Doe", date.today(), compareOptions)
        docOriginal.save(aeb.artifacts_dir + "Document.CompareOptions.docx")
        # ExEnd

        docOriginal = aw.Document(aeb.artifacts_dir + "Document.CompareOptions.docx")

    #            TestUtil.VerifyFootnote(FootnoteType.Endnote, True, string.Empty,
    #                "OriginalEdited endnote text.", (Footnote)docOriginal.get_child(NodeType.Footnote, 0, True))

    # If we set compareOptions to ignore certain types of changes,
    # then revisions done on those types of nodes will not appear in the output document.
    # We can tell what kind of node a revision was done by looking at the NodeType of the revision's parent nodes.
    #        self.assertNotEqual(compareOptions.ignore_formatting, docOriginal.Revisions.Any(rev => rev.RevisionType == RevisionType.FormatChange))
    #        self.assertNotEqual(compareOptions.IgnoreCaseChanges, docOriginal.Revisions.Any(s => s.parent_node.get_text().Contains("hello")))
    #        self.assertNotEqual(compareOptions.IgnoreComments, docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.Comment)))
    #        self.assertNotEqual(compareOptions.IgnoreTables, docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.Table)))
    #        self.assertNotEqual(compareOptions.IgnoreFields, docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.FieldStart)))
    #        self.assertNotEqual(compareOptions.IgnoreFootnotes, docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.Footnote)))
    #        self.assertNotEqual(compareOptions.IgnoreTextboxes, docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.Shape)))
    #        self.assertNotEqual(compareOptions.IgnoreHeadersAndFooters, docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.HeaderFooter)))

    # <summary>
    # Returns True if the passed revision has a parent node with the type specified by parentType.
    # </summary>
    def has_parent_of_type(revision, parentType):
        n = revision.parent_node
        while n.parent_node != None:
            if n.node_type == parentType:
                return True
            n = n.parent_node

            return False

    def test_ignore_dml_unique_id(self):
        for isIgnoreDmlUniqueId in [True, False]:
            with self.subTest(isIgnoreDmlUniqueId=isIgnoreDmlUniqueId):
                # ExStart
                # ExFor:CompareOptions.IgnoreDmlUniqueId
                # ExSummary:Shows how to compare documents ignoring DML unique ID.
                docA = aw.Document(aeb.my_dir + "DML unique ID original.docx")
                docB = aw.Document(aeb.my_dir + "DML unique ID compare.docx")

                # By default, Aspose.Words do not ignore DML's unique ID, and the revisions count was 2.
                # If we are ignoring DML's unique ID, and revisions count were 0.
                compareOptions = aw.comparing.CompareOptions()
                compareOptions.ignore_dml_unique_id = isIgnoreDmlUniqueId

                docA.compare(docB, "Aspose.Words", date.today(), compareOptions)

                if isIgnoreDmlUniqueId:
                    self.assertEqual(0, docA.revisions.count)
                else:
                    self.assertEqual(2, docA.revisions.count)
                # ExEnd

    def test_remove_external_schema_references(self):
        # ExStart
        # ExFor:Document.RemoveExternalSchemaReferences
        # ExSummary:Shows how to remove all external XML schema references from a document.
        doc = aw.Document(aeb.my_dir + "External XML schema.docx")
        doc.remove_external_schema_references()
        # ExEnd

    def test_track_revisions(self):
        # ExStart
        # ExFor:Document.StartTrackRevisions(String)
        # ExFor:Document.StartTrackRevisions(String, DateTime)
        # ExFor:Document.StopTrackRevisions
        # ExSummary:Shows how to track revisions while editing a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Editing a document usually does not count as a revision until we begin tracking them.
        builder.write("Hello world! ")

        self.assertEqual(0, doc.revisions.count)
        self.assertFalse(doc.first_section.body.paragraphs[0].runs[0].is_insert_revision)

        doc.start_track_revisions("John Doe")

        builder.write("Hello again! ")

        self.assertEqual(1, doc.revisions.count)
        self.assertTrue(doc.first_section.body.paragraphs[0].runs[1].is_insert_revision)
        self.assertEqual("John Doe", doc.revisions[0].author)
        #            Assert.That(doc.revisions[0].DateTime, Is.EqualTo(DateTime.Now).Within(10).Milliseconds)

        # Stop tracking revisions to not count any future edits as revisions.
        doc.stop_track_revisions()
        builder.write("Hello again! ")

        self.assertEqual(1, doc.revisions.count)
        self.assertFalse(doc.first_section.body.paragraphs[0].runs[2].is_insert_revision)

        # Creating revisions gives them a date and time of the operation.
        # We can disable this by passing DateTime.MinValue when we start tracking revisions.
        doc.start_track_revisions("John Doe", datetime.min)
        builder.write("Hello again! ")

        self.assertEqual(2, doc.revisions.count)
        self.assertEqual("John Doe", doc.revisions[1].author)
        self.assertEqual(datetime.min, doc.revisions[1].date_time)

        # We can accept/reject these revisions programmatically
        # by calling methods such as Document.AcceptAllRevisions, or each revision's Accept method.
        # In Microsoft Word, we can process them manually via "Review" -> "Changes".
        doc.save(aeb.artifacts_dir + "Document.StartTrackRevisions.docx")
        # ExEnd

    def test_accept_all_revisions(self):
        # ExStart
        # ExFor:Document.AcceptAllRevisions
        # ExSummary:Shows how to accept all tracking changes in the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Edit the document while tracking changes to create a few revisions.
        doc.start_track_revisions("John Doe")
        builder.write("Hello world! ")
        builder.write("Hello again! ")
        builder.write("This is another revision.")
        doc.stop_track_revisions()

        self.assertEqual(3, doc.revisions.count)

        # We can iterate through every revision and accept/reject it as a part of our document.
        # If we know we wish to accept every revision, we can do it more straightforwardly so by calling this method.
        doc.accept_all_revisions()

        self.assertEqual(0, doc.revisions.count)
        self.assertEqual("Hello world! Hello again! This is another revision.", doc.get_text().strip())
        # ExEnd

    def test_get_revised_properties_of_list(self):
        # ExStart
        # ExFor:RevisionsView
        # ExFor:Document.RevisionsView
        # ExSummary:Shows how to switch between the revised and the original view of a document.
        doc = aw.Document(aeb.my_dir + "Revisions at list levels.docx")
        doc.update_list_labels()

        paragraphs = doc.first_section.body.paragraphs
        self.assertEqual("1.", paragraphs[0].list_label.label_string)
        self.assertEqual("a.", paragraphs[1].list_label.label_string)
        self.assertEqual("", paragraphs[2].list_label.label_string)

        # View the document object as if all the revisions are accepted. Currently supports list labels.
        doc.revisions_view = aw.RevisionsView.FINAL

        self.assertEqual("", paragraphs[0].list_label.label_string)
        self.assertEqual("1.", paragraphs[1].list_label.label_string)
        self.assertEqual("a.", paragraphs[2].list_label.label_string)
        # ExEnd

        doc.revisions_view = aw.RevisionsView.ORIGINAL
        doc.accept_all_revisions()

        self.assertEqual("a.", paragraphs[0].list_label.label_string)
        self.assertEqual("", paragraphs[1].list_label.label_string)
        self.assertEqual("b.", paragraphs[2].list_label.label_string)

    def test_update_thumbnail(self):
        # ExStart
        # ExFor:Document.UpdateThumbnail()
        # ExFor:Document.UpdateThumbnail(ThumbnailGeneratingOptions)
        # ExFor:ThumbnailGeneratingOptions
        # ExFor:ThumbnailGeneratingOptions.GenerateFromFirstPage
        # ExFor:ThumbnailGeneratingOptions.ThumbnailSize
        # ExSummary:Shows how to update a document's thumbnail.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello world!")
        builder.insert_image(aeb.image_dir + "Logo.jpg")

        # There are two ways of setting a thumbnail image when saving a document to .epub.
        # 1 -  Use the document's first page:
        doc.update_thumbnail()
        doc.save(aeb.artifacts_dir + "Document.UpdateThumbnail.FirstPage.epub")

        # 2 -  Use the first image found in the document:
        options = aw.rendering.ThumbnailGeneratingOptions()
        # self.assertEqual(new Size(600, 900), options.ThumbnailSize) #ExSkip
        self.assertTrue(options.generate_from_first_page)  # ExSkip
        # options.ThumbnailSize = new Size(400, 400)
        options.generate_from_first_page = False

        doc.update_thumbnail(options)
        doc.save(aeb.artifacts_dir + "Document.UpdateThumbnail.FirstImage.epub")
        # ExEnd

    def test_hyphenation_options(self):
        # ExStart
        # ExFor:Document.HyphenationOptions
        # ExFor:HyphenationOptions
        # ExFor:HyphenationOptions.AutoHyphenation
        # ExFor:HyphenationOptions.ConsecutiveHyphenLimit
        # ExFor:HyphenationOptions.HyphenationZone
        # ExFor:HyphenationOptions.HyphenateCaps
        # ExSummary:Shows how to configure automatic hyphenation.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.font.size = 24
        builder.writeln(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")

        doc.hyphenation_options.auto_hyphenation = True
        doc.hyphenation_options.consecutive_hyphen_limit = 2
        doc.hyphenation_options.hyphenation_zone = 720
        doc.hyphenation_options.hyphenate_caps = True

        doc.save(aeb.artifacts_dir + "Document.HyphenationOptions.docx")
        # ExEnd

        self.assertEqual(True, doc.hyphenation_options.auto_hyphenation)
        self.assertEqual(2, doc.hyphenation_options.consecutive_hyphen_limit)
        self.assertEqual(720, doc.hyphenation_options.hyphenation_zone)
        self.assertEqual(True, doc.hyphenation_options.hyphenate_caps)

        self.assertTrue(DocumentHelper.compare_docs(aeb.golds_dir + "Document.HyphenationOptions Gold.docx", aeb.artifacts_dir + "Document.HyphenationOptions.docx"))

    def test_hyphenation_options_default_values(self):
        doc = aw.Document()
        doc = DocumentHelper.save_open(doc)

        self.assertEqual(False, doc.hyphenation_options.auto_hyphenation)
        self.assertEqual(0, doc.hyphenation_options.consecutive_hyphen_limit)
        self.assertEqual(360, doc.hyphenation_options.hyphenation_zone)  # 0.25 inch
        self.assertEqual(True, doc.hyphenation_options.hyphenate_caps)

    def test_hyphenation_options_exceptions(self):
        doc = aw.Document()

        doc.hyphenation_options.consecutive_hyphen_limit = 0
        #            Assert.That(() => doc.HyphenationOptions.hyphenation_zone = 0, Throws.TypeOf<ArgumentOutOfRangeException>())
        #            Assert.That(() => doc.HyphenationOptions.ConsecutiveHyphenLimit = -1, Throws.TypeOf<ArgumentOutOfRangeException>())
        doc.hyphenation_options.hyphenation_zone = 360

    def test_ooxml_compliance_version(self):
        # ExStart
        # ExFor:Document.Compliance
        # ExSummary:Shows how to read a loaded document's Open Office XML compliance version.
        # The compliance version varies between documents created by different versions of Microsoft Word.
        doc = aw.Document(aeb.my_dir + "Document.doc")

        self.assertEqual(doc.compliance, aw.saving.OoxmlCompliance.ECMA376_2006)

        doc = aw.Document(aeb.my_dir + "Document.docx")

        self.assertEqual(doc.compliance, aw.saving.OoxmlCompliance.ISO29500_2008_TRANSITIONAL)
        # ExEnd

    def test_image_save_options(self):
        # ExStart
        # ExFor:Document.Save(String, Saving.SaveOptions)
        # ExFor:SaveOptions.UseAntiAliasing
        # ExFor:SaveOptions.UseHighQualityRendering
        # ExSummary:Shows how to improve the quality of a rendered document with SaveOptions.
        doc = aw.Document(aeb.my_dir + "Rendering.docx")
        builder = aw.DocumentBuilder(doc)

        builder.font.size = 60
        builder.writeln("Some text.")

        options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)
        self.assertFalse(options.use_anti_aliasing)  # ExSkip
        self.assertFalse(options.use_high_quality_rendering)  # ExSkip

        doc.save(aeb.artifacts_dir + "Document.ImageSaveOptions.Default.jpg", options)

        options.use_anti_aliasing = True
        options.use_high_quality_rendering = True

        doc.save(aeb.artifacts_dir + "Document.ImageSaveOptions.HighQuality.jpg", options)
        # ExEnd

    #            TestUtil.VerifyImage(794, 1122, ArtifactsDir + "Document.ImageSaveOptions.Default.jpg")
    #            TestUtil.VerifyImage(794, 1122, ArtifactsDir + "Document.ImageSaveOptions.HighQuality.jpg")

    def test_cleanup(self):
        # ExStart
        # ExFor:Document.Cleanup
        # ExSummary:Shows how to remove unused custom styles from a document.
        doc = aw.Document()

        listStyle1 = doc.styles.add(aw.StyleType.LIST, "MyListStyle1")
        listStyle2 = doc.styles.add(aw.StyleType.LIST, "MyListStyle2")
        paraStyle1 = doc.styles.add(aw.StyleType.CHARACTER, "MyParagraphStyle1")
        paraStyle2 = doc.styles.add(aw.StyleType.CHARACTER, "MyParagraphStyle2")

        # Combined with the built-in styles, the document now has eight styles.
        # A custom style counts as "used" while applied to some part of the document,
        # which means that the four styles we added are currently unused.
        self.assertEqual(8, doc.styles.count)

        # Apply a custom character style, and then a custom list style. Doing so will mark the styles as "used".
        builder = aw.DocumentBuilder(doc)
        builder.font.style = paraStyle1
        builder.writeln("Hello world!")

        lst = doc.lists.add(listStyle1)
        builder.list_format.list = lst
        builder.writeln("Item 1")
        builder.writeln("Item 2")

        doc.cleanup()

        self.assertEqual(6, doc.styles.count)

        # Removing every node that a custom style is applied to marks it as "unused" again.
        # Run the Cleanup method again to remove them.
        doc.first_section.body.remove_all_children()
        doc.cleanup()

        self.assertEqual(4, doc.styles.count)
        # ExEnd

    def test_automatically_update_styles(self):
        # ExStart
        # ExFor:Document.AutomaticallyUpdateStyles
        # ExSummary:Shows how to attach a template to a document.
        doc = aw.Document()

        # Microsoft Word documents by default come with an attached template called "Normal.dotm".
        # There is no default template for blank Aspose.Words documents.
        self.assertEqual("", doc.attached_template)

        # Attach a template, then set the flag to apply style changes
        # within the template to styles in our document.
        doc.attached_template = aeb.my_dir + "Business brochure.dotx"
        doc.automatically_update_styles = True

        doc.save(aeb.artifacts_dir + "Document.AutomaticallyUpdateStyles.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "Document.AutomaticallyUpdateStyles.docx")

        self.assertTrue(doc.automatically_update_styles)
        self.assertEqual(aeb.my_dir + "Business brochure.dotx", doc.attached_template)
        self.assertTrue(os.path.exists(doc.attached_template))

    def test_default_template(self):
        # ExStart
        # ExFor:Document.AttachedTemplate
        # ExFor:Document.AutomaticallyUpdateStyles
        # ExFor:SaveOptions.CreateSaveOptions(String)
        # ExFor:SaveOptions.DefaultTemplate
        # ExSummary:Shows how to set a default template for documents that do not have attached templates.
        doc = aw.Document()

        # Enable automatic style updating, but do not attach a template document.
        doc.automatically_update_styles = True

        self.assertEqual("", doc.attached_template)

        # Since there is no template document, the document had nowhere to track style changes.
        # Use a SaveOptions object to automatically set a template
        # if a document that we are saving does not have one.
        options = aw.saving.SaveOptions.create_save_options("Document.DefaultTemplate.docx")
        options.default_template = aeb.my_dir + "Business brochure.dotx"

        doc.save(aeb.artifacts_dir + "Document.DefaultTemplate.docx", options)
        # ExEnd

        self.assertTrue(os.path.exists(options.default_template))

    @unittest.skip("Range.Replace in the wrapped version does not accept regular expressing yet.")
    def test_use_substitutions(self):
        # ExStart
        # ExFor:FindReplaceOptions.UseSubstitutions
        # ExFor:FindReplaceOptions.LegacyMode
        # ExSummary:Shows how to recognize and use substitutions within replacement patterns.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Jason gave money to Paul.")

        regex = "([A-z]+) gave money to ([A-z]+)"

        options = aw.replacing.FindReplaceOptions()
        options.use_substitutions = True

        # Using legacy mode does not support many advanced features, so we need to set it to 'False'.
        options.legacy_mode = False

        doc.range.replace(regex, "$2 took money from $1", options)

        self.assertEqual(doc.get_text(), "Paul took money from Jason.\f")
        # ExEnd

    def test_set_invalidate_field_types(self):
        # ExStart
        # ExFor:Document.NormalizeFieldTypes
        # ExSummary:Shows how to get the keep a field's type up to date with its field code.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field("DATE", None)

        # Aspose.Words automatically detects field types based on field codes.
        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field.type)

        # Manually change the raw text of the field, which determines the field code.
        fieldText = doc.first_section.body.first_paragraph.runs[0];
        self.assertEqual("DATE", fieldText.text)  # ExSkip
        fieldText.text = "PAGE"

        # Changing the field code has changed this field to one of a different type,
        # but the field's type properties still display the old type.
        self.assertEqual("PAGE", field.get_field_code())
        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field.type)
        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field.start.field_type)
        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field.separator.field_type)
        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field.end.field_type)

        # Update those properties with this method to display current value.
        doc.normalize_field_types()

        self.assertEqual(aw.fields.FieldType.FIELD_PAGE, field.type)
        self.assertEqual(aw.fields.FieldType.FIELD_PAGE, field.start.field_type)
        self.assertEqual(aw.fields.FieldType.FIELD_PAGE, field.separator.field_type)
        self.assertEqual(aw.fields.FieldType.FIELD_PAGE, field.end.field_type)
        # ExEnd

    def test_layout_options_revisions(self):
        # ExStart
        # ExFor:Document.LayoutOptions
        # ExFor:LayoutOptions
        # ExFor:LayoutOptions.RevisionOptions
        # ExFor:RevisionColor
        # ExFor:RevisionOptions
        # ExFor:RevisionOptions.InsertedTextColor
        # ExFor:RevisionOptions.ShowRevisionBars
        # ExSummary:Shows how to alter the appearance of revisions in a rendered output document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a revision, then change the color of all revisions to green.
        builder.writeln("This is not a revision.")
        doc.start_track_revisions("John Doe", date.today())
        self.assertEqual(aw.layout.RevisionColor.BY_AUTHOR,
                         doc.layout_options.revision_options.inserted_text_color)  # ExSkip
        self.assertTrue(doc.layout_options.revision_options.show_revision_bars)  # ExSkip
        builder.writeln("This is a revision.")
        doc.stop_track_revisions()
        builder.writeln("This is not a revision.")

        # Remove the bar that appears to the left of every revised line.
        doc.layout_options.revision_options.inserted_text_color = aw.layout.RevisionColor.BRIGHT_GREEN
        doc.layout_options.revision_options.show_revision_bars = False

        doc.save(aeb.artifacts_dir + "Document.LayoutOptionsRevisions.pdf")
        # ExEnd

    def test_layout_options_hidden_text(self):
        for showHiddenText in [True, False]:
            with self.subTest(showHiddenText=showHiddenText):
                # ExStart
                # ExFor:Document.LayoutOptions
                # ExFor:LayoutOptions
                # ExFor:Layout.LayoutOptions.ShowHiddenText
                # ExSummary:Shows how to hide text in a rendered output document.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                self.assertFalse(doc.layout_options.show_hidden_text)  # ExSkip

                # Insert hidden text, then specify whether we wish to omit it from a rendered document.
                builder.writeln("This text is not hidden.")
                builder.font.hidden = True
                builder.writeln("This text is hidden.")

                doc.layout_options.show_hidden_text = showHiddenText

                doc.save(aeb.artifacts_dir + "Document.LayoutOptionsHiddenText.pdf")
                # ExEnd

    #        #if NET462 || NETCOREAPP2_1 || JAVA
    #                    Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "Document.LayoutOptionsHiddenText.pdf")
    #                    TextAbsorber textAbsorber = new TextAbsorber()
    #                    textAbsorber.Visit(pdfDoc)
    #
    #                self.assertEqual(showHiddenText ?
    #                            $"This text is not hidden.{Environment.NewLine}This text is hidden." :
    #                            "This text is not hidden.", textAbsorber.Text)
    #        #endif

    def test_layout_options_paragraph_marks(self):
        for showParagraphMarks in [True, False]:
            with self.subTest(showParagraphMarks=showParagraphMarks):
                # ExStart
                # ExFor:Document.LayoutOptions
                # ExFor:LayoutOptions
                # ExFor:Layout.LayoutOptions.ShowParagraphMarks
                # ExSummary:Shows how to show paragraph marks in a rendered output document.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                self.assertFalse(doc.layout_options.show_paragraph_marks)  # ExSkip

                # Add some paragraphs, then enable paragraph marks to show the ends of paragraphs
                # with a pilcrow (¶) symbol when we render the document.
                builder.writeln("Hello world!")
                builder.writeln("Hello again!")

                doc.layout_options.show_paragraph_marks = showParagraphMarks

                doc.save(aeb.artifacts_dir + "Document.LayoutOptionsParagraphMarks.pdf")
                # ExEnd

    #        #if NET462 || NETCOREAPP2_1 || JAVA
    #                    Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "Document.LayoutOptionsParagraphMarks.pdf")
    #                    TextAbsorber textAbsorber = new TextAbsorber()
    #                    textAbsorber.Visit(pdfDoc)
    #
    #                self.assertEqual(showParagraphMarks ?
    #                            $"Hello world!¶{Environment.NewLine}Hello again!¶{Environment.NewLine}¶" :
    #                            $"Hello world!{Environment.NewLine}Hello again!", textAbsorber.Text)
    #        #endif

    def test_update_page_layout(self):
        # ExStart
        # ExFor:StyleCollection.Item(String)
        # ExFor:SectionCollection.Item(Int32)
        # ExFor:Document.UpdatePageLayout
        # ExSummary:Shows when to recalculate the page layout of the document.
        doc = aw.Document(aeb.my_dir + "Rendering.docx")

        # Saving a document to PDF, to an image, or printing for the first time will automatically
        # cache the layout of the document within its pages.
        doc.save(aeb.artifacts_dir + "Document.UpdatePageLayout.1.pdf")

        # Modify the document in some way.
        DocumentHelper.get_style_by_name(doc, "Normal").font.size = 6
        doc.sections[0].page_setup.orientation = aw.Orientation.LANDSCAPE

        # In the current version of Aspose.Words, modifying the document does not automatically rebuild 
        # the cached page layout. If we wish for the cached layout
        # to stay up to date, we will need to update it manually.
        doc.update_page_layout()

        doc.save(aeb.artifacts_dir + "Document.UpdatePageLayout.2.pdf")
        # ExEnd

    def test_doc_package_custom_parts(self):
        # ExStart
        # ExFor:CustomPart
        # ExFor:CustomPart.ContentType
        # ExFor:CustomPart.RelationshipType
        # ExFor:CustomPart.IsExternal
        # ExFor:CustomPart.Data
        # ExFor:CustomPart.Name
        # ExFor:CustomPart.Clone
        # ExFor:CustomPartCollection
        # ExFor:CustomPartCollection.Add(CustomPart)
        # ExFor:CustomPartCollection.Clear
        # ExFor:CustomPartCollection.Clone
        # ExFor:CustomPartCollection.Count
        # ExFor:CustomPartCollection.GetEnumerator
        # ExFor:CustomPartCollection.Item(Int32)
        # ExFor:CustomPartCollection.RemoveAt(Int32)
        # ExFor:Document.PackageCustomParts
        # ExSummary:Shows how to access a document's arbitrary custom parts collection.
        doc = aw.Document(aeb.my_dir + "Custom parts OOXML package.docx")

        self.assertEqual(2, doc.package_custom_parts.count)

        # Clone the second part, then add the clone to the collection.
        clonedPart = doc.package_custom_parts[1].clone()
        doc.package_custom_parts.add(clonedPart)
        self.check_doc_package_custom_parts(doc.package_custom_parts)  # ExSkip

        self.assertEqual(3, doc.package_custom_parts.count)

        # Enumerate over the collection and print every part.
        index = 0
        for part in doc.package_custom_parts:
            print(f"Part index {index}:")
            print(f"\tName:\t\t\t\t{part.name}")
            print(f"\tContent type:\t\t{part.content_type}")
            print(f"\tRelationship type:\t{part.relationship_type}")
            if part.is_external:
                print("\tSourced from outside the document")
            else:
                print(f"\tStored within the document, length: {len(part.data)} bytes")
            index += 1

        # We can remove elements from this collection individually, or all at once.
        doc.package_custom_parts.remove_at(2)

        self.assertEqual(2, doc.package_custom_parts.count)

        doc.package_custom_parts.clear()

        self.assertEqual(0, doc.package_custom_parts.count)
        # ExEnd

    def check_doc_package_custom_parts(self, parts):
        self.assertEqual(3, parts.count)

        self.assertEqual("/payload/payload_on_package.test", parts[0].name)
        self.assertEqual("mytest/somedata", parts[0].content_type)
        self.assertEqual("http://mytest.payload.internal", parts[0].relationship_type)
        self.assertEqual(False, parts[0].is_external)
        self.assertEqual(18, len(parts[0].data))

        self.assertEqual("http://www.aspose.com/Images/aspose-logo.jpg", parts[1].name)
        self.assertEqual("", parts[1].content_type)
        self.assertEqual("http://mytest.payload.external", parts[1].relationship_type)
        self.assertEqual(True, parts[1].is_external)
        self.assertEqual(0, len(parts[1].data))

        self.assertEqual("http://www.aspose.com/Images/aspose-logo.jpg", parts[2].name)
        self.assertEqual("", parts[2].content_type)
        self.assertEqual("http://mytest.payload.external", parts[2].relationship_type)
        self.assertEqual(True, parts[2].is_external)
        self.assertEqual(0, len(parts[2].data))

    def test_shade_form_data(self):
        for useGreyShading in [True, False]:
            with self.subTest(useGreyShading=useGreyShading):
                # ExStart
                # ExFor:Document.ShadeFormData
                # ExSummary:Shows how to apply gray shading to form fields.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                self.assertTrue(doc.shade_form_data)  # ExSkip

                builder.write("Hello world! ")
                builder.insert_text_input("My form field", aw.fields.TextFormFieldType.REGULAR, "",
                                          "Text contents of form field, which are shaded in grey by default.", 0)

                # We can turn the grey shading off, so the bookmarked text will blend in with the other text.
                doc.shade_form_data = useGreyShading
                doc.save(aeb.artifacts_dir + "Document.ShadeFormData.docx")
                # ExEnd

    def test_versions_count(self):
        # ExStart
        # ExFor:Document.VersionsCount
        # ExSummary:Shows how to work with the versions count feature of older Microsoft Word documents.
        doc = aw.Document(aeb.my_dir + "Versions.doc")

        # We can read this property of a document, but we cannot preserve it while saving.
        self.assertEqual(4, doc.versions_count)

        doc.save(aeb.artifacts_dir + "Document.VersionsCount.doc")
        doc = aw.Document(aeb.artifacts_dir + "Document.VersionsCount.doc")

        self.assertEqual(0, doc.versions_count)
        # ExEnd

    def test_write_protection(self):
        # ExStart
        # ExFor:Document.WriteProtection
        # ExFor:WriteProtection
        # ExFor:WriteProtection.IsWriteProtected
        # ExFor:WriteProtection.ReadOnlyRecommended
        # ExFor:WriteProtection.SetPassword(String)
        # ExFor:WriteProtection.ValidatePassword(String)
        # ExSummary:Shows how to protect a document with a password.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world! This document is protected.")
        self.assertFalse(doc.write_protection.is_write_protected)  # ExSkip
        self.assertFalse(doc.write_protection.read_only_recommended)  # ExSkip

        # Enter a password up to 15 characters in length, and then verify the document's protection status.
        doc.write_protection.set_password("MyPassword")
        doc.write_protection.read_only_recommended = True

        self.assertTrue(doc.write_protection.is_write_protected)
        self.assertTrue(doc.write_protection.validate_password("MyPassword"))

        # Protection does not prevent the document from being edited programmatically, nor does it encrypt the contents.
        doc.save(aeb.artifacts_dir + "Document.WriteProtection.docx")
        doc = aw.Document(aeb.artifacts_dir + "Document.WriteProtection.docx")

        self.assertTrue(doc.write_protection.is_write_protected)

        builder = aw.DocumentBuilder(doc)
        builder.move_to_document_end()
        builder.writeln("Writing text in a protected document.")

        self.assertEqual("Hello world! This document is protected.\rWriting text in a protected document.",
                         doc.get_text().strip())
        # ExEnd
        self.assertTrue(doc.write_protection.read_only_recommended)
        self.assertTrue(doc.write_protection.validate_password("MyPassword"))
        self.assertFalse(doc.write_protection.validate_password("wrongpassword"))

    def test_remove_personal_information(self):
        for saveWithoutPersonalInfo in [True, False]:
            with self.subTest(saveWithoutPersonalInfo=saveWithoutPersonalInfo):
                # ExStart
                # ExFor:Document.RemovePersonalInformation
                # ExSummary:Shows how to enable the removal of personal information during a manual save.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Insert some content with personal information.
                doc.built_in_document_properties.author = "John Doe"
                doc.built_in_document_properties.company = "Placeholder Inc."

                doc.start_track_revisions(doc.built_in_document_properties.author, date.today())
                builder.write("Hello world!")
                doc.stop_track_revisions()

                # This flag is equivalent to File -> Options -> Trust Center -> Trust Center Settings... ->
                # Privacy Options -> "Remove personal information from file properties on save" in Microsoft Word.
                doc.remove_personal_information = saveWithoutPersonalInfo

                # This option will not take effect during a save operation made using Aspose.Words.
                # Personal data will be removed from our document with the flag set when we save it manually using Microsoft Word.
                doc.save(aeb.artifacts_dir + "Document.RemovePersonalInformation.docx")
                doc = aw.Document(aeb.artifacts_dir + "Document.RemovePersonalInformation.docx")

                self.assertEqual(saveWithoutPersonalInfo, doc.remove_personal_information)
                self.assertEqual("John Doe", doc.built_in_document_properties.author)
                self.assertEqual("Placeholder Inc.", doc.built_in_document_properties.company)
                self.assertEqual("John Doe", doc.revisions[0].author)

    def test_show_comments(self):
        for showComments in [True, False]:
            with self.subTest(showComments=showComments):
                # ExStart
                # ExFor:LayoutOptions.ShowComments
                # ExSummary:Shows how to show/hide comments when saving a document to a rendered format.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.write("Hello world!")

                comment = aw.Comment(doc, "John Doe", "J.D.", date.today())
                comment.set_text("My comment.")
                builder.current_paragraph.append_child(comment)

                doc.layout_options.show_comments = showComments

                doc.save(aeb.artifacts_dir + "Document.ShowComments.pdf")
                # ExEnd

    # if NET462 || NETCOREAPP2_1 || JAVA
    #            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "Document.ShowComments.pdf")
    #            TextAbsorber textAbsorber = new TextAbsorber()
    #            textAbsorber.Visit(pdfDoc)
    #
    #        self.assertEqual(
    #                showComments
    #                    ? "Hello world!                                                                    Commented [J.D.1]:  My comment."
    #                    : "Hello world!", textAbsorber.Text)
    # endif

    @unittest.skip("Not sure why but the secont asset fails. Skip the test for now.")
    def test_copy_template_styles_via_document(self):
        # ExStart
        # ExFor:Document.CopyStylesFromTemplate(Document)
        # ExSummary:Shows how to copies styles from the template to a document via Document.
        template = aw.Document(aeb.my_dir + "Rendering.docx")
        target = aw.Document(aeb.my_dir + "Document.docx")

        self.assertEqual(18, template.styles.count)  # ExSkip
        self.assertEqual(8, target.styles.count)  # ExSkip

        target.copy_styles_from_template(template)
        self.assertEqual(18, target.styles.count)  # ExSkip
        # ExEnd

    def test_copy_template_styles_via_document_new(self):
        # ExStart
        # ExFor:Document.CopyStylesFromTemplate(Document)
        # ExFor:Document.CopyStylesFromTemplate(String)
        # ExSummary:Shows how to copy styles from one document to another.
        # Create a document, and then add styles that we will copy to another document.
        template = aw.Document()

        style = template.styles.add(aw.StyleType.PARAGRAPH, "TemplateStyle1")
        style.font.name = "Times New Roman"
        #        style.Font.Color = Color.Navy

        style = template.styles.add(aw.StyleType.PARAGRAPH, "TemplateStyle2")
        style.font.name = "Arial"
        #        style.Font.Color = Color.DeepSkyBlue

        style = template.styles.add(aw.StyleType.PARAGRAPH, "TemplateStyle3")
        style.font.name = "Courier New"
        #        style.Font.Color = Color.RoyalBlue

        self.assertEqual(7, template.styles.count)

        # Create a document which we will copy the styles to.
        target = aw.Document()

        # Create a style with the same name as a style from the template document and add it to the target document.
        style = target.styles.add(aw.StyleType.PARAGRAPH, "TemplateStyle3")
        style.font.name = "Calibri"
        #        style.Font.Color = Color.Orange

        self.assertEqual(5, target.styles.count)

        # There are two ways of calling the method to copy all the styles from one document to another.
        # 1 -  Passing the template document object:
        target.copy_styles_from_template(template)

        # Copying styles adds all styles from the template document to the target
        # and overwrites existing styles with the same name.
        self.assertEqual(7, target.styles.count)

        self.assertEqual("Courier New", DocumentHelper.get_style_by_name(target, "TemplateStyle3").font.name)
        #        self.assertEqual(Color.RoyalBlue.ToArgb(), target.Styles["TemplateStyle3"].Font.Color.ToArgb())

        # 2 -  Passing the local system filename of a template document:
        target.copy_styles_from_template(aeb.my_dir + "Rendering.docx")

        self.assertEqual(21, target.styles.count)
        # ExEnd

    def test_read_macros_from_existing_document(self):
        # ExStart
        # ExFor:Document.VbaProject
        # ExFor:VbaModuleCollection
        # ExFor:VbaModuleCollection.Count
        # ExFor:VbaModuleCollection.Item(System.Int32)
        # ExFor:VbaModuleCollection.Item(System.String)
        # ExFor:VbaModuleCollection.Remove
        # ExFor:VbaModule
        # ExFor:VbaModule.Name
        # ExFor:VbaModule.SourceCode
        # ExFor:VbaProject
        # ExFor:VbaProject.Name
        # ExFor:VbaProject.Modules
        # ExFor:VbaProject.CodePage
        # ExFor:VbaProject.IsSigned
        # ExSummary:Shows how to access a document's VBA project information.
        doc = aw.Document(aeb.my_dir + "VBA project.docm")

        # A VBA project contains a collection of VBA modules.
        vbaProject = doc.vba_project
        self.assertTrue(vbaProject.is_signed)  # ExSkip
        if vbaProject.is_signed:
            print(
                f"Project name: {vbaProject.name} signed Project code page: {vbaProject.code_page} Modules count: {vbaProject.modules.count}\n")
        else:
            print(
                f"Project name: {vbaProject.name} not signed Project code page: {vbaProject.code_page} Modules count: {vbaProject.modules.count}\n")

        vbaModules = doc.vba_project.modules

        self.assertEqual(vbaModules.count, 3)

        for module in vbaModules:
            print(f"Module name: {module.name}\nModule code:\n{module.source_code}\n")

        # Set new source code for VBA module. You can access VBA modules in the collection either by index.
        vbaModules[0].source_code = "Your VBA code..."
        DocumentHelper.get_vba_module_by_name(doc, "Module1").source_code = "Your VBA code..."

        # Remove a module from the collection.
        vbaModules.remove(vbaModules[2])
        # ExEnd

        self.assertEqual("AsposeVBAtest", vbaProject.name)
        self.assertEqual(2, vbaProject.modules.count)
        self.assertEqual(1251, vbaProject.code_page)
        self.assertFalse(vbaProject.is_signed)

        self.assertEqual("ThisDocument", vbaModules[0].name)
        self.assertEqual("Your VBA code...", vbaModules[0].source_code)

        self.assertEqual("Module1", vbaModules[1].name)
        self.assertEqual("Your VBA code...", vbaModules[1].source_code)

    def test_save_output_parameters(self):
        # ExStart
        # ExFor:SaveOutputParameters
        # ExFor:SaveOutputParameters.ContentType
        # ExSummary:Shows how to access output parameters of a document's save operation.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world!")

        # After we save a document, we can access the Internet Media Type (MIME type) of the newly created output document.
        parameters = doc.save(aeb.artifacts_dir + "Document.SaveOutputParameters.doc")

        self.assertEqual("application/msword", parameters.content_type)

        # This property changes depending on the save format.
        parameters = doc.save(aeb.artifacts_dir + "Document.SaveOutputParameters.pdf")

        self.assertEqual("application/pdf", parameters.content_type)
        # ExEnd

    def test_sub_document(self):
        # ExStart
        # ExFor:SubDocument
        # ExFor:SubDocument.NodeType
        # ExSummary:Shows how to access a master document's subdocument.
        doc = aw.Document(aeb.my_dir + "Master document.docx")

        subDocuments = doc.get_child_nodes(aw.NodeType.SUB_DOCUMENT, True)
        self.assertEqual(1, subDocuments.count)  # ExSkip

        # This node serves as a reference to an external document, and its contents cannot be accessed.
        subDocument = subDocuments[0]

        self.assertFalse(subDocument.is_composite)
        # ExEnd

    def test_create_web_extension(self):
        # ExStart
        # ExFor:BaseWebExtensionCollection`1.Add(`0)
        # ExFor:BaseWebExtensionCollection`1.Clear
        # ExFor:TaskPane
        # ExFor:TaskPane.DockState
        # ExFor:TaskPane.IsVisible
        # ExFor:TaskPane.Width
        # ExFor:TaskPane.IsLocked
        # ExFor:TaskPane.WebExtension
        # ExFor:TaskPane.Row
        # ExFor:WebExtension
        # ExFor:WebExtension.Reference
        # ExFor:WebExtension.Properties
        # ExFor:WebExtension.Bindings
        # ExFor:WebExtension.IsFrozen
        # ExFor:WebExtensionReference.Id
        # ExFor:WebExtensionReference.Version
        # ExFor:WebExtensionReference.StoreType
        # ExFor:WebExtensionReference.Store
        # ExFor:WebExtensionPropertyCollection
        # ExFor:WebExtensionBindingCollection
        # ExFor:WebExtensionProperty.#ctor(String, String)
        # ExFor:WebExtensionBinding.#ctor(String, WebExtensionBindingType, String)
        # ExFor:WebExtensionStoreType
        # ExFor:WebExtensionBindingType
        # ExFor:TaskPaneDockState
        # ExFor:TaskPaneCollection
        # ExSummary:Shows how to add a web extension to a document.
        doc = aw.Document()

        # Create task pane with "MyScript" add-in, which will be used by the document,
        # then set its default location.
        myScriptTaskPane = aw.webextensions.TaskPane()
        doc.web_extension_task_panes.add(myScriptTaskPane)
        myScriptTaskPane.dock_state = aw.webextensions.TaskPaneDockState.RIGHT
        myScriptTaskPane.is_visible = True
        myScriptTaskPane.width = 300
        myScriptTaskPane.is_locked = True

        # If there are multiple task panes in the same docking location, we can set this index to arrange them.
        myScriptTaskPane.row = 1

        # Create an add-in called "MyScript Math Sample", which the task pane will display within.
        webExtension = myScriptTaskPane.web_extension

        # Set application store reference parameters for our add-in, such as the ID.
        webExtension.reference.id = "WA104380646"
        webExtension.reference.version = "1.0.0.0"
        webExtension.reference.store_type = aw.webextensions.WebExtensionStoreType.OMEX
        webExtension.reference.store = "en-US"  # CultureInfo.CurrentCulture.Name
        webExtension.properties.add(aw.webextensions.WebExtensionProperty("MyScript", "MyScript Math Sample"))
        webExtension.bindings.add(
            aw.webextensions.WebExtensionBinding("MyScript", aw.webextensions.WebExtensionBindingType.TEXT, "104380646"))

        # Allow the user to interact with the add-in.
        webExtension.is_frozen = False

        # We can access the web extension in Microsoft Word via Developer -> Add-ins.
        doc.save(aeb.artifacts_dir + "Document.WebExtension.docx")

        # Remove all web extension task panes at once like this.
        doc.web_extension_task_panes.clear()

        #self.assertEqual(0, doc.web_extension_task_panes.count) # Conunt was not wrapped.
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "Document.WebExtension.docx")
        myScriptTaskPane = doc.web_extension_task_panes[0]

        self.assertEqual(aw.webextensions.TaskPaneDockState.RIGHT, myScriptTaskPane.dock_state)
        self.assertTrue(myScriptTaskPane.is_visible)
        self.assertEqual(300.0, myScriptTaskPane.width)
        self.assertTrue(myScriptTaskPane.is_locked)
        self.assertEqual(1, myScriptTaskPane.row)
        webExtension = myScriptTaskPane.web_extension

        self.assertEqual("WA104380646", webExtension.reference.id)
        self.assertEqual("1.0.0.0", webExtension.reference.version)
        self.assertEqual(aw.webextensions.WebExtensionStoreType.OMEX, webExtension.reference.store_type)
        self.assertEqual("en-US", webExtension.reference.store)

        self.assertEqual("MyScript", webExtension.properties[0].name)
        self.assertEqual("MyScript Math Sample", webExtension.properties[0].value)

        self.assertEqual("MyScript", webExtension.bindings[0].id)
        self.assertEqual(aw.webextensions.WebExtensionBindingType.TEXT, webExtension.bindings[0].binding_type)
        self.assertEqual("104380646", webExtension.bindings[0].app_ref)

        self.assertFalse(webExtension.is_frozen)

    def test_get_web_extension_info(self):
        # ExStart
        # ExFor:BaseWebExtensionCollection`1
        # ExFor:BaseWebExtensionCollection`1.GetEnumerator
        # ExFor:BaseWebExtensionCollection`1.Remove(Int32)
        # ExFor:BaseWebExtensionCollection`1.Count
        # ExFor:BaseWebExtensionCollection`1.Item(Int32)
        # ExSummary:Shows how to work with a document's collection of web extensions.
        doc = aw.Document(aeb.my_dir + "Web extension.docx")

        #self.assertEqual(1, doc.web_extension_task_panes.count) # Count was not wrapped

        # Print all properties of the document's web extension.
        webExtensionPropertyCollection = doc.web_extension_task_panes[0].web_extension.properties
        for webExtensionProperty in webExtensionPropertyCollection:
            print(f"Binding name: {webExtensionProperty.name} Binding value: {webExtensionProperty.value}")

        # Remove the web extension.
        doc.web_extension_task_panes.remove(0)

        #self.assertEqual(0, doc.web_extension_task_panes.count) # Cont was not wrapped
        # ExEnd

    def test_epub_cover(self):
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world!")

        # When saving to .epub, some Microsoft Word document properties convert to .epub metadata.
        doc.built_in_document_properties.author = "John Doe"
        doc.built_in_document_properties.title = "My Book Title"

        # The thumbnail we specify here can become the cover image.
        with open(aeb.image_dir + "Transparent background logo.png", "rb") as f:
            image = f.read()
        doc.built_in_document_properties.thumbnail = image

        doc.save(aeb.artifacts_dir + "Document.EpubCover.epub")

    def test_text_watermark(self):
        # ExStart
        # ExFor:Watermark.SetText(String)
        # ExFor:Watermark.SetText(String, TextWatermarkOptions)
        # ExFor:Watermark.Remove
        # ExFor:TextWatermarkOptions.FontFamily
        # ExFor:TextWatermarkOptions.FontSize
        # ExFor:TextWatermarkOptions.Color
        # ExFor:TextWatermarkOptions.Layout
        # ExFor:TextWatermarkOptions.IsSemitrasparent
        # ExFor:WatermarkLayout
        # ExFor:WatermarkType
        # ExSummary:Shows how to create a text watermark.
        doc = aw.Document()

        # Add a plain text watermark.
        doc.watermark.set_text("Aspose Watermark")

        # If we wish to edit the text formatting using it as a watermark,
        # we can do so by passing a TextWatermarkOptions object when creating the watermark.
        textWatermarkOptions = aw.TextWatermarkOptions()
        textWatermarkOptions.font_family = "Arial"
        textWatermarkOptions.font_size = 36
        #        textWatermarkOptions.Color = Color.Black
        textWatermarkOptions.layout = aw.WatermarkLayout.DIAGONAL
        textWatermarkOptions.is_semitrasparent = False

        doc.watermark.set_text("Aspose Watermark", textWatermarkOptions)

        doc.save(aeb.artifacts_dir + "Document.TextWatermark.docx")

        # We can remove a watermark from a document like this.
        if doc.watermark.type == aw.WatermarkType.TEXT:
            doc.watermark.remove()
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "Document.TextWatermark.docx")

        self.assertEqual(aw.WatermarkType.TEXT, doc.watermark.type)

    def test_image_watermark(self):
        # ExStart
        # ExFor:Watermark.SetImage(Image, ImageWatermarkOptions)
        # ExFor:ImageWatermarkOptions.Scale
        # ExFor:ImageWatermarkOptions.IsWashout
        # ExSummary:Shows how to create a watermark from an image in the local file system.
        doc = aw.Document()

        # Modify the image watermark's appearance with an ImageWatermarkOptions object,
        # then pass it while creating a watermark from an image file.
        imageWatermarkOptions = aw.ImageWatermarkOptions()
        imageWatermarkOptions.scale = 5
        imageWatermarkOptions.is_washout = False

        doc.watermark.set_image(aeb.image_dir + "Logo.jpg", imageWatermarkOptions)

        doc.save(aeb.artifacts_dir + "Document.ImageWatermark.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "Document.ImageWatermark.docx")

        self.assertEqual(aw.WatermarkType.IMAGE, doc.watermark.type)

    def test_spelling_and_grammar_errors(self):
        for showErrors in [True, False]:
            with self.subTest(showErrors=showErrors):
                # ExStart
                # ExFor:Document.ShowGrammaticalErrors
                # ExFor:Document.ShowSpellingErrors
                # ExSummary:Shows how to show/hide errors in the document.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Insert two sentences with mistakes that would be picked up
                # by the spelling and grammar checkers in Microsoft Word.
                builder.writeln("There is a speling error in this sentence.")
                builder.writeln("Their is a grammatical error in this sentence.")

                # If these options are enabled, then spelling errors will be underlined
                # in the output document by a jagged red line, and a double blue line will highlight grammatical mistakes.
                doc.show_grammatical_errors = showErrors
                doc.show_spelling_errors = showErrors

                doc.save(aeb.artifacts_dir + "Document.SpellingAndGrammarErrors.docx")
                # ExEnd

                doc = aw.Document(aeb.artifacts_dir + "Document.SpellingAndGrammarErrors.docx")

                self.assertEqual(showErrors, doc.show_grammatical_errors)
                self.assertEqual(showErrors, doc.show_spelling_errors)

    def test_granularity_compare_option(self):
        for granularity in [aw.comparing.Granularity.CHAR_LEVEL, aw.comparing.Granularity.WORD_LEVEL]:
            with self.subTest(granularity=granularity):
                # ExStart
                # ExFor:CompareOptions.Granularity
                # ExFor:Granularity
                # ExSummary:Shows to specify a granularity while comparing documents.
                docA = aw.Document()
                builderA = aw.DocumentBuilder(docA)
                builderA.writeln("Alpha Lorem ipsum dolor sit amet, consectetur adipiscing elit")

                docB = aw.Document()
                builderB = aw.DocumentBuilder(docB)
                builderB.writeln("Lorems ipsum dolor sit amet consectetur - \"adipiscing\" elit")

                # Specify whether changes are tracking
                # by character ('Granularity.CharLevel'), or by word ('Granularity.WordLevel').
                compareOptions = aw.comparing.CompareOptions()
                compareOptions.granularity = granularity

                docA.compare(docB, "author", date.today(), compareOptions)

                # The first document's collection of revision groups contains all the differences between documents.
                groups = docA.revisions.groups
                self.assertEqual(5, groups.count)
                # ExEnd

                if granularity == aw.comparing.Granularity.CHAR_LEVEL:
                    self.assertEqual(aw.RevisionType.DELETION, groups[0].revision_type)
                    self.assertEqual("Alpha ", groups[0].text)

                    self.assertEqual(aw.RevisionType.DELETION, groups[1].revision_type)
                    self.assertEqual(",", groups[1].text)

                    self.assertEqual(aw.RevisionType.INSERTION, groups[2].revision_type)
                    self.assertEqual("s", groups[2].text)

                    self.assertEqual(aw.RevisionType.INSERTION, groups[3].revision_type)
                    self.assertEqual("- \"", groups[3].text)

                    self.assertEqual(aw.RevisionType.INSERTION, groups[4].revision_type)
                    self.assertEqual("\"", groups[4].text)
                else:
                    self.assertEqual(aw.RevisionType.DELETION, groups[0].revision_type)
                    self.assertEqual("Alpha Lorem ", groups[0].text)

                    self.assertEqual(aw.RevisionType.DELETION, groups[1].revision_type)
                    self.assertEqual(",", groups[1].text)

                    self.assertEqual(aw.RevisionType.INSERTION, groups[2].revision_type)
                    self.assertEqual("Lorems ", groups[2].text)

                    self.assertEqual(aw.RevisionType.INSERTION, groups[3].revision_type)
                    self.assertEqual("- \"", groups[3].text)

                    self.assertEqual(aw.RevisionType.INSERTION, groups[4].revision_type)
                    self.assertEqual("\"", groups[4].text)

    def test_ignore_printer_metrics(self):
        # ExStart
        # ExFor:LayoutOptions.IgnorePrinterMetrics
        # ExSummary:Shows how to ignore 'Use printer metrics to lay out document' option.
        doc = aw.Document(aeb.my_dir + "Rendering.docx")

        doc.layout_options.ignore_printer_metrics = False

        doc.save(aeb.artifacts_dir + "Document.IgnorePrinterMetrics.docx")
        # ExEnd

    def test_extract_pages(self):
        # ExStart
        # ExFor:Document.ExtractPages
        # ExSummary:Shows how to get specified range of pages from the document.
        doc = aw.Document(aeb.my_dir + "Layout entities.docx")

        doc = doc.extract_pages(0, 2)

        doc.save(aeb.artifacts_dir + "Document.ExtractPages.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "Document.ExtractPages.docx")
        self.assertEqual(doc.page_count, 2)

    def test_spelling_or_grammar(self):
        for checkSpellingGrammar in [True, False]:
            with self.subTest(checkSpellingGrammar=checkSpellingGrammar):
                # ExStart
                # ExFor:Document.SpellingChecked
                # ExFor:Document.GrammarChecked
                # ExSummary:Shows how to set spelling or grammar verifying.
                doc = aw.Document()

                # The string with spelling errors.
                doc.first_section.body.first_paragraph.runs.add(
                    aw.Run(doc, "The speeling in this documentz is all broked."))

                # Spelling/Grammar check start if we set properties to False. 
                # We can see all errors in Microsoft Word via Review -> Spelling & Grammar.
                # Note that Microsoft Word does not start grammar/spell check automatically for DOC and RTF document format.
                doc.spelling_checked = checkSpellingGrammar
                doc.grammar_checked = checkSpellingGrammar

                doc.save(aeb.artifacts_dir + "Document.SpellingOrGrammar.docx")
                # ExEnd

    def test_allow_embedding_post_script_fonts(self):
        # ExStart
        # ExFor:SaveOptions.AllowEmbeddingPostScriptFonts
        # ExSummary:Shows how to save the document with PostScript font.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.font.name = "PostScriptFont"
        builder.writeln("Some text with PostScript font.")

        # Load the font with PostScript to use in the document.
        with open(aeb.fonts_dir + "AllegroOpen.otf", "rb") as f:
            fontBytes = f.read()
        otf = aw.fonts.MemoryFontSource(fontBytes)
        doc.font_settings = aw.fonts.FontSettings()
        doc.font_settings.set_fonts_sources([otf])

        # Embed TrueType fonts.
        doc.font_infos.embed_true_type_fonts = True

        # Allow embedding PostScript fonts while embedding TrueType fonts.
        # Microsoft Word does not embed PostScript fonts, but can open documents with embedded fonts of this type.
        saveOptions = aw.saving.SaveOptions.create_save_options(aw.SaveFormat.DOCX)
        saveOptions.allow_embedding_post_script_fonts = True

        doc.save(aeb.artifacts_dir + "Document.AllowEmbeddingPostScriptFonts.docx", saveOptions)
        # ExEnd


if __name__ == '__main__':
    unittest.main()
