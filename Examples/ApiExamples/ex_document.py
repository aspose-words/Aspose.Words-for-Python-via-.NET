# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from urllib.request import urlopen, Request
import glob
import sys
import os
import aspose.words.drawing
import base64
from document_helper import DocumentHelper
from datetime import timedelta, timezone
import aspose.pydrawing
import aspose.words as aw
import aspose.words.digitalsignatures
import aspose.words.fields
import aspose.words.fonts
import aspose.words.layout
import aspose.words.loading
import aspose.words.notes
import aspose.words.rendering
import aspose.words.saving
import aspose.words.settings
import aspose.words.webextensions
import datetime
import document_helper
import io
import system_helper
import test_util
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, FONTS_DIR, GOLDS_DIR, IMAGE_DIR, MY_DIR

class ExDocument(ApiExampleBase):

    def test_constructor(self):
        #ExStart
        #ExFor:Document.__init__()
        #ExFor:Document.__init__(str,LoadOptions)
        #ExSummary:Shows how to create and load documents.
        # There are two ways of creating a Document object using Aspose.Words.
        # 1 -  Create a blank document:
        doc = aw.Document()
        # New Document objects by default come with the minimal set of nodes
        # required to begin adding content such as text and shapes: a Section, a Body, and a Paragraph.
        doc.first_section.body.first_paragraph.append_child(aw.Run(doc=doc, text='Hello world!'))
        # 2 -  Load a document that exists in the local file system:
        doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        # Loaded documents will have contents that we can access and edit.
        self.assertEqual('Hello World!', doc.first_section.body.first_paragraph.get_text().strip())
        # Some operations that need to occur during loading, such as using a password to decrypt a document,
        # can be done by passing a LoadOptions object when loading the document.
        doc = aw.Document(file_name=MY_DIR + 'Encrypted.docx', load_options=aw.loading.LoadOptions(password='docPassword'))
        self.assertEqual('Test encrypted document.', doc.first_section.body.first_paragraph.get_text().strip())
        #ExEnd

    def test_load_from_stream(self):
        #ExStart
        #ExFor:Document.__init__(BytesIO)
        #ExSummary:Shows how to load a document using a stream.
        with system_helper.io.File.open_read(MY_DIR + 'Document.docx') as stream:
            doc = aw.Document(stream=stream)
            self.assertEqual('Hello World!\r\rHello Word!\r\r\rHello World!', doc.get_text().strip())
        #ExEnd

    def test_convert_to_pdf(self):
        #ExStart
        #ExFor:Document.__init__(str)
        #ExFor:Document.save(str)
        #ExSummary:Shows how to open a document and convert it to .PDF.
        doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        doc.save(file_name=ARTIFACTS_DIR + 'Document.ConvertToPdf.pdf')
        #ExEnd

    def test_detect_mobi_document_format(self):
        info = aw.FileFormatUtil.detect_file_format(file_name=MY_DIR + 'Document.mobi')
        self.assertEqual(info.load_format, aw.LoadFormat.MOBI)

    def test_detect_pdf_document_format(self):
        info = aw.FileFormatUtil.detect_file_format(file_name=MY_DIR + 'Pdf Document.pdf')
        self.assertEqual(info.load_format, aw.LoadFormat.PDF)

    def test_open_pdf_document(self):
        doc = aw.Document(file_name=MY_DIR + 'Pdf Document.pdf')
        self.assertEqual('Heading 1\rHeading 1.1.1.1 Heading 1.1.1.2\rHeading 1.1.1.1.1.1.1.1.1 Heading 1.1.1.1.1.1.1.1.2\x0c', doc.range.text)

    def test_open_protected_pdf_document(self):
        doc = aw.Document(file_name=MY_DIR + 'Pdf Document.pdf')
        save_options = aw.saving.PdfSaveOptions()
        save_options.encryption_details = aw.saving.PdfEncryptionDetails(user_password='Aspose', owner_password=None)
        doc.save(file_name=ARTIFACTS_DIR + 'Document.PdfDocumentEncrypted.pdf', save_options=save_options)
        load_options = aw.loading.PdfLoadOptions()
        load_options.password = 'Aspose'
        load_options.load_format = aw.LoadFormat.PDF
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Document.PdfDocumentEncrypted.pdf', load_options=load_options)

    def test_open_from_stream_with_base_uri(self):
        #ExStart
        #ExFor:Document.__init__(BytesIO,LoadOptions)
        #ExFor:LoadOptions.__init__
        #ExFor:LoadOptions.base_uri
        #ExFor:ShapeBase.is_image
        #ExSummary:Shows how to open an HTML document with images from a stream using a base URI.
        with system_helper.io.File.open_read(MY_DIR + 'Document.html') as stream:
            # Pass the URI of the base folder while loading it
            # so that any images with relative URIs in the HTML document can be found.
            load_options = aw.loading.LoadOptions()
            load_options.base_uri = IMAGE_DIR
            doc = aw.Document(stream=stream, load_options=load_options)
            # Verify that the first shape of the document contains a valid image.
            shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
            self.assertTrue(shape.is_image)
            self.assertIsNotNone(shape.image_data.image_bytes)
            self.assertAlmostEqual(32, aw.ConvertUtil.point_to_pixel(points=shape.width), delta=0.01)
            self.assertAlmostEqual(32, aw.ConvertUtil.point_to_pixel(points=shape.height), delta=0.01)
        #ExEnd

    def test_load_encrypted(self):
        #ExStart
        #ExFor:Document.__init__(BytesIO,LoadOptions)
        #ExFor:Document.__init__(str,LoadOptions)
        #ExFor:LoadOptions
        #ExFor:LoadOptions.__init__(str)
        #ExSummary:Shows how to load an encrypted Microsoft Word document.
        doc = None
        # Aspose.Words throw an exception if we try to open an encrypted document without its password.
        with self.assertRaises(Exception):
            doc = aw.Document(file_name=MY_DIR + 'Encrypted.docx')
        # When loading such a document, the password is passed to the document's constructor using a LoadOptions object.
        options = aw.loading.LoadOptions(password='docPassword')
        # There are two ways of loading an encrypted document with a LoadOptions object.
        # 1 -  Load the document from the local file system by filename:
        doc = aw.Document(file_name=MY_DIR + 'Encrypted.docx', load_options=options)
        self.assertEqual('Test encrypted document.', doc.get_text().strip())  #ExSkip
        # 2 -  Load the document from a stream:
        with system_helper.io.File.open_read(MY_DIR + 'Encrypted.docx') as stream:
            doc = aw.Document(stream=stream, load_options=options)
            self.assertEqual('Test encrypted document.', doc.get_text().strip())  #ExSkip
        #ExEnd

    def test_temp_folder(self):
        #ExStart
        #ExFor:LoadOptions.temp_folder
        #ExSummary:Shows how to load a document using temporary files.
        # Note that such an approach can reduce memory usage but degrades speed
        load_options = aw.loading.LoadOptions()
        load_options.temp_folder = 'C:\\TempFolder\\'
        # Ensure that the directory exists and load
        system_helper.io.Directory.create_directory(load_options.temp_folder)
        doc = aw.Document(file_name=MY_DIR + 'Document.docx', load_options=load_options)
        #ExEnd

    def test_convert_to_html(self):
        #ExStart
        #ExFor:Document.save(str,SaveFormat)
        #ExFor:SaveFormat
        #ExSummary:Shows how to convert from DOCX to HTML format.
        doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        doc.save(file_name=ARTIFACTS_DIR + 'Document.ConvertToHtml.html', save_format=aw.SaveFormat.HTML)
        #ExEnd

    def test_convert_to_mhtml(self):
        doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        doc.save(file_name=ARTIFACTS_DIR + 'Document.ConvertToMhtml.mht')

    def test_convert_to_txt(self):
        doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        doc.save(file_name=ARTIFACTS_DIR + 'Document.ConvertToTxt.txt')

    def test_convert_to_epub(self):
        doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
        doc.save(file_name=ARTIFACTS_DIR + 'Document.ConvertToEpub.epub')

    def test_save_to_stream(self):
        #ExStart
        #ExFor:Document.save(BytesIO,SaveFormat)
        #ExSummary:Shows how to save a document to a stream.
        doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        with io.BytesIO() as dst_stream:
            doc.save(stream=dst_stream, save_format=aw.SaveFormat.DOCX)
            # Verify that the stream contains the document.
            self.assertEqual('Hello World!\r\rHello Word!\r\r\rHello World!', aw.Document(stream=dst_stream).get_text().strip())
        #ExEnd

    def test_append_document(self):
        #ExStart
        #ExFor:Document.append_document(Document,ImportFormatMode)
        #ExSummary:Shows how to append a document to the end of another document.
        src_doc = aw.Document()
        src_doc.first_section.body.append_paragraph('Source document text. ')
        dst_doc = aw.Document()
        dst_doc.first_section.body.append_paragraph('Destination document text. ')
        # Append the source document to the destination document while preserving its formatting,
        # then save the source document to the local file system.
        dst_doc.append_document(src_doc=src_doc, import_format_mode=aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
        self.assertEqual(2, dst_doc.sections.count)  #ExSkip
        dst_doc.save(file_name=ARTIFACTS_DIR + 'Document.AppendDocument.docx')
        #ExEnd
        out_doc_text = aw.Document(file_name=ARTIFACTS_DIR + 'Document.AppendDocument.docx').get_text()
        self.assertTrue(out_doc_text.startswith(dst_doc.get_text()))
        self.assertTrue(out_doc_text.endswith(src_doc.get_text()))

    def test_append_document_from_automation(self):
        doc = aw.Document()
        # We should call this method to clear this document of any existing content.
        doc.remove_all_children()
        record_count = 5
        i = 1
        while i <= record_count:
            src_doc = aw.Document()
            self.assertRaises(Exception, lambda: aw.Document(file_name='C:\\DetailsList.doc'))
            # Append the source document at the end of the destination document.
            doc.append_document(src_doc=src_doc, import_format_mode=aw.ImportFormatMode.USE_DESTINATION_STYLES)
            # Automation required you to insert a new section break at this point, however, in Aspose.Words we
            # do not need to do anything here as the appended document is imported as separate sections already
            # Unlink all headers/footers in this section from the previous section headers/footers
            # if this is the second document or above being appended.
            if i > 1:
                self.assertRaises(Exception, lambda: doc.sections[i].headers_footers.link_to_previous(is_link_to_previous=False))
            i += 1

    def test_keep_source_numbering_same_list_ids(self):
        #ExStart
        #ExFor:ImportFormatOptions.keep_source_numbering
        #ExFor:NodeImporter.__init__(DocumentBase,DocumentBase,ImportFormatMode,ImportFormatOptions)
        #ExSummary:Shows how resolve a clash when importing documents that have lists with the same list definition identifier.
        src_doc = aw.Document(file_name=MY_DIR + 'List with the same definition identifier - source.docx')
        dst_doc = aw.Document(file_name=MY_DIR + 'List with the same definition identifier - destination.docx')
        # Set the "KeepSourceNumbering" property to "true" to apply a different list definition ID
        # to identical styles as Aspose.Words imports them into destination documents.
        import_format_options = aw.ImportFormatOptions()
        import_format_options.keep_source_numbering = True
        dst_doc.append_document(src_doc=src_doc, import_format_mode=aw.ImportFormatMode.USE_DESTINATION_STYLES, import_format_options=import_format_options)
        dst_doc.update_list_labels()
        #ExEnd
        para_text = dst_doc.sections[1].body.last_paragraph.get_text()
        self.assertTrue(para_text.startswith('13->13'), msg=para_text)
        self.assertEqual('1.', dst_doc.sections[1].body.last_paragraph.list_label.label_string)

    def test_merge_pasted_lists(self):
        #ExStart
        #ExFor:ImportFormatOptions.merge_pasted_lists
        #ExSummary:Shows how to merge lists from a documents.
        src_doc = aw.Document(file_name=MY_DIR + 'List item.docx')
        dst_doc = aw.Document(file_name=MY_DIR + 'List destination.docx')
        options = aw.ImportFormatOptions()
        options.merge_pasted_lists = True
        # Set the "MergePastedLists" property to "true" pasted lists will be merged with surrounding lists.
        dst_doc.append_document(src_doc=src_doc, import_format_mode=aw.ImportFormatMode.USE_DESTINATION_STYLES, import_format_options=options)
        dst_doc.save(file_name=ARTIFACTS_DIR + 'Document.MergePastedLists.docx')
        #ExEnd

    def test_force_copy_styles(self):
        #ExStart
        #ExFor:ImportFormatOptions.force_copy_styles
        #ExSummary:Shows how to copy source styles with unique names forcibly.
        # Both documents contain MyStyle1 and MyStyle2, MyStyle3 exists only in a source document.
        src_doc = aw.Document(file_name=MY_DIR + 'Styles source.docx')
        dst_doc = aw.Document(file_name=MY_DIR + 'Styles destination.docx')
        options = aw.ImportFormatOptions()
        options.force_copy_styles = True
        dst_doc.append_document(src_doc=src_doc, import_format_mode=aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, import_format_options=options)
        paras = dst_doc.sections[1].body.paragraphs
        self.assertEqual(paras[0].paragraph_format.style.name, 'MyStyle1_0')
        self.assertEqual(paras[1].paragraph_format.style.name, 'MyStyle2_0')
        self.assertEqual(paras[2].paragraph_format.style.name, 'MyStyle3')
        #ExEnd

    def test_adjust_sentence_and_word_spacing(self):
        #ExStart
        #ExFor:ImportFormatOptions.adjust_sentence_and_word_spacing
        #ExSummary:Shows how to adjust sentence and word spacing automatically.
        src_doc = aw.Document()
        dst_doc = aw.Document()
        builder = aw.DocumentBuilder(doc=src_doc)
        builder.write('Dolor sit amet.')
        builder = aw.DocumentBuilder(doc=dst_doc)
        builder.write('Lorem ipsum.')
        options = aw.ImportFormatOptions()
        options.adjust_sentence_and_word_spacing = True
        builder.insert_document(src_doc=src_doc, import_format_mode=aw.ImportFormatMode.USE_DESTINATION_STYLES, import_format_options=options)
        self.assertEqual('Lorem ipsum. Dolor sit amet.', dst_doc.first_section.body.first_paragraph.get_text().strip())
        #ExEnd

    @unittest.skip('DigitalSignatureUtil.sing method is not supported')
    def test_digital_signature(self):
        #ExStart
        #ExFor:DigitalSignature.certificate_holder
        #ExFor:DigitalSignature.issuer_name
        #ExFor:DigitalSignature.subject_name
        #ExFor:DigitalSignatureCollection
        #ExFor:DigitalSignatureCollection.is_valid
        #ExFor:DigitalSignatureCollection.count
        #ExFor:DigitalSignatureCollection.__getitem__(int)
        #ExFor:DigitalSignatureUtil.sign(BytesIO,BytesIO,CertificateHolder)
        #ExFor:DigitalSignatureUtil.sign(str,str,CertificateHolder)
        #ExFor:DigitalSignatureType
        #ExFor:Document.digital_signatures
        #ExSummary:Shows how to sign documents with X.509 certificates.
        # Verify that a document is not signed.
        self.assertFalse(aw.FileFormatUtil.detect_file_format(file_name=MY_DIR + 'Document.docx').has_digital_signature)
        # Create a CertificateHolder object from a PKCS12 file, which we will use to sign the document.
        certificate_holder = aw.digitalsignatures.CertificateHolder.create(file_name=MY_DIR + 'morzal.pfx', password='aw', alias=None)
        # There are two ways of saving a signed copy of a document to the local file system:
        # 1 - Designate a document by a local system filename and save a signed copy at a location specified by another filename.
        sign_options = aw.digitalsignatures.SignOptions()
        sign_options.sign_time = datetime.datetime.now()
        aw.digitalsignatures.DigitalSignatureUtil.sign(src_file_name=MY_DIR + 'Document.docx', dst_file_name=ARTIFACTS_DIR + 'Document.DigitalSignature.docx', cert_holder=certificate_holder, sign_options=sign_options)
        self.assertTrue(aw.FileFormatUtil.detect_file_format(file_name=ARTIFACTS_DIR + 'Document.DigitalSignature.docx').has_digital_signature)
        # 2 - Take a document from a stream and save a signed copy to another stream.
        with system_helper.io.FileStream(MY_DIR + 'Document.docx', system_helper.io.FileMode.OPEN) as in_doc:
            with system_helper.io.FileStream(ARTIFACTS_DIR + 'Document.DigitalSignature.docx', system_helper.io.FileMode.CREATE) as out_doc:
                aw.digitalsignatures.DigitalSignatureUtil.sign(src_stream=in_doc, dst_stream=out_doc, cert_holder=certificate_holder)
        self.assertTrue(aw.FileFormatUtil.detect_file_format(file_name=ARTIFACTS_DIR + 'Document.DigitalSignature.docx').has_digital_signature)
        # Please verify that all of the document's digital signatures are valid and check their details.
        signed_doc = aw.Document(file_name=ARTIFACTS_DIR + 'Document.DigitalSignature.docx')
        digital_signature_collection = signed_doc.digital_signatures
        self.assertTrue(digital_signature_collection.is_valid)
        self.assertEqual(1, digital_signature_collection.count)
        self.assertEqual(aw.digitalsignatures.DigitalSignatureType.XML_DSIG, digital_signature_collection[0].signature_type)
        self.assertEqual('CN=Morzal.Me', signed_doc.digital_signatures[0].issuer_name)
        self.assertEqual('CN=Morzal.Me', signed_doc.digital_signatures[0].subject_name)
        #ExEnd

    def test_append_all_documents_in_folder(self):
        #ExStart
        #ExFor:Document.append_document(Document,ImportFormatMode)
        #ExSummary:Shows how to append all the documents in a folder to the end of a template document.
        dst_doc = aw.Document()
        builder = aw.DocumentBuilder(doc=dst_doc)
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
        builder.writeln('Template Document')
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.NORMAL
        builder.writeln('Some content here')
        self.assertEqual(5, dst_doc.styles.count)  #ExSkip
        self.assertEqual(1, dst_doc.sections.count)  #ExSkip
        # Append all unencrypted documents with the .doc extension
        # from our local file system directory to the base document.
        doc_files = list(filter(lambda item: item.endswith('.doc'), list(system_helper.io.Directory.get_files(MY_DIR, '*.doc'))))
        for file_name in doc_files:
            info = aw.FileFormatUtil.detect_file_format(file_name=file_name)
            if info.is_encrypted:
                continue
            src_doc = aw.Document(file_name=file_name)
            dst_doc.append_document(src_doc=src_doc, import_format_mode=aw.ImportFormatMode.USE_DESTINATION_STYLES)
        dst_doc.save(file_name=ARTIFACTS_DIR + 'Document.AppendAllDocumentsInFolder.doc')
        #ExEnd
        self.assertEqual(7, dst_doc.styles.count)
        self.assertEqual(10, dst_doc.sections.count)

    def test_join_runs_with_same_formatting(self):
        #ExStart
        #ExFor:Document.join_runs_with_same_formatting
        #ExSummary:Shows how to join runs in a document to reduce unneeded runs.
        # Open a document that contains adjacent runs of text with identical formatting,
        # which commonly occurs if we edit the same paragraph multiple times in Microsoft Word.
        doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
        # If any number of these runs are adjacent with identical formatting,
        # then the document may be simplified.
        self.assertEqual(317, doc.get_child_nodes(aw.NodeType.RUN, True).count)
        # Combine such runs with this method and verify the number of run joins that will take place.
        self.assertEqual(121, doc.join_runs_with_same_formatting())
        # The number of joins and the number of runs we have after the join
        # should add up the number of runs we had initially.
        self.assertEqual(196, doc.get_child_nodes(aw.NodeType.RUN, True).count)
        #ExEnd

    def test_clone_document(self):
        #ExStart
        #ExFor:Document.clone
        #ExSummary:Shows how to deep clone a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.write('Hello world!')
        # Cloning will produce a new document with the same contents as the original,
        # but with a unique copy of each of the original document's nodes.
        clone = doc.clone()
        self.assertEqual(doc.first_section.body.first_paragraph.runs[0].get_text(), clone.first_section.body.first_paragraph.runs[0].text)
        self.assertNotEqual(hash(doc.first_section.body.first_paragraph.runs[0]), hash(clone.first_section.body.first_paragraph.runs[0]))
        #ExEnd

    def test_document_get_text_to_string(self):
        #ExStart
        #ExFor:CompositeNode.get_text
        #ExFor:Node.__str__(SaveFormat)
        #ExSummary:Shows the difference between calling the GetText and ToString methods on a node.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.insert_field(field_code='MERGEFIELD Field')
        # GetText will retrieve the visible text as well as field codes and special characters.
        self.assertEqual('\x13MERGEFIELD Field\x14«Field»\x15', doc.get_text().strip())
        # ToString will give us the document's appearance if saved to a passed save format.
        self.assertEqual('«Field»', doc.to_string(save_format=aw.SaveFormat.TEXT).strip())
        #ExEnd

    def test_protect_unprotect(self):
        #ExStart
        #ExFor:Document.protect(ProtectionType,str)
        #ExFor:Document.protection_type
        #ExFor:Document.unprotect
        #ExFor:Document.unprotect(str)
        #ExSummary:Shows how to protect and unprotect a document.
        doc = aw.Document()
        doc.protect(type=aw.ProtectionType.READ_ONLY, password='password')
        self.assertEqual(aw.ProtectionType.READ_ONLY, doc.protection_type)
        # If we open this document with Microsoft Word intending to edit it,
        # we will need to apply the password to get through the protection.
        doc.save(file_name=ARTIFACTS_DIR + 'Document.Protect.docx')
        # Note that the protection only applies to Microsoft Word users opening our document.
        # We have not encrypted the document in any way, and we do not need the password to open and edit it programmatically.
        protected_doc = aw.Document(file_name=ARTIFACTS_DIR + 'Document.Protect.docx')
        self.assertEqual(aw.ProtectionType.READ_ONLY, protected_doc.protection_type)
        builder = aw.DocumentBuilder(doc=protected_doc)
        builder.writeln('Text added to a protected document.')
        self.assertEqual('Text added to a protected document.', protected_doc.range.text.strip())  #ExSkip
        # There are two ways of removing protection from a document.
        # 1 - With no password:
        doc.unprotect()
        self.assertEqual(aw.ProtectionType.NO_PROTECTION, doc.protection_type)
        doc.protect(type=aw.ProtectionType.READ_ONLY, password='NewPassword')
        self.assertEqual(aw.ProtectionType.READ_ONLY, doc.protection_type)
        doc.unprotect('WrongPassword')
        self.assertEqual(aw.ProtectionType.READ_ONLY, doc.protection_type)
        # 2 - With the correct password:
        doc.unprotect('NewPassword')
        self.assertEqual(aw.ProtectionType.NO_PROTECTION, doc.protection_type)
        #ExEnd

    def test_document_ensure_minimum(self):
        #ExStart
        #ExFor:Document.ensure_minimum
        #ExSummary:Shows how to ensure that a document contains the minimal set of nodes required for editing its contents.
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
        nodes[2].as_paragraph().runs.add(aw.Run(doc=doc, text='Hello world!'))
        #ExEnd
        self.assertEqual('Hello world!', doc.get_text().strip())

    def test_remove_macros_from_document(self):
        #ExStart
        #ExFor:Document.remove_macros
        #ExSummary:Shows how to remove all macros from a document.
        doc = aw.Document(file_name=MY_DIR + 'Macro.docm')
        self.assertTrue(doc.has_macros)
        self.assertEqual('Project', doc.vba_project.name)
        # Remove the document's VBA project, along with all its macros.
        doc.remove_macros()
        self.assertFalse(doc.has_macros)
        self.assertIsNone(doc.vba_project)
        #ExEnd

    def test_get_page_count(self):
        #ExStart
        #ExFor:Document.page_count
        #ExSummary:Shows how to count the number of pages in the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.write('Page 1')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.write('Page 2')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.write('Page 3')
        # Verify the expected page count of the document.
        self.assertEqual(3, doc.page_count)
        # Getting the PageCount property invoked the document's page layout to calculate the value.
        # This operation will not need to be re-done when rendering the document to a fixed page save format,
        # such as .pdf. So you can save some time, especially with more complex documents.
        doc.save(file_name=ARTIFACTS_DIR + 'Document.GetPageCount.pdf')
        #ExEnd

    def test_get_updated_page_properties(self):
        #ExStart
        #ExFor:Document.update_word_count()
        #ExFor:Document.update_word_count(bool)
        #ExFor:BuiltInDocumentProperties.characters
        #ExFor:BuiltInDocumentProperties.words
        #ExFor:BuiltInDocumentProperties.paragraphs
        #ExFor:BuiltInDocumentProperties.lines
        #ExSummary:Shows how to update all list labels in a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Lorem ipsum dolor sit amet, consectetur adipiscing elit, ' + 'sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.')
        builder.write('Ut enim ad minim veniam, ' + 'quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.')
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
        #ExEnd

    def test_table_style_to_direct_formatting(self):
        #ExStart
        #ExFor:CompositeNode.get_child
        #ExFor:Document.expand_table_styles_to_direct_formatting
        #ExSummary:Shows how to apply the properties of a table's style directly to the table's elements.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        table = builder.start_table()
        builder.insert_cell()
        builder.write('Hello world!')
        builder.end_table()
        table_style = doc.styles.add(aw.StyleType.TABLE, 'MyTableStyle1').as_table_style()
        table_style.row_stripe = 3
        table_style.cell_spacing = 5
        table_style.shading.background_pattern_color = aspose.pydrawing.Color.antique_white
        table_style.borders.color = aspose.pydrawing.Color.blue
        table_style.borders.line_style = aw.LineStyle.DOT_DASH
        table.style = table_style
        # This method concerns table style properties such as the ones we set above.
        doc.expand_table_styles_to_direct_formatting()
        doc.save(file_name=ARTIFACTS_DIR + 'Document.TableStyleToDirectFormatting.docx')
        #ExEnd
        test_util.TestUtil.doc_package_file_contains_string('<w:tblStyleRowBandSize w:val="3" />', ARTIFACTS_DIR + 'Document.TableStyleToDirectFormatting.docx', 'document.xml')
        test_util.TestUtil.doc_package_file_contains_string('<w:tblCellSpacing w:w="100" w:type="dxa" />', ARTIFACTS_DIR + 'Document.TableStyleToDirectFormatting.docx', 'document.xml')
        test_util.TestUtil.doc_package_file_contains_string('<w:tblBorders><w:top w:val="dotDash" w:sz="2" w:space="0" w:color="0000FF" /><w:left w:val="dotDash" w:sz="2" w:space="0" w:color="0000FF" /><w:bottom w:val="dotDash" w:sz="2" w:space="0" w:color="0000FF" /><w:right w:val="dotDash" w:sz="2" w:space="0" w:color="0000FF" /><w:insideH w:val="dotDash" w:sz="2" w:space="0" w:color="0000FF" /><w:insideV w:val="dotDash" w:sz="2" w:space="0" w:color="0000FF" /></w:tblBorders>', ARTIFACTS_DIR + 'Document.TableStyleToDirectFormatting.docx', 'document.xml')

    def test_get_original_file_info(self):
        #ExStart
        #ExFor:Document.original_file_name
        #ExFor:Document.original_load_format
        #ExSummary:Shows how to retrieve details of a document's load operation.
        doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        self.assertEqual(MY_DIR + 'Document.docx', doc.original_file_name)
        self.assertEqual(aw.LoadFormat.DOCX, doc.original_load_format)
        #ExEnd

    def test_footnote_columns(self):
        #ExStart
        #ExFor:FootnoteOptions
        #ExFor:FootnoteOptions.columns
        #ExSummary:Shows how to split the footnote section into a given number of columns.
        doc = aw.Document(file_name=MY_DIR + 'Footnotes and endnotes.docx')
        self.assertEqual(0, doc.footnote_options.columns)  #ExSkip
        doc.footnote_options.columns = 2
        doc.save(file_name=ARTIFACTS_DIR + 'Document.FootnoteColumns.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Document.FootnoteColumns.docx')
        self.assertEqual(2, doc.first_section.page_setup.footnote_options.columns)

    def test_remove_external_schema_references(self):
        #ExStart
        #ExFor:Document.remove_external_schema_references
        #ExSummary:Shows how to remove all external XML schema references from a document.
        doc = aw.Document(file_name=MY_DIR + 'External XML schema.docx')
        doc.remove_external_schema_references()
        #ExEnd

    def test_update_thumbnail(self):
        #ExStart
        #ExFor:Document.update_thumbnail()
        #ExFor:Document.update_thumbnail(ThumbnailGeneratingOptions)
        #ExFor:ThumbnailGeneratingOptions
        #ExFor:ThumbnailGeneratingOptions.generate_from_first_page
        #ExFor:ThumbnailGeneratingOptions.thumbnail_size
        #ExSummary:Shows how to update a document's thumbnail.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Hello world!')
        builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
        # There are two ways of setting a thumbnail image when saving a document to .epub.
        # 1 -  Use the document's first page:
        doc.update_thumbnail()
        doc.save(file_name=ARTIFACTS_DIR + 'Document.UpdateThumbnail.FirstPage.epub')
        # 2 -  Use the first image found in the document:
        options = aw.rendering.ThumbnailGeneratingOptions()
        self.assertEqual(aspose.pydrawing.Size(600, 900), options.thumbnail_size)  #ExSkip
        self.assertTrue(options.generate_from_first_page)  #ExSkip
        options.thumbnail_size = aspose.pydrawing.Size(400, 400)
        options.generate_from_first_page = False
        doc.update_thumbnail(options)
        doc.save(file_name=ARTIFACTS_DIR + 'Document.UpdateThumbnail.FirstImage.epub')
        #ExEnd

    def test_hyphenation_options(self):
        #ExStart
        #ExFor:Document.hyphenation_options
        #ExFor:HyphenationOptions
        #ExFor:HyphenationOptions.auto_hyphenation
        #ExFor:HyphenationOptions.consecutive_hyphen_limit
        #ExFor:HyphenationOptions.hyphenation_zone
        #ExFor:HyphenationOptions.hyphenate_caps
        #ExSummary:Shows how to configure automatic hyphenation.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.size = 24
        builder.writeln('Lorem ipsum dolor sit amet, consectetur adipiscing elit, ' + 'sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.')
        doc.hyphenation_options.auto_hyphenation = True
        doc.hyphenation_options.consecutive_hyphen_limit = 2
        doc.hyphenation_options.hyphenation_zone = 720
        doc.hyphenation_options.hyphenate_caps = True
        doc.save(file_name=ARTIFACTS_DIR + 'Document.HyphenationOptions.docx')
        #ExEnd
        self.assertEqual(True, doc.hyphenation_options.auto_hyphenation)
        self.assertEqual(2, doc.hyphenation_options.consecutive_hyphen_limit)
        self.assertEqual(720, doc.hyphenation_options.hyphenation_zone)
        self.assertEqual(True, doc.hyphenation_options.hyphenate_caps)
        self.assertTrue(document_helper.DocumentHelper.compare_docs(ARTIFACTS_DIR + 'Document.HyphenationOptions.docx', GOLDS_DIR + 'Document.HyphenationOptions Gold.docx'))

    def test_hyphenation_options_default_values(self):
        doc = aw.Document()
        doc = document_helper.DocumentHelper.save_open(doc)
        self.assertEqual(False, doc.hyphenation_options.auto_hyphenation)
        self.assertEqual(0, doc.hyphenation_options.consecutive_hyphen_limit)
        self.assertEqual(360, doc.hyphenation_options.hyphenation_zone)  # 0.25 inch
        self.assertEqual(True, doc.hyphenation_options.hyphenate_caps)

    def test_ooxml_compliance_version(self):
        #ExStart
        #ExFor:Document.compliance
        #ExSummary:Shows how to read a loaded document's Open Office XML compliance version.
        # The compliance version varies between documents created by different versions of Microsoft Word.
        doc = aw.Document(file_name=MY_DIR + 'Document.doc')
        self.assertEqual(doc.compliance, aw.saving.OoxmlCompliance.ECMA376_2006)
        doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        self.assertEqual(doc.compliance, aw.saving.OoxmlCompliance.ISO29500_2008_TRANSITIONAL)
        #ExEnd

    @unittest.skip('WORDSNET-20342')
    def test_image_save_options(self):
        #ExStart
        #ExFor:Document.save(str,SaveOptions)
        #ExFor:SaveOptions.use_anti_aliasing
        #ExFor:SaveOptions.use_high_quality_rendering
        #ExSummary:Shows how to improve the quality of a rendered document with SaveOptions.
        doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.size = 60
        builder.writeln('Some text.')
        options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)
        self.assertFalse(options.use_anti_aliasing)  #ExSkip
        self.assertFalse(options.use_high_quality_rendering)  #ExSkip
        doc.save(file_name=ARTIFACTS_DIR + 'Document.ImageSaveOptions.Default.jpg', save_options=options)
        options.use_anti_aliasing = True
        options.use_high_quality_rendering = True
        doc.save(file_name=ARTIFACTS_DIR + 'Document.ImageSaveOptions.HighQuality.jpg', save_options=options)
        #ExEnd
        test_util.TestUtil.verify_image(794, 1122, ARTIFACTS_DIR + 'Document.ImageSaveOptions.Default.jpg')
        test_util.TestUtil.verify_image(794, 1122, ARTIFACTS_DIR + 'Document.ImageSaveOptions.HighQuality.jpg')

    def test_cleanup(self):
        #ExStart
        #ExFor:Document.cleanup
        #ExSummary:Shows how to remove unused custom styles from a document.
        doc = aw.Document()
        doc.styles.add(aw.StyleType.LIST, 'MyListStyle1')
        doc.styles.add(aw.StyleType.LIST, 'MyListStyle2')
        doc.styles.add(aw.StyleType.CHARACTER, 'MyParagraphStyle1')
        doc.styles.add(aw.StyleType.CHARACTER, 'MyParagraphStyle2')
        # Combined with the built-in styles, the document now has eight styles.
        # A custom style counts as "used" while applied to some part of the document,
        # which means that the four styles we added are currently unused.
        self.assertEqual(8, doc.styles.count)
        # Apply a custom character style, and then a custom list style. Doing so will mark the styles as "used".
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.style = doc.styles.get_by_name('MyParagraphStyle1')
        builder.writeln('Hello world!')
        list = doc.lists.add(list_style=doc.styles.get_by_name('MyListStyle1'))
        builder.list_format.list = list
        builder.writeln('Item 1')
        builder.writeln('Item 2')
        doc.cleanup()
        self.assertEqual(6, doc.styles.count)
        # Removing every node that a custom style is applied to marks it as "unused" again.
        # Run the Cleanup method again to remove them.
        doc.first_section.body.remove_all_children()
        doc.cleanup()
        self.assertEqual(4, doc.styles.count)
        #ExEnd

    def test_automatically_update_styles(self):
        #ExStart
        #ExFor:Document.automatically_update_styles
        #ExSummary:Shows how to attach a template to a document.
        doc = aw.Document()
        # Microsoft Word documents by default come with an attached template called "Normal.dotm".
        # There is no default template for blank Aspose.Words documents.
        self.assertEqual('', doc.attached_template)
        # Attach a template, then set the flag to apply style changes
        # within the template to styles in our document.
        doc.attached_template = MY_DIR + 'Business brochure.dotx'
        doc.automatically_update_styles = True
        doc.save(file_name=ARTIFACTS_DIR + 'Document.AutomaticallyUpdateStyles.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Document.AutomaticallyUpdateStyles.docx')
        self.assertTrue(doc.automatically_update_styles)
        self.assertEqual(MY_DIR + 'Business brochure.dotx', doc.attached_template)
        self.assertTrue(system_helper.io.File.exist(doc.attached_template))

    def test_default_template(self):
        #ExStart
        #ExFor:Document.attached_template
        #ExFor:Document.automatically_update_styles
        #ExFor:SaveOptions.create_save_options(str)
        #ExFor:SaveOptions.default_template
        #ExSummary:Shows how to set a default template for documents that do not have attached templates.
        doc = aw.Document()
        # Enable automatic style updating, but do not attach a template document.
        doc.automatically_update_styles = True
        self.assertEqual('', doc.attached_template)
        # Since there is no template document, the document had nowhere to track style changes.
        # Use a SaveOptions object to automatically set a template
        # if a document that we are saving does not have one.
        options = aw.saving.SaveOptions.create_save_options(file_name='Document.DefaultTemplate.docx')
        options.default_template = MY_DIR + 'Business brochure.dotx'
        doc.save(file_name=ARTIFACTS_DIR + 'Document.DefaultTemplate.docx', save_options=options)
        #ExEnd
        self.assertTrue(system_helper.io.File.exist(options.default_template))

    def test_set_invalidate_field_types(self):
        #ExStart
        #ExFor:Document.normalize_field_types
        #ExFor:Range.normalize_field_types
        #ExSummary:Shows how to get the keep a field's type up to date with its field code.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        field = builder.insert_field(field_code='DATE', field_value=None)
        # Aspose.Words automatically detects field types based on field codes.
        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field.type)
        # Manually change the raw text of the field, which determines the field code.
        field_text = doc.first_section.body.first_paragraph.get_child_nodes(aw.NodeType.RUN, True)[0].as_run()
        self.assertEqual('DATE', field_text.text)  #ExSkip
        field_text.text = 'PAGE'
        # Changing the field code has changed this field to one of a different type,
        # but the field's type properties still display the old type.
        self.assertEqual('PAGE', field.get_field_code())
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
        #ExEnd

    def test_layout_options_hidden_text(self):
        for show_hidden_text in [False, True]:
            #ExStart
            #ExFor:Document.layout_options
            #ExFor:LayoutOptions
            #ExFor:LayoutOptions.show_hidden_text
            #ExSummary:Shows how to hide text in a rendered output document.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            self.assertFalse(doc.layout_options.show_hidden_text)  #ExSkip
            # Insert hidden text, then specify whether we wish to omit it from a rendered document.
            builder.writeln('This text is not hidden.')
            builder.font.hidden = True
            builder.writeln('This text is hidden.')
            doc.layout_options.show_hidden_text = show_hidden_text
            doc.save(file_name=ARTIFACTS_DIR + 'Document.LayoutOptionsHiddenText.pdf')
            #ExEnd

    def test_layout_options_paragraph_marks(self):
        for show_paragraph_marks in [False, True]:
            #ExStart
            #ExFor:Document.layout_options
            #ExFor:LayoutOptions
            #ExFor:LayoutOptions.show_paragraph_marks
            #ExSummary:Shows how to show paragraph marks in a rendered output document.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            self.assertFalse(doc.layout_options.show_paragraph_marks)  #ExSkip
            # Add some paragraphs, then enable paragraph marks to show the ends of paragraphs
            # with a pilcrow (¶) symbol when we render the document.
            builder.writeln('Hello world!')
            builder.writeln('Hello again!')
            doc.layout_options.show_paragraph_marks = show_paragraph_marks
            doc.save(file_name=ARTIFACTS_DIR + 'Document.LayoutOptionsParagraphMarks.pdf')
            #ExEnd

    def test_update_page_layout(self):
        #ExStart
        #ExFor:StyleCollection.__getitem__(str)
        #ExFor:SectionCollection.__getitem__(int)
        #ExFor:Document.update_page_layout
        #ExFor:Margins
        #ExFor:PageSetup.margins
        #ExSummary:Shows when to recalculate the page layout of the document.
        doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
        # Saving a document to PDF, to an image, or printing for the first time will automatically
        # cache the layout of the document within its pages.
        doc.save(file_name=ARTIFACTS_DIR + 'Document.UpdatePageLayout.1.pdf')
        # Modify the document in some way.
        doc.styles.get_by_name('Normal').font.size = 6
        doc.sections[0].page_setup.orientation = aw.Orientation.LANDSCAPE
        doc.sections[0].page_setup.margins = aw.Margins.MIRRORED
        # In the current version of Aspose.Words, modifying the document does not automatically rebuild
        # the cached page layout. If we wish for the cached layout
        # to stay up to date, we will need to update it manually.
        doc.update_page_layout()
        doc.save(file_name=ARTIFACTS_DIR + 'Document.UpdatePageLayout.2.pdf')
        #ExEnd

    def test_shade_form_data(self):
        for use_grey_shading in [False, True]:
            #ExStart
            #ExFor:Document.shade_form_data
            #ExSummary:Shows how to apply gray shading to form fields.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            self.assertTrue(doc.shade_form_data)  #ExSkip
            builder.write('Hello world! ')
            builder.insert_text_input('My form field', aw.fields.TextFormFieldType.REGULAR, '', 'Text contents of form field, which are shaded in grey by default.', 0)
            # We can turn the grey shading off, so the bookmarked text will blend in with the other text.
            doc.shade_form_data = use_grey_shading
            doc.save(file_name=ARTIFACTS_DIR + 'Document.ShadeFormData.docx')
            #ExEnd

    def test_versions_count(self):
        #ExStart
        #ExFor:Document.versions_count
        #ExSummary:Shows how to work with the versions count feature of older Microsoft Word documents.
        doc = aw.Document(file_name=MY_DIR + 'Versions.doc')
        # We can read this property of a document, but we cannot preserve it while saving.
        self.assertEqual(4, doc.versions_count)
        doc.save(file_name=ARTIFACTS_DIR + 'Document.VersionsCount.doc')
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Document.VersionsCount.doc')
        self.assertEqual(0, doc.versions_count)
        #ExEnd

    def test_write_protection(self):
        #ExStart
        #ExFor:Document.write_protection
        #ExFor:WriteProtection
        #ExFor:WriteProtection.is_write_protected
        #ExFor:WriteProtection.read_only_recommended
        #ExFor:WriteProtection.set_password(str)
        #ExFor:WriteProtection.validate_password(str)
        #ExSummary:Shows how to protect a document with a password.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Hello world! This document is protected.')
        self.assertFalse(doc.write_protection.is_write_protected)  #ExSkip
        self.assertFalse(doc.write_protection.read_only_recommended)  #ExSkip
        # Enter a password up to 15 characters in length, and then verify the document's protection status.
        doc.write_protection.set_password('MyPassword')
        doc.write_protection.read_only_recommended = True
        self.assertTrue(doc.write_protection.is_write_protected)
        self.assertTrue(doc.write_protection.validate_password('MyPassword'))
        # Protection does not prevent the document from being edited programmatically, nor does it encrypt the contents.
        doc.save(file_name=ARTIFACTS_DIR + 'Document.WriteProtection.docx')
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Document.WriteProtection.docx')
        self.assertTrue(doc.write_protection.is_write_protected)
        builder = aw.DocumentBuilder(doc=doc)
        builder.move_to_document_end()
        builder.writeln('Writing text in a protected document.')
        self.assertEqual('Hello world! This document is protected.' + '\rWriting text in a protected document.', doc.get_text().strip())
        #ExEnd
        self.assertTrue(doc.write_protection.read_only_recommended)
        self.assertTrue(doc.write_protection.validate_password('MyPassword'))
        self.assertFalse(doc.write_protection.validate_password('wrongpassword'))

    def test_remove_personal_information(self):
        for save_without_personal_info in [False, True]:
            #ExStart
            #ExFor:Document.remove_personal_information
            #ExSummary:Shows how to enable the removal of personal information during a manual save.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            # Insert some content with personal information.
            doc.built_in_document_properties.author = 'John Doe'
            doc.built_in_document_properties.company = 'Placeholder Inc.'
            doc.start_track_revisions(author=doc.built_in_document_properties.author, date_time=datetime.datetime.now())
            builder.write('Hello world!')
            doc.stop_track_revisions()
            # This flag is equivalent to File -> Options -> Trust Center -> Trust Center Settings... ->
            # Privacy Options -> "Remove personal information from file properties on save" in Microsoft Word.
            doc.remove_personal_information = save_without_personal_info
            # This option will not take effect during a save operation made using Aspose.Words.
            # Personal data will be removed from our document with the flag set when we save it manually using Microsoft Word.
            doc.save(file_name=ARTIFACTS_DIR + 'Document.RemovePersonalInformation.docx')
            doc = aw.Document(file_name=ARTIFACTS_DIR + 'Document.RemovePersonalInformation.docx')
            self.assertEqual(save_without_personal_info, doc.remove_personal_information)
            self.assertEqual('John Doe', doc.built_in_document_properties.author)
            self.assertEqual('Placeholder Inc.', doc.built_in_document_properties.company)
            self.assertEqual('John Doe', doc.revisions[0].author)
            #ExEnd

    def test_show_comments(self):
        #ExStart
        #ExFor:LayoutOptions.comment_display_mode
        #ExFor:CommentDisplayMode
        #ExSummary:Shows how to show comments when saving a document to a rendered format.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.write('Hello world!')
        comment = aw.Comment(doc=doc, author='John Doe', initial='J.D.', date_time=datetime.datetime.now())
        comment.set_text('My comment.')
        builder.current_paragraph.append_child(comment)
        # ShowInAnnotations is only available in Pdf1.7 and Pdf1.5 formats.
        # In other formats, it will work similarly to Hide.
        doc.layout_options.comment_display_mode = aw.layout.CommentDisplayMode.SHOW_IN_ANNOTATIONS
        doc.save(file_name=ARTIFACTS_DIR + 'Document.ShowCommentsInAnnotations.pdf')
        # Note that it's required to rebuild the document page layout (via Document.UpdatePageLayout() method)
        # after changing the Document.LayoutOptions values.
        doc.layout_options.comment_display_mode = aw.layout.CommentDisplayMode.SHOW_IN_BALLOONS
        doc.update_page_layout()
        doc.save(file_name=ARTIFACTS_DIR + 'Document.ShowCommentsInBalloons.pdf')
        #ExEnd

    def test_copy_template_styles_via_document(self):
        #ExStart
        #ExFor:Document.copy_styles_from_template(Document)
        #ExSummary:Shows how to copies styles from the template to a document via Document.
        template = aw.Document(file_name=MY_DIR + 'Rendering.docx')
        target = aw.Document(file_name=MY_DIR + 'Document.docx')
        self.assertEqual(18, template.styles.count)  #ExSkip
        self.assertEqual(12, target.styles.count)  #ExSkip
        target.copy_styles_from_template(template=template)
        self.assertEqual(22, target.styles.count)  #ExSkip
        #ExEnd

    def test_copy_template_styles_via_document_new(self):
        #ExStart
        #ExFor:Document.copy_styles_from_template(Document)
        #ExFor:Document.copy_styles_from_template(str)
        #ExSummary:Shows how to copy styles from one document to another.
        # Create a document, and then add styles that we will copy to another document.
        template = aw.Document()
        style = template.styles.add(aw.StyleType.PARAGRAPH, 'TemplateStyle1')
        style.font.name = 'Times New Roman'
        style.font.color = aspose.pydrawing.Color.navy
        style = template.styles.add(aw.StyleType.PARAGRAPH, 'TemplateStyle2')
        style.font.name = 'Arial'
        style.font.color = aspose.pydrawing.Color.deep_sky_blue
        style = template.styles.add(aw.StyleType.PARAGRAPH, 'TemplateStyle3')
        style.font.name = 'Courier New'
        style.font.color = aspose.pydrawing.Color.royal_blue
        self.assertEqual(7, template.styles.count)
        # Create a document which we will copy the styles to.
        target = aw.Document()
        # Create a style with the same name as a style from the template document and add it to the target document.
        style = target.styles.add(aw.StyleType.PARAGRAPH, 'TemplateStyle3')
        style.font.name = 'Calibri'
        style.font.color = aspose.pydrawing.Color.orange
        self.assertEqual(5, target.styles.count)
        # There are two ways of calling the method to copy all the styles from one document to another.
        # 1 -  Passing the template document object:
        target.copy_styles_from_template(template=template)
        # Copying styles adds all styles from the template document to the target
        # and overwrites existing styles with the same name.
        self.assertEqual(7, target.styles.count)
        self.assertEqual('Courier New', target.styles.get_by_name('TemplateStyle3').font.name)
        self.assertEqual(aspose.pydrawing.Color.royal_blue.to_argb(), target.styles.get_by_name('TemplateStyle3').font.color.to_argb())
        # 2 -  Passing the local system filename of a template document:
        target.copy_styles_from_template(template=MY_DIR + 'Rendering.docx')
        self.assertEqual(21, target.styles.count)
        #ExEnd

    def test_save_output_parameters(self):
        #ExStart
        #ExFor:SaveOutputParameters
        #ExFor:SaveOutputParameters.content_type
        #ExSummary:Shows how to access output parameters of a document's save operation.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Hello world!')
        # After we save a document, we can access the Internet Media Type (MIME type) of the newly created output document.
        parameters = doc.save(file_name=ARTIFACTS_DIR + 'Document.SaveOutputParameters.doc')
        self.assertEqual('application/msword', parameters.content_type)
        # This property changes depending on the save format.
        parameters = doc.save(file_name=ARTIFACTS_DIR + 'Document.SaveOutputParameters.pdf')
        self.assertEqual('application/pdf', parameters.content_type)
        #ExEnd

    def test_sub_document(self):
        #ExStart
        #ExFor:SubDocument
        #ExFor:SubDocument.node_type
        #ExSummary:Shows how to access a master document's subdocument.
        doc = aw.Document(file_name=MY_DIR + 'Master document.docx')
        sub_documents = doc.get_child_nodes(aw.NodeType.SUB_DOCUMENT, True)
        self.assertEqual(1, sub_documents.count)  #ExSkip
        # This node serves as a reference to an external document, and its contents cannot be accessed.
        sub_document = sub_documents[0].as_sub_document()
        self.assertFalse(sub_document.is_composite)
        #ExEnd

    def test_epub_cover(self):
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Hello world!')
        # When saving to .epub, some Microsoft Word document properties convert to .epub metadata.
        doc.built_in_document_properties.author = 'John Doe'
        doc.built_in_document_properties.title = 'My Book Title'
        # The thumbnail we specify here can become the cover image.
        image = system_helper.io.File.read_all_bytes(IMAGE_DIR + 'Transparent background logo.png')
        doc.built_in_document_properties.thumbnail = image
        doc.save(file_name=ARTIFACTS_DIR + 'Document.EpubCover.epub')

    def test_text_watermark(self):
        #ExStart
        #ExFor:Document.watermark
        #ExFor:Watermark
        #ExFor:Watermark.set_text(str)
        #ExFor:Watermark.set_text(str,TextWatermarkOptions)
        #ExFor:Watermark.remove
        #ExFor:TextWatermarkOptions
        #ExFor:TextWatermarkOptions.font_family
        #ExFor:TextWatermarkOptions.font_size
        #ExFor:TextWatermarkOptions.color
        #ExFor:TextWatermarkOptions.layout
        #ExFor:TextWatermarkOptions.is_semitrasparent
        #ExFor:WatermarkLayout
        #ExFor:WatermarkType
        #ExFor:Watermark.type
        #ExSummary:Shows how to create a text watermark.
        doc = aw.Document()
        # Add a plain text watermark.
        doc.watermark.set_text(text='Aspose Watermark')
        # If we wish to edit the text formatting using it as a watermark,
        # we can do so by passing a TextWatermarkOptions object when creating the watermark.
        text_watermark_options = aw.TextWatermarkOptions()
        text_watermark_options.font_family = 'Arial'
        text_watermark_options.font_size = 36
        text_watermark_options.color = aspose.pydrawing.Color.black
        text_watermark_options.layout = aw.WatermarkLayout.DIAGONAL
        text_watermark_options.is_semitrasparent = False
        doc.watermark.set_text(text='Aspose Watermark', options=text_watermark_options)
        doc.save(file_name=ARTIFACTS_DIR + 'Document.TextWatermark.docx')
        # We can remove a watermark from a document like this.
        if doc.watermark.type == aw.WatermarkType.TEXT:
            doc.watermark.remove()
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Document.TextWatermark.docx')
        self.assertEqual(aw.WatermarkType.TEXT, doc.watermark.type)

    def test_spelling_and_grammar_errors(self):
        for show_errors in [False, True]:
            #ExStart
            #ExFor:Document.show_grammatical_errors
            #ExFor:Document.show_spelling_errors
            #ExSummary:Shows how to show/hide errors in the document.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            # Insert two sentences with mistakes that would be picked up
            # by the spelling and grammar checkers in Microsoft Word.
            builder.writeln('There is a speling error in this sentence.')
            builder.writeln('Their is a grammatical error in this sentence.')
            # If these options are enabled, then spelling errors will be underlined
            # in the output document by a jagged red line, and a double blue line will highlight grammatical mistakes.
            doc.show_grammatical_errors = show_errors
            doc.show_spelling_errors = show_errors
            doc.save(file_name=ARTIFACTS_DIR + 'Document.SpellingAndGrammarErrors.docx')
            #ExEnd
            doc = aw.Document(file_name=ARTIFACTS_DIR + 'Document.SpellingAndGrammarErrors.docx')
            self.assertEqual(show_errors, doc.show_grammatical_errors)
            self.assertEqual(show_errors, doc.show_spelling_errors)

    def test_ignore_printer_metrics(self):
        #ExStart
        #ExFor:LayoutOptions.ignore_printer_metrics
        #ExSummary:Shows how to ignore 'Use printer metrics to lay out document' option.
        doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
        doc.layout_options.ignore_printer_metrics = False
        doc.save(file_name=ARTIFACTS_DIR + 'Document.IgnorePrinterMetrics.docx')
        #ExEnd

    def test_extract_pages(self):
        #ExStart
        #ExFor:Document.extract_pages
        #ExSummary:Shows how to get specified range of pages from the document.
        doc = aw.Document(file_name=MY_DIR + 'Layout entities.docx')
        doc = doc.extract_pages(0, 2)
        doc.save(file_name=ARTIFACTS_DIR + 'Document.ExtractPages.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Document.ExtractPages.docx')
        self.assertEqual(doc.page_count, 2)

    def test_spelling_or_grammar(self):
        for check_spelling_grammar in [True, False]:
            #ExStart
            #ExFor:Document.spelling_checked
            #ExFor:Document.grammar_checked
            #ExSummary:Shows how to set spelling or grammar verifying.
            doc = aw.Document()
            # The string with spelling errors.
            doc.first_section.body.first_paragraph.runs.add(aw.Run(doc=doc, text='The speeling in this documentz is all broked.'))
            # Spelling/Grammar check start if we set properties to false.
            # We can see all errors in Microsoft Word via Review -> Spelling & Grammar.
            # Note that Microsoft Word does not start grammar/spell check automatically for DOC and RTF document format.
            doc.spelling_checked = check_spelling_grammar
            doc.grammar_checked = check_spelling_grammar
            doc.save(file_name=ARTIFACTS_DIR + 'Document.SpellingOrGrammar.docx')
            #ExEnd

    def test_allow_embedding_post_script_fonts(self):
        #ExStart
        #ExFor:SaveOptions.allow_embedding_post_script_fonts
        #ExSummary:Shows how to save the document with PostScript font.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.name = 'PostScriptFont'
        builder.writeln('Some text with PostScript font.')
        # Load the font with PostScript to use in the document.
        otf = aw.fonts.MemoryFontSource(font_data=system_helper.io.File.read_all_bytes(FONTS_DIR + 'AllegroOpen.otf'))
        doc.font_settings = aw.fonts.FontSettings()
        doc.font_settings.set_fonts_sources(sources=[otf])
        # Embed TrueType fonts.
        doc.font_infos.embed_true_type_fonts = True
        # Allow embedding PostScript fonts while embedding TrueType fonts.
        # Microsoft Word does not embed PostScript fonts, but can open documents with embedded fonts of this type.
        save_options = aw.saving.SaveOptions.create_save_options(save_format=aw.SaveFormat.DOCX)
        save_options.allow_embedding_post_script_fonts = True
        doc.save(file_name=ARTIFACTS_DIR + 'Document.AllowEmbeddingPostScriptFonts.docx', save_options=save_options)
        #ExEnd

    def test_frameset(self):
        #ExStart
        #ExFor:Document.frameset
        #ExFor:Frameset
        #ExFor:Frameset.frame_default_url
        #ExFor:Frameset.is_frame_link_to_file
        #ExFor:Frameset.child_framesets
        #ExFor:FramesetCollection
        #ExFor:FramesetCollection.count
        #ExFor:FramesetCollection.__getitem__(int)
        #ExSummary:Shows how to access frames on-page.
        # Document contains several frames with links to other documents.
        doc = aw.Document(file_name=MY_DIR + 'Frameset.docx')
        self.assertEqual(3, doc.frameset.child_framesets.count)
        # We can check the default URL (a web page URL or local document) or if the frame is an external resource.
        self.assertEqual('https://file-examples-com.github.io/uploads/2017/02/file-sample_100kB.docx', doc.frameset.child_framesets[0].child_framesets[0].frame_default_url)
        self.assertTrue(doc.frameset.child_framesets[0].child_framesets[0].is_frame_link_to_file)
        self.assertEqual('Document.docx', doc.frameset.child_framesets[1].frame_default_url)
        self.assertFalse(doc.frameset.child_framesets[1].is_frame_link_to_file)
        # Change properties for one of our frames.
        doc.frameset.child_framesets[0].child_framesets[0].frame_default_url = 'https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/Examples/Data/Absolute%20position%20tab.docx'
        doc.frameset.child_framesets[0].child_framesets[0].is_frame_link_to_file = False
        #ExEnd
        doc = document_helper.DocumentHelper.save_open(doc)
        self.assertEqual('https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/Examples/Data/Absolute%20position%20tab.docx', doc.frameset.child_framesets[0].child_framesets[0].frame_default_url)
        self.assertFalse(doc.frameset.child_framesets[0].child_framesets[0].is_frame_link_to_file)

    def test_open_azw(self):
        info = aw.FileFormatUtil.detect_file_format(file_name=MY_DIR + 'Azw3 document.azw3')
        self.assertEqual(info.load_format, aw.LoadFormat.AZW3)
        doc = aw.Document(file_name=MY_DIR + 'Azw3 document.azw3')
        self.assertTrue('Hachette Book Group USA' in doc.get_text())

    def test_open_epub(self):
        info = aw.FileFormatUtil.detect_file_format(file_name=MY_DIR + 'Epub document.epub')
        self.assertEqual(info.load_format, aw.LoadFormat.EPUB)
        doc = aw.Document(file_name=MY_DIR + 'Epub document.epub')
        self.assertTrue('Down the Rabbit-Hole' in doc.get_text())

    def test_open_xml(self):
        info = aw.FileFormatUtil.detect_file_format(file_name=MY_DIR + 'Mail merge data - Customers.xml')
        self.assertEqual(info.load_format, aw.LoadFormat.XML)
        doc = aw.Document(file_name=MY_DIR + 'Mail merge data - Purchase order.xml')
        self.assertTrue('Ellen Adams\r123 Maple Street' in doc.get_text())

    def test_move_to_structured_document_tag(self):
        #ExStart
        #ExFor:DocumentBuilder.move_to_structured_document_tag(int,int)
        #ExFor:DocumentBuilder.move_to_structured_document_tag(StructuredDocumentTag,int)
        #ExFor:DocumentBuilder.is_at_end_of_structured_document_tag
        #ExFor:DocumentBuilder.current_structured_document_tag
        #ExSummary:Shows how to move cursor of DocumentBuilder inside a structured document tag.
        doc = aw.Document(file_name=MY_DIR + 'Structured document tags.docx')
        builder = aw.DocumentBuilder(doc=doc)
        # There is a several ways to move the cursor:
        # 1 -  Move to the first character of structured document tag by index.
        builder.move_to_structured_document_tag(structured_document_tag_index=1, character_index=1)
        # 2 -  Move to the first character of structured document tag by object.
        tag = doc.get_child(aw.NodeType.STRUCTURED_DOCUMENT_TAG, 2, True).as_structured_document_tag()
        builder.move_to_structured_document_tag(structured_document_tag=tag, character_index=1)
        builder.write(' New text.')
        self.assertEqual('R New text.ichText', tag.get_text().strip())
        # 3 -  Move to the end of the second structured document tag.
        builder.move_to_structured_document_tag(structured_document_tag_index=1, character_index=-1)
        self.assertTrue(builder.is_at_end_of_structured_document_tag)
        # Get currently selected structured document tag.
        builder.current_structured_document_tag.color = aspose.pydrawing.Color.green
        doc.save(file_name=ARTIFACTS_DIR + 'Document.MoveToStructuredDocumentTag.docx')
        #ExEnd

    def test_include_textboxes_footnotes_endnotes_in_stat(self):
        #ExStart
        #ExFor:Document.include_textboxes_footnotes_endnotes_in_stat
        #ExSummary: Shows how to include or exclude textboxes, footnotes and endnotes from word count statistics.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Lorem ipsum')
        builder.insert_footnote(footnote_type=aw.notes.FootnoteType.FOOTNOTE, footnote_text='sit amet')
        # By default option is set to 'false'.
        doc.update_word_count()
        # Words count without textboxes, footnotes and endnotes.
        self.assertEqual(2, doc.built_in_document_properties.words)
        doc.include_textboxes_footnotes_endnotes_in_stat = True
        doc.update_word_count()
        # Words count with textboxes, footnotes and endnotes.
        self.assertEqual(4, doc.built_in_document_properties.words)
        #ExEnd

    def test_set_justification_mode(self):
        #ExStart
        #ExFor:Document.justification_mode
        #ExFor:JustificationMode
        #ExSummary:Shows how to manage character spacing control.
        doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        justification_mode = doc.justification_mode
        if justification_mode == aw.settings.JustificationMode.EXPAND:
            doc.justification_mode = aw.settings.JustificationMode.COMPRESS
        doc.save(file_name=ARTIFACTS_DIR + 'Document.SetJustificationMode.docx')
        #ExEnd

    def test_page_is_in_color(self):
        #ExStart
        #ExFor:PageInfo.colored
        #ExFor:Document.get_page_info(int)
        #ExSummary:Shows how to check whether the page is in color or not.
        doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        # Check that the first page of the document is not colored.
        self.assertFalse(doc.get_page_info(0).colored)
        #ExEnd

    def test_insert_document_inline(self):
        #ExStart:InsertDocumentInline
        #ExFor:DocumentBuilder.insert_document_inline(Document,ImportFormatMode,ImportFormatOptions)
        #ExSummary:Shows how to insert a document inline at the cursor position.
        src_doc = aw.DocumentBuilder()
        src_doc.write('[src content]')
        # Create destination document.
        dst_doc = aw.DocumentBuilder()
        dst_doc.write('Before ')
        dst_doc.insert_node(aw.BookmarkStart(dst_doc.document, 'src_place'))
        dst_doc.insert_node(aw.BookmarkEnd(dst_doc.document, 'src_place'))
        dst_doc.write(' after')
        self.assertEqual('Before  after', dst_doc.document.get_text().rstrip())
        # Insert source document into destination inline.
        dst_doc.move_to_bookmark(bookmark_name='src_place')
        dst_doc.insert_document_inline(src_doc.document, aw.ImportFormatMode.USE_DESTINATION_STYLES, aw.ImportFormatOptions())
        self.assertEqual('Before [src content] after', dst_doc.document.get_text().rstrip())
        #ExEnd:InsertDocumentInline

    def test_save_document_to_stream(self):
        for save_format in [aw.SaveFormat.DOC, aw.SaveFormat.DOT, aw.SaveFormat.DOCX, aw.SaveFormat.DOCM, aw.SaveFormat.DOTX, aw.SaveFormat.DOTM, aw.SaveFormat.FLAT_OPC, aw.SaveFormat.FLAT_OPC_MACRO_ENABLED, aw.SaveFormat.FLAT_OPC_TEMPLATE, aw.SaveFormat.FLAT_OPC_TEMPLATE_MACRO_ENABLED, aw.SaveFormat.RTF, aw.SaveFormat.WORD_ML, aw.SaveFormat.PDF, aw.SaveFormat.XPS, aw.SaveFormat.XAML_FIXED, aw.SaveFormat.SVG, aw.SaveFormat.HTML_FIXED, aw.SaveFormat.OPEN_XPS, aw.SaveFormat.PS, aw.SaveFormat.PCL, aw.SaveFormat.HTML, aw.SaveFormat.MHTML, aw.SaveFormat.EPUB, aw.SaveFormat.AZW3, aw.SaveFormat.MOBI, aw.SaveFormat.ODT, aw.SaveFormat.OTT, aw.SaveFormat.TEXT, aw.SaveFormat.XAML_FLOW, aw.SaveFormat.XAML_FLOW_PACK, aw.SaveFormat.MARKDOWN, aw.SaveFormat.XLSX, aw.SaveFormat.TIFF, aw.SaveFormat.PNG, aw.SaveFormat.BMP, aw.SaveFormat.EMF, aw.SaveFormat.JPEG, aw.SaveFormat.GIF, aw.SaveFormat.EPS]:
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.writeln('Lorem ipsum')
            with io.BytesIO() as stream:
                if save_format == aw.SaveFormat.HTML_FIXED:
                    save_options = aw.saving.HtmlFixedSaveOptions()
                    save_options.export_embedded_css = True
                    save_options.export_embedded_fonts = True
                    save_options.save_format = save_format
                    doc.save(stream=stream, save_options=save_options)
                elif save_format == aw.SaveFormat.XAML_FIXED:
                    save_options = aw.saving.XamlFixedSaveOptions()
                    save_options.resources_folder = ARTIFACTS_DIR
                    save_options.save_format = save_format
                    doc.save(stream=stream, save_options=save_options)
                else:
                    doc.save(stream=stream, save_format=save_format)

    def test_has_macros(self):
        #ExStart:HasMacros
        #ExFor:FileFormatInfo.has_macros
        #ExSummary:Shows how to check VBA macro presence without loading document.
        file_format_info = aw.FileFormatUtil.detect_file_format(file_name=MY_DIR + 'Macro.docm')
        self.assertTrue(file_format_info.has_macros)
        #ExEnd:HasMacros

    def test_punctuation_kerning(self):
        #ExStart
        #ExFor:Document.punctuation_kerning
        #ExSummary:Shows how to work with kerning applies to both Latin text and punctuation.
        doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        self.assertTrue(doc.punctuation_kerning)
        #ExEnd

    def test_remove_blank_pages(self):
        #ExStart
        #ExFor:Document.remove_blank_pages
        #ExSummary:Shows how to remove blank pages from the document.
        doc = aw.Document(file_name=MY_DIR + 'Blank pages.docx')
        self.assertEqual(2, doc.page_count)
        doc.remove_blank_pages()
        doc.update_page_layout()
        self.assertEqual(1, doc.page_count)
        #ExEnd

    def test_create_simple_document(self):
        #ExStart:CreateSimpleDocument
        #ExFor:Document.__init__()
        #ExSummary:Shows how to create simple document.
        doc = aw.Document()
        # New Document objects by default come with the minimal set of nodes
        # required to begin adding content such as text and shapes: a Section, a Body, and a Paragraph.
        section = doc.append_child(aw.Section(doc)).as_section()
        body = section.append_child(aw.Body(doc)).as_body()
        para = body.append_child(aw.Paragraph(doc)).as_paragraph()
        para.append_child(aw.Run(doc=doc, text='Hello world!'))
        #ExEnd:CreateSimpleDocument

    def test_load_from_web(self):
        #ExStart
        #ExFor:Document.__init__(BytesIO)
        #ExSummary:Shows how to load a document from a URL.
        # Create a URL that points to a Microsoft Word document.
        url = 'https://filesamples.com/samples/document/docx/sample3.docx'
        # Download the document into a byte array, then load that array into a document using a memory stream.
        request_site = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        data_bytes = urlopen(request_site).read()
        with io.BytesIO(data_bytes) as byte_stream:
            doc = aw.Document(byte_stream)
            # At this stage, we can read and edit the document's contents and then save it to the local file system.
            self.assertEqual('There are eight section headings in this document. At the beginning, "Sample Document" is a level 1 heading. ' + 'The main section headings, such as "Headings" and "Lists" are level 2 headings. ' + 'The Tables section contains two sub-headings, "Simple Table" and "Complex Table," which are both level 3 headings.', doc.first_section.body.paragraphs[3].get_text().strip())
            doc.save(ARTIFACTS_DIR + 'Document.load_from_web.docx')
        #ExEnd

    @unittest.skipUnless(sys.platform.startswith('win'), 'requires windows')
    def test_save_to_image_stream(self):
        #ExStart
        #ExFor:Document.save(BytesIO,SaveFormat)
        #ExSummary:Shows how to save a document to an image via stream, and then read the image from that stream.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.font.name = 'Times New Roman'
        builder.font.size = 24
        builder.writeln('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.')
        builder.insert_image(IMAGE_DIR + 'Logo.jpg')
        with io.BytesIO() as stream:
            doc.save(stream, aw.SaveFormat.BMP)
            stream.seek(0, os.SEEK_SET)
            # Read the stream back into an image.
            with aspose.pydrawing.Image.from_stream(stream) as image:
                self.assertEqual(aspose.pydrawing.imaging.ImageFormat.bmp, image.raw_format)
                self.assertEqual(816, image.width)
                self.assertEqual(1056, image.height)
        #ExEnd

    def test_insert_html_from_web_page(self):
        #ExStart
        #ExFor:Document.__init__(BytesIO,LoadOptions)
        #ExFor:LoadOptions.__init__(LoadFormat,str,str)
        #ExFor:LoadFormat
        #ExSummary:Shows how save a web page as a .docx file.
        url = 'https://products.aspose.com/words/'
        with io.BytesIO(urlopen(url).read()) as stream:
            # The URL is used again as a "base_uri" to ensure that any relative image paths are retrieved correctly.
            options = aw.loading.LoadOptions(aw.LoadFormat.HTML, '', url)
            # Load the HTML document from stream and pass the LoadOptions object.
            doc = aw.Document(stream, options)
            # At this stage, we can read and edit the document's contents and then save it to the local file system.
            self.assertTrue(doc.get_text().find('HYPERLINK "https://products.aspose.com/words/net/" \\o "Aspose.Words"') > 0)  #ExSkip
            doc.save(ARTIFACTS_DIR + 'Document.insert_html_from_web_page.docx')
        #ExEnd
        self.verify_web_response_status_code(200, url)

    def test_import_list(self):
        for is_keep_source_numbering in (True, False):
            with self.subTest(is_keep_source_numbering=is_keep_source_numbering):
                #ExStart
                #ExFor:ImportFormatOptions.keep_source_numbering
                #ExSummary:Shows how to import a document with numbered lists.
                src_doc = aw.Document(MY_DIR + 'List source.docx')
                dst_doc = aw.Document(MY_DIR + 'List destination.docx')
                self.assertEqual(4, dst_doc.lists.count)
                options = aw.ImportFormatOptions()
                # If there is a clash of list styles, apply the list format of the source document.
                # Set the "keep_source_numbering" property to "False" to not import any list numbers into the destination document.
                # Set the "keep_source_numbering" property to "True" import all clashing
                # list style numbering with the same appearance that it had in the source document.
                options.keep_source_numbering = is_keep_source_numbering
                dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, options)
                dst_doc.update_list_labels()
                if is_keep_source_numbering:
                    self.assertEqual(5, dst_doc.lists.count)
                else:
                    self.assertEqual(4, dst_doc.lists.count)
                #ExEnd

    def test_validate_individual_document_signatures(self):
        #ExStart
        #ExFor:CertificateHolder.certificate
        #ExFor:Document.digital_signatures
        #ExFor:DigitalSignature
        #ExFor:DigitalSignatureCollection
        #ExFor:DigitalSignature.is_valid
        #ExFor:DigitalSignature.comments
        #ExFor:DigitalSignature.sign_time
        #ExFor:DigitalSignature.signature_type
        #ExSummary:Shows how to validate and display information about each signature in a document.
        doc = aw.Document(MY_DIR + 'Digitally signed.docx')
        for signature in doc.digital_signatures:
            print(f"\n{('Valid' if signature.is_valid else 'Invalid')} signature: ")
            print(f'\tReason:\t{signature.comments}')
            print(f'\tType:\t{signature.signature_type}')
            print(f'\tSign time:\t{signature.sign_time}')
            # System.Security.Cryptography.X509Certificates.X509Certificate2 is not supported. That is why the following information is not accesible.
            #print(f"\tSubject name:\t{signature.certificate_holder.certificate.subject_name}")
            #print(f"\tIssuer name:\t{signature.certificate_holder.certificate.issuer_name.name}")
            print()
        #ExEnd
        self.assertEqual(1, doc.digital_signatures.count)
        digital_sig = doc.digital_signatures[0]
        self.assertTrue(digital_sig.is_valid)
        self.assertEqual('Test Sign', digital_sig.comments)
        self.assertEqual(aw.digitalsignatures.DigitalSignatureType.XML_DSIG, digital_sig.signature_type)
        # System.Security.Cryptography.X509Certificates.X509Certificate2 is not supported. That is why the following information is not accesible.
        # self.assertTrue(digital_sig.certificate_holder.certificate.subject.contains("Aspose Pty Ltd"))
        # self.assertIsNotNone(digital_sig.certificate_holder.certificate.issuer_name.name is not None)
        # self.assertIn("VeriSign", digital_sig.certificate_holder.certificate.issuer_name.name)

    def test_signature_value(self):
        #ExStart
        #ExFor:DigitalSignature.signature_value
        #ExSummary:Shows how to get a digital signature value from a digitally signed document.
        doc = aw.Document(MY_DIR + 'Digitally signed.docx')
        for digital_signature_val in doc.digital_signatures:
            signature_value = base64.b64encode(digital_signature_val.signature_value)
            self.assertEqual(b'K1cVLLg2kbJRAzT5WK+m++G8eEO+l7S+5ENdjMxxTXkFzGUfvwxREuJdSFj9AbDMhnGvDURv9KEhC25DDF1al8NRVR71TF3CjHVZXpYu7edQS5/yLw/k5CiFZzCp1+MmhOdYPcVO+Fm+9fKr2iNLeyYB+fgEeZHfTqTFM2WwAqo=', signature_value)
        #ExEnd

    def test_default_tab_stop(self):
        #ExStart
        #ExFor:Document.default_tab_stop
        #ExFor:ControlChar.tab
        #ExFor:ControlChar.tab_char
        #ExSummary:Shows how to set a custom interval for tab stop positions.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Set tab stops to appear every 72 points (1 inch).
        builder.document.default_tab_stop = 72
        # Each tab character snaps the text after it to the next closest tab stop position.
        builder.writeln('Hello' + aw.ControlChar.TAB + 'World!')
        #ExEnd
        doc = document_helper.DocumentHelper.save_open(doc)
        self.assertEqual(72, doc.default_tab_stop)

    def test_use_substitutions(self):
        #ExStart
        #ExFor:FindReplaceOptions.use_substitutions
        #ExFor:FindReplaceOptions.legacy_mode
        #ExSummary:Shows how to recognize and use substitutions within replacement patterns.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.write('Jason gave money to Paul.')
        options = aw.replacing.FindReplaceOptions()
        options.use_substitutions = True
        # Using legacy mode does not support many advanced features, so we need to set it to 'False'.
        options.legacy_mode = False
        doc.range.replace_regex('([A-z]+) gave money to ([A-z]+)', '$2 took money from $1', options)
        self.assertEqual(doc.get_text(), 'Paul took money from Jason.\x0c')
        #ExEnd

    def test_doc_package_custom_parts(self):
        #ExStart
        #ExFor:CustomPart
        #ExFor:CustomPart.content_type
        #ExFor:CustomPart.relationship_type
        #ExFor:CustomPart.is_external
        #ExFor:CustomPart.data
        #ExFor:CustomPart.name
        #ExFor:CustomPart.clone
        #ExFor:CustomPartCollection
        #ExFor:CustomPartCollection.add(CustomPart)
        #ExFor:CustomPartCollection.clear
        #ExFor:CustomPartCollection.clone
        #ExFor:CustomPartCollection.count
        #ExFor:CustomPartCollection.__iter__
        #ExFor:CustomPartCollection.__getitem__(int)
        #ExFor:CustomPartCollection.remove_at(int)
        #ExFor:Document.package_custom_parts
        #ExSummary:Shows how to access a document's arbitrary custom parts collection.
        doc = aw.Document(MY_DIR + 'Custom parts OOXML package.docx')
        self.assertEqual(2, doc.package_custom_parts.count)
        # Clone the second part, then add the clone to the collection.
        cloned_part = doc.package_custom_parts[1].clone()
        doc.package_custom_parts.add(cloned_part)
        self._test_doc_package_custom_parts(doc.package_custom_parts)  #ExSkip
        self.assertEqual(3, doc.package_custom_parts.count)
        # Enumerate over the collection and print every part.
        for index, part in enumerate(doc.package_custom_parts):
            print(f'Part index {index}:')
            print(f'\tName:\t\t\t\t{part.name}')
            print(f'\tContent type:\t\t{part.content_type}')
            print(f'\tRelationship type:\t{part.relationship_type}')
            if part.is_external:
                print('\tSourced from outside the document')
            else:
                print(f'\tStored within the document, length: {len(part.data)} bytes')
        # We can remove elements from this collection individually, or all at once.
        doc.package_custom_parts.remove_at(2)
        self.assertEqual(2, doc.package_custom_parts.count)
        doc.package_custom_parts.clear()
        self.assertEqual(0, doc.package_custom_parts.count)
        #ExEnd

    def test_read_macros_from_existing_document(self):
        #ExStart
        #ExFor:Document.vba_project
        #ExFor:VbaModuleCollection
        #ExFor:VbaModuleCollection.count
        #ExFor:VbaModuleCollection.__getitem__(int)
        #ExFor:VbaModuleCollection.__getitem__(string)
        #ExFor:VbaModuleCollection.remove
        #ExFor:VbaModule
        #ExFor:VbaModule.name
        #ExFor:VbaModule.source_code
        #ExFor:VbaProject
        #ExFor:VbaProject.name
        #ExFor:VbaProject.modules
        #ExFor:VbaProject.code_page
        #ExFor:VbaProject.is_signed
        #ExSummary:Shows how to access a document's VBA project information.
        doc = aw.Document(MY_DIR + 'VBA project.docm')
        # A VBA project contains a collection of VBA modules.
        vba_project = doc.vba_project
        self.assertTrue(vba_project.is_signed)  #ExSkip
        if vba_project.is_signed:
            print(f'Project name: {vba_project.name} signed; Project code page: {vba_project.code_page}; Modules count: {vba_project.modules.count}\n')
        else:
            print(f'Project name: {vba_project.name} not signed; Project code page: {vba_project.code_page}; Modules count: {vba_project.modules.count}\n')
        vba_modules = doc.vba_project.modules
        self.assertEqual(vba_modules.count, 3)
        for module in vba_modules:
            print(f'Module name: {module.name};\nModule code:\n{module.source_code}\n')
        # Set new source code for VBA module. You can access VBA modules in the collection either by index or by name.
        vba_modules[0].source_code = 'Your VBA code...'
        vba_modules.get_by_name('Module1').source_code = 'Your VBA code...'
        # Remove a module from the collection.
        vba_modules.remove(vba_modules[2])
        #ExEnd
        self.assertEqual('AsposeVBAtest', vba_project.name)
        self.assertEqual(2, vba_project.modules.count)
        self.assertEqual(1251, vba_project.code_page)
        self.assertFalse(vba_project.is_signed)
        self.assertEqual('ThisDocument', vba_modules[0].name)
        self.assertEqual('Your VBA code...', vba_modules[0].source_code)
        self.assertEqual('Module1', vba_modules[1].name)
        self.assertEqual('Your VBA code...', vba_modules[1].source_code)

    def test_create_web_extension(self):
        #ExStart
        #ExFor:BaseWebExtensionCollection.add()
        #ExFor:BaseWebExtensionCollection.clear
        #ExFor:TaskPane
        #ExFor:TaskPane.dock_state
        #ExFor:TaskPane.is_visible
        #ExFor:TaskPane.width
        #ExFor:TaskPane.is_locked
        #ExFor:TaskPane.web_extension
        #ExFor:TaskPane.row
        #ExFor:WebExtension
        #ExFor:WebExtension.reference
        #ExFor:WebExtension.properties
        #ExFor:WebExtension.bindings
        #ExFor:WebExtension.is_frozen
        #ExFor:WebExtensionReference.id
        #ExFor:WebExtensionReference.version
        #ExFor:WebExtensionReference.store_type
        #ExFor:WebExtensionReference.store
        #ExFor:WebExtensionPropertyCollection
        #ExFor:WebExtensionBindingCollection
        #ExFor:WebExtensionProperty.__init__(str,str)
        #ExFor:WebExtensionBinding.__init__(str,WebExtensionBindingType,str)
        #ExFor:WebExtensionStoreType
        #ExFor:WebExtensionBindingType
        #ExFor:TaskPaneDockState
        #ExFor:TaskPaneCollection
        #ExSummary:Shows how to add a web extension to a document.
        doc = aw.Document()
        # Create task pane with "MyScript" add-in, which will be used by the document,
        # then set its default location.
        my_script_task_pane = aw.webextensions.TaskPane()
        doc.web_extension_task_panes.add(my_script_task_pane)
        my_script_task_pane.dock_state = aw.webextensions.TaskPaneDockState.RIGHT
        my_script_task_pane.is_visible = True
        my_script_task_pane.width = 300
        my_script_task_pane.is_locked = True
        # If there are multiple task panes in the same docking location, we can set this index to arrange them.
        my_script_task_pane.row = 1
        # Create an add-in called "MyScript Math Sample", which the task pane will display within.
        web_extension = my_script_task_pane.web_extension
        # Set application store reference parameters for our add-in, such as the ID.
        web_extension.reference.id = 'WA104380646'
        web_extension.reference.version = '1.0.0.0'
        web_extension.reference.store_type = aw.webextensions.WebExtensionStoreType.OMEX
        web_extension.reference.store = 'en-US'
        web_extension.properties.add(aw.webextensions.WebExtensionProperty('MyScript', 'MyScript Math Sample'))
        web_extension.bindings.add(aw.webextensions.WebExtensionBinding('MyScript', aw.webextensions.WebExtensionBindingType.TEXT, '104380646'))
        # Allow the user to interact with the add-in.
        web_extension.is_frozen = False
        # We can access the web extension in Microsoft Word via Developer -> Add-ins.
        doc.save(ARTIFACTS_DIR + 'Document.create_web_extension.docx')
        # Remove all web extension task panes at once like this.
        doc.web_extension_task_panes.clear()
        self.assertEqual(0, doc.web_extension_task_panes.count)
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Document.create_web_extension.docx')
        my_script_task_pane = doc.web_extension_task_panes[0]
        self.assertEqual(aw.webextensions.TaskPaneDockState.RIGHT, my_script_task_pane.dock_state)
        self.assertTrue(my_script_task_pane.is_visible)
        self.assertEqual(300.0, my_script_task_pane.width)
        self.assertTrue(my_script_task_pane.is_locked)
        self.assertEqual(1, my_script_task_pane.row)
        web_extension = my_script_task_pane.web_extension
        self.assertEqual('WA104380646', web_extension.reference.id)
        self.assertEqual('1.0.0.0', web_extension.reference.version)
        self.assertEqual(aw.webextensions.WebExtensionStoreType.OMEX, web_extension.reference.store_type)
        self.assertEqual('en-US', web_extension.reference.store)
        self.assertEqual('MyScript', web_extension.properties[0].name)
        self.assertEqual('MyScript Math Sample', web_extension.properties[0].value)
        self.assertEqual('MyScript', web_extension.bindings[0].id)
        self.assertEqual(aw.webextensions.WebExtensionBindingType.TEXT, web_extension.bindings[0].binding_type)
        self.assertEqual('104380646', web_extension.bindings[0].app_ref)
        self.assertFalse(web_extension.is_frozen)

    def test_get_web_extension_info(self):
        #ExStart
        #ExFor:BaseWebExtensionCollection
        #ExFor:BaseWebExtensionCollection.__iter__
        #ExFor:BaseWebExtensionCollection.remove(int)
        #ExFor:BaseWebExtensionCollection.count
        #ExFor:BaseWebExtensionCollection.__getitem__(int)
        #ExSummary:Shows how to work with a document's collection of web extensions.
        doc = aw.Document(MY_DIR + 'Web extension.docx')
        self.assertEqual(1, doc.web_extension_task_panes.count)
        #print all properties of the document's web extension.
        web_extension_property_collection = doc.web_extension_task_panes[0].web_extension.properties
        for web_extension_property in web_extension_property_collection:
            print(f'Binding name: {web_extension_property.name}; Binding value: {web_extension_property.value}')
        # Remove the web extension.
        doc.web_extension_task_panes.remove(0)
        self.assertEqual(0, doc.web_extension_task_panes.count)
        #ExEnd

    @unittest.skip("drawing.Image type isn't supported yet")
    def test_image_watermark(self):
        #ExStart
        #ExFor:Watermark.set_image(Image,ImageWatermarkOptions)
        #ExFor:ImageWatermarkOptions.scale
        #ExFor:ImageWatermarkOptions.is_washout
        #ExSummary:Shows how to create a watermark from an image in the local file system.
        doc = aw.Document()
        # Modify the image watermark's appearance with an ImageWatermarkOptions object,
        # then pass it while creating a watermark from an image file.
        image_watermark_options = aw.ImageWatermarkOptions()
        image_watermark_options.scale = 5
        image_watermark_options.is_washout = False
        doc.watermark.set_image(drawing.Image.from_file(IMAGE_DIR + 'Logo.jpg'), image_watermark_options)
        doc.save(ARTIFACTS_DIR + 'Document.image_watermark.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Document.image_watermark.docx')
        self.assertEqual(aw.WatermarkType.IMAGE, doc.watermark.type)

    def _test_doc_package_custom_parts(self, parts: aw.markup.CustomPartCollection):
        self.assertEqual(3, parts.count)
        self.assertEqual('/payload/payload_on_package.test', parts[0].name)
        self.assertEqual('mytest/somedata', parts[0].content_type)
        self.assertEqual('http://mytest.payload.internal', parts[0].relationship_type)
        self.assertEqual(False, parts[0].is_external)
        self.assertEqual(18, len(parts[0].data))
        self.assertEqual('http://www.aspose.com/Images/aspose-logo.jpg', parts[1].name)
        self.assertEqual('', parts[1].content_type)
        self.assertEqual('http://mytest.payload.external', parts[1].relationship_type)
        self.assertTrue(parts[1].is_external)
        self.assertEqual(0, len(parts[1].data))
        self.assertEqual('http://www.aspose.com/Images/aspose-logo.jpg', parts[2].name)
        self.assertEqual('', parts[2].content_type)
        self.assertEqual('http://mytest.payload.external', parts[2].relationship_type)
        self.assertTrue(parts[2].is_external)
        self.assertEqual(0, len(parts[2].data))