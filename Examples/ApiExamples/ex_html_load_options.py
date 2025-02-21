# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import sys
import io
from datetime import datetime
import aspose.words as aw
import aspose.words.digitalsignatures
import aspose.words.drawing
import aspose.words.loading
import datetime
import test_util
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, IMAGE_DIR, MY_DIR, IMAGE_URL

class ExHtmlLoadOptions(ApiExampleBase):

    @unittest.skipUnless(sys.platform.startswith('win'), 'different calculation on Linux')
    def test_support_vml(self):
        for support_vml in [True, False]:
            #ExStart
            #ExFor:HtmlLoadOptions
            #ExFor:HtmlLoadOptions.__init__
            #ExFor:HtmlLoadOptions.support_vml
            #ExSummary:Shows how to support conditional comments while loading an HTML document.
            load_options = aw.loading.HtmlLoadOptions()
            # If the value is true, then we take VML code into account while parsing the loaded document.
            load_options.support_vml = support_vml
            # This document contains a JPEG image within "<!--[if gte vml 1]>" tags,
            # and a different PNG image within "<![if !vml]>" tags.
            # If we set the "SupportVml" flag to "true", then Aspose.Words will load the JPEG.
            # If we set this flag to "false", then Aspose.Words will only load the PNG.
            doc = aw.Document(file_name=MY_DIR + 'VML conditional.htm', load_options=load_options)
            if support_vml:
                self.assertEqual(aw.drawing.ImageType.JPEG, doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().image_data.image_type)
            else:
                self.assertEqual(aw.drawing.ImageType.PNG, doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().image_data.image_type)
            #ExEnd
            image_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
            if support_vml:
                test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
            else:
                test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.PNG, image_shape)

    def test_encrypted_html(self):
        #ExStart
        #ExFor:HtmlLoadOptions.__init__(str)
        #ExSummary:Shows how to encrypt an Html document, and then open it using a password.
        # Create and sign an encrypted HTML document from an encrypted .docx.
        certificate_holder = aw.digitalsignatures.CertificateHolder.create(file_name=MY_DIR + 'morzal.pfx', password='aw')
        sign_options = aw.digitalsignatures.SignOptions()
        sign_options.comments = 'Comment'
        sign_options.sign_time = datetime.datetime.now()
        sign_options.decryption_password = 'docPassword'
        input_file_name = MY_DIR + 'Encrypted.docx'
        output_file_name = ARTIFACTS_DIR + 'HtmlLoadOptions.EncryptedHtml.html'
        aw.digitalsignatures.DigitalSignatureUtil.sign(src_file_name=input_file_name, dst_file_name=output_file_name, cert_holder=certificate_holder, sign_options=sign_options)
        # To load and read this document, we will need to pass its decryption
        # password using a HtmlLoadOptions object.
        load_options = aw.loading.HtmlLoadOptions(password='docPassword')
        self.assertEqual(sign_options.decryption_password, load_options.password)
        doc = aw.Document(file_name=output_file_name, load_options=load_options)
        self.assertEqual('Test encrypted document.', doc.get_text().strip())
        #ExEnd

    def test_base_uri(self):
        #ExStart
        #ExFor:HtmlLoadOptions.__init__(LoadFormat,str,str)
        #ExFor:LoadOptions.__init__(LoadFormat,str,str)
        #ExFor:LoadOptions.load_format
        #ExFor:LoadFormat
        #ExSummary:Shows how to specify a base URI when opening an html document.
        # Suppose we want to load an .html document that contains an image linked by a relative URI
        # while the image is in a different location. In that case, we will need to resolve the relative URI into an absolute one.
        # We can provide a base URI using an HtmlLoadOptions object.
        load_options = aw.loading.HtmlLoadOptions(load_format=aw.LoadFormat.HTML, password='', base_uri=IMAGE_DIR)
        self.assertEqual(aw.LoadFormat.HTML, load_options.load_format)
        doc = aw.Document(file_name=MY_DIR + 'Missing image.html', load_options=load_options)
        # While the image was broken in the input .html, our custom base URI helped us repair the link.
        image_shape = doc.get_child_nodes(aw.NodeType.SHAPE, True)[0].as_shape()
        self.assertTrue(image_shape.is_image)
        # This output document will display the image that was missing.
        doc.save(file_name=ARTIFACTS_DIR + 'HtmlLoadOptions.BaseUri.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'HtmlLoadOptions.BaseUri.docx')
        self.assertTrue(len(doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().image_data.image_bytes) > 0)

    def test_font_face_rules(self):
        #ExStart:FontFaceRules
        #ExFor:HtmlLoadOptions.support_font_face_rules
        #ExSummary:Shows how to load declared "@font-face" rules.
        load_options = aw.loading.HtmlLoadOptions()
        load_options.support_font_face_rules = True
        doc = aw.Document(file_name=MY_DIR + 'Html with FontFace.html', load_options=load_options)
        self.assertEqual('Squarish Sans CT Regular', doc.font_infos[0].name)
        #ExEnd:FontFaceRules

    def test_get_select_as_sdt(self):
        #ExStart
        #ExFor:HtmlLoadOptions.preferred_control_type
        #ExSummary:Shows how to set preferred type of document nodes that will represent imported <input> and <select> elements.
        html = "\n            <html>\n                <select name='ComboBox' size='1'>\n                    <option value='val1'>item1</option>\n                    <option value='val2'></option>\n                </select>\n            </html>"
        html_load_options = aw.loading.HtmlLoadOptions()
        html_load_options.preferred_control_type = aw.loading.HtmlControlType.STRUCTURED_DOCUMENT_TAG
        doc = aw.Document(io.BytesIO(html.encode('utf-8')), html_load_options)
        nodes = doc.get_child_nodes(aw.NodeType.STRUCTURED_DOCUMENT_TAG, True)
        tag = nodes[0].as_structured_document_tag()
        #ExEnd
        self.assertEqual(2, tag.list_items.count)
        self.assertEqual('val1', tag.list_items[0].value)
        self.assertEqual('val2', tag.list_items[1].value)

    def test_get_input_as_form_field(self):
        html = "\n            <html>\n                <input type='text' value='Input value text' />\n            </html>"
        # By default, "HtmlLoadOptions.preferred_control_type" value is "HtmlControlType.FORM_FIELD".
        # So, we do not set this value.
        html_load_options = aw.loading.HtmlLoadOptions()
        doc = aw.Document(io.BytesIO(html.encode('utf-8')), html_load_options)
        nodes = doc.get_child_nodes(aw.NodeType.FORM_FIELD, True)
        self.assertEqual(1, nodes.count)
        form_field = nodes[0].as_form_field()
        self.assertEqual('Input value text', form_field.result)

    def test_ignore_noscript_elements(self):
        for ignore_noscript_elements in (True, False):
            with self.subTest(ignore_noscript_elements=ignore_noscript_elements):
                #ExStart
                #ExFor:HtmlLoadOptions.ignore_noscript_elements
                #ExSummary:Shows how to ignore <noscript> HTML elements.
                html = '\n                    <html>\n                      <head>\n                        <title>NOSCRIPT</title>\n                          <meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">\n                          <script type=""text/javascript"">\n                            alert(""Hello, world!"")\n                          </script>\n                      </head>\n                    <body>\n                      <noscript><p>Your browser does not support JavaScript!</p></noscript>\n                    </body>\n                    </html>'
                html_load_options = aw.loading.HtmlLoadOptions()
                html_load_options.ignore_noscript_elements = ignore_noscript_elements
                doc = aw.Document(io.BytesIO(html.encode('utf-8')), html_load_options)
                doc.save(ARTIFACTS_DIR + 'HtmlLoadOptions.ignore_noscript_elements.pdf')
                #ExEnd

    def test_block_import(self):
        for block_import_mode in [aw.loading.BlockImportMode.PRESERVE, aw.loading.BlockImportMode.MERGE]:
            html = "<html><div style='border:dotted'><div style='border:solid'><p>paragraph 1</p><p>paragraph 2</p></div></div></html>"
            with io.BytesIO(html.encode('utf-8')) as stream:
                load_options = aw.loading.HtmlLoadOptions()
                # Set the new mode of import HTML block-level elements.
                load_options.block_import_mode = block_import_mode
                doc = aw.Document(stream, load_options)
                doc.save(ARTIFACTS_DIR + 'HtmlLoadOptions.BlockImport.docx')