# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import io
from datetime import datetime

import aspose.words as aw
import aspose.words.loading as awl

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, IMAGE_URL, IMAGE_DIR

class ExHtmlLoadOptions(ApiExampleBase):

    def test_support_vml(self):

        for support_vml in (True, False):
            with self.subTest(support_vml=support_vml):
                #ExStart
                #ExFor:HtmlLoadOptions.__init__()
                #ExFor:HtmlLoadOptions.support_vml
                #ExSummary:Shows how to support conditional comments while loading an HTML document.
                load_options = aw.loading.HtmlLoadOptions()

                # If the value is True, then we take VML code into account while parsing the loaded document.
                load_options.support_vml = support_vml

                # This document contains a JPEG image within "<!--[if gte vml 1]>" tags,
                # and a different PNG image within "<![if !vml]>" tags.
                # If we set the "support_vml" flag to "True", then Aspose.Words will load the JPEG.
                # If we set this flag to "False", then Aspose.Words will only load the PNG.
                doc = aw.Document(MY_DIR + "VML conditional.htm", load_options)

                if support_vml:
                    self.assertEqual(aw.drawing.ImageType.JPEG, doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().image_data.image_type)
                else:
                    self.assertEqual(aw.drawing.ImageType.PNG, doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().image_data.image_type)
                #ExEnd

                image_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

                if support_vml:
                    self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
                else:
                    self.verify_image_in_shape(400, 400, aw.drawing.ImageType.PNG, image_shape)

    ##ExStart
    ##ExFor:HtmlLoadOptions.web_request_timeout
    ##ExSummary:Shows how to set a time limit for web requests when loading a document with external resources linked by URLs.
    #def test_web_request_timeout(self):

    #    # Create a new HtmlLoadOptions object and verify its timeout threshold for a web request.
    #    options = aw.loading.HtmlLoadOptions()

    #    # When loading an Html document with resources externally linked by a web address URL,
    #    # Aspose.Words will abort web requests that fail to fetch the resources within this time limit, in milliseconds.
    #    self.assertEqual(100000, options.web_request_timeout)

    #    # Set a WarningCallback that will record all warnings that occur during loading.
    #    warning_callback = ExHtmlLoadOptions.ListDocumentWarnings()
    #    options.warning_callback = warning_callback

    #    # Load such a document and verify that a shape with image data has been created.
    #    # This linked image will require a web request to load, which will have to complete within our time limit.
    #    html = '<html><img src="{IMAGE_URL}" alt="Aspose logo" style="width:400px;height:400px;"></html>'

    #    # Set an unreasonable timeout limit and try load the document again.
    #    options.web_request_timeout = 0
    #    doc = aw.Document(io.BytesIO(html.encode("utf-8")), options)
    #    self.assertEqual(2, warning_callback.warnings().count)

    #    # A web request that fails to obtain an image within the time limit will still produce an image.
    #    # However, the image will be the red 'x' that commonly signifies missing images.
    #    image_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
    #    self.assertEqual(924, len(image_shape.image_data.image_bytes))

    #    # We can also configure a custom callback to pick up any warnings from timed out web requests.
    #    self.assertEqual(aw.WarningSource.HTML, warning_callback.warnings[0].source)
    #    self.assertEqual(aw.WarningType.DATA_LOSS, warning_callback.warnings[0].warning_type)
    #    self.assertEqual(f"Couldn't load a resource from \'{IMAGE_URL}\'.", warning_callback.warnings[0].description)

    #    self.assertEqual(aw.WarningSource.HTML, warning_callback.warnings[1].source)
    #    self.assertEqual(aw.WarningType.DATA_LOSS, warning_callback.warnings[1].warning_type)
    #    self.assertEqual("Image has been replaced with a placeholder.", warning_callback.warnings[1].description)

    #    doc.save(ARTIFACTS_DIR + "HtmlLoadOptions.web_request_timeout.docx")

    #class ListDocumentWarnings(aw.IWarningCallback):
    #    """Stores all warnings that occur during a document loading operation in a List."""

    #    def __init__(self):
    #        self.warnings = []

    #    def warning(self, info: aw.WarningInfo):
    #        self.warnings.add(info)

    ##ExEnd

    def test_encrypted_html(self):

        #ExStart
        #ExFor:HtmlLoadOptions.__init__(str)
        #ExSummary:Shows how to encrypt an Html document, and then open it using a password.
        # Create and sign an encrypted HTML document from an encrypted .docx.
        certificate_holder = aw.digitalsignatures.CertificateHolder.create(MY_DIR + "morzal.pfx", "aw")

        sign_options = aw.digitalsignatures.SignOptions()
        sign_options.comments = "Comment"
        sign_options.sign_time = datetime.now()
        sign_options.decryption_password = "docPassword"

        input_file_name = MY_DIR + "Encrypted.docx"
        output_file_name = ARTIFACTS_DIR + "HtmlLoadOptions.encrypted_html.html"
        aw.digitalsignatures.DigitalSignatureUtil.sign(input_file_name, output_file_name, certificate_holder, sign_options)

        # To load and read this document, we will need to pass its decryption
        # password using a HtmlLoadOptions object.
        load_options = aw.loading.HtmlLoadOptions("docPassword")

        self.assertEqual(sign_options.decryption_password, load_options.password)

        doc = aw.Document(output_file_name, load_options)

        self.assertEqual("Test encrypted document.", doc.get_text().strip())
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
        load_options = aw.loading.HtmlLoadOptions(aw.LoadFormat.HTML, "", IMAGE_DIR)

        self.assertEqual(aw.LoadFormat.HTML, load_options.load_format)

        doc = aw.Document(MY_DIR + "Missing image.html", load_options)

        # While the image was broken in the input .html, our custom base URI helped us repair the link.
        image_shape = doc.get_child_nodes(aw.NodeType.SHAPE, True)[0].as_shape()
        self.assertTrue(image_shape.is_image)

        # This output document will display the image that was missing.
        doc.save(ARTIFACTS_DIR + "HtmlLoadOptions.base_uri.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "HtmlLoadOptions.base_uri.docx")

        self.assertGreater(len(doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().image_data.image_bytes), 0)

    def test_get_select_as_sdt(self):

        #ExStart
        #ExFor:HtmlLoadOptions.preferred_control_type
        #ExSummary:Shows how to set preferred type of document nodes that will represent imported <input> and <select> elements.
        html = """
            <html>
                <select name='ComboBox' size='1'>
                    <option value='val1'>item1</option>
                    <option value='val2'></option>
                </select>
            </html>"""

        html_load_options = aw.loading.HtmlLoadOptions()
        html_load_options.preferred_control_type = aw.loading.HtmlControlType.STRUCTURED_DOCUMENT_TAG

        doc = aw.Document(io.BytesIO(html.encode("utf-8")), html_load_options)
        nodes = doc.get_child_nodes(aw.NodeType.STRUCTURED_DOCUMENT_TAG, True)

        tag = nodes[0].as_structured_document_tag()
        #ExEnd

        self.assertEqual(2, tag.list_items.count)

        self.assertEqual("val1", tag.list_items[0].value)
        self.assertEqual("val2", tag.list_items[1].value)

    def test_get_input_as_form_field(self):

        html = """
            <html>
                <input type='text' value='Input value text' />
            </html>"""

        # By default, "HtmlLoadOptions.preferred_control_type" value is "HtmlControlType.FORM_FIELD".
        # So, we do not set this value.
        html_load_options = aw.loading.HtmlLoadOptions()

        doc = aw.Document(io.BytesIO(html.encode("utf-8")), html_load_options)
        nodes = doc.get_child_nodes(aw.NodeType.FORM_FIELD, True)

        self.assertEqual(1, nodes.count)

        form_field = nodes[0].as_form_field()
        self.assertEqual("Input value text", form_field.result)

    def test_ignore_noscript_elements(self):

        for ignore_noscript_elements in (True, False):
            with self.subTest(ignore_noscript_elements=ignore_noscript_elements):
                #ExStart
                #ExFor:HtmlLoadOptions.ignore_noscript_elements
                #ExSummary:Shows how to ignore <noscript> HTML elements.
                html = """
                    <html>
                      <head>
                        <title>NOSCRIPT</title>
                          <meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">
                          <script type=""text/javascript"">
                            alert(""Hello, world!"")
                          </script>
                      </head>
                    <body>
                      <noscript><p>Your browser does not support JavaScript!</p></noscript>
                    </body>
                    </html>"""

                html_load_options = aw.loading.HtmlLoadOptions()
                html_load_options.ignore_noscript_elements = ignore_noscript_elements

                doc = aw.Document(io.BytesIO(html.encode("utf-8")), html_load_options)
                doc.save(ARTIFACTS_DIR + "HtmlLoadOptions.ignore_noscript_elements.pdf")
                #ExEnd

                #pdf_doc = aspose.pdf.Document(ARTIFACTS_DIR + "HtmlLoadOptions.ignore_noscript_elements.pdf")
                #text_absorber = aspose.pdf.text.TextAbsorber()
                #text_absorber.visit(pdf_doc)

                #self.assertEqual("" if ignore_noscript_elements else "Your browser does not support JavaScript!", text_absorber.text)

    def test_block_import(self):
        for block_import_mode in [awl.BlockImportMode.PRESERVE, awl.BlockImportMode.MERGE]:
            html = "<html><div style='border:dotted'><div style='border:solid'><p>paragraph 1</p><p>paragraph 2</p></div></div></html>"
            with io.BytesIO(html.encode("utf-8")) as stream:
                load_options = awl.HtmlLoadOptions()
                # Set the new mode of import HTML block-level elements.
                load_options.block_import_mode = block_import_mode
                doc = aw.Document(stream, load_options)
                doc.save(ARTIFACTS_DIR + "HtmlLoadOptions.BlockImport.docx")
