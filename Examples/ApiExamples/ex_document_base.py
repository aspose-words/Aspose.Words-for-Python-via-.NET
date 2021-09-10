import unittest

import api_example_base as aeb

import aspose.words as aw


class ExDocumentBase(aeb.ApiExampleBase):

    def test_constructor(self):
        # ExStart
        # ExFor:DocumentBase
        # ExSummary:Shows how to initialize the subclasses of DocumentBase.
        doc = aw.Document()

        self.assertEqual(type(aw.DocumentBase), doc.get_type().base_type)

        glossary_doc = aw.GlossaryDocument()
        doc.glossary_document = glossary_doc

        self.assertEqual(type(aw.DocumentBase), glossary_doc.get_type().base_type)
        # ExEnd

    def test_set_page_color(self):
        # ExStart
        # ExFor:DocumentBase.page_color
        # ExSummary:Shows how to set the background color for all pages of a document.

        print("Color is not available (commented all lines with color)")

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world!")

        # doc.page_color = System.drawing.color.light_gray

        doc.save(aeb.artifacts_dir + "DocumentBase.set_page_color.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBase.set_page_color.docx")

        # self.assertEqual(System.drawing.color.light_gray.to_argb(), doc.page_color.to_argb())

    def test_import_node(self):
        # ExStart
        # ExFor:DocumentBase.import_node(Node, Boolean)
        # ExSummary:Shows how to import a node from one document to another.
        src_doc = aw.Document()
        dst_doc = aw.Document()

        src_doc.first_section.body.first_paragraph.append_child(
            aw.Run(src_doc, "Source document first paragraph text."))
        dst_doc.first_section.body.first_paragraph.append_child(
            aw.Run(dst_doc, "Destination document first paragraph text."))

        # Every node has a parent document, which is the document that contains the node.
        # Inserting a node into a document that the node does not belong to will throw an exception.
        self.assertNotEqual(dst_doc, src_doc.first_section.document)
        # Assert.throws<ArgumentException>(() =>  dst_doc.append_child(src_doc.first_section) )

        # Use the ImportNode method to create a copy of a node, which will have the document
        # that called the ImportNode method set as its new owner document.
        imported_section = dst_doc.import_node(src_doc.first_section, True).as_section()

        self.assertEqual(dst_doc, imported_section.document)

        # We can now insert the node into the document.
        dst_doc.append_child(imported_section)

        self.assertEqual("Destination document first paragraph text.\r\nSource document first paragraph text.\r\n",
                         dst_doc.to_string(aw.SaveFormat.TEXT))
        # ExEnd

        self.assertNotEqual(imported_section, src_doc.first_section)
        self.assertNotEqual(imported_section.document, src_doc.first_section.document)
        self.assertEqual(imported_section.body.first_paragraph.get_text(),
                         src_doc.first_section.body.first_paragraph.get_text())

    @unittest.skip("Item properties can use only int (lines 113-114)")
    def test_import_node_custom(self):

        #ExStart
        #ExFor:DocumentBase.import_node(Node, System.boolean, ImportFormatMode)
        #ExSummary:Shows how to import node from source document to destination document with specific options.
        # Create two documents and add a character style to each document.
        # Configure the styles to have the same name, but different text formatting.
        src_doc = aw.Document()
        src_style = src_doc.styles.add(aw.StyleType.CHARACTER, "My style")
        src_style.font.name = "Courier New"
        src_builder = aw.DocumentBuilder(src_doc)
        src_builder.font.style = src_style
        src_builder.writeln("Source document text.")

        dst_doc = aw.Document()
        dst_style = dst_doc.styles.add(aw.StyleType.CHARACTER, "My style")
        dst_style.font.name = "Calibri"
        dst_builder = aw.DocumentBuilder(dst_doc)
        dst_builder.font.style = dst_style
        dst_builder.writeln("Destination document text.")

        # Import the Section from the destination document into the source document, causing a style name collision.
        # If we use destination styles, then the imported source text with the same style name
        # as destination text will adopt the destination style.
        imported_section = dst_doc.import_node(src_doc.first_section, True, aw.ImportFormatMode.USE_DESTINATION_STYLES).as_section()
        self.assertEqual("Source document text.", imported_section.body.paragraphs[0].runs[0].get_text().strip()) #ExSkip
        self.assertIsNone(dst_doc.styles["My style_0"]) #ExSkip
        self.assertEqual(dst_style.font.name, imported_section.body.first_paragraph.runs[0].font.name)
        self.assertEqual(dst_style.name, imported_section.body.first_paragraph.runs[0].font.style_name)

        # If we use ImportFormatMode.keep_different_styles, the source style is preserved,
        # and the naming clash resolves by adding a suffix.
        dst_doc.import_node(src_doc.first_section, True, aw.ImportFormatMode.KEEP_DIFFERENT_STYLES)
        self.assertEqual(dst_style.font.name, dst_doc.styles["My style"].font.name)
        self.assertEqual(src_style.font.name, dst_doc.styles["My style_0"].font.name)
        #ExEnd


    def test_background_shape(self):

        #ExStart
        #ExFor:DocumentBase.background_shape
        #ExSummary:Shows how to set a background shape for every page of a document.

        print("Color is not available (commented all lines with color)")

        doc = aw.Document()

        self.assertIsNone(doc.background_shape)

        # The only shape type that we can use as a background is a rectangle.
        shape_rectangle = aw.drawing.Shape(doc, aw.drawing.ShapeType.RECTANGLE)

        # There are two ways of using this shape as a page background.
        # 1 -  A flat color:
        # shape_rectangle.fill_color = System.drawing.color.light_blue
        doc.background_shape = shape_rectangle

        doc.save(aeb.artifacts_dir + "DocumentBase.background_shape.flat_color.docx")

        # 2 -  An image:
        shape_rectangle = aw.drawing.Shape(doc, aw.drawing.ShapeType.RECTANGLE)
        shape_rectangle.image_data.set_image(aeb.image_dir + "Transparent background logo.png")

        # Adjust the image's appearance to make it more suitable as a watermark.
        shape_rectangle.image_data.contrast = 0.2
        shape_rectangle.image_data.brightness = 0.7

        doc.background_shape = shape_rectangle

        self.assertTrue(doc.background_shape.has_image)

        # Microsoft Word does not support shapes with images as backgrounds,
        # but we can still see these backgrounds in other save formats such as .pdf.
        doc.save(aeb.artifacts_dir + "DocumentBase.background_shape.image.pdf")
        #ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBase.background_shape.flat_color.docx")

        # self.assertEqual(System.drawing.color.light_blue.to_argb(), doc.background_shape.fill_color.to_argb())
        # Assert.throws<ArgumentException>(() =>
        #
        #     doc.background_shape = new Shape(doc, ShapeType.triangle)
        # )


    #ExStart
    #ExFor:DocumentBase.resource_loading_callback
    #ExFor:IResourceLoadingCallback
    #ExFor:IResourceLoadingCallback.resource_loading(ResourceLoadingArgs)
    #ExFor:ResourceLoadingAction
    #ExFor:ResourceLoadingArgs
    #ExFor:ResourceLoadingArgs.original_uri
    #ExFor:ResourceLoadingArgs.resource_type
    #ExFor:ResourceLoadingArgs.set_data(Byte[])
    #ExFor:ResourceType
    #ExSummary:Shows how to customize the process of loading external resources into a document.
    # def test_resource_loading_callback(self):
    #
    #     doc = aw.Document()
    #     doc.resource_loading_callback = new ImageNameHandler()
    #
    #     builder = aw.DocumentBuilder(doc)
    #
    #     # Images usually are inserted using a URI, or a byte array.
    #     # Every instance of a resource load will call our callback's ResourceLoading method.
    #     builder.insert_image("Google logo")
    #     builder.insert_image("Aspose logo")
    #     builder.insert_image("Watermark")
    #
    #     self.assertEqual(3, doc.get_child_nodes(NodeType.shape, true).count)
    #
    #     doc.save(ArtifactsDir + "DocumentBase.resource_loading_callback.docx")
    #     TestResourceLoadingCallback(new Document(ArtifactsDir + "DocumentBase.resource_loading_callback.docx")) #ExSkip
    #
    #
    # # <summary>
    # # Allows us to load images into a document using predefined shorthands, as opposed to URIs.
    # # This will separate image loading logic from the rest of the document construction.
    # # </summary>
    # class _ImageNameHandler : aw.loading.IResourceLoadingCallback
    #
    #     aw.loading.ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args):
    #         # If this callback encounters one of the image shorthands while loading an image,
    #         # it will apply unique logic for each defined shorthand instead of treating it as a URI.
    #         if (args.resource_type == ResourceType.image)
    #             switch (args.original_uri)
    #
    #                 case "Google logo":
    #                     using (WebClient webClient = new WebClient())
    #
    #                         args.set_data(webClient.download_data("http:#www.google.com/images/logos/ps_logo2.png"))
    #
    #
    #                     return ResourceLoadingAction.user_provided
    #
    #                 case "Aspose logo":
    #                     args.set_data(File.read_all_bytes(ImageDir + "Logo.jpg"))
    #
    #                     return ResourceLoadingAction.user_provided
    #
    #                 case "Watermark":
    #                     args.set_data(File.read_all_bytes(ImageDir + "Transparent background logo.png"))
    #
    #                     return ResourceLoadingAction.user_provided
    #
    #         return ResourceLoadingAction.default

    #
    # #ExEnd
    #
    # def TestResourceLoadingCallback(Document doc)
    #
    #     foreach (Shape shape in doc.get_child_nodes(NodeType.shape, true))
    #
    #         self.assertTrue(shape.has_image)
    #         Assert.is_not_empty(shape.image_data.image_bytes)
    #
    #
    #     TestUtil.verify_web_response_status_code(HttpStatusCode.ok, "http:#www.google.com/images/logos/ps_logo2.png")

#endif


if __name__ == '__main__':
    unittest.main()
