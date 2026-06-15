# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.pydrawing
import aspose.words as aw
import aspose.words.buildingblocks
import aspose.words.drawing
import aspose.words.loading
import aspose.words.saving
import aspose.words.themes
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, IMAGE_DIR

class ExDocumentBase(ApiExampleBase):

    def test_constructor(self):
        #ExStart
        #ExFor:DocumentBase
        #ExSummary:Shows how to initialize the subclasses of DocumentBase.
        doc = aw.Document()
        self.assertEqual(aw.DocumentBase, type(doc).__bases__[0])
        glossary_doc = aw.buildingblocks.GlossaryDocument()
        doc.glossary_document = glossary_doc
        self.assertEqual(aw.DocumentBase, type(glossary_doc).__bases__[0])
        #ExEnd

    def test_set_page_color(self):
        #ExStart
        #ExFor:DocumentBase.page_color
        #ExSummary:Shows how to set the background color for all pages of a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Hello world!')
        doc.page_color = aspose.pydrawing.Color.light_gray
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBase.SetPageColor.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBase.SetPageColor.docx')
        self.assertEqual(aspose.pydrawing.Color.light_gray.to_argb(), doc.page_color.to_argb())

    def test_import_node(self):
        from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR
        import aspose.words as aw
        #ExStart
        #ExFor:DocumentBase.import_node(Node,bool)
        #ExSummary:Shows how to import a node from one document to another.
        # Every node has a parent document, which is the document that contains the node.
        # Inserting a node into a document that the node does not belong to will throw an exception.
        src_doc = aw.Document()
        dst_doc = aw.Document()
        src_doc.first_section.body.first_paragraph.append_child(aw.Run(doc=src_doc, text='Source document first paragraph text.'))
        dst_doc.first_section.body.first_paragraph.append_child(aw.Run(doc=dst_doc, text='Destination document first paragraph text.'))
        self.assertNotEqual(dst_doc, src_doc.first_section.document)
        # Use the ImportNode method to create a copy of a node, which will have the document
        # that called the ImportNode method set as its new owner document.
        imported_section = dst_doc.import_node(src_node=src_doc.first_section, is_import_children=True).as_section()
        self.assertEqual(dst_doc, imported_section.document)
        # We can now insert the node into the document.
        dst_doc.append_child(imported_section)
        self.assertEqual('Destination document first paragraph text.\r\nSource document first paragraph text.\r\n', dst_doc.to_string(save_format=aw.SaveFormat.TEXT))
        #ExEnd
        self.assertNotEqual(imported_section, src_doc.first_section)
        self.assertNotEqual(imported_section.document, src_doc.first_section.document)
        self.assertEqual(imported_section.body.first_paragraph.get_text(), src_doc.first_section.body.first_paragraph.get_text())

    def test_import_node_custom(self):
        #ExStart
        #ExFor:DocumentBase.import_node(Node,bool,ImportFormatMode)
        #ExSummary:Shows how to import node from source document to destination document with specific options.
        # Create two documents and add a character style to each document.
        # Configure the styles to have the same name, but different text formatting.
        src_doc = aw.Document()
        src_style = src_doc.styles.add(aw.StyleType.CHARACTER, 'My style')
        src_style.font.name = 'Courier New'
        src_builder = aw.DocumentBuilder(doc=src_doc)
        src_builder.font.style = src_style
        src_builder.writeln('Source document text.')
        dst_doc = aw.Document()
        dst_style = dst_doc.styles.add(aw.StyleType.CHARACTER, 'My style')
        dst_style.font.name = 'Calibri'
        dst_builder = aw.DocumentBuilder(doc=dst_doc)
        dst_builder.font.style = dst_style
        dst_builder.writeln('Destination document text.')
        # Import the Section from the destination document into the source document, causing a style name collision.
        # If we use destination styles, then the imported source text with the same style name
        # as destination text will adopt the destination style.
        imported_section = dst_doc.import_node(src_node=src_doc.first_section, is_import_children=True, import_format_mode=aw.ImportFormatMode.USE_DESTINATION_STYLES).as_section()
        self.assertEqual('Source document text.', imported_section.body.paragraphs[0].runs[0].get_text().strip())  #ExSkip
        self.assertIsNone(dst_doc.styles.get_by_name('My style_0'))  #ExSkip
        self.assertEqual(dst_style.font.name, imported_section.body.first_paragraph.runs[0].font.name)
        self.assertEqual(dst_style.name, imported_section.body.first_paragraph.runs[0].font.style_name)
        # If we use ImportFormatMode.KeepDifferentStyles, the source style is preserved,
        # and the naming clash resolves by adding a suffix.
        dst_doc.import_node(src_node=src_doc.first_section, is_import_children=True, import_format_mode=aw.ImportFormatMode.KEEP_DIFFERENT_STYLES)
        self.assertEqual(dst_style.font.name, dst_doc.styles.get_by_name('My style').font.name)
        self.assertEqual(src_style.font.name, dst_doc.styles.get_by_name('My style_0').font.name)
        #ExEnd

    def test_background_shape(self):
        #ExStart
        #ExFor:DocumentBase.background_shape
        #ExSummary:Shows how to set a background shape for every page of a document.
        doc = aw.Document()
        self.assertIsNone(doc.background_shape)
        # The only shape type that we can use as a background is a rectangle.
        shape_rectangle = aw.drawing.Shape(doc, aw.drawing.ShapeType.RECTANGLE)
        # There are two ways of using this shape as a page background.
        # 1 -  A flat color:
        shape_rectangle.fill_color = aspose.pydrawing.Color.light_blue
        doc.background_shape = shape_rectangle
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBase.BackgroundShape.FlatColor.docx')
        # 2 -  An image:
        shape_rectangle = aw.drawing.Shape(doc, aw.drawing.ShapeType.RECTANGLE)
        shape_rectangle.image_data.set_image(file_name=IMAGE_DIR + 'Transparent background logo.png')
        # Adjust the image's appearance to make it more suitable as a watermark.
        shape_rectangle.image_data.contrast = 0.2
        shape_rectangle.image_data.brightness = 0.7
        doc.background_shape = shape_rectangle
        self.assertTrue(doc.background_shape.has_image)
        save_options = aw.saving.PdfSaveOptions()
        save_options.cache_background_graphics = False
        # Microsoft Word does not support shapes with images as backgrounds,
        # but we can still see these backgrounds in other save formats such as .pdf.
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBase.BackgroundShape.Image.pdf', save_options=save_options)
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBase.BackgroundShape.FlatColor.docx')
        self.assertEqual(aspose.pydrawing.Color.light_blue.to_argb(), doc.background_shape.fill_color.to_argb())
        with self.assertRaises(Exception):
            doc.background_shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.TRIANGLE)

    def _test_resource_loading_callback(self, doc):
        for shape in doc.get_child_nodes(aw.NodeType.SHAPE, True):
            shape = shape.as_shape()
            self.assertTrue(shape.has_image)
            self.assertTrue(shape.image_data.image_bytes)

    def test_import_node_with_resolve_theme_colors(self):
        #ExStart:ImportNodeWithResolveThemeColors
        #ExFor:DocumentBase.import_node(Node,bool,ImportFormatMode,ImportFormatOptions)
        #ExFor:ImportFormatOptions.resolve_theme_colors
        #ExSummary:Shows how to import a node with resolving source theme colors of shapes.
        src_doc = aw.Document()
        builder = aw.DocumentBuilder(doc=src_doc)
        # Move to the primary footer and insert a shape that uses theme colors.
        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)
        shape = builder.insert_shape(shape_type=aw.drawing.ShapeType.RECTANGLE, width=100, height=50)
        shape.stroke.fore_theme_color = aw.themes.ThemeColor.DARK1
        dst_doc = aw.Document()
        # Import the source footer into the destination document with theme colors resolved,
        # so the shape preserves its actual color from the source document.
        footer = src_doc.first_section.headers_footers.get_by_header_footer_type(aw.HeaderFooterType.FOOTER_PRIMARY)
        options = aw.ImportFormatOptions()
        options.resolve_theme_colors = True
        imported_footer = dst_doc.import_node(src_node=footer, is_import_children=True, import_format_mode=aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, import_format_options=options).as_header_footer()
        dst_doc.first_section.headers_footers.add(imported_footer)
        dst_doc.save(file_name=ARTIFACTS_DIR + 'DocumentBase.ImportNodeWithResolveThemeColors.docx')
        #ExEnd:ImportNodeWithResolveThemeColors
    #ExStart
    #ExFor:DocumentBase.resource_loading_callback
    #ExFor:IResourceLoadingCallback
    #ExFor:IResourceLoadingCallback.resource_loading(ResourceLoadingArgs)
    #ExFor:ResourceLoadingAction
    #ExFor:ResourceLoadingArgs
    #ExFor:ResourceLoadingArgs.original_uri
    #ExFor:ResourceLoadingArgs.resource_type
    #ExFor:ResourceLoadingArgs.set_data(bytes)
    #ExFor:ResourceType
    #ExSummary:Shows how to customize the process of loading external resources into a document (ImageNameHandler).

    class ImageNameHandler(aw.loading.IResourceLoadingCallback):

        def resource_loading(self, args):
            # If this callback encounters one of the image shorthands while loading an image,
            # it will apply unique logic for each defined shorthand instead of treating it as a URI.
            if args.resource_type == aw.loading.ResourceType.IMAGE:
                switch_condition = args.original_uri
                if switch_condition == 'Google logo':
                    import requests
                    image_data = requests.get('http://www.google.com/images/logos/ps_logo2.png').content
                    args.set_data(image_data)
                    return aw.loading.ResourceLoadingAction.USER_PROVIDED
                elif switch_condition == 'Aspose logo':
                    from api_example_base import IMAGE_DIR
                    args.set_data(open(IMAGE_DIR + 'Logo.jpg', 'rb').read())
                    return aw.loading.ResourceLoadingAction.USER_PROVIDED
                elif switch_condition == 'Watermark':
                    args.set_data(open(IMAGE_DIR + 'Transparent background logo.png', 'rb').read())
                    return aw.loading.ResourceLoadingAction.USER_PROVIDED
            return aw.loading.ResourceLoadingAction.DEFAULT
    #ExEnd