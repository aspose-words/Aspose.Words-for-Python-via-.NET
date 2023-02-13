# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import io
import unittest

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, ARTIFACTS_DIR, IMAGE_DIR

class ExDocumentBuilderImages(ApiExampleBase):

    def test_insert_image_from_stream(self):

        #ExStart
        #ExFor:DocumentBuilder.insert_image(BytesIO)
        #ExFor:DocumentBuilder.insert_image(BytesIO,float,float)
        #ExFor:DocumentBuilder.insert_image(BytesIO,RelativeHorizontalPosition,float,RelativeVerticalPosition,float,float,float,WrapType)
        #ExSummary:Shows how to insert an image from a stream into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        with open(IMAGE_DIR + "Logo.jpg", "rb") as stream:

            # Below are three ways of inserting an image from a stream.
            # 1 -  Inline shape with a default size based on the image's original dimensions:
            builder.insert_image(stream)

            builder.insert_break(aw.BreakType.PAGE_BREAK)

            # 2 -  Inline shape with custom dimensions:
            builder.insert_image(stream, aw.ConvertUtil.pixel_to_point(250), aw.ConvertUtil.pixel_to_point(144))

            builder.insert_break(aw.BreakType.PAGE_BREAK)

            # 3 -  Floating shape with custom dimensions:
            builder.insert_image(stream, aw.drawing.RelativeHorizontalPosition.MARGIN, 100, aw.drawing.RelativeVerticalPosition.MARGIN,
                100, 200, 100, aw.drawing.WrapType.SQUARE)

        doc.save(ARTIFACTS_DIR + "DocumentBuilderImages.insert_image_from_stream.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "DocumentBuilderImages.insert_image_from_stream.docx")

        image_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.assertEqual(300.0, image_shape.height)
        self.assertEqual(300.0, image_shape.width)
        self.assertEqual(0.0, image_shape.left)
        self.assertEqual(0.0, image_shape.top)

        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)

        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(300.0, image_shape.image_data.image_size.height_points)
        self.assertEqual(300.0, image_shape.image_data.image_size.width_points)

        image_shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()

        self.assertEqual(108.0, image_shape.height)
        self.assertEqual(187.5, image_shape.width)
        self.assertEqual(0.0, image_shape.left)
        self.assertEqual(0.0, image_shape.top)

        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)

        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(300.0, image_shape.image_data.image_size.height_points)
        self.assertEqual(300.0, image_shape.image_data.image_size.width_points)

        image_shape = doc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()

        self.assertEqual(100.0, image_shape.height)
        self.assertEqual(200.0, image_shape.width)
        self.assertEqual(100.0, image_shape.left)
        self.assertEqual(100.0, image_shape.top)

        self.assertEqual(aw.drawing.WrapType.SQUARE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, image_shape.relative_vertical_position)

        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(300.0, image_shape.image_data.image_size.height_points)
        self.assertEqual(300.0, image_shape.image_data.image_size.width_points)

    def test_insert_image_from_filename(self):

        #ExStart
        #ExFor:DocumentBuilder.insert_image(str)
        #ExFor:DocumentBuilder.insert_image(str,float,float)
        #ExFor:DocumentBuilder.insert_image(str,RelativeHorizontalPosition,float,RelativeVerticalPosition,float,float,float,WrapType)
        #ExSummary:Shows how to insert an image from the local file system into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Below are three ways of inserting an image from a local system filename.
        # 1 -  Inline shape with a default size based on the image's original dimensions:
        builder.insert_image(IMAGE_DIR + "Logo.jpg")

        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # 2 -  Inline shape with custom dimensions:
        builder.insert_image(IMAGE_DIR + "Transparent background logo.png", aw.ConvertUtil.pixel_to_point(250),
            aw.ConvertUtil.pixel_to_point(144))

        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # 3 -  Floating shape with custom dimensions:
        builder.insert_image(IMAGE_DIR + "Windows MetaFile.wmf", aw.drawing.RelativeHorizontalPosition.MARGIN, 100,
            aw.drawing.RelativeVerticalPosition.MARGIN, 100, 200, 100, aw.drawing.WrapType.SQUARE)

        doc.save(ARTIFACTS_DIR + "DocumentBuilderImages.insert_image_from_filename.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "DocumentBuilderImages.insert_image_from_filename.docx")

        image_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.assertEqual(300.0, image_shape.height)
        self.assertEqual(300.0, image_shape.width)
        self.assertEqual(0.0, image_shape.left)
        self.assertEqual(0.0, image_shape.top)

        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)

        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(300.0, image_shape.image_data.image_size.height_points)
        self.assertEqual(300.0, image_shape.image_data.image_size.width_points)

        image_shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()

        self.assertEqual(108.0, image_shape.height)
        self.assertEqual(187.5, image_shape.width)
        self.assertEqual(0.0, image_shape.left)
        self.assertEqual(0.0, image_shape.top)

        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)

        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.PNG, image_shape)
        self.assertEqual(300.0, image_shape.image_data.image_size.height_points)
        self.assertEqual(300.0, image_shape.image_data.image_size.width_points)

        image_shape = doc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()

        self.assertEqual(100.0, image_shape.height)
        self.assertEqual(200.0, image_shape.width)
        self.assertEqual(100.0, image_shape.left)
        self.assertEqual(100.0, image_shape.top)

        self.assertEqual(aw.drawing.WrapType.SQUARE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, image_shape.relative_vertical_position)

        self.verify_image_in_shape(1600, 1600, aw.drawing.ImageType.WMF, image_shape)
        self.assertEqual(400.0, image_shape.image_data.image_size.height_points)
        self.assertEqual(400.0, image_shape.image_data.image_size.width_points)

    def test_insert_svg_image(self):

        #ExStart
        #ExFor:DocumentBuilder.insert_image(str)
        #ExSummary:Shows how to determine which image will be inserted.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_image(IMAGE_DIR + "Scalable Vector Graphics.svg")

        # Aspose.Words insert SVG image to the document as PNG with svgBlip extension
        # that contains the original vector SVG image representation.
        doc.save(ARTIFACTS_DIR + "DocumentBuilderImages.insert_svg_image.svg_with_svg_blip.docx")

        # Aspose.Words insert SVG image to the document as PNG, just like Microsoft Word does for old format.
        doc.save(ARTIFACTS_DIR + "DocumentBuilderImages.insert_svg_image.svg.doc")

        doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2003)

        # Aspose.Words insert SVG image to the document as EMF metafile to keep the image in vector representation.
        doc.save(ARTIFACTS_DIR + "DocumentBuilderImages.insert_svg_image.emf.docx")
        #ExEnd

    @unittest.skip("drawing.Image type isn't supported yet")
    def test_insert_image_from_image_object(self):

        #ExStart
        #ExFor:DocumentBuilder.insert_image(Image)
        #ExFor:DocumentBuilder.insert_image(Image,float,float)
        #ExFor:DocumentBuilder.insert_image(Image,RelativeHorizontalPosition,float,RelativeVerticalPosition,float,float,float,WrapType)
        #ExSummary:Shows how to insert an image from an object into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        image = drawing.Image.from_file(IMAGE_DIR + "Logo.jpg")

        # Below are three ways of inserting an image from an Image object instance.
        # 1 -  Inline shape with a default size based on the image's original dimensions:
        builder.insert_image(image)

        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # 2 -  Inline shape with custom dimensions:
        builder.insert_image(image, aw.ConvertUtil.pixel_to_point(250), aw.ConvertUtil.pixel_to_point(144))

        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # 3 -  Floating shape with custom dimensions:
        builder.insert_image(image, aw.drawing.RelativeHorizontalPosition.MARGIN, 100, aw.drawing.RelativeVerticalPosition.MARGIN,
                             100, 200, 100, aw.drawing.WrapType.SQUARE)

        doc.save(ARTIFACTS_DIR + "DocumentBuilderImages.insert_image_from_image_object.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "DocumentBuilderImages.insert_image_from_image_object.docx")

        image_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.assertEqual(300.0, image_shape.height)
        self.assertEqual(300.0, image_shape.width)
        self.assertEqual(0.0, image_shape.left)
        self.assertEqual(0.0, image_shape.top)

        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)

        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(300.0, image_shape.image_data.image_size.height_points)
        self.assertEqual(300.0, image_shape.image_data.image_size.width_points)

        image_shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()

        self.assertEqual(108.0, image_shape.height)
        self.assertEqual(187.5, image_shape.width)
        self.assertEqual(0.0, image_shape.left)
        self.assertEqual(0.0, image_shape.top)

        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)

        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(300.0, image_shape.image_data.image_size.height_points)
        self.assertEqual(300.0, image_shape.image_data.image_size.width_points)

        image_shape = doc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()

        self.assertEqual(100.0, image_shape.height)
        self.assertEqual(200.0, image_shape.width)
        self.assertEqual(100.0, image_shape.left)
        self.assertEqual(100.0, image_shape.top)

        self.assertEqual(aw.drawing.WrapType.SQUARE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, image_shape.relative_vertical_position)

        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(300.0, image_shape.image_data.image_size.height_points)
        self.assertEqual(300.0, image_shape.image_data.image_size.width_points)

    def test_insert_image_from_byte_array(self):

        #ExStart
        #ExFor:DocumentBuilder.insert_image(bytes)
        #ExFor:DocumentBuilder.insert_image(bytes,float,float)
        #ExFor:DocumentBuilder.insert_image(bytes,RelativeHorizontalPosition,float,RelativeVerticalPosition,float,float,float,WrapType)
        #ExSummary:Shows how to insert an image from a byte array into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        image = drawing.Image.from_file(IMAGE_DIR + "Logo.jpg")

        with io.BytesIO() as stream:

            image.save(stream, drawing.imaging.ImageFormat.png)
            image_byte_array = bytes(stream.getvalue())

            # Below are three ways of inserting an image from a byte array.
            # 1 -  Inline shape with a default size based on the image's original dimensions:
            builder.insert_image(image_byte_array)

            builder.insert_break(aw.BreakType.PAGE_BREAK)

            # 2 -  Inline shape with custom dimensions:
            builder.insert_image(image_byte_array, aw.ConvertUtil.pixel_to_point(250), aw.ConvertUtil.pixel_to_point(144))

            builder.insert_break(aw.BreakType.PAGE_BREAK)

            # 3 -  Floating shape with custom dimensions:
            builder.insert_image(image_byte_array, aw.drawing.RelativeHorizontalPosition.MARGIN, 100, aw.drawing.RelativeVerticalPosition.MARGIN,
                                 100, 200, 100, aw.drawing.WrapType.SQUARE)

        doc.save(ARTIFACTS_DIR + "DocumentBuilderImages.insert_image_from_byte_array.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "DocumentBuilderImages.insert_image_from_byte_array.docx")

        image_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.assertAlmostEqual(300.0, image_shape.height, delta=0.1)
        self.assertAlmostEqual(300.0, image_shape.width, delta=0.1)
        self.assertEqual(0.0, image_shape.left)
        self.assertEqual(0.0, image_shape.top)

        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)

        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.PNG, image_shape)
        self.assertAlmostEqual(300.0, image_shape.image_data.image_size.height_points, delta=0.1)
        self.assertAlmostEqual(300.0, image_shape.image_data.image_size.width_points, delta=0.1)

        image_shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()

        self.assertEqual(108.0, image_shape.height)
        self.assertEqual(187.5, image_shape.width)
        self.assertEqual(0.0, image_shape.left)
        self.assertEqual(0.0, image_shape.top)

        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)

        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.PNG, image_shape)
        self.assertAlmostEqual(300.0, image_shape.image_data.image_size.height_points, delta=0.1)
        self.assertAlmostEqual(300.0, image_shape.image_data.image_size.width_points, delta=0.1)

        image_shape = doc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()

        self.assertEqual(100.0, image_shape.height)
        self.assertEqual(200.0, image_shape.width)
        self.assertEqual(100.0, image_shape.left)
        self.assertEqual(100.0, image_shape.top)

        self.assertEqual(aw.drawing.WrapType.SQUARE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, image_shape.relative_vertical_position)

        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.PNG, image_shape)
        self.assertAlmostEqual(300.0, image_shape.image_data.image_size.height_points, delta=0.1)
        self.assertAlmostEqual(300.0, image_shape.image_data.image_size.width_points, delta=0.1)

    def test_insert_gif(self):

        #ExStart
        #ExFor:DocumentBuilder.insert_image(str)
        #ExSummary:Shows how to insert gif image to the document.
        builder = aw.DocumentBuilder()

        # We can insert gif image using path or bytes array.
        # It works only if DocumentBuilder optimized to Word version 2010 or higher.
        # Note, that access to the image bytes causes conversion Gif to Png.
        gif_image = builder.insert_image(IMAGE_DIR + "Graphics Interchange Format.gif")

        with open(IMAGE_DIR + "Graphics Interchange Format.gif", "rb") as file:
            gif_image = builder.insert_image(file.read())

        builder.document.save(ARTIFACTS_DIR + "DocumentBuilderImages.insert_gif.docx")
        #ExEnd
