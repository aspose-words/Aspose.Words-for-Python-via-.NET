# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.pydrawing as drawing
import io
import aspose.words as aw
import aspose.words.drawing
import aspose.words.settings
import system_helper
import test_util
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, IMAGE_DIR

class ExDocumentBuilderImages(ApiExampleBase):

    def test_insert_image_from_stream(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_image(BytesIO)
        #ExFor:DocumentBuilder.insert_image(BytesIO,float,float)
        #ExFor:DocumentBuilder.insert_image(BytesIO,RelativeHorizontalPosition,float,RelativeVerticalPosition,float,float,float,WrapType)
        #ExSummary:Shows how to insert an image from a stream into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        with system_helper.io.File.open_read(IMAGE_DIR + 'Logo.jpg') as stream:
            # Below are three ways of inserting an image from a stream.
            # 1 -  Inline shape with a default size based on the image's original dimensions:
            builder.insert_image(stream=stream)
            builder.insert_break(aw.BreakType.PAGE_BREAK)
            # 2 -  Inline shape with custom dimensions:
            builder.insert_image(stream=stream, width=aw.ConvertUtil.pixel_to_point(pixels=250), height=aw.ConvertUtil.pixel_to_point(pixels=144))
            builder.insert_break(aw.BreakType.PAGE_BREAK)
            # 3 -  Floating shape with custom dimensions:
            builder.insert_image(stream=stream, horz_pos=aw.drawing.RelativeHorizontalPosition.MARGIN, left=100, vert_pos=aw.drawing.RelativeVerticalPosition.MARGIN, top=100, width=200, height=100, wrap_type=aw.drawing.WrapType.SQUARE)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilderImages.InsertImageFromStream.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilderImages.InsertImageFromStream.docx')
        image_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual(300, image_shape.height)
        self.assertEqual(300, image_shape.width)
        self.assertEqual(0, image_shape.left)
        self.assertEqual(0, image_shape.top)
        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)
        test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(300, image_shape.image_data.image_size.height_points)
        self.assertEqual(300, image_shape.image_data.image_size.width_points)
        image_shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()
        self.assertEqual(108, image_shape.height)
        self.assertEqual(187.5, image_shape.width)
        self.assertEqual(0, image_shape.left)
        self.assertEqual(0, image_shape.top)
        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)
        test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(300, image_shape.image_data.image_size.height_points)
        self.assertEqual(300, image_shape.image_data.image_size.width_points)
        image_shape = doc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()
        self.assertEqual(100, image_shape.height)
        self.assertEqual(200, image_shape.width)
        self.assertEqual(100, image_shape.left)
        self.assertEqual(100, image_shape.top)
        self.assertEqual(aw.drawing.WrapType.SQUARE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, image_shape.relative_vertical_position)
        test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(300, image_shape.image_data.image_size.height_points)
        self.assertEqual(300, image_shape.image_data.image_size.width_points)

    def test_insert_image_from_filename(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_image(str)
        #ExFor:DocumentBuilder.insert_image(str,float,float)
        #ExFor:DocumentBuilder.insert_image(str,RelativeHorizontalPosition,float,RelativeVerticalPosition,float,float,float,WrapType)
        #ExSummary:Shows how to insert an image from the local file system into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Below are three ways of inserting an image from a local system filename.
        # 1 -  Inline shape with a default size based on the image's original dimensions:
        builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        # 2 -  Inline shape with custom dimensions:
        builder.insert_image(file_name=IMAGE_DIR + 'Transparent background logo.png', width=aw.ConvertUtil.pixel_to_point(pixels=250), height=aw.ConvertUtil.pixel_to_point(pixels=144))
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        # 3 -  Floating shape with custom dimensions:
        builder.insert_image(file_name=IMAGE_DIR + 'Windows MetaFile.wmf', horz_pos=aw.drawing.RelativeHorizontalPosition.MARGIN, left=100, vert_pos=aw.drawing.RelativeVerticalPosition.MARGIN, top=100, width=200, height=100, wrap_type=aw.drawing.WrapType.SQUARE)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilderImages.InsertImageFromFilename.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilderImages.InsertImageFromFilename.docx')
        image_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual(300, image_shape.height)
        self.assertEqual(300, image_shape.width)
        self.assertEqual(0, image_shape.left)
        self.assertEqual(0, image_shape.top)
        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)
        test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(300, image_shape.image_data.image_size.height_points)
        self.assertEqual(300, image_shape.image_data.image_size.width_points)
        image_shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()
        self.assertEqual(108, image_shape.height)
        self.assertEqual(187.5, image_shape.width)
        self.assertEqual(0, image_shape.left)
        self.assertEqual(0, image_shape.top)
        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)
        test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.PNG, image_shape)
        self.assertEqual(300, image_shape.image_data.image_size.height_points)
        self.assertEqual(300, image_shape.image_data.image_size.width_points)
        image_shape = doc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()
        self.assertEqual(100, image_shape.height)
        self.assertEqual(200, image_shape.width)
        self.assertEqual(100, image_shape.left)
        self.assertEqual(100, image_shape.top)
        self.assertEqual(aw.drawing.WrapType.SQUARE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, image_shape.relative_vertical_position)
        test_util.TestUtil.verify_image_in_shape(1600, 1600, aw.drawing.ImageType.WMF, image_shape)
        self.assertEqual(400, image_shape.image_data.image_size.height_points)
        self.assertEqual(400, image_shape.image_data.image_size.width_points)

    def test_insert_svg_image(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_image(str)
        #ExSummary:Shows how to determine which image will be inserted.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.insert_image(file_name=IMAGE_DIR + 'Scalable Vector Graphics.svg')
        # Aspose.Words insert SVG image to the document as PNG with svgBlip extension
        # that contains the original vector SVG image representation.
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilderImages.InsertSvgImage.SvgWithSvgBlip.docx')
        # Aspose.Words insert SVG image to the document as PNG, just like Microsoft Word does for old format.
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilderImages.InsertSvgImage.Svg.doc')
        doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2003)
        # Aspose.Words insert SVG image to the document as EMF metafile to keep the image in vector representation.
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilderImages.InsertSvgImage.Emf.docx')
        #ExEnd

    @unittest.skip("drawing.Image type isn't supported yet")
    def test_insert_image_from_image_object(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_image(Image)
        #ExFor:DocumentBuilder.insert_image(Image,float,float)
        #ExFor:DocumentBuilder.insert_image(Image,RelativeHorizontalPosition,float,RelativeVerticalPosition,float,float,float,WrapType)
        #ExSummary:Shows how to insert an image from an object into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        image_file = IMAGE_DIR + 'Logo.jpg'
        # Below are three ways of inserting an image from an Image object instance.
        # 1 -  Inline shape with a default size based on the image's original dimensions:
        builder.insert_image(file_name=image_file)
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        # 2 -  Inline shape with custom dimensions:
        builder.insert_image(file_name=image_file, width=aw.ConvertUtil.pixel_to_point(pixels=250), height=aw.ConvertUtil.pixel_to_point(pixels=144))
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        # 3 -  Floating shape with custom dimensions:
        builder.insert_image(file_name=image_file, horz_pos=aw.drawing.RelativeHorizontalPosition.MARGIN, left=100, vert_pos=aw.drawing.RelativeVerticalPosition.MARGIN, top=100, width=200, height=100, wrap_type=aw.drawing.WrapType.SQUARE)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilderImages.InsertImageFromImageObject.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilderImages.InsertImageFromImageObject.docx')
        image_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual(300, image_shape.height)
        self.assertEqual(300, image_shape.width)
        self.assertEqual(0, image_shape.left)
        self.assertEqual(0, image_shape.top)
        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)
        test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(300, image_shape.image_data.image_size.height_points)
        self.assertEqual(300, image_shape.image_data.image_size.width_points)
        image_shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()
        self.assertEqual(108, image_shape.height)
        self.assertEqual(187.5, image_shape.width)
        self.assertEqual(0, image_shape.left)
        self.assertEqual(0, image_shape.top)
        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)
        test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(300, image_shape.image_data.image_size.height_points)
        self.assertEqual(300, image_shape.image_data.image_size.width_points)
        image_shape = doc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()
        self.assertEqual(100, image_shape.height)
        self.assertEqual(200, image_shape.width)
        self.assertEqual(100, image_shape.left)
        self.assertEqual(100, image_shape.top)
        self.assertEqual(aw.drawing.WrapType.SQUARE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, image_shape.relative_vertical_position)
        test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertEqual(300, image_shape.image_data.image_size.height_points)
        self.assertEqual(300, image_shape.image_data.image_size.width_points)

    def test_insert_image_from_byte_array(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_image(bytes)
        #ExFor:DocumentBuilder.insert_image(bytes,float,float)
        #ExFor:DocumentBuilder.insert_image(bytes,RelativeHorizontalPosition,float,RelativeVerticalPosition,float,float,float,WrapType)
        #ExSummary:Shows how to insert an image from a byte array into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        image_byte_array = test_util.TestUtil.image_to_byte_array(IMAGE_DIR + 'Logo.jpg')
        # Below are three ways of inserting an image from a byte array.
        # 1 -  Inline shape with a default size based on the image's original dimensions:
        builder.insert_image(image_bytes=image_byte_array)
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        # 2 -  Inline shape with custom dimensions:
        builder.insert_image(image_bytes=image_byte_array, width=aw.ConvertUtil.pixel_to_point(pixels=250), height=aw.ConvertUtil.pixel_to_point(pixels=144))
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        # 3 -  Floating shape with custom dimensions:
        builder.insert_image(image_bytes=image_byte_array, horz_pos=aw.drawing.RelativeHorizontalPosition.MARGIN, left=100, vert_pos=aw.drawing.RelativeVerticalPosition.MARGIN, top=100, width=200, height=100, wrap_type=aw.drawing.WrapType.SQUARE)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentBuilderImages.InsertImageFromByteArray.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentBuilderImages.InsertImageFromByteArray.docx')
        image_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertAlmostEqual(300, image_shape.height, delta=0.1)
        self.assertAlmostEqual(300, image_shape.width, delta=0.1)
        self.assertEqual(0, image_shape.left)
        self.assertEqual(0, image_shape.top)
        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)
        test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertAlmostEqual(300, image_shape.image_data.image_size.height_points, delta=0.1)
        self.assertAlmostEqual(300, image_shape.image_data.image_size.width_points, delta=0.1)
        image_shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()
        self.assertEqual(108, image_shape.height)
        self.assertEqual(187.5, image_shape.width)
        self.assertEqual(0, image_shape.left)
        self.assertEqual(0, image_shape.top)
        self.assertEqual(aw.drawing.WrapType.INLINE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.COLUMN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, image_shape.relative_vertical_position)
        test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertAlmostEqual(300, image_shape.image_data.image_size.height_points, delta=0.1)
        self.assertAlmostEqual(300, image_shape.image_data.image_size.width_points, delta=0.1)
        image_shape = doc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()
        self.assertEqual(100, image_shape.height)
        self.assertEqual(200, image_shape.width)
        self.assertEqual(100, image_shape.left)
        self.assertEqual(100, image_shape.top)
        self.assertEqual(aw.drawing.WrapType.SQUARE, image_shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, image_shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.MARGIN, image_shape.relative_vertical_position)
        test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, image_shape)
        self.assertAlmostEqual(300, image_shape.image_data.image_size.height_points, delta=0.1)
        self.assertAlmostEqual(300, image_shape.image_data.image_size.width_points, delta=0.1)

    def test_insert_gif(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_image(str)
        #ExSummary:Shows how to insert gif image to the document.
        builder = aw.DocumentBuilder()
        # We can insert gif image using path or bytes array.
        # It works only if DocumentBuilder optimized to Word version 2010 or higher.
        # Note, that access to the image bytes causes conversion Gif to Png.
        gif_image = builder.insert_image(file_name=IMAGE_DIR + 'Graphics Interchange Format.gif')
        gif_image = builder.insert_image(image_bytes=system_helper.io.File.read_all_bytes(IMAGE_DIR + 'Graphics Interchange Format.gif'))
        builder.document.save(file_name=ARTIFACTS_DIR + 'InsertGif.docx')
        #ExEnd