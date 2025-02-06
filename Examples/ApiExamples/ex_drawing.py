# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from document_helper import DocumentHelper
import glob
import os
import aspose.pydrawing
import aspose.words as aw
import aspose.words.drawing
import io
import system_helper
import test_util
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, IMAGE_DIR, MY_DIR

class ExDrawing(ApiExampleBase):

    @unittest.skip("drawing.Image type isn't supported yet")
    def test_type_of_image(self):
        #ExStart
        #ExFor:ImageType
        #ExSummary:Shows how to add an image to a shape and check its type.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        img_shape = builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
        self.assertEqual(aw.drawing.ImageType.JPEG, img_shape.image_data.image_type)
        #ExEnd

    def test_fill_solid(self):
        #ExStart
        #ExFor:Fill.color()
        #ExFor:FillType
        #ExFor:Fill.fill_type
        #ExFor:Fill.solid
        #ExFor:Fill.transparency
        #ExFor:Font.fill
        #ExSummary:Shows how to convert any of the fills back to solid fill.
        doc = aw.Document(file_name=MY_DIR + 'Two color gradient.docx')
        # Get Fill object for Font of the first Run.
        fill = doc.first_section.body.paragraphs[0].runs[0].font.fill
        # Check Fill properties of the Font.
        print('The type of the fill is: {0}'.format(fill.fill_type))
        print('The foreground color of the fill is: {0}'.format(fill.fore_color))
        print('The fill is transparent at {0}%'.format(fill.transparency * 100))
        # Change type of the fill to Solid with uniform green color.
        fill.solid()
        print('\nThe fill is changed:')
        print('The type of the fill is: {0}'.format(fill.fill_type))
        print('The foreground color of the fill is: {0}'.format(fill.fore_color))
        print('The fill transparency is {0}%'.format(fill.transparency * 100))
        doc.save(file_name=ARTIFACTS_DIR + 'Drawing.FillSolid.docx')
        #ExEnd

    def test_stroke_pattern(self):
        #ExStart
        #ExFor:Stroke.color2
        #ExFor:Stroke.image_bytes
        #ExSummary:Shows how to process shape stroke features.
        doc = aw.Document(file_name=MY_DIR + 'Shape stroke pattern border.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        stroke = shape.stroke
        # Strokes can have two colors, which are used to create a pattern defined by two-tone image data.
        # Strokes with a single color do not use the Color2 property.
        self.assertEqual(aspose.pydrawing.Color.from_argb(255, 128, 0, 0), stroke.color)
        self.assertEqual(aspose.pydrawing.Color.from_argb(255, 255, 255, 0), stroke.color2)
        self.assertIsNotNone(stroke.image_bytes)
        system_helper.io.File.write_all_bytes(ARTIFACTS_DIR + 'Drawing.StrokePattern.png', stroke.image_bytes)
        #ExEnd
        test_util.TestUtil.verify_image(8, 8, ARTIFACTS_DIR + 'Drawing.StrokePattern.png')

    def test_text_box(self):
        #ExStart
        #ExFor:LayoutFlow
        #ExSummary:Shows how to add text to a text box, and change its orientation
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        textbox = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_BOX)
        textbox.width = 100
        textbox.height = 100
        textbox.text_box.layout_flow = aw.drawing.LayoutFlow.BOTTOM_TO_TOP
        textbox.append_child(aw.Paragraph(doc))
        builder.insert_node(textbox)
        builder.move_to(textbox.first_paragraph)
        builder.write('This text is flipped 90 degrees to the left.')
        doc.save(file_name=ARTIFACTS_DIR + 'Drawing.TextBox.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Drawing.TextBox.docx')
        textbox = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual(aw.drawing.ShapeType.TEXT_BOX, textbox.shape_type)
        self.assertEqual(100, textbox.width)
        self.assertEqual(100, textbox.height)
        self.assertEqual(aw.drawing.LayoutFlow.BOTTOM_TO_TOP, textbox.text_box.layout_flow)
        self.assertEqual('This text is flipped 90 degrees to the left.', textbox.get_text().strip())

    def test_image_data(self):
        #ExStart
        #ExFor:ImageData.bi_level
        #ExFor:ImageData.borders
        #ExFor:ImageData.brightness
        #ExFor:ImageData.chroma_key
        #ExFor:ImageData.contrast
        #ExFor:ImageData.crop_bottom
        #ExFor:ImageData.crop_left
        #ExFor:ImageData.crop_right
        #ExFor:ImageData.crop_top
        #ExFor:ImageData.gray_scale
        #ExFor:ImageData.is_link
        #ExFor:ImageData.is_link_only
        #ExFor:ImageData.title
        #ExSummary:Shows how to edit a shape's image data.
        img_source_doc = aw.Document(file_name=MY_DIR + 'Images.docx')
        source_shape = img_source_doc.get_child_nodes(aw.NodeType.SHAPE, True)[0].as_shape()
        dst_doc = aw.Document()
        # Import a shape from the source document and append it to the first paragraph.
        imported_shape = dst_doc.import_node(src_node=source_shape, is_import_children=True).as_shape()
        dst_doc.first_section.body.first_paragraph.append_child(imported_shape)
        # The imported shape contains an image. We can access the image's properties and raw data via the ImageData object.
        image_data = imported_shape.image_data
        image_data.title = 'Imported Image'
        self.assertTrue(image_data.has_image)
        # If an image has no borders, its ImageData object will define the border color as empty.
        self.assertEqual(4, image_data.borders.count)
        self.assertEqual(aspose.pydrawing.Color.empty(), image_data.borders[0].color)
        # This image does not link to another shape or image file in the local file system.
        self.assertFalse(image_data.is_link)
        self.assertFalse(image_data.is_link_only)
        # The "Brightness" and "Contrast" properties define image brightness and contrast
        # on a 0-1 scale, with the default value at 0.5.
        image_data.brightness = 0.8
        image_data.contrast = 1
        # The above brightness and contrast values have created an image with a lot of white.
        # We can select a color with the ChromaKey property to replace with transparency, such as white.
        image_data.chroma_key = aspose.pydrawing.Color.white
        # Import the source shape again and set the image to monochrome.
        imported_shape = dst_doc.import_node(src_node=source_shape, is_import_children=True).as_shape()
        dst_doc.first_section.body.first_paragraph.append_child(imported_shape)
        imported_shape.image_data.gray_scale = True
        # Import the source shape again to create a third image and set it to BiLevel.
        # BiLevel sets every pixel to either black or white, whichever is closer to the original color.
        imported_shape = dst_doc.import_node(src_node=source_shape, is_import_children=True).as_shape()
        dst_doc.first_section.body.first_paragraph.append_child(imported_shape)
        imported_shape.image_data.bi_level = True
        # Cropping is determined on a 0-1 scale. Cropping a side by 0.3
        # will crop 30% of the image out at the cropped side.
        imported_shape.image_data.crop_bottom = 0.3
        imported_shape.image_data.crop_left = 0.3
        imported_shape.image_data.crop_top = 0.3
        imported_shape.image_data.crop_right = 0.3
        dst_doc.save(file_name=ARTIFACTS_DIR + 'Drawing.ImageData.docx')
        #ExEnd
        img_source_doc = aw.Document(file_name=ARTIFACTS_DIR + 'Drawing.ImageData.docx')
        source_shape = img_source_doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        test_util.TestUtil.verify_image_in_shape(2467, 1500, aw.drawing.ImageType.JPEG, source_shape)
        self.assertEqual('Imported Image', source_shape.image_data.title)
        self.assertAlmostEqual(0.8, source_shape.image_data.brightness, delta=0.1)
        self.assertAlmostEqual(1, source_shape.image_data.contrast, delta=0.1)
        self.assertEqual(aspose.pydrawing.Color.white.to_argb(), source_shape.image_data.chroma_key.to_argb())
        source_shape = img_source_doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()
        test_util.TestUtil.verify_image_in_shape(2467, 1500, aw.drawing.ImageType.JPEG, source_shape)
        self.assertTrue(source_shape.image_data.gray_scale)
        source_shape = img_source_doc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()
        test_util.TestUtil.verify_image_in_shape(2467, 1500, aw.drawing.ImageType.JPEG, source_shape)
        self.assertTrue(source_shape.image_data.bi_level)
        self.assertAlmostEqual(0.3, source_shape.image_data.crop_bottom, delta=0.1)
        self.assertAlmostEqual(0.3, source_shape.image_data.crop_left, delta=0.1)
        self.assertAlmostEqual(0.3, source_shape.image_data.crop_top, delta=0.1)
        self.assertAlmostEqual(0.3, source_shape.image_data.crop_right, delta=0.1)

    def test_image_size(self):
        #ExStart
        #ExFor:ImageSize.height_pixels
        #ExFor:ImageSize.horizontal_resolution
        #ExFor:ImageSize.vertical_resolution
        #ExFor:ImageSize.width_pixels
        #ExSummary:Shows how to read the properties of an image in a shape.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a shape into the document which contains an image taken from our local file system.
        shape = builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
        # If the shape contains an image, its ImageData property will be valid,
        # and it will contain an ImageSize object.
        image_size = shape.image_data.image_size
        # The ImageSize object contains read-only information about the image within the shape.
        self.assertEqual(400, image_size.height_pixels)
        self.assertEqual(400, image_size.width_pixels)
        delta = 0.05
        self.assertAlmostEqual(95.98, image_size.horizontal_resolution, delta=delta)
        self.assertAlmostEqual(95.98, image_size.vertical_resolution, delta=delta)
        # We can base the size of the shape on the size of its image to avoid stretching the image.
        shape.width = image_size.width_points * 2
        shape.height = image_size.height_points * 2
        doc.save(file_name=ARTIFACTS_DIR + 'Drawing.ImageSize.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Drawing.ImageSize.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        test_util.TestUtil.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, shape)
        self.assertEqual(600, shape.width)
        self.assertEqual(600, shape.height)
        image_size = shape.image_data.image_size
        self.assertEqual(400, image_size.height_pixels)
        self.assertEqual(400, image_size.width_pixels)
        self.assertAlmostEqual(95.98, image_size.horizontal_resolution, delta=delta)
        self.assertAlmostEqual(95.98, image_size.vertical_resolution, delta=delta)

    @unittest.skip("drawing.Image type isn't supported yet")
    def test_various_shapes(self):
        #ExStart
        #ExFor:ArrowLength
        #ExFor:ArrowType
        #ExFor:ArrowWidth
        #ExFor:DashStyle
        #ExFor:EndCap
        #ExFor:Fill.fore_color
        #ExFor:Fill.image_bytes
        #ExFor:Fill.visible
        #ExFor:Shape.stroke
        #ExFor:Stroke.color
        #ExFor:Stroke.start_arrow_length
        #ExFor:Stroke.start_arrow_type
        #ExFor:Stroke.start_arrow_width
        #ExFor:Stroke.end_arrow_length
        #ExFor:Stroke.end_arrow_width
        #ExFor:Stroke.dash_style
        #ExFor:Stroke.end_arrow_type
        #ExFor:Stroke.end_cap
        #ExFor:Stroke.opacity
        #ExFor:Stroke.join_style
        #ExSummary:Shows to create a variety of shapes.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Below are four examples of shapes that we can insert into our documents.
        # 1 -  Dotted, horizontal, half-transparent red line
        # with an arrow on the left end and a diamond on the right end:
        arrow = aw.drawing.Shape(doc, aw.drawing.ShapeType.LINE)
        arrow.width = 200
        arrow.stroke.color = drawing.Color.red
        arrow.stroke.start_arrow_type = aw.drawing.ArrowType.ARROW
        arrow.stroke.start_arrow_length = aw.drawing.ArrowLength.LONG
        arrow.stroke.start_arrow_width = aw.drawing.ArrowWidth.WIDE
        arrow.stroke.end_arrow_type = aw.drawing.ArrowType.DIAMOND
        arrow.stroke.end_arrow_length = aw.drawing.ArrowLength.LONG
        arrow.stroke.end_arrow_width = aw.drawing.ArrowWidth.WIDE
        arrow.stroke.dash_style = aw.drawing.DashStyle.DASH
        arrow.stroke.opacity = 0.5
        self.assertEqual(aw.drawing.JoinStyle.MITER, arrow.stroke.join_style)
        builder.insert_node(arrow)
        # 2 -  Thick black diagonal line with rounded ends:
        line = aw.drawing.Shape(doc, aw.drawing.ShapeType.LINE)
        line.top = 40
        line.width = 200
        line.height = 20
        line.stroke_weight = 5.0
        line.stroke.end_cap = aw.drawing.EndCap.ROUND
        builder.insert_node(line)
        # 3 -  Arrow with a green fill:
        filled_in_arrow = aw.drawing.Shape(doc, aw.drawing.ShapeType.ARROW)
        filled_in_arrow.width = 200
        filled_in_arrow.height = 40
        filled_in_arrow.top = 100
        filled_in_arrow.fill.fore_color = drawing.Color.green
        filled_in_arrow.fill.visible = True
        builder.insert_node(filled_in_arrow)
        # 4 -  Arrow with a flipped orientation filled in with the Aspose logo:
        filled_in_arrow_img = aw.drawing.Shape(doc, aw.drawing.ShapeType.ARROW)
        filled_in_arrow_img.width = 200
        filled_in_arrow_img.height = 40
        filled_in_arrow_img.top = 160
        filled_in_arrow_img.flip_orientation = aw.drawing.FlipOrientation.BOTH
        with open(IMAGE_DIR + 'Logo.jpg', 'rb') as stream:
            image = drawing.Image.from_stream(stream)
            # When we flip the orientation of our arrow, we also flip the image that the arrow contains.
            # Flip the image the other way to cancel this out before getting the shape to display it.
            image.rotate_flip(drawing.RotateFlipType.ROTATE_NONE_FLIP_XY)
            filled_in_arrow_img.image_data.set_image(image)
            filled_in_arrow_img.stroke.join_style = aw.drawing.JoinStyle.ROUND
            builder.insert_node(filled_in_arrow_img)
        doc.save(ARTIFACTS_DIR + 'Drawing.various_shapes.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Drawing.various_shapes.docx')
        self.assertEqual(4, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)
        arrow = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual(aw.drawing.ShapeType.LINE, arrow.shape_type)
        self.assertEqual(200.0, arrow.width)
        self.assertEqual(drawing.Color.red.to_argb(), arrow.stroke.color.to_argb())
        self.assertEqual(aw.drawing.ArrowType.ARROW, arrow.stroke.start_arrow_type)
        self.assertEqual(aw.drawing.ArrowLength.LONG, arrow.stroke.start_arrow_length)
        self.assertEqual(aw.drawing.ArrowWidth.WIDE, arrow.stroke.start_arrow_width)
        self.assertEqual(aw.drawing.ArrowType.DIAMOND, arrow.stroke.end_arrow_type)
        self.assertEqual(aw.drawing.ArrowLength.LONG, arrow.stroke.end_arrow_length)
        self.assertEqual(aw.drawing.ArrowWidth.WIDE, arrow.stroke.end_arrow_width)
        self.assertEqual(aw.drawing.DashStyle.DASH, arrow.stroke.dash_style)
        self.assertEqual(0.5, arrow.stroke.opacity)
        line = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()
        self.assertEqual(aw.drawing.ShapeType.LINE, line.shape_type)
        self.assertEqual(40.0, line.top)
        self.assertEqual(200.0, line.width)
        self.assertEqual(20.0, line.height)
        self.assertEqual(5.0, line.stroke_weight)
        self.assertEqual(aw.drawing.EndCap.ROUND, line.stroke.end_cap)
        filled_in_arrow = doc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()
        self.assertEqual(aw.drawing.ShapeType.ARROW, filled_in_arrow.shape_type)
        self.assertEqual(200.0, filled_in_arrow.width)
        self.assertEqual(40.0, filled_in_arrow.height)
        self.assertEqual(100.0, filled_in_arrow.top)
        self.assertEqual(drawing.Color.green.to_argb(), filled_in_arrow.fill.fore_color.to_argb())
        self.assertTrue(filled_in_arrow.fill.visible)
        filled_in_arrow_img = doc.get_child(aw.NodeType.SHAPE, 3, True).as_shape()
        self.assertEqual(aw.drawing.ShapeType.ARROW, filled_in_arrow_img.shape_type)
        self.assertEqual(200.0, filled_in_arrow_img.width)
        self.assertEqual(40.0, filled_in_arrow_img.height)
        self.assertEqual(160.0, filled_in_arrow_img.top)
        self.assertEqual(aw.drawing.FlipOrientation.BOTH, filled_in_arrow_img.flip_orientation)

    @unittest.skip("drawing.Image type isn't supported yet")
    def test_import_image(self):
        #ExStart
        #ExFor:ImageData.set_image(Image)
        #ExFor:ImageData.set_image(BytesIO)
        #ExSummary:Shows how to display images from the local file system in a document.
        doc = aw.Document()
        # To display an image in a document, we will need to create a shape
        # which will contain an image, and then append it to the document's body.
        img_shape = None
        # Below are two ways of getting an image from a file in the local file system.
        # 1 -  Create an image object from an image file:
        with drawing.Image.from_file(IMAGE_DIR + 'Logo.jpg') as src_image:
            img_shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.IMAGE)
            doc.first_section.body.first_paragraph.append_child(img_shape)
            img_shape.image_data.set_image(src_image)
        # 2 -  Open an image file from the local file system using a stream:
        with open(IMAGE_DIR + 'Logo.jpg', 'rb') as stream:
            img_shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.IMAGE)
            doc.first_section.body.first_paragraph.append_child(img_shape)
            img_shape.image_data.set_image(stream)
            img_shape.left = 150.0
        doc.save(ARTIFACTS_DIR + 'Drawing.import_image.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Drawing.import_image.docx')
        self.assertEqual(2, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)
        img_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, img_shape)
        self.assertEqual(0.0, img_shape.left)
        self.assertEqual(0.0, img_shape.top)
        self.assertEqual(300.0, img_shape.height)
        self.assertEqual(300.0, img_shape.width)
        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, img_shape)
        img_shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()
        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, img_shape)
        self.assertEqual(150.0, img_shape.left)
        self.assertEqual(0.0, img_shape.top)
        self.assertEqual(300.0, img_shape.height)
        self.assertEqual(300.0, img_shape.width)

    @unittest.skip('Unknown method aspose.words.drawing.ImageData.to_image')
    @unittest.skip("drawing.ImageFormatConverter type isn't supported yet")
    def test_save_all_images(self):
        #ExStart
        #ExFor:ImageData.has_image
        #ExFor:ImageData.to_image
        #ExFor:ImageData.save(BytesIO)
        #ExSummary:Shows how to save all images from a document to the file system.
        img_source_doc = aw.Document(MY_DIR + 'Images.docx')
        # Shapes with the "has_image" flag set store and display all the document's images.
        shapes_with_images = [node.as_shape() for node in img_source_doc.get_child_nodes(aw.NodeType.SHAPE, True) if node.as_shape().has_image]
        # Go through each shape and save its image.
        format_converter = drawing.ImageFormatConverter()
        for shape_index, shape in enumerate(shapes_with_images):
            image_data = shape.image_data
            format = image_data.to_image().raw_format
            file_extension = format_converter.convert_to_string(format)
            with open(ARTIFACTS_DIR + f'Drawing.save_all_images.{shape_index}.{file_extension}', 'wb') as file_stream:
                file_stream.write(file_stream)
        #ExEnd
        image_file_names = glob.glob(ARTIFACTS_DIR + 'Drawing.save_all_images.*')
        image_file_ext = [os.path.splitext(name)[-1] for name in image_file_names]
        self.verify_image(2467, 1500, image_file_names[0])
        self.assertEqual('.Jpeg', image_file_ext[0])
        self.verify_image(400, 400, image_file_names[1])
        self.assertEqual('.Png', image_file_ext[1])
        self.verify_image(382, 138, image_file_names[2])
        self.assertEqual('.Emf', image_file_ext[2])
        self.verify_image(1600, 1600, image_file_names[3])
        self.assertEqual('.Wmf', image_file_ext[3])
        self.verify_image(534, 534, image_file_names[4])
        self.assertEqual('.Emf', image_file_ext[4])
        self.verify_image(1260, 660, image_file_names[5])
        self.assertEqual('.Jpeg', image_file_ext[5])
        self.verify_image(1125, 1500, image_file_names[6])
        self.assertEqual('.Jpeg', image_file_ext[6])
        self.verify_image(1027, 1500, image_file_names[7])
        self.assertEqual('.Jpeg', image_file_ext[7])
        self.verify_image(1200, 1500, image_file_names[8])
        self.assertEqual('.Jpeg', image_file_ext[8])

    def test_get_data_from_image(self):
        #ExStart
        #ExFor:ImageData.image_bytes
        #ExFor:ImageData.to_byte_array
        #ExFor:ImageData.to_stream
        #ExSummary:Shows how to create an image file from a shape's raw image data.
        img_source_doc = aw.Document(MY_DIR + 'Images.docx')
        self.assertEqual(10, img_source_doc.get_child_nodes(aw.NodeType.SHAPE, True).count)  #ExSkip
        img_shape = img_source_doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertTrue(img_shape.has_image)
        # to_byte_array() returns the array stored in the "image_bytes" property.
        self.assertEqual(img_shape.image_data.image_bytes, img_shape.image_data.to_byte_array())
        # Save the shape's image data to an image file in the local file system.
        with img_shape.image_data.to_stream() as img_stream:
            with open(ARTIFACTS_DIR + 'Drawing.get_data_from_image.png', 'wb') as out_stream:
                out_stream.write(img_stream.read())
        #ExEnd
        self.verify_image(2467, 1500, ARTIFACTS_DIR + 'Drawing.get_data_from_image.png')

    def _test_group_shapes(self, doc: aw.Document):
        doc = DocumentHelper.save_open(doc)
        shapes = doc.get_child(aw.NodeType.GROUP_SHAPE, 0, True).as_group_shape()
        self.assertEqual(2, shapes.get_child_nodes(aw.NodeType.ANY, False).count)
        shape = shapes.get_child_nodes(aw.NodeType.ANY, False)[0].as_shape()
        self.assertEqual(aw.drawing.ShapeType.BALLOON, shape.shape_type)
        self.assertEqual(200.0, shape.width)
        self.assertEqual(200.0, shape.height)
        self.assertEqual(drawing.Color.red.to_argb(), shape.stroke_color.to_argb())
        shape = shapes.get_child_nodes(aw.NodeType.ANY, False)[1].as_shape()
        self.assertEqual(aw.drawing.ShapeType.CUBE, shape.shape_type)
        self.assertEqual(100.0, shape.width)
        self.assertEqual(100.0, shape.height)
        self.assertEqual(drawing.Color.blue.to_argb(), shape.stroke_color.to_argb())