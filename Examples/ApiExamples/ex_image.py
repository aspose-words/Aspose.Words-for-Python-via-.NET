# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from aspose.words import Document, DocumentBuilder, NodeType
from aspose.words.drawing import ImageType
import os
import aspose.pydrawing as drawing
import aspose.words as aw
import aspose.words.drawing
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, IMAGE_DIR, IMAGE_URL, MY_DIR

class ExImage(ApiExampleBase):

    def test_insert_webp_image(self):
        #ExStart:InsertWebpImage
        #ExFor:DocumentBuilder.insert_image(str)
        #ExSummary:Shows how to insert WebP image.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.insert_image(file_name=IMAGE_DIR + 'WebP image.webp')
        doc.save(file_name=ARTIFACTS_DIR + 'Image.InsertWebpImage.docx')
        #ExEnd:InsertWebpImage

    def test_read_webp_image(self):
        #ExStart:ReadWebpImage
        #ExFor:ImageType
        #ExSummary:Shows how to read WebP image.
        doc = aw.Document(file_name=MY_DIR + 'Document with WebP image.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual(aw.drawing.ImageType.WEB_P, shape.image_data.image_type)
        #ExEnd:ReadWebpImage

    def test_from_file(self):
        #ExStart
        #ExFor:Shape.__init__(DocumentBase,ShapeType)
        #ExFor:ShapeType
        #ExSummary:Shows how to insert a shape with an image from the local file system into a document.
        doc = aw.Document()
        # The "Shape" class's public constructor will create a shape with "ShapeMarkupLanguage.VML" markup type.
        # If you need to create a shape of a non-primitive type, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
        # TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, or DiagonalCornersRounded,
        # please use DocumentBuilder.insert_shape.
        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.IMAGE)
        shape.image_data.set_image(IMAGE_DIR + 'Windows MetaFile.wmf')
        shape.width = 100
        shape.height = 100
        doc.first_section.body.first_paragraph.append_child(shape)
        doc.save(ARTIFACTS_DIR + 'Image.from_file.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Image.from_file.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.verify_image_in_shape(1600, 1600, aw.drawing.ImageType.WMF, shape)
        self.assertEqual(100.0, shape.height)
        self.assertEqual(100.0, shape.width)

    def test_from_url(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_image(str)
        #ExSummary:Shows how to insert a shape with an image into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Below are two locations where the document builder's "insert_image" method
        # can source the image that the shape will display.
        # 1 -  Pass a local file system filename of an image file:
        builder.write('Image from local file: ')
        builder.insert_image(IMAGE_DIR + 'Logo.jpg')
        builder.writeln()
        # 2 -  Pass a URL which points to an image.
        builder.write('Image from a URL: ')
        builder.insert_image(IMAGE_URL)
        builder.writeln()
        doc.save(ARTIFACTS_DIR + 'Image.from_url.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Image.from_url.docx')
        shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)
        self.assertEqual(2, shapes.count)
        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, shapes[0].as_shape())
        self.verify_image_in_shape(272, 92, aw.drawing.ImageType.PNG, shapes[1].as_shape())

    def test_from_stream(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_image(BytesIO)
        #ExSummary:Shows how to insert a shape with an image from a stream into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        with open(IMAGE_DIR + 'Logo.jpg', 'rb') as stream:
            builder.write('Image from stream: ')
            builder.insert_image(stream)
        doc.save(ARTIFACTS_DIR + 'Image.from_stream.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Image.from_stream.docx')
        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, doc.get_child_nodes(aw.NodeType.SHAPE, True)[0].as_shape())

    def test_create_floating_page_center(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_image(str)
        #ExFor:Shape
        #ExFor:ShapeBase
        #ExFor:ShapeBase.wrap_type
        #ExFor:ShapeBase.behind_text
        #ExFor:ShapeBase.relative_horizontal_position
        #ExFor:ShapeBase.relative_vertical_position
        #ExFor:ShapeBase.horizontal_alignment
        #ExFor:ShapeBase.vertical_alignment
        #ExFor:WrapType
        #ExFor:RelativeHorizontalPosition
        #ExFor:RelativeVerticalPosition
        #ExFor:HorizontalAlignment
        #ExFor:VerticalAlignment
        #ExSummary:Shows how to insert a floating image to the center of a page.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Insert a floating image that will appear behind the overlapping text and align it to the page's center.
        shape = builder.insert_image(IMAGE_DIR + 'Logo.jpg')
        shape.wrap_type = aw.drawing.WrapType.NONE
        shape.behind_text = True
        shape.relative_horizontal_position = aw.drawing.RelativeHorizontalPosition.PAGE
        shape.relative_vertical_position = aw.drawing.RelativeVerticalPosition.PAGE
        shape.horizontal_alignment = aw.drawing.HorizontalAlignment.CENTER
        shape.vertical_alignment = aw.drawing.VerticalAlignment.CENTER
        doc.save(ARTIFACTS_DIR + 'Image.create_floating_page_center.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Image.create_floating_page_center.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, shape)
        self.assertEqual(aw.drawing.WrapType.NONE, shape.wrap_type)
        self.assertTrue(shape.behind_text)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.PAGE, shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PAGE, shape.relative_vertical_position)
        self.assertEqual(aw.drawing.HorizontalAlignment.CENTER, shape.horizontal_alignment)
        self.assertEqual(aw.drawing.VerticalAlignment.CENTER, shape.vertical_alignment)

    def test_create_floating_position_size(self):
        #ExStart
        #ExFor:ShapeBase.left
        #ExFor:ShapeBase.right
        #ExFor:ShapeBase.top
        #ExFor:ShapeBase.bottom
        #ExFor:ShapeBase.width
        #ExFor:ShapeBase.height
        #ExFor:DocumentBuilder.current_section
        #ExFor:PageSetup.page_width
        #ExSummary:Shows how to insert a floating image, and specify its position and size.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        shape = builder.insert_image(IMAGE_DIR + 'Logo.jpg')
        shape.wrap_type = aw.drawing.WrapType.NONE
        # Configure the shape's "relative_horizontal_position" property to treat the value of the "left" property
        # as the shape's horizontal distance, in points, from the left side of the page.
        shape.relative_horizontal_position = aw.drawing.RelativeHorizontalPosition.PAGE
        # Set the shape's horizontal distance from the left side of the page to 100.
        shape.left = 100
        # Use the "relative_vertical_position" property in a similar way to position the shape 80pt below the top of the page.
        shape.relative_vertical_position = aw.drawing.RelativeVerticalPosition.PAGE
        shape.top = 80
        # Set the shape's height, which will automatically scale the width to preserve dimensions.
        shape.height = 125
        self.assertEqual(125.0, shape.width)
        # The "bottom" and "right" properties contain the bottom and right edges of the image.
        self.assertEqual(shape.top + shape.height, shape.bottom)
        self.assertEqual(shape.left + shape.width, shape.right)
        doc.save(ARTIFACTS_DIR + 'Image.create_floating_position_size.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Image.create_floating_position_size.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, shape)
        self.assertEqual(aw.drawing.WrapType.NONE, shape.wrap_type)
        self.assertEqual(aw.drawing.RelativeHorizontalPosition.PAGE, shape.relative_horizontal_position)
        self.assertEqual(aw.drawing.RelativeVerticalPosition.PAGE, shape.relative_vertical_position)
        self.assertEqual(100.0, shape.left)
        self.assertEqual(80.0, shape.top)
        self.assertEqual(125.0, shape.height)
        self.assertEqual(125.0, shape.width)
        self.assertEqual(shape.top + shape.height, shape.bottom)
        self.assertEqual(shape.left + shape.width, shape.right)

    def test_insert_image_with_hyperlink(self):
        #ExStart
        #ExFor:ShapeBase.href
        #ExFor:ShapeBase.screen_tip
        #ExFor:ShapeBase.target
        #ExSummary:Shows how to insert a shape which contains an image, and is also a hyperlink.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        shape = builder.insert_image(IMAGE_DIR + 'Logo.jpg')
        shape.href = 'https://forum.aspose.com/'
        shape.target = 'New Window'
        shape.screen_tip = 'Aspose.Words Support Forums'
        # Ctrl + left-clicking the shape in Microsoft Word will open a new web browser window
        # and take us to the hyperlink in the "href" property.
        doc.save(ARTIFACTS_DIR + 'Image.insert_image_with_hyperlink.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Image.insert_image_with_hyperlink.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.verify_web_response_status_code(200, shape.href)
        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, shape)
        self.assertEqual('New Window', shape.target)
        self.assertEqual('Aspose.Words Support Forums', shape.screen_tip)

    def test_create_linked_image(self):
        #ExStart
        #ExFor:Shape.image_data
        #ExFor:ImageData
        #ExFor:ImageData.source_full_name
        #ExFor:ImageData.set_image(str)
        #ExFor:DocumentBuilder.insert_node
        #ExSummary:Shows how to insert a linked image into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        image_file_name = IMAGE_DIR + 'Windows MetaFile.wmf'
        # Below are two ways of applying an image to a shape so that it can display it.
        # 1 -  Set the shape to contain the image.
        shape = aw.drawing.Shape(builder.document, aw.drawing.ShapeType.IMAGE)
        shape.wrap_type = aw.drawing.WrapType.INLINE
        shape.image_data.set_image(image_file_name)
        builder.insert_node(shape)
        doc.save(ARTIFACTS_DIR + 'Image.create_linked_image.embedded.docx')
        # Every image that we store in shape will increase the size of our document.
        self.assertLess(70000, os.path.getsize(ARTIFACTS_DIR + 'Image.create_linked_image.embedded.docx'))
        doc.first_section.body.first_paragraph.remove_all_children()
        # 2 -  Set the shape to link to an image file in the local file system.
        shape = aw.drawing.Shape(builder.document, aw.drawing.ShapeType.IMAGE)
        shape.wrap_type = aw.drawing.WrapType.INLINE
        shape.image_data.source_full_name = image_file_name
        builder.insert_node(shape)
        doc.save(ARTIFACTS_DIR + 'Image.create_linked_image.linked.docx')
        # Linking to images will save space and result in a smaller document.
        # However, the document can only display the image correctly while
        # the image file is present at the location that the shape's "source_full_name" property points to.
        self.assertGreater(10000, os.path.getsize(ARTIFACTS_DIR + 'Image.create_linked_image.linked.docx'))
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Image.create_linked_image.embedded.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.verify_image_in_shape(1600, 1600, aw.drawing.ImageType.WMF, shape)
        self.assertEqual(aw.drawing.WrapType.INLINE, shape.wrap_type)
        self.assertEqual('', shape.image_data.source_full_name.replace('%20', ' '))
        doc = aw.Document(ARTIFACTS_DIR + 'Image.create_linked_image.linked.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.verify_image_in_shape(0, 0, aw.drawing.ImageType.WMF, shape)
        self.assertEqual(aw.drawing.WrapType.INLINE, shape.wrap_type)
        self.assertEqual(image_file_name, shape.image_data.source_full_name.replace('%20', ' '))

    def test_delete_all_images(self):
        #ExStart
        #ExFor:Shape.has_image
        #ExFor:Node.remove
        #ExSummary:Shows how to delete all shapes with images from a document.
        doc = aw.Document(MY_DIR + 'Images.docx')
        shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)
        self.assertEqual(9, len([node for node in shapes if node.as_shape().has_image]))
        for node in shapes:
            shape = node.as_shape()
            if shape.has_image:
                shape.remove()
        self.assertEqual(0, len([node for node in shapes if node.as_shape().has_image]))
        #ExEnd

    def test_delete_all_images_pre_order(self):
        #ExStart
        #ExFor:Node.next_pre_order(Node)
        #ExFor:Node.previous_pre_order(Node)
        #ExSummary:Shows how to traverse the document's node tree using the pre-order traversal algorithm, and delete any encountered shape with an image.
        doc = aw.Document(MY_DIR + 'Images.docx')
        self.assertEqual(9, len([node for node in doc.get_child_nodes(aw.NodeType.SHAPE, True) if node.as_shape().has_image]))
        cur_node = doc
        while cur_node is not None:
            next_node = cur_node.next_pre_order(doc)
            if cur_node.previous_pre_order(doc) is not None and next_node is not None:
                self.assertEqual(cur_node, next_node.previous_pre_order(doc))
            if cur_node.node_type == aw.NodeType.SHAPE and cur_node.as_shape().has_image:
                cur_node.remove()
            cur_node = next_node
        self.assertEqual(0, len([node for node in doc.get_child_nodes(aw.NodeType.SHAPE, True) if node.as_shape().has_image]))
        #ExEnd

    def test_scale_image(self):
        #ExStart
        #ExFor:ImageData.image_size
        #ExFor:ImageSize
        #ExFor:ImageSize.width_points
        #ExFor:ImageSize.height_points
        #ExFor:ShapeBase.width
        #ExFor:ShapeBase.height
        #ExSummary:Shows how to resize a shape with an image.
        # When we insert an image using the "insert_image" method, the builder scales the shape that displays the image so that,
        # when we view the document using 100% zoom in Microsoft Word, the shape displays the image in its actual size.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        shape = builder.insert_image(IMAGE_DIR + 'Logo.jpg')
        # A 400x400 image will create an ImageData object with an image size of 300x300pt.
        image_size = shape.image_data.image_size
        self.assertEqual(300.0, image_size.width_points)
        self.assertEqual(300.0, image_size.height_points)
        # If a shape's dimensions match the image data's dimensions,
        # then the shape is displaying the image in its original size.
        self.assertEqual(300.0, shape.width)
        self.assertEqual(300.0, shape.height)
        # Reduce the overall size of the shape by 50%.
        shape.width *= 0.5
        # Scaling factors apply to both the width and the height at the same time to preserve the shape's proportions.
        self.assertEqual(150.0, shape.width)
        self.assertEqual(150.0, shape.height)
        # When we resize the shape, the size of the image data remains the same.
        self.assertEqual(300.0, image_size.width_points)
        self.assertEqual(300.0, image_size.height_points)
        # We can reference the image data dimensions to apply a scaling based on the size of the image.
        shape.width = image_size.width_points * 1.1
        self.assertEqual(330.0, shape.width)
        self.assertEqual(330.0, shape.height)
        doc.save(ARTIFACTS_DIR + 'Image.scale_image.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Image.scale_image.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual(330.0, shape.width)
        self.assertEqual(330.0, shape.height)
        image_size = shape.image_data.image_size
        self.assertEqual(300.0, image_size.width_points)
        self.assertEqual(300.0, image_size.height_points)