import unittest

import api_example_base as aeb

import aspose.words as aw


class ExDocumentBuilderImages(aeb.ApiExampleBase):
    @unittest.skip("Streams are not supported")
    def test_insert_image_from_stream(self):

        #ExStart
        #ExFor:DocumentBuilder.insert_image(Stream)
        #ExFor:DocumentBuilder.insert_image(Stream, Double, Double)
        #ExFor:DocumentBuilder.insert_image(Stream, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        #ExSummary:Shows how to insert an image from a stream into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # using (Stream stream = File.open_read(aeb.image_dir + "Logo.jpg"))
        #
        #     # Below are three ways of inserting an image from a stream.
        #     # 1 -  Inline shape with a default size based on the image's original dimensions:
        #     builder.insert_image(stream)
        #
        #     builder.insert_break(BreakType.page_break)
        #
        #     # 2 -  Inline shape with custom dimensions:
        #     builder.insert_image(stream, ConvertUtil.pixel_to_point(250), ConvertUtil.pixel_to_point(144))
        #
        #     builder.insert_break(BreakType.page_break)
        #
        #     # 3 -  Floating shape with custom dimensions:
        #     builder.insert_image(stream, RelativeHorizontalPosition.margin, 100, RelativeVerticalPosition.margin,
        #         100, 200, 100, WrapType.square)
        #
        #
        # doc.save(artifacts_dir + "DocumentBuilderImages.insert_image_from_stream.docx")
        # #ExEnd
        #
        # doc = new Document(artifacts_dir + "DocumentBuilderImages.insert_image_from_stream.docx")
        #
        # Shape imageShape = (Shape)doc.get_child(NodeType.shape, 0, true)
        #
        # self.assertEqual(300.0, imageShape.height)
        # self.assertEqual(300.0, imageShape.width)
        # self.assertEqual(0.0, imageShape.left)
        # self.assertEqual(0.0, imageShape.top)
        #
        # self.assertEqual(WrapType.inline, imageShape.wrap_type)
        # self.assertEqual(RelativeHorizontalPosition.column, imageShape.relative_horizontal_position)
        # self.assertEqual(RelativeVerticalPosition.paragraph, imageShape.relative_vertical_position)
        #
        # TestUtil.verify_image_in_shape(400, 400, ImageType.jpeg, imageShape)
        # self.assertEqual(300.0, imageShape.image_data.image_size.height_points)
        # self.assertEqual(300.0, imageShape.image_data.image_size.width_points)
        #
        # imageShape = (Shape)doc.get_child(NodeType.shape, 1, true)
        #
        # self.assertEqual(108.0, imageShape.height)
        # self.assertEqual(187.5, imageShape.width)
        # self.assertEqual(0.0, imageShape.left)
        # self.assertEqual(0.0, imageShape.top)
        #
        # self.assertEqual(WrapType.inline, imageShape.wrap_type)
        # self.assertEqual(RelativeHorizontalPosition.column, imageShape.relative_horizontal_position)
        # self.assertEqual(RelativeVerticalPosition.paragraph, imageShape.relative_vertical_position)
        #
        # TestUtil.verify_image_in_shape(400, 400, ImageType.jpeg, imageShape)
        # self.assertEqual(300.0, imageShape.image_data.image_size.height_points)
        # self.assertEqual(300.0, imageShape.image_data.image_size.width_points)
        #
        # imageShape = (Shape)doc.get_child(NodeType.shape, 2, true)
        #
        # self.assertEqual(100.0, imageShape.height)
        # self.assertEqual(200.0, imageShape.width)
        # self.assertEqual(100.0, imageShape.left)
        # self.assertEqual(100.0, imageShape.top)
        #
        # self.assertEqual(WrapType.square, imageShape.wrap_type)
        # self.assertEqual(RelativeHorizontalPosition.margin, imageShape.relative_horizontal_position)
        # self.assertEqual(RelativeVerticalPosition.margin, imageShape.relative_vertical_position)
        #
        # TestUtil.verify_image_in_shape(400, 400, ImageType.jpeg, imageShape)
        # self.assertEqual(300.0, imageShape.image_data.image_size.height_points)
        # self.assertEqual(300.0, imageShape.image_data.image_size.width_points)

    @unittest.skip("No type casting (lines 120, 135)")
    def test_insert_image_from_filename(self):

        #ExStart
        #ExFor:DocumentBuilder.insert_image(String)
        #ExFor:DocumentBuilder.insert_image(String, Double, Double)
        #ExFor:DocumentBuilder.insert_image(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        #ExSummary:Shows how to insert an image from the local file system into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Below are three ways of inserting an image from a local system filename.
        # 1 -  Inline shape with a default size based on the image's original dimensions:
        builder.insert_image(aeb.image_dir + "Logo.jpg")

        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # 2 -  Inline shape with custom dimensions:
        builder.insert_image(aeb.image_dir + "Transparent background logo.png", aw.ConvertUtil.pixel_to_point(250),
            aw.ConvertUtil.pixel_to_point(144))

        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # 3 -  Floating shape with custom dimensions:
        builder.insert_image(aeb.image_dir + "Windows MetaFile.wmf", aw.drawing.RelativeHorizontalPosition.MARGIN, 100,
            aw.drawing.RelativeVerticalPosition.MARGIN, 100, 200, 100, aw.drawing.WrapType.SQUARE)

        doc.save(aeb.artifacts_dir + "DocumentBuilderImages.insert_image_from_filename.docx")
        #ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilderImages.insert_image_from_filename.docx")

        # imageShape = (Shape)doc.get_child(NodeType.shape, 0, true)
        #
        # self.assertEqual(300.0, imageShape.height)
        # self.assertEqual(300.0, imageShape.width)
        # self.assertEqual(0.0, imageShape.left)
        # self.assertEqual(0.0, imageShape.top)
        #
        # self.assertEqual(WrapType.inline, imageShape.wrap_type)
        # self.assertEqual(RelativeHorizontalPosition.column, imageShape.relative_horizontal_position)
        # self.assertEqual(RelativeVerticalPosition.paragraph, imageShape.relative_vertical_position)
        #
        # TestUtil.verify_image_in_shape(400, 400, ImageType.jpeg, imageShape)
        # self.assertEqual(300.0, imageShape.image_data.image_size.height_points)
        # self.assertEqual(300.0, imageShape.image_data.image_size.width_points)
        #
        # imageShape = (Shape)doc.get_child(NodeType.shape, 1, true)
        #
        # self.assertEqual(108.0, imageShape.height)
        # self.assertEqual(187.5, imageShape.width)
        # self.assertEqual(0.0, imageShape.left)
        # self.assertEqual(0.0, imageShape.top)
        #
        # self.assertEqual(WrapType.inline, imageShape.wrap_type)
        # self.assertEqual(RelativeHorizontalPosition.column, imageShape.relative_horizontal_position)
        # self.assertEqual(RelativeVerticalPosition.paragraph, imageShape.relative_vertical_position)
        #
        # TestUtil.verify_image_in_shape(400, 400, ImageType.png, imageShape)
        # self.assertEqual(300.0, imageShape.image_data.image_size.height_points)
        # self.assertEqual(300.0, imageShape.image_data.image_size.width_points)
        #
        # imageShape = (Shape)doc.get_child(NodeType.shape, 2, true)
        #
        # self.assertEqual(100.0, imageShape.height)
        # self.assertEqual(200.0, imageShape.width)
        # self.assertEqual(100.0, imageShape.left)
        # self.assertEqual(100.0, imageShape.top)
        #
        # self.assertEqual(WrapType.square, imageShape.wrap_type)
        # self.assertEqual(RelativeHorizontalPosition.margin, imageShape.relative_horizontal_position)
        # self.assertEqual(RelativeVerticalPosition.margin, imageShape.relative_vertical_position)
        #
        # TestUtil.verify_image_in_shape(1600, 1600, ImageType.wmf, imageShape)
        # self.assertEqual(400.0, imageShape.image_data.image_size.height_points)
        # self.assertEqual(400.0, imageShape.image_data.image_size.width_points)

    def test_insert_svg_image(self):

        #ExStart
        #ExFor:DocumentBuilder.insert_image(String)
        #ExSummary:Shows how to determine which image will be inserted.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_image(aeb.image_dir + "Scalable Vector Graphics.svg")

        # Aspose.words insert SVG image to the document as PNG with svgBlip extension
        # that contains the original vector SVG image representation.
        doc.save(aeb.artifacts_dir + "DocumentBuilderImages.insert_svg_image.svg_with_svg_blip.docx")

        # Aspose.words insert SVG image to the document as PNG, just like Microsoft Word does for old format.
        doc.save(aeb.artifacts_dir + "DocumentBuilderImages.insert_svg_image.svg.doc")

        doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2003)

        # Aspose.words insert SVG image to the document as EMF metafile to keep the image in vector representation.
        doc.save(aeb.artifacts_dir + "DocumentBuilderImages.insert_svg_image.emf.docx")
        #ExEnd

    @unittest.skip("No type casting (lines 221, 236, 251)")
    def test_insert_image_from_image_object(self):

        #ExStart
        #ExFor:DocumentBuilder.insert_image(Image)
        #ExFor:DocumentBuilder.insert_image(Image, Double, Double)
        #ExFor:DocumentBuilder.insert_image(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        #ExSummary:Shows how to insert an image from an object into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        image = aw.Image.from_file(aeb.image_dir + "Logo.jpg")

        # Below are three ways of inserting an image from an Image object instance.
        # 1 -  Inline shape with a default size based on the image's original dimensions:
        builder.insert_image(image)

        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # 2 -  Inline shape with custom dimensions:
        builder.insert_image(image, aw.ConvertUtil.pixel_to_point(250), aw.ConvertUtil.pixel_to_point(144))

        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # 3 -  Floating shape with custom dimensions:
        builder.insert_image(image, aw.RelativeHorizontalPosition.MARGIN, 100, aw.RelativeVerticalPosition.MARGIN,
        100, 200, 100, aw.WrapType.SQUARE)

        doc.save(aeb.artifacts_dir + "DocumentBuilderImages.insert_image_from_image_object.docx")
        #ExEnd

        doc = aw.Document(aeb.artifacts_dir + "DocumentBuilderImages.insert_image_from_image_object.docx")

        # Shape imageShape = (Shape)doc.get_child(NodeType.shape, 0, true)
        # 
        # self.assertEqual(300.0, imageShape.height)
        # self.assertEqual(300.0, imageShape.width)
        # self.assertEqual(0.0, imageShape.left)
        # self.assertEqual(0.0, imageShape.top)
        # 
        # self.assertEqual(WrapType.inline, imageShape.wrap_type)
        # self.assertEqual(RelativeHorizontalPosition.column, imageShape.relative_horizontal_position)
        # self.assertEqual(RelativeVerticalPosition.paragraph, imageShape.relative_vertical_position)
        # 
        # TestUtil.verify_image_in_shape(400, 400, ImageType.jpeg, imageShape)
        # self.assertEqual(300.0, imageShape.image_data.image_size.height_points)
        # self.assertEqual(300.0, imageShape.image_data.image_size.width_points)
        # 
        # imageShape = (Shape)doc.get_child(NodeType.shape, 1, true)
        # 
        # self.assertEqual(108.0, imageShape.height)
        # self.assertEqual(187.5d, imageShape.width)
        # self.assertEqual(0.0, imageShape.left)
        # self.assertEqual(0.0, imageShape.top)
        # 
        # self.assertEqual(WrapType.inline, imageShape.wrap_type)
        # self.assertEqual(RelativeHorizontalPosition.column, imageShape.relative_horizontal_position)
        # self.assertEqual(RelativeVerticalPosition.paragraph, imageShape.relative_vertical_position)
        # 
        # TestUtil.verify_image_in_shape(400, 400, ImageType.jpeg, imageShape)
        # self.assertEqual(300.0, imageShape.image_data.image_size.height_points)
        # self.assertEqual(300.0, imageShape.image_data.image_size.width_points)
        # 
        # imageShape = (Shape)doc.get_child(NodeType.shape, 2, true)
        # 
        # self.assertEqual(100.0, imageShape.height)
        # self.assertEqual(200.0, imageShape.width)
        # self.assertEqual(100.0, imageShape.left)
        # self.assertEqual(100.0, imageShape.top)
        # 
        # self.assertEqual(WrapType.square, imageShape.wrap_type)
        # self.assertEqual(RelativeHorizontalPosition.margin, imageShape.relative_horizontal_position)
        # self.assertEqual(RelativeVerticalPosition.margin, imageShape.relative_vertical_position)
        # 
        # TestUtil.verify_image_in_shape(400, 400, ImageType.jpeg, imageShape)
        # self.assertEqual(300.0, imageShape.image_data.image_size.height_points)
        # self.assertEqual(300.0, imageShape.image_data.image_size.width_points)

    @unittest.skip("No type casting (lines 305, 320, 335), memory streams are not supported")
    def test_insert_image_from_byte_array(self):

        #ExStart
        #ExFor:DocumentBuilder.insert_image(Byte[])
        #ExFor:DocumentBuilder.insert_image(Byte[], Double, Double)
        #ExFor:DocumentBuilder.insert_image(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        #ExSummary:Shows how to insert an image from a byte array into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        image = aw.Image.from_file(aeb.image_dir + "Logo.jpg")

        # using (MemoryStream ms = new MemoryStream())
        #
        #     image.save(ms, ImageFormat.png)
        #     byte[] imageByteArray = ms.to_array()
        #
        #     # Below are three ways of inserting an image from a byte array.
        #     # 1 -  Inline shape with a default size based on the image's original dimensions:
        #     builder.insert_image(imageByteArray)
        #
        #     builder.insert_break(BreakType.page_break)
        #
        #     # 2 -  Inline shape with custom dimensions:
        #     builder.insert_image(imageByteArray, ConvertUtil.pixel_to_point(250), ConvertUtil.pixel_to_point(144))
        #
        #     builder.insert_break(BreakType.page_break)
        #
        #     # 3 -  Floating shape with custom dimensions:
        #     builder.insert_image(imageByteArray, RelativeHorizontalPosition.margin, 100, RelativeVerticalPosition.margin,
        #     100, 200, 100, WrapType.square)
        #
        #
        # doc.save(artifacts_dir + "DocumentBuilderImages.insert_image_from_byte_array.docx")
        # #ExEnd
        #
        # doc = new Document(artifacts_dir + "DocumentBuilderImages.insert_image_from_byte_array.docx")
        #
        # Shape imageShape = (Shape)doc.get_child(NodeType.shape, 0, true)
        #
        # self.assertEqual(300.0, imageShape.height, 0.1d)
        # self.assertEqual(300.0, imageShape.width, 0.1d)
        # self.assertEqual(0.0, imageShape.left)
        # self.assertEqual(0.0, imageShape.top)
        #
        # self.assertEqual(WrapType.inline, imageShape.wrap_type)
        # self.assertEqual(RelativeHorizontalPosition.column, imageShape.relative_horizontal_position)
        # self.assertEqual(RelativeVerticalPosition.paragraph, imageShape.relative_vertical_position)
        #
        # TestUtil.verify_image_in_shape(400, 400, ImageType.png, imageShape)
        # self.assertEqual(300.0, imageShape.image_data.image_size.height_points, 0.1d)
        # self.assertEqual(300.0, imageShape.image_data.image_size.width_points, 0.1d)
        #
        # imageShape = (Shape)doc.get_child(NodeType.shape, 1, true)
        #
        # self.assertEqual(108.0, imageShape.height)
        # self.assertEqual(187.5d, imageShape.width)
        # self.assertEqual(0.0, imageShape.left)
        # self.assertEqual(0.0, imageShape.top)
        #
        # self.assertEqual(WrapType.inline, imageShape.wrap_type)
        # self.assertEqual(RelativeHorizontalPosition.column, imageShape.relative_horizontal_position)
        # self.assertEqual(RelativeVerticalPosition.paragraph, imageShape.relative_vertical_position)
        #
        # TestUtil.verify_image_in_shape(400, 400, ImageType.png, imageShape)
        # self.assertEqual(300.0, imageShape.image_data.image_size.height_points, 0.1d)
        # self.assertEqual(300.0, imageShape.image_data.image_size.width_points, 0.1d)
        #
        # imageShape = (Shape)doc.get_child(NodeType.shape, 2, true)
        #
        # self.assertEqual(100.0, imageShape.height)
        # self.assertEqual(200.0, imageShape.width)
        # self.assertEqual(100.0, imageShape.left)
        # self.assertEqual(100.0, imageShape.top)
        #
        # self.assertEqual(WrapType.square, imageShape.wrap_type)
        # self.assertEqual(RelativeHorizontalPosition.margin, imageShape.relative_horizontal_position)
        # self.assertEqual(RelativeVerticalPosition.margin, imageShape.relative_vertical_position)
        #
        # TestUtil.verify_image_in_shape(400, 400, ImageType.png, imageShape)
        # self.assertEqual(300.0, imageShape.image_data.image_size.height_points, 0.1d)
        # self.assertEqual(300.0, imageShape.image_data.image_size.width_points, 0.1d)


if __name__ == '__main__':
    unittest.main()
