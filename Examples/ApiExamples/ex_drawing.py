import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

import api_example_base as aeb


class ExDrawing(aeb.ApiExampleBase):

    def test_various_shapes(self):
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
        filledInArrow = aw.drawing.Shape(doc, aw.drawing.ShapeType.ARROW)
        filledInArrow.width = 200
        filledInArrow.height = 40
        filledInArrow.top = 100
        filledInArrow.fill.fore_color = drawing.Color.green
        filledInArrow.fill.visible = True

        builder.insert_node(filledInArrow)

        # 4 -  Arrow with a flipped orientation filled in with the Aspose logo:
        filledInArrowImg = aw.drawing.Shape(doc, aw.drawing.ShapeType.ARROW)
        filledInArrowImg.width = 200
        filledInArrowImg.height = 40
        filledInArrowImg.top = 160
        filledInArrowImg.flip_orientation = aw.drawing.FlipOrientation.BOTH

        # imageBytes = open(aeb.image_dir + "Logo.jpg")
        # with io.BytesIO(imageBytes) as stream:
        #
        # with open(aeb.image_dir + "Logo.jpg") as imageBytes:
        #
        #     image = drawing.Image.from_stream(imageBytes)
        #     # When we flip the orientation of our arrow, we also flip the image that the arrow contains.
        #     # Flip the image the other way to cancel this out before getting the shape to display it.
        #     image.rotate_flip(drawing.RotateFlipType.rotate_none_flip_xy)
        #
        #     filledInArrowImg.image_data.set_image(image)
        #     filledInArrowImg.stroke.join_style = aw.drawing.JoinStyle.ROUND
        #
        #     builder.insert_node(filledInArrowImg)

        doc.save(aeb.artifacts_dir + "Drawing.various_shapes.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "Drawing.various_shapes.docx")

        # self.assertEqual(4, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)

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

        filledInArrow = doc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()

        self.assertEqual(aw.drawing.ShapeType.ARROW, filledInArrow.shape_type)
        self.assertEqual(200.0, filledInArrow.width)
        self.assertEqual(40.0, filledInArrow.height)
        self.assertEqual(100.0, filledInArrow.top)
        self.assertEqual(drawing.Color.green.to_argb(), filledInArrow.fill.fore_color.to_argb())
        self.assertTrue(filledInArrow.fill.visible)

        # filledInArrowImg = doc.get_child(aw.NodeType.SHAPE, 3, True).as_shape()
        #
        # self.assertEqual(aw.drawing.ShapeType.ARROW, filledInArrowImg.shape_type)
        # self.assertEqual(200.0, filledInArrowImg.width)
        # self.assertEqual(40.0, filledInArrowImg.height)
        # self.assertEqual(160.0, filledInArrowImg.top)
        # self.assertEqual(aw.drawing.FlipOrientation.both, filledInArrowImg.flip_orientation)

    # def test_type_of_image(self):
    #
    #     #ExStart
    #     #ExFor:Drawing.image_type
    #     #ExSummary:Shows how to add an image to a shape and check its type.
    #     doc = aw.Document()
    #     builder = aw.DocumentBuilder(doc)
    #
    #     byte[] imageBytes = File.read_all_bytes(aeb.image_dir + "Logo.jpg")
    #
    #     using (MemoryStream stream = new MemoryStream(imageBytes))
    #
    #         Image image = Image.from_stream(stream)
    #
    #         # The image in the URL is a .gif. Inserting it into a document converts it into a .png.
    #         Shape imgShape = builder.insert_image(image)
    #         self.assertEqual(ImageType.jpeg, imgShape.image_data.image_type)

    # ExEnd

    def test_fill_solid(self):
        # ExStart
        # ExFor:Fill.color()
        # ExFor:Fill.color(Color)
        # ExSummary:Shows how to convert any of the fills back to solid fill.
        doc = aw.Document(aeb.my_dir + "Two color gradient.docx")

        # Get Fill object for Font of the first Run.
        fill = doc.first_section.body.paragraphs[0].runs[0].font.fill

        # Check Fill properties of the Font.
        print("The type of the fill is: 0", fill.fill_type)
        print("The foreground color of the fill is: 0", fill.fore_color)
        print("The fill is transparent at 0%", fill.transparency * 100)

        # Change type of the fill to Solid with uniform green color.
        fill.solid(drawing.Color.green)
        print("\nThe fill is changed:")
        print("The type of the fill is: 0", fill.fill_type)
        print("The foreground color of the fill is: 0", fill.fore_color)
        print("The fill transparency is 0%", fill.transparency * 100)

        doc.save(aeb.artifacts_dir + "Drawing.fill_solid.docx")
        # ExEnd

    # def test_save_all_images(self):
    #
    #     #ExStart
    #     #ExFor:ImageData.has_image
    #     #ExFor:ImageData.to_image
    #     #ExFor:ImageData.save(Stream)
    #     #ExSummary:Shows how to save all images from a document to the file system.
    #     Document imgSourceDoc = new Document(aeb.my_dir + "Images.docx")
    #
    #     # Shapes with the "HasImage" flag set store and display all the document's images.
    #     IEnumerable<Shape> shapesWithImages =
    #         imgSourceDoc.get_child_nodes(NodeType.shape, true).cast<Shape>().where(s => s.has_image)
    #
    #     # Go through each shape and save its image.
    #     ImageFormatConverter formatConverter = new ImageFormatConverter()
    #
    #     using (IEnumerator<Shape> enumerator = shapesWithImages.get_enumerator())
    #
    #         int shapeIndex = 0
    #
    #         while (enumerator.move_next())
    #
    #             ImageData imageData = enumerator.current.image_data
    #             ImageFormat format = imageData.to_image().raw_format
    #             string fileExtension = formatConverter.convert_to_string(format)
    #
    #             using (FileStream fileStream = File.create(aeb.artifacts_dir + $"Drawing.save_all_images.++shapeIndex.file_extension"))
    #                 imageData.save(fileStream)
    #
    #
    #     #ExEnd
    #
    #     string[] imageFileNames = Directory.get_files(aeb.artifacts_dir).where(s => s.starts_with(aeb.artifacts_dir + "Drawing.save_all_images.")).order_by(s => s).to_array()
    #     List<FileInfo> fileInfos = imageFileNames.select(s => new FileInfo(s)).to_list()
    #
    #     TestUtil.verify_image(2467, 1500, fileInfos[0].full_name)
    #     self.assertEqual(".jpeg", fileInfos[0].extension)
    #     TestUtil.verify_image(400, 400, fileInfos[1].full_name)
    #     self.assertEqual(".png", fileInfos[1].extension)
    #     TestUtil.verify_image(382, 138, fileInfos[2].full_name)
    #     self.assertEqual(".emf", fileInfos[2].extension)
    #     TestUtil.verify_image(1600, 1600, fileInfos[3].full_name)
    #     self.assertEqual(".wmf", fileInfos[3].extension)
    #     TestUtil.verify_image(534, 534, fileInfos[4].full_name)
    #     self.assertEqual(".emf", fileInfos[4].extension)
    #     TestUtil.verify_image(1260, 660, fileInfos[5].full_name)
    #     self.assertEqual(".jpeg", fileInfos[5].extension)
    #     TestUtil.verify_image(1125, 1500, fileInfos[6].full_name)
    #     self.assertEqual(".jpeg", fileInfos[6].extension)
    #     TestUtil.verify_image(1027, 1500, fileInfos[7].full_name)
    #     self.assertEqual(".jpeg", fileInfos[7].extension)
    #     TestUtil.verify_image(1200, 1500, fileInfos[8].full_name)
    #     self.assertEqual(".jpeg", fileInfos[8].extension)

    # def test_import_image(self):
    #
    #     #ExStart
    #     #ExFor:ImageData.set_image(Image)
    #     #ExFor:ImageData.set_image(Stream)
    #     #ExSummary:Shows how to display images from the local file system in a document.
    #     doc = aw.Document()
    #
    #     # To display an image in a document, we will need to create a shape
    #     # which will contain an image, and then append it to the document's body.
    #     Shape imgShape
    #
    #     # Below are two ways of getting an image from a file in the local file system.
    #     # 1 -  Create an image object from an image file:
    #     using (Image srcImage = Image.from_file(aeb.image_dir + "Logo.jpg"))
    #
    #         imgShape = new Shape(doc, ShapeType.image)
    #         doc.first_section.body.first_paragraph.append_child(imgShape)
    #         imgShape.image_data.set_image(srcImage)
    #
    #
    #     # 2 -  Open an image file from the local file system using a stream:
    #     using (Stream stream = new FileStream(aeb.image_dir + "Logo.jpg", FileMode.open, FileAccess.read))
    #
    #         imgShape = new Shape(doc, ShapeType.image)
    #         doc.first_section.body.first_paragraph.append_child(imgShape)
    #         imgShape.image_data.set_image(stream)
    #         imgShape.left = 150.0f
    #
    #
    #     doc.save(aeb.artifacts_dir + "Drawing.import_image.docx")
    #     #ExEnd
    #
    #     doc = new Document(aeb.artifacts_dir + "Drawing.import_image.docx")
    #
    #     self.assertEqual(2, doc.get_child_nodes(NodeType.shape, true).count)
    #
    #     imgShape = (Shape)doc.get_child(NodeType.shape, 0, true)
    #
    #     TestUtil.verify_image_in_shape(400, 400, ImageType.jpeg, imgShape)
    #     self.assertEqual(0.0d, imgShape.left)
    #     self.assertEqual(0.0d, imgShape.top)
    #     self.assertEqual(300.0d, imgShape.height)
    #     self.assertEqual(300.0d, imgShape.width)
    #     TestUtil.verify_image_in_shape(400, 400, ImageType.jpeg, imgShape)
    #
    #     imgShape = (Shape)doc.get_child(NodeType.shape, 1, true)
    #
    #     TestUtil.verify_image_in_shape(400, 400, ImageType.jpeg, imgShape)
    #     self.assertEqual(150.0d, imgShape.left)
    #     self.assertEqual(0.0d, imgShape.top)
    #     self.assertEqual(300.0d, imgShape.height)
    #     self.assertEqual(300.0d, imgShape.width)

    # endif

    # [Test]
    # public void StrokePattern()
    #
    #     #ExStart
    #     #ExFor:Stroke.color_2
    #     #ExFor:Stroke.image_bytes
    #     #ExSummary:Shows how to process shape stroke features.
    #     Document doc = new Document(aeb.my_dir + "Shape stroke pattern border.docx")
    #     Shape shape = (Shape)doc.get_child(NodeType.shape, 0, true)
    #     Stroke stroke = shape.stroke
    #
    #     # Strokes can have two colors, which are used to create a pattern defined by two-tone image data.
    #     # Strokes with a single color do not use the Color2 property.
    #     self.assertEqual(Color.from_argb(255, 128, 0, 0), stroke.color)
    #     self.assertEqual(Color.from_argb(255, 255, 255, 0), stroke.color_2)
    #
    #     Assert.not_null(stroke.image_bytes)
    #     File.write_all_bytes(aeb.artifacts_dir + "Drawing.stroke_pattern.png", stroke.image_bytes)
    #     #ExEnd
    #
    #     TestUtil.verify_image(8, 8, aeb.artifacts_dir + "Drawing.stroke_pattern.png")
    #
    #
    # #ExStart
    # #ExFor:DocumentVisitor.visit_shape_end(Shape)
    # #ExFor:DocumentVisitor.visit_shape_start(Shape)
    # #ExFor:DocumentVisitor.visit_group_shape_end(GroupShape)
    # #ExFor:DocumentVisitor.visit_group_shape_start(GroupShape)
    # #ExFor:Drawing.group_shape
    # #ExFor:Drawing.group_shape.#ctor(DocumentBase)
    # #ExFor:Drawing.group_shape.accept(DocumentVisitor)
    # #ExFor:ShapeBase.is_group
    # #ExFor:ShapeBase.shape_type
    # #ExSummary:Shows how to create a group of shapes, and print its contents using a document visitor.
    # [Test] #ExSkip
    # public void GroupOfShapes()
    #
    #     doc = aw.Document()
    #     builder = aw.DocumentBuilder(doc)
    #
    #     # If you need to create "NonPrimitive" shapes, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
    #     # TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
    #     # please use DocumentBuilder.insert_shape methods.
    #     Shape balloon = new Shape(doc, ShapeType.balloon)
    #
    #         Width = 200,
    #         Height = 200,
    #         Stroke =  Color = Color.red
    #
    #
    #     Shape cube = new Shape(doc, ShapeType.cube)
    #
    #         Width = 100,
    #         Height = 100,
    #         Stroke =  Color = Color.blue
    #
    #
    #     GroupShape group = new GroupShape(doc)
    #     group.append_child(balloon)
    #     group.append_child(cube)
    #
    #     self.assertTrue(group.is_group)
    #
    #     builder.insert_node(group)
    #
    #     ShapeGroupPrinter printer = new ShapeGroupPrinter()
    #     group.accept(printer)
    #
    #     print(printer.get_text())
    #     TestGroupShapes(doc) #ExSkip
    #
    #
    # # <summary>
    # # Prints the contents of a visited shape group to the console.
    # # </summary>
    # public class ShapeGroupPrinter : DocumentVisitor
    #
    #     public ShapeGroupPrinter()
    #
    #         mBuilder = new StringBuilder()
    #
    #
    #     public string GetText()
    #
    #         return mBuilder.to_string()
    #
    #
    #     public override VisitorAction VisitGroupShapeStart(GroupShape groupShape)
    #
    #         mBuilder.append_line("Shape group started:")
    #         return VisitorAction.continue
    #
    #
    #     public override VisitorAction VisitGroupShapeEnd(GroupShape groupShape)
    #
    #         mBuilder.append_line("End of shape group")
    #         return VisitorAction.continue
    #
    #
    #     public override VisitorAction VisitShapeStart(Shape shape)
    #
    #         mBuilder.append_line("\tShape - " + shape.shape_type + ":")
    #         mBuilder.append_line("\t\tWidth: " + shape.width)
    #         mBuilder.append_line("\t\tHeight: " + shape.height)
    #         mBuilder.append_line("\t\tStroke color: " + shape.stroke.color)
    #         mBuilder.append_line("\t\tFill color: " + shape.fill.fore_color)
    #         return VisitorAction.continue
    #
    #
    #     public override VisitorAction VisitShapeEnd(Shape shape)
    #
    #         mBuilder.append_line("\tEnd of shape")
    #         return VisitorAction.continue
    #
    #
    #     private readonly StringBuilder mBuilder
    #
    # #ExEnd
    #
    # private static void TestGroupShapes(Document doc)
    #
    #     doc = DocumentHelper.save_open(doc)
    #     GroupShape shapes = (GroupShape)doc.get_child(NodeType.group_shape, 0, true)
    #
    #     self.assertEqual(2, shapes.child_nodes.count)
    #
    #     Shape shape = (Shape)shapes.child_nodes[0]
    #
    #     self.assertEqual(ShapeType.balloon, shape.shape_type)
    #     self.assertEqual(200.0d, shape.width)
    #     self.assertEqual(200.0d, shape.height)
    #     self.assertEqual(Color.red.to_argb(), shape.stroke_color.to_argb())
    #
    #     shape = (Shape)shapes.child_nodes[1]
    #
    #     self.assertEqual(ShapeType.cube, shape.shape_type)
    #     self.assertEqual(100.0d, shape.width)
    #     self.assertEqual(100.0d, shape.height)
    #     self.assertEqual(Color.blue.to_argb(), shape.stroke_color.to_argb())
    #

    # def test_text_box(self):
    #
    #     #ExStart
    #     #ExFor:Drawing.layout_flow
    #     #ExSummary:Shows how to add text to a text box, and change its orientation
    #     doc = aw.Document()
    #     builder = aw.DocumentBuilder(doc)
    #
    #     textbox = aw.Shape(doc, aw.ShapeType.TEXT_BOX)
    #     Width = 100,
    #     Height = 100,
    #     TextBox =  LayoutFlow = LayoutFlow.bottom_to_top
    #
    #
    #     textbox.append_child(new Paragraph(doc))
    #     builder.insert_node(textbox)
    #
    #     builder.move_to(textbox.first_paragraph)
    #     builder.write("This text is flipped 90 degrees to the left.")
    #
    #     doc.save(aeb.artifacts_dir + "Drawing.text_box.docx")
    #     #ExEnd
    #
    #     doc = new Document(aeb.artifacts_dir + "Drawing.text_box.docx")
    #     textbox = (Shape)doc.get_child(NodeType.shape, 0, true)
    #
    #     self.assertEqual(ShapeType.text_box, textbox.shape_type)
    #     self.assertEqual(100.0d, textbox.width)
    #     self.assertEqual(100.0d, textbox.height)
    #     self.assertEqual(LayoutFlow.bottom_to_top, textbox.text_box.layout_flow)
    #     self.assertEqual("This text is flipped 90 degrees to the left.", textbox.get_text().strip())

    # def test_get_data_from_image(self):
    #
    #     #ExStart
    #     #ExFor:ImageData.image_bytes
    #     #ExFor:ImageData.to_byte_array
    #     #ExFor:ImageData.to_stream
    #     #ExSummary:Shows how to create an image file from a shape's raw image data.
    #     Document imgSourceDoc = new Document(aeb.my_dir + "Images.docx")
    #     self.assertEqual(10, imgSourceDoc.get_child_nodes(NodeType.shape, true).count) #ExSkip
    #
    #     Shape imgShape = (Shape) imgSourceDoc.get_child(NodeType.shape, 0, true)
    #
    #     self.assertTrue(imgShape.has_image)
    #
    #     # ToByteArray() returns the array stored in the ImageBytes property.
    #     self.assertEqual(imgShape.image_data.image_bytes, imgShape.image_data.to_byte_array())
    #
    #     # Save the shape's image data to an image file in the local file system.
    #     using (Stream imgStream = imgShape.image_data.to_stream())
    #
    #         using (FileStream outStream = new FileStream(aeb.artifacts_dir + "Drawing.get_data_from_image.png",
    #             FileMode.create, FileAccess.read_write))
    #
    #             imgStream.copy_to(outStream)
    #
    #
    #     #ExEnd
    #
    #     TestUtil.verify_image(2467, 1500, aeb.artifacts_dir + "Drawing.get_data_from_image.png")

    def test_image_data(self):
        # ExStart
        # ExFor:ImageData.bi_level
        # ExFor:ImageData.borders
        # ExFor:ImageData.brightness
        # ExFor:ImageData.chroma_key
        # ExFor:ImageData.contrast
        # ExFor:ImageData.crop_bottom
        # ExFor:ImageData.crop_left
        # ExFor:ImageData.crop_right
        # ExFor:ImageData.crop_top
        # ExFor:ImageData.gray_scale
        # ExFor:ImageData.is_link
        # ExFor:ImageData.is_link_only
        # ExFor:ImageData.title
        # ExSummary:Shows how to edit a shape's image data.
        imgSourceDoc = aw.Document(aeb.my_dir + "Images.docx")
        sourceShape = imgSourceDoc.get_child_nodes(aw.NodeType.SHAPE, True)[0].as_shape()

        dstDoc = aw.Document()

        # Import a shape from the source document and append it to the first paragraph.
        importedShape = dstDoc.import_node(sourceShape, True).as_shape()
        dstDoc.first_section.body.first_paragraph.append_child(importedShape)

        # The imported shape contains an image. We can access the image's properties and raw data via the ImageData object.
        imageData = importedShape.image_data
        imageData.title = "Imported Image"

        self.assertTrue(imageData.has_image)

        # If an image has no borders, its ImageData object will define the border color as empty.
        self.assertEqual(4, imageData.borders.count)
        # self.assertEqual(drawing.Color.empty, imageData.borders[0].color)

        # This image does not link to another shape or image file in the local file system.
        self.assertFalse(imageData.is_link)
        self.assertFalse(imageData.is_link_only)

        # The "Brightness" and "Contrast" properties define image brightness and contrast
        # on a 0-1 scale, with the default value at 0.5.
        imageData.brightness = 0.8
        imageData.contrast = 1.0

        # The above brightness and contrast values have created an image with a lot of white.
        # We can select a color with the ChromaKey property to replace with transparency, such as white.
        imageData.chroma_key = drawing.Color.white

        # Import the source shape again and set the image to monochrome.
        importedShape = dstDoc.import_node(sourceShape, True).as_shape()
        dstDoc.first_section.body.first_paragraph.append_child(importedShape)

        importedShape.image_data.gray_scale = True

        # Import the source shape again to create a third image and set it to BiLevel.
        # BiLevel sets every pixel to either black or white, whichever is closer to the original color.
        importedShape = dstDoc.import_node(sourceShape, True).as_shape()
        dstDoc.first_section.body.first_paragraph.append_child(importedShape)

        importedShape.image_data.bi_level = True

        # Cropping is determined on a 0-1 scale. Cropping a side by 0.3
        # will crop 30% of the image out at the cropped side.
        importedShape.image_data.crop_bottom = 0.3
        importedShape.image_data.crop_left = 0.3
        importedShape.image_data.crop_top = 0.3
        importedShape.image_data.crop_right = 0.3

        dstDoc.save(aeb.artifacts_dir + "Drawing.image_data.docx")
        # ExEnd

        imgSourceDoc = aw.Document(aeb.artifacts_dir + "Drawing.image_data.docx")
        sourceShape = imgSourceDoc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        # TestUtil.verify_image_in_shape(2467, 1500, ImageType.jpeg, sourceShape)
        self.assertEqual("Imported Image", sourceShape.image_data.title)
        self.assertAlmostEqual(0.8, sourceShape.image_data.brightness, 1)
        self.assertAlmostEqual(1.0, sourceShape.image_data.contrast, 1)
        self.assertEqual(drawing.Color.white.to_argb(), sourceShape.image_data.chroma_key.to_argb())

        sourceShape = imgSourceDoc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()

        # TestUtil.verify_image_in_shape(2467, 1500, ImageType.jpeg, sourceShape)
        self.assertTrue(sourceShape.image_data.gray_scale)

        sourceShape = imgSourceDoc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()

        # TestUtil.verify_image_in_shape(2467, 1500, ImageType.jpeg, sourceShape)
        self.assertTrue(sourceShape.image_data.bi_level)
        self.assertAlmostEqual(0.3, sourceShape.image_data.crop_bottom, 1)
        self.assertAlmostEqual(0.3, sourceShape.image_data.crop_left, 1)
        self.assertAlmostEqual(0.3, sourceShape.image_data.crop_top, 1)
        self.assertAlmostEqual(0.3, sourceShape.image_data.crop_right, 1)

    def test_image_size(self):
        # ExStart
        # ExFor:ImageSize.height_pixels
        # ExFor:ImageSize.horizontal_resolution
        # ExFor:ImageSize.vertical_resolution
        # ExFor:ImageSize.width_pixels
        # ExSummary:Shows how to read the properties of an image in a shape.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a shape into the document which contains an image taken from our local file system.
        shape = builder.insert_image(aeb.image_dir + "Logo.jpg")

        # If the shape contains an image, its ImageData property will be valid,
        # and it will contain an ImageSize object.
        imageSize = shape.image_data.image_size

        # The ImageSize object contains read-only information about the image within the shape.
        self.assertEqual(400, imageSize.height_pixels)
        self.assertEqual(400, imageSize.width_pixels)

        delta = 1
        self.assertAlmostEqual(95.98, imageSize.horizontal_resolution, delta)
        self.assertAlmostEqual(95.98, imageSize.vertical_resolution, delta)

        # We can base the size of the shape on the size of its image to avoid stretching the image.
        shape.width = imageSize.width_points * 2
        shape.height = imageSize.height_points * 2

        doc.save(aeb.artifacts_dir + "Drawing.image_size.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "Drawing.image_size.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        # TestUtil.verify_image_in_shape(400, 400, ImageType.jpeg, shape)
        self.assertEqual(600.0, shape.width)
        self.assertEqual(600.0, shape.height)

        imageSize = shape.image_data.image_size

        self.assertEqual(400, imageSize.height_pixels)
        self.assertEqual(400, imageSize.width_pixels)
        self.assertAlmostEqual(95.98, imageSize.horizontal_resolution, delta)
        self.assertAlmostEqual(95.98, imageSize.vertical_resolution, delta)


if __name__ == '__main__':
    unittest.main()
