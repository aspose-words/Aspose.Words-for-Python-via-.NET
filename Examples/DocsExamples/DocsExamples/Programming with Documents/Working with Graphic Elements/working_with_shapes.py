import unittest
import os
import sys
import uuid
import io

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw
import aspose.pydrawing as drawing

class WorkingWithShapes(docs_base.DocsExamplesBase):

    def test_add_group_shape(self):

        #ExStart:AddGroupShape
        doc = aw.Document()
        doc.ensure_minimum()

        group_shape = aw.drawing.GroupShape(doc)
        accent_border_shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.ACCENT_BORDER_CALLOUT1)
        accent_border_shape.width = 100
        accent_border_shape.height = 100
        group_shape.append_child(accent_border_shape)

        action_button_shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.ACTION_BUTTON_BEGINNING)
        action_button_shape.left = 100
        action_button_shape.width = 100
        action_button_shape.height = 200

        group_shape.append_child(action_button_shape)

        group_shape.width = 200
        group_shape.height = 200
        group_shape.coord_size = drawing.Size(200, 200)

        builder = aw.DocumentBuilder(doc)
        builder.insert_node(group_shape)

        doc.save(docs_base.artifacts_dir + "WorkingWithShapes.add_group_shape.docx")
        #ExEnd:AddGroupShape


    def test_insert_shape(self):

        #ExStart:InsertShape
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, aw.drawing.RelativeHorizontalPosition.PAGE, 100,
            aw.drawing.RelativeVerticalPosition.PAGE, 100, 50, 50, aw.drawing.WrapType.NONE)
        shape.rotation = 30.0

        builder.writeln()

        shape = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 50, 50)
        shape.rotation = 30.0

        save_options = aw.saving.OoxmlSaveOptions(aw.SaveFormat.DOCX)
        save_options.compliance = aw.saving.OoxmlCompliance.ISO29500_2008_TRANSITIONAL


        doc.save(docs_base.artifacts_dir + "WorkingWithShapes.insert_shape.docx", save_options)
        #ExEnd:InsertShape


    def test_aspect_ratio_locked(self):

        #ExStart:AspectRatioLocked
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_image(docs_base.images_dir + "Transparent background logo.png")
        shape.aspect_ratio_locked = False

        doc.save(docs_base.artifacts_dir + "WorkingWithShapes.aspect_ratio_locked.docx")
        #ExEnd:AspectRatioLocked


    def test_layout_in_cell(self):

        #ExStart:LayoutInCell
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_table()
        builder.row_format.height = 100
        builder.row_format.height_rule = aw.HeightRule.EXACTLY

        for i in range(0, 31):

            if (i != 0 and i % 7 == 0):
                builder.end_row()
            builder.insert_cell()
            builder.write("Cell contents")


        builder.end_table()

        watermark = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_PLAIN_TEXT)

        watermark.relative_horizontal_position = aw.drawing.RelativeHorizontalPosition.PAGE
        watermark.relative_vertical_position = aw.drawing.RelativeVerticalPosition.PAGE
        watermark.is_layout_in_cell = True # Display the shape outside of the table cell if it will be placed into a cell.
        watermark.width = 300
        watermark.height = 70
        watermark.horizontal_alignment = aw.drawing.HorizontalAlignment.CENTER
        watermark.vertical_alignment = aw.drawing.VerticalAlignment.CENTER
        watermark.rotation = -40


        watermark.fill_color = drawing.Color.gray
        watermark.stroke_color = drawing.Color.gray

        watermark.text_path.text = "watermarkText"
        watermark.text_path.font_family = "Arial"

        watermark.name = "WaterMark_" + str(uuid.uuid4())
        watermark.wrap_type = aw.drawing.WrapType.NONE

        run = doc.get_child_nodes(aw.NodeType.RUN, True)[doc.get_child_nodes(aw.NodeType.RUN, True).count - 1].as_run()

        builder.move_to(run)
        builder.insert_node(watermark)
        doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2010)

        doc.save(docs_base.artifacts_dir + "WorkingWithShapes.layout_in_cell.docx")
        #ExEnd:LayoutInCell


    def test_add_corners_snipped(self):

        #ExStart:AddCornersSnipped
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_shape(aw.drawing.ShapeType.TOP_CORNERS_SNIPPED, 50, 50)

        save_options = aw.saving.OoxmlSaveOptions(aw.SaveFormat.DOCX)
        save_options.compliance = aw.saving.OoxmlCompliance.ISO29500_2008_TRANSITIONAL

        doc.save(docs_base.artifacts_dir + "WorkingWithShapes.add_corners_snipped.docx", save_options)
        #ExEnd:AddCornersSnipped


    def test_get_actual_shape_bounds_points(self):

        #ExStart:GetActualShapeBoundsPoints
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_image(docs_base.images_dir + "Transparent background logo.png")
        shape.aspect_ratio_locked = False

        print("\nGets the actual bounds of the shape in points: ")
        print(shape.get_shape_renderer().bounds_in_points)
        #ExEnd:GetActualShapeBoundsPoints


    def test_vertical_anchor(self):

        #ExStart:VerticalAnchor
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        text_box = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 200, 200)
        text_box.text_box.vertical_anchor = aw.drawing.TextBoxAnchor.BOTTOM

        builder.move_to(text_box.first_paragraph)
        builder.write("Textbox contents")

        doc.save(docs_base.artifacts_dir + "WorkingWithShapes.vertical_anchor.docx")
        #ExEnd:VerticalAnchor


    def test_detect_smart_art_shape(self):

        #ExStart:DetectSmartArtShape
        doc = aw.Document(docs_base.my_dir + "SmartArt.docx")

        count = 0
        for shape in doc.get_child_nodes(aw.NodeType.SHAPE, True):
            shape = shape.as_shape()
            if(shape.has_smart_art):
                count += 1

        print("The document has 0 shapes with SmartArt.", count)
        #ExEnd:DetectSmartArtShape


    def test_update_smart_art_drawing(self):

        doc = aw.Document(docs_base.my_dir + "SmartArt.docx")

        #ExStart:UpdateSmartArtDrawing
        for shape in doc.get_child_nodes(aw.NodeType.SHAPE, True):
            shape = shape.as_shape()
            if (shape.has_smart_art):
                shape.update_smart_art_drawing()
        #ExEnd:UpdateSmartArtDrawing

    def test_render_shape_to_disk(self):

        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        #ExStart:RenderShapeToDisk
        renderer = shape.get_shape_renderer()

        # Define custom options which control how the image is rendered. Render the shape to the JPEG raster format.
        image_options = aw.saving.ImageSaveOptions(aw.SaveFormat.EMF)
        image_options.scale = 1.5

        # Save the rendered image to disk.
        renderer.save(docs_base.artifacts_dir + "TestFile.RenderToDisk_out.emf", image_options)
        #ExEnd:RenderShapeToDisk

    def test_render_shape_to_stream(self):

        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        #ExStart:RenderShapeToStream
        renderer = shape.get_shape_renderer()

        # Define custom options which control how the image is rendered. Render the shape to the vector format EMF.
        image_options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)

        # Output the image in gray scale
        image_options.image_color_mode = aw.saving.ImageColorMode.GRAYSCALE

        # Reduce the brightness a bit (default is 0.5f).
        image_options.image_brightness = 0.45

        stream =  io.FileIO(docs_base.artifacts_dir + "TestFile.RenderToStream_out.jpg", "w+b")

        # Save the rendered image to the stream using different options.
        renderer.save(stream, image_options)

        # Close the stream
        stream.close()
        #ExEnd:RenderShapeToStream

    def test_render_shape_to_disk(self):

        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        #ExStart:RenderShapeImage
        # Save the rendered image to disk.
        shape.get_shape_renderer().save(docs_base.artifacts_dir + "TestFile.RenderShapeImage.jpeg", None)
        #ExEnd:RenderShapeImage


if __name__ == '__main__':
    unittest.main()
