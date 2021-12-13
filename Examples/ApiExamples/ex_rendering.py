import unittest
import io
from typing import Dict, List, Tuple

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class ExRendering(ApiExampleBase):

    ##ExStart
    ##ExFor:NodeRendererBase.RenderToScale(Graphics, Single, Single, Single)
    ##ExFor:NodeRendererBase.RenderToSize(Graphics, Single, Single, Single, Single)
    ##ExFor:ShapeRenderer
    ##ExFor:ShapeRenderer.#ctor(ShapeBase)
    ##ExSummary:Shows how to render a shape with a Graphics object and display it using a Windows Form.
    #def test_render_shapes_on_form(self):

    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    shape_form = ShapeForm(drawing.Size(1017, 840))

    #    # Below are two ways to use the "ShapeRenderer" class to render a shape to a Graphics object.
    #    # 1 -  Create a shape with a chart, and render it to a specific scale.
    #    chart = builder.insert_chart(aw.drawing.charts.ChartType.PIE, 500, 400).chart
    #    chart.series.clear()
    #    chart.series.add("Desktop Browser Market Share (Oct. 2020)",
    #        ["Google Chrome", "Apple Safari", "Mozilla Firefox", "Microsoft Edge", "Other"],
    #        [70.33, 8.87, 7.69, 5.83, 7.28])

    #    chart_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

    #    shape_form.add_shape_to_render_to_scale(chart_shape, 0, 0, 1.5)

    #    # 2 -  Create a shape group, and render it to a specific size.
    #    group = aw.drawing.GroupShape(doc)
    #    group.bounds = drawing.RectangleF(0, 0, 100, 100)
    #    group.coord_size = drawing.Size(500, 500)

    #    sub_shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.RECTANGLE)
    #    sub_shape.width = 500
    #    sub_shape.height = 500
    #    sub_shape.left = 0
    #    sub_shape.top = 0
    #    sub_shape.fill_color = drawing.Color.royal_blue
    #    group.append_child(sub_shape)

    #    sub_shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.IMAGE)
    #    sub_shape.width = 450
    #    sub_shape.height = 450
    #    sub_shape.left = 25
    #    sub_shape.top = 25
    #    sub_shape.image_data.set_image(IMAGE_DIR + "Logo.jpg")
    #    group.append_child(sub_shape)

    #    builder.insert_node(group)

    #    group_shape = doc.get_child(aw.NodeType.GROUP_SHAPE, 0, True).as_group_shape()
    #    shape_form.add_shape_to_render_to_size(groupShape, 880, 680, 100, 100)

    #    shape_form.show_dialog()

    #class ShapeForm(Form):
    #    """Renders and displays a list of shapes."""

    #    def __init__(self, size: drawing.Size):
    #        #timer = Timer() #ExSKip
    #        #timer.interval = 10000 #ExSKip
    #        #timer.Tick += TimerTick #ExSKip
    #        #timer.start() #ExSKip
    #        self.size = size
    #        self.shapes_to_render: List[Tuple[aw.drawing.ShapeBase, List[float]]] = []

    #    def add_shape_to_render_to_scale(self, shape: aw.drawing.ShapeBase, x: float, y: float, scale: float):
    #        self.shapes_to_render.append(shape, [x, y, scale])

    #    def add_shape_to_render_to_size(self, shape: aw.drawing.ShapeBase, x: float, y: float, width: float, height: float):
    #        self.shapes_to_render.append(shape, [x, y, width, height])

    #    def on_paint(self, e: PaintEventArgs):

    #        for shape, values in self.shapes_to_render:
    #            if len(values) == 3:
    #                self.render_shape_to_scale(shape, *values)
    #            elif len(values) == 4:
    #                self.render_shape_to_size(shape, *values)

    #    def render_shape_to_scale(self, shape: aw.drawing.ShapeBase, x: float, y: float, scale: float):
    #        renderer = aw.rendering.ShapeRenderer(shape)
    #        with create_graphics() as form_graphics:
    #            renderer.render_to_scale(form_graphics, x, y, scale)

    #    def render_shape_to_size(self, shape: aw.drawing.ShapeBase, x: float, y: float, width: float, height: float):

    #        renderer = aw.rendering.ShapeRenderer(shape)
    #        with create_graphics() as form_graphics:
    #            renderer.render_to_size(form_graphics, x, y, width, height)

    ##ExEnd

    def test_render_to_size(self):

        #ExStart
        #ExFor:Document.RenderToSize
        #ExSummary:Shows how to render a document to a bitmap at a specified location and size.
        doc = aw.Document(MY_DIR + "Rendering.docx")

        with drawing.Bitmap(700, 700) as bmp:

            with drawing.Graphics.from_image(bmp) as gr:

                gr.text_rendering_hint = drawing.text.TextRenderingHint.ANTI_ALIAS_GRID_FIT

                # Set the "page_unit" property to "GraphicsUnit.INCH" to use inches as the
                # measurement unit for any transformations and dimensions that we will define.
                gr.page_unit = drawing.GraphicsUnit.INCH

                # Offset the output 0.5" from the edge.
                gr.translate_transform(0.5, 0.5)

                # Rotate the output by 10 degrees.
                gr.rotate_transform(10)

                # Draw a 3"x3" rectangle.
                gr.draw_rectangle(drawing.Pen(drawing.Color.black, 3 / 72), 0, 0, 3, 3)

                # Draw the first page of our document with the same dimensions and transformation as the rectangle.
                # The rectangle will frame the first page.
                returned_scale = doc.render_to_size(0, gr, 0, 0, 3, 3)

                # This is the scaling factor that the "render_to_size" method applied to the first page to fit the specified size.
                self.assertEqual(0.2566, returned_scale, 0.0001)

                # Set the "page_unit" property to "GraphicsUnit.MILLIMETER" to use millimeters as the
                # measurement unit for any transformations and dimensions that we will define.
                gr.page_unit = drawing.GraphicsUnit.MILLIMETER

                # Reset the transformations that we used from the previous rendering.
                gr.reset_transform()

                # Apply another set of transformations.
                gr.translate_transform(10, 10)
                gr.scale_transform(0.5, 0.5)
                gr.page_scale = 2

                # Create another rectangle and use it to frame another page from the document.
                gr.draw_rectangle(drawing.Pen(drawing.Color.black, 1), 90, 10, 50, 100)
                doc.render_to_size(1, gr, 90, 10, 50, 100)

                bmp.save(ARTIFACTS_DIR + "Rendering.render_to_size.png")

        #ExEnd

    def test_thumbnails(self):

        #ExStart
        #ExFor:Document.RenderToScale
        #ExSummary:Shows how to the individual pages of a document to graphics to create one image with thumbnails of all pages.
        doc = aw.Document(MY_DIR + "Rendering.docx")

        # Calculate the number of rows and columns that we will fill with thumbnails.
        thumb_columns = 2
        thumb_rows = doc.page_count // thumb_columns
        remainder = doc.page_count % thumb_columns

        if remainder > 0:
            thumb_rows += 1

        # Scale the thumbnails relative to the size of the first page.
        scale = 0.25
        thumb_size = doc.get_page_info(0).get_size_in_pixels(scale, 96)

        # Calculate the size of the image that will contain all the thumbnails.
        img_width = thumb_size.width * thumb_columns
        img_height = thumb_size.height * thumb_rows

        with drawing.Bitmap(img_width, img_height) as img:

            with drawing.Graphics.from_image(img) as gr:

                gr.text_rendering_hint = drawing.text.TextRenderingHint.ANTI_ALIAS_GRID_FIT

                # Fill the background, which is transparent by default, in white.
                gr.fill_rectangle(drawing.SolidBrush(drawing.Color.white), 0, 0, img_width, img_height)

                for page_index in range(doc.page_count):

                    row_idx = page_index // thumb_columns
                    column_idx = page_index % thumb_columns

                    # Specify where we want the thumbnail to appear.
                    thumb_left = column_idx * thumb_size.width
                    thumb_top = row_idx * thumb_size.height

                    # Render a page as a thumbnail, and then frame it in a rectangle of the same size.
                    size = doc.render_to_scale(page_index, gr, thumb_left, thumb_top, scale)
                    gr.draw_rectangle(drawing.Pens.black, thumb_left, thumb_top, size.width, size.height)

                img.save(ARTIFACTS_DIR + "Rendering.thumbnails.png")

        #ExEnd
