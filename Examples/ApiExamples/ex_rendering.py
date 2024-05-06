# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.pydrawing as drawing
import unittest
from typing import Dict, List, Tuple
import aspose.words as aw
from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExRendering(ApiExampleBase):

    @unittest.skip("drawing.Bitmap type isn't supported yet")
    @unittest.skip("drawing.Graphics type isn't supported yet")
    @unittest.skip("aspose.words.Document.render_to_size method isn't supported yet")
    def test_render_to_size(self):
        #ExStart
        #ExFor:Document.render_to_size
        #ExSummary:Shows how to render a document to a bitmap at a specified location and size.
        doc = aw.Document(MY_DIR + 'Rendering.docx')
        with drawing.Bitmap(700, 700) as bmp:
            with drawing.Graphics.from_image(bmp) as graphics:
                graphics.text_rendering_hint = drawing.text.TextRenderingHint.ANTI_ALIAS_GRID_FIT
                # Set the "page_unit" property to "GraphicsUnit.INCH" to use inches as the
                # measurement unit for any transformations and dimensions that we will define.
                graphics.page_unit = drawing.GraphicsUnit.INCH
                # Offset the output 0.5" from the edge.
                graphics.translate_transform(0.5, 0.5)
                # Rotate the output by 10 degrees.
                graphics.rotate_transform(10)
                # Draw a 3"x3" rectangle.
                graphics.draw_rectangle(drawing.Pen(drawing.Color.black, 3 / 72), 0, 0, 3, 3)
                # Draw the first page of our document with the same dimensions and transformation as the rectangle.
                # The rectangle will frame the first page.
                returned_scale = doc.render_to_size(0, graphics, 0, 0, 3, 3)
                # This is the scaling factor that the "render_to_size" method applied to the first page to fit the specified size.
                self.assertAlmostEqual(0.2566, returned_scale, delta=0.0001)
                # Set the "page_unit" property to "GraphicsUnit.MILLIMETER" to use millimeters as the
                # measurement unit for any transformations and dimensions that we will define.
                graphics.page_unit = drawing.GraphicsUnit.MILLIMETER
                # Reset the transformations that we used from the previous rendering.
                graphics.reset_transform()
                # Apply another set of transformations.
                graphics.translate_transform(10, 10)
                graphics.scale_transform(0.5, 0.5)
                graphics.page_scale = 2
                # Create another rectangle and use it to frame another page from the document.
                graphics.draw_rectangle(drawing.Pen(drawing.Color.black, 1), 90, 10, 50, 100)
                doc.render_to_size(1, graphics, 90, 10, 50, 100)
                bmp.save(ARTIFACTS_DIR + 'Rendering.render_to_size.png')

    @unittest.skip("drawing.Bitmap type isn't supported yet")
    @unittest.skip("drawing.Graphics type isn't supported yet")
    @unittest.skip("aspose.words.Document.render_to_size method isn't supported yet")
    def test_thumbnails(self):
        #ExStart
        #ExFor:Document.render_to_scale
        #ExSummary:Shows how to the individual pages of a document to graphics to create one image with thumbnails of all pages.
        doc = aw.Document(MY_DIR + 'Rendering.docx')
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
            with drawing.Graphics.from_image(img) as graphics:
                graphics.text_rendering_hint = drawing.text.TextRenderingHint.ANTI_ALIAS_GRID_FIT
                # Fill the background, which is transparent by default, in white.
                graphics.fill_rectangle(drawing.SolidBrush(drawing.Color.white), 0, 0, img_width, img_height)
                for page_index in range(doc.page_count):
                    row_idx = page_index // thumb_columns
                    column_idx = page_index % thumb_columns
                    # Specify where we want the thumbnail to appear.
                    thumb_left = column_idx * thumb_size.width
                    thumb_top = row_idx * thumb_size.height
                    # Render a page as a thumbnail, and then frame it in a rectangle of the same size.
                    size = doc.render_to_scale(page_index, graphics, thumb_left, thumb_top, scale)
                    graphics.draw_rectangle(drawing.Pens.black, thumb_left, thumb_top, size.width, size.height)
                img.save(ARTIFACTS_DIR + 'Rendering.thumbnails.png')