# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import os

import aspose.words as aw
import aspose.words.drawing as awd
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, GOLDS_DIR, IMAGE_DIR
from document_helper import DocumentHelper

class ExShape(ApiExampleBase):
    """Examples using shapes in documents."""

    def test_alt_text(self):

        #ExStart
        #ExFor:ShapeBase.alternative_text
        #ExFor:ShapeBase.name
        #ExSummary:Shows how to use a shape's alternative text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        shape = builder.insert_shape(aw.drawing.ShapeType.CUBE, 150, 150)
        shape.name = "MyCube"

        shape.alternative_text = "Alt text for MyCube."

        # We can access the alternative text of a shape by right-clicking it, and then via "Format AutoShape" -> "Alt Text".
        doc.save(ARTIFACTS_DIR + "Shape.alt_text.docx")

        # Save the document to HTML, and then delete the linked image that belongs to our shape.
        # The browser that is reading our HTML will display the alt text in place of the missing image.
        doc.save(ARTIFACTS_DIR + "Shape.alt_text.html")
        self.assertTrue(os.path.exists(ARTIFACTS_DIR + "Shape.alt_text.001.png")) #ExSkip
        os.remove(ARTIFACTS_DIR + "Shape.alt_text.001.png")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.alt_text.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.CUBE, "MyCube", 150.0, 150.0, 0, 0, shape)
        self.assertEqual("Alt text for MyCube.", shape.alternative_text)
        self.assertEqual("Times New Roman", shape.font.name)

        doc = aw.Document(ARTIFACTS_DIR + "Shape.alt_text.html")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.IMAGE, "", 151.5, 151.5, 0, 0, shape)
        self.assertEqual("Alt text for MyCube.", shape.alternative_text)

        with open(ARTIFACTS_DIR + "Shape.alt_text.html", 'rb') as file:
            self.assertIn(
                "<img src=\"Shape.alt_text.001.png\" width=\"202\" height=\"202\" alt=\"Alt text for MyCube.\" " +
                "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />",
                file.read().decode('utf-8'))

    def test_font(self):

        for hide_shape in (False, True):
            with self.subTest(hide_shape=hide_shape):
                #ExStart
                #ExFor:ShapeBase.font
                #ExFor:ShapeBase.parent_paragraph
                #ExSummary:Shows how to insert a text box, and set the font of its contents.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.writeln("Hello world!")

                shape = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 300, 50)
                builder.move_to(shape.last_paragraph)
                builder.write("This text is inside the text box.")

                # Set the "hidden" property of the shape's "font" object to "True" to hide the text box from sight
                # and collapse the space that it would normally occupy.
                # Set the "hidden" property of the shape's "font" object to "False" to leave the text box visible.
                shape.font.hidden = hide_shape

                # If the shape is visible, we will modify its appearance via the font object.
                if not hide_shape:
                    shape.font.highlight_color = drawing.Color.light_gray
                    shape.font.color = drawing.Color.red
                    shape.font.underline = aw.Underline.DASH

                # Move the builder out of the text box back into the main document.
                builder.move_to(shape.parent_paragraph)

                builder.writeln("\nThis text is outside the text box.")

                doc.save(ARTIFACTS_DIR + "Shape.font.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "Shape.font.docx")
                shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

                self.assertEqual(hide_shape, shape.font.hidden)

                if hide_shape:
                    self.assertEqual(drawing.Color.empty().to_argb(), shape.font.highlight_color.to_argb())
                    self.assertEqual(drawing.Color.empty().to_argb(), shape.font.color.to_argb())
                    self.assertEqual(aw.Underline.NONE, shape.font.underline)
                else:
                    self.assertEqual(drawing.Color.silver.to_argb(), shape.font.highlight_color.to_argb())
                    self.assertEqual(drawing.Color.red.to_argb(), shape.font.color.to_argb())
                    self.assertEqual(aw.Underline.DASH, shape.font.underline)

                self.verify_shape(aw.drawing.ShapeType.TEXT_BOX, "TextBox 100002", 300.0, 50.0, 0, 0, shape)
                self.assertEqual("This text is inside the text box.", shape.get_text().strip())
                self.assertEqual("Hello world!\rThis text is inside the text box.\r\rThis text is outside the text box.", doc.get_text().strip())

    def test_rotate(self):

        #ExStart
        #ExFor:ShapeBase.can_have_image
        #ExFor:ShapeBase.rotation
        #ExSummary:Shows how to insert and rotate an image.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a shape with an image.
        shape = builder.insert_image(drawing.Image.from_file(IMAGE_DIR + "Logo.jpg"))
        self.assertTrue(shape.can_have_image)
        self.assertTrue(shape.has_image)

        # Rotate the image 45 degrees clockwise.
        shape.rotation = 45

        doc.save(ARTIFACTS_DIR + "Shape.rotate.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.rotate.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.IMAGE, "", 300.0, 300.0, 0, 0, shape)
        self.assertTrue(shape.can_have_image)
        self.assertTrue(shape.has_image)
        self.assertEqual(45.0, shape.rotation)

    def test_aspect_ratio_locked_default_value(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        image = drawing.Image.from_file(IMAGE_DIR + "Transparent background logo.png")

        shape = builder.insert_image(image)
        shape.wrap_type = aw.drawing.wrap_type.NONE
        shape.behind_text = True

        shape.relative_horizontal_position = aw.drawing.RelativeHorizontalPosition.PAGE
        shape.relative_vertical_position = aw.drawing.RelativeVerticalPosition.PAGE

        # Calculate image left and top position so it appears in the center of the page.
        shape.left = (builder.page_setup.page_width - shape.width) / 2
        shape.top = (builder.page_setup.page_height - shape.height) / 2

        doc = DocumentHelper.save_open(doc)

        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual(True, shape.aspect_ratio_locked)

    def test_coordinates(self):

        #ExStart
        #ExFor:ShapeBase.distance_bottom
        #ExFor:ShapeBase.distance_left
        #ExFor:ShapeBase.distance_right
        #ExFor:ShapeBase.distance_top
        #ExSummary:Shows how to set the wrapping distance for a text that surrounds a shape.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a rectangle and, get the text to wrap tightly around its bounds.
        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, 150, 150)
        shape.wrap_type = aw.drawing.WrapType.TIGHT

        # Set the minimum distance between the shape and surrounding text to 40pt from all sides.
        shape.distance_top = 40
        shape.distance_bottom = 40
        shape.distance_left = 40
        shape.distance_right = 40

        # Move the shape closer to the center of the page, and then rotate the shape 60 degrees clockwise.
        shape.top = 75
        shape.left = 150
        shape.rotation = 60

        # Add text that will wrap around the shape.
        builder.font.size = 24
        builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                      "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.")

        doc.save(ARTIFACTS_DIR + "Shape.coordinates.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.coordinates.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.RECTANGLE, "Rectangle 100002", 150.0, 150.0, 75.0, 150.0, shape)
        self.assertEqual(40.0, shape.distance_bottom)
        self.assertEqual(40.0, shape.distance_left)
        self.assertEqual(40.0, shape.distance_right)
        self.assertEqual(40.0, shape.distance_top)
        self.assertEqual(60.0, shape.rotation)

    def test_group_shape(self):

        #ExStart
        #ExFor:ShapeBase.bounds
        #ExFor:ShapeBase.coord_origin
        #ExFor:ShapeBase.coord_size
        #ExSummary:Shows how to create and populate a group shape.
        doc = aw.Document()

        # Create a group shape. A group shape can display a collection of child shape nodes.
        # In Microsoft Word, clicking within the group shape's boundary or on one of the group shape's child shapes will
        # select all the other child shapes within this group and allow us to scale and move all the shapes at once.
        group = aw.drawing.GroupShape(doc)

        self.assertEqual(aw.drawing.WrapType.NONE, group.wrap_type)

        # Create a 400pt x 400pt group shape and place it at the document's floating shape coordinate origin.
        group.bounds = drawing.RectangleF(0, 0, 400, 400)

        # Set the group's internal coordinate plane size to 500 x 500pt.
        # The top left corner of the group will have an x and y coordinate of (0, 0),
        # and the bottom right corner will have an x and y coordinate of (500, 500).
        group.coord_size = drawing.Size(500, 500)

        # Set the coordinates of the top left corner of the group to (-250, -250).
        # The group's center will now have an x and y coordinate value of (0, 0),
        # and the bottom right corner will be at (250, 250).
        group.coord_origin = drawing.Point(-250, -250)

        # Create a rectangle that will display the boundary of this group shape and add it to the group.
        rectangle = aw.drawing.Shape(doc, aw.drawing.ShapeType.RECTANGLE)
        rectangle.width = group.coord_size.width
        rectangle.height = group.coord_size.height
        rectangle.left = group.coord_origin.x
        rectangle.top = group.coord_origin.y
        group.append_child(rectangle)

        # Once a shape is a part of a group shape, we can access it as a child node and then modify it.
        group.get_child(aw.NodeType.SHAPE, 0, True).as_shape().stroke.dash_style = aw.drawing.DashStyle.DASH

        # Create a small red star and insert it into the group.
        # Line up the shape with the group's coordinate origin, which we have moved to the center.
        red_star = aw.drawing.Shape(doc, aw.drawing.ShapeType.STAR)
        red_star.width = 20
        red_star.height = 20
        red_star.left = -10
        red_star.top = -10
        red_star.fill_color = drawing.Color.red
        group.append_child(red_star)

        # Insert a rectangle, and then insert a slightly smaller rectangle in the same place with an image.
        # Newer shapes that we add to the group overlap older shapes. The light blue rectangle will partially overlap the red star,
        # and then the shape with the image will overlap the light blue rectangle, using it as a frame.
        # We cannot use the "z_order" properties of shapes to manipulate their arrangement within a group shape.
        blue_rectangle = aw.drawing.Shape(doc, aw.drawing.ShapeType.RECTANGLE)
        blue_rectangle.width = 250
        blue_rectangle.height = 250
        blue_rectangle.left = -250
        blue_rectangle.top = -250
        blue_rectangle.fill_color = drawing.Color.light_blue
        group.append_child(blue_rectangle)

        image = aw.drawing.Shape(doc, aw.drawing.ShapeType.IMAGE)
        image.width = 200
        image.height = 200
        image.left = -225
        image.top = -225
        group.append_child(image)

        group.get_child(aw.NodeType.SHAPE, 3, True).as_shape().image_data.set_image(IMAGE_DIR + "Logo.jpg")

        # Insert a text box into the group shape. Set the "left" property so that the text box's right edge
        # touches the right boundary of the group shape. Set the "top" property so that the text box sits outside
        # the boundary of the group shape, with its top size lined up along the group shape's bottom margin.
        text_box = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_BOX)
        text_box.width = 200
        text_box.height = 50
        text_box.left = group.coord_size.width + group.coord_origin.x - 200
        text_box.top = group.coord_size.height + group.coord_origin.y
        group.append_child(text_box)

        builder = aw.DocumentBuilder(doc)
        builder.insert_node(group)
        builder.move_to(group.get_child(aw.NodeType.SHAPE, 4, True).as_shape().append_child(aw.Paragraph(doc)))
        builder.write("Hello world!")

        doc.save(ARTIFACTS_DIR + "Shape.group_shape.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.group_shape.docx")
        group = doc.get_child(aw.NodeType.GROUP_SHAPE, 0, True).as_group_shape()

        self.assertEqual(drawing.RectangleF(0, 0, 400, 400), group.bounds)
        self.assertEqual(drawing.Size(500, 500), group.coord_size)
        self.assertEqual(drawing.Point(-250, -250), group.coord_origin)

        self.verify_shape(aw.drawing.ShapeType.RECTANGLE, "", 500.0, 500.0, -250.0, -250.0, group.get_child(aw.NodeType.SHAPE, 0, True).as_shape())
        self.verify_shape(aw.drawing.ShapeType.STAR, "", 20.0, 20.0, -10.0, -10.0, group.get_child(aw.NodeType.SHAPE, 1, True).as_shape())
        self.verify_shape(aw.drawing.ShapeType.RECTANGLE, "", 250.0, 250.0, -250.0, -250.0, group.get_child(aw.NodeType.SHAPE, 2, True).as_shape())
        self.verify_shape(aw.drawing.ShapeType.IMAGE, "", 200.0, 200.0, -225.0, -225.0, group.get_child(aw.NodeType.SHAPE, 3, True).as_shape())
        self.verify_shape(aw.drawing.ShapeType.TEXT_BOX, "", 200.0, 50.0, 250.0, 50.0, group.get_child(aw.NodeType.SHAPE, 4, True).as_shape())

    def test_is_top_level(self):

        #ExStart
        #ExFor:ShapeBase.is_top_level
        #ExSummary:Shows how to tell whether a shape is a part of a group shape.
        doc = aw.Document()

        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.RECTANGLE)
        shape.width = 200
        shape.height = 200
        shape.wrap_type = aw.drawing.WrapType.NONE

        # A shape by default is not part of any group shape, and therefore has the "is_top_level" property set to "True".
        self.assertTrue(shape.is_top_level)

        group = aw.drawing.GroupShape(doc)
        group.append_child(shape)

        # Once we assimilate a shape into a group shape, the "is_top_level" property changes to "False".
        self.assertFalse(shape.is_top_level)
        #ExEnd

    def test_local_to_parent(self):

        #ExStart
        #ExFor:ShapeBase.coord_origin
        #ExFor:ShapeBase.coord_size
        #ExFor:ShapeBase.local_to_parent(PointF)
        #ExSummary:Shows how to translate the x and y coordinate location on a shape's coordinate plane to a location on the parent shape's coordinate plane.
        doc = aw.Document()

        # Insert a group shape, and place it 100 points below and to the right of
        # the document's x and Y coordinate origin point.
        group = aw.drawing.GroupShape(doc)
        group.bounds = drawing.RectangleF(100, 100, 500, 500)

        # Use the "local_to_parent" method to determine that (0, 0) on the group's internal x and y coordinates
        # lies on (100, 100) of its parent shape's coordinate system. The group shape's parent is the document itself.
        self.assertEqual(drawing.PointF(100, 100), group.local_to_parent(drawing.PointF(0, 0)))

        # By default, a shape's internal coordinate plane has the top left corner at (0, 0),
        # and the bottom right corner at (1000, 1000). Due to its size, our group shape covers an area of 500pt x 500pt
        # in the document's plane. This means that a movement of 1pt on the document's coordinate plane will translate
        # to a movement of 2pts on the group shape's coordinate plane.
        self.assertEqual(drawing.PointF(150, 150), group.local_to_parent(drawing.PointF(100, 100)))
        self.assertEqual(drawing.PointF(200, 200), group.local_to_parent(drawing.PointF(200, 200)))
        self.assertEqual(drawing.PointF(250, 250), group.local_to_parent(drawing.PointF(300, 300)))

        # Move the group shape's x and y axis origin from the top left corner to the center.
        # This will offset the group's internal coordinates relative to the document's coordinates even further.
        group.coord_origin = drawing.Point(-250, -250)

        self.assertEqual(drawing.PointF(375, 375), group.local_to_parent(drawing.PointF(300, 300)))

        # Changing the scale of the coordinate plane will also affect relative locations.
        group.coord_size = drawing.Size(500, 500)

        self.assertEqual(drawing.PointF(650, 650), group.local_to_parent(drawing.PointF(300, 300)))

        # If we wish to add a shape to this group while defining its location based on a location in the document,
        # we will need to first confirm a location in the group shape that will match the document's location.
        self.assertEqual(drawing.PointF(700, 700), group.local_to_parent(drawing.PointF(350, 350)))

        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.RECTANGLE)
        shape.width = 100
        shape.height = 100
        shape.left = 700
        shape.top = 700

        group.append_child(shape)
        doc.first_section.body.first_paragraph.append_child(group)

        doc.save(ARTIFACTS_DIR + "Shape.local_to_parent.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.local_to_parent.docx")
        group = doc.get_child(aw.NodeType.GROUP_SHAPE, 0, True).as_group_shape()

        self.assertEqual(drawing.RectangleF(100, 100, 500, 500), group.bounds)
        self.assertEqual(drawing.Size(500, 500), group.coord_size)
        self.assertEqual(drawing.Point(-250, -250), group.coord_origin)

    def test_anchor_locked(self):

        for anchor_locked in (False, True):
            with self.subTest(anchor_locked=anchor_locked):
                #ExStart
                #ExFor:ShapeBase.anchor_locked
                #ExSummary:Shows how to lock or unlock a shape's paragraph anchor.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.writeln("Hello world!")

                builder.write("Our shape will have an anchor attached to this paragraph.")
                shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, 200, 160)
                shape.wrap_type = aw.drawing.WrapType.NONE
                builder.insert_break(aw.BreakType.PARAGRAPH_BREAK)

                builder.writeln("Hello again!")

                # Set the "anchor_locked" property to "True" to prevent the shape's anchor
                # from moving when moving the shape in Microsoft Word.
                # Set the "anchor_locked" property to "False" to allow any movement of the shape
                # to also move its anchor to any other paragraph that the shape ends up close to.
                shape.anchor_locked = anchor_locked

                # If the shape does not have a visible anchor symbol to its left,
                # we will need to enable visible anchors via "Options" -> "Display" -> "Object Anchors".
                doc.save(ARTIFACTS_DIR + "Shape.anchor_locked.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "Shape.anchor_locked.docx")
                shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

                self.assertEqual(anchor_locked, shape.anchor_locked)

    def test_delete_all_shapes(self):

        #ExStart
        #ExFor:Shape
        #ExSummary:Shows how to delete all shapes from a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert two shapes along with a group shape with another shape inside it.
        builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, 400, 200)
        builder.insert_shape(aw.drawing.ShapeType.STAR, 300, 300)

        group = aw.drawing.GroupShape(doc)
        group.bounds = drawing.RectangleF(100, 50, 200, 100)
        group.coord_origin = drawing.Point(-1000, -500)

        sub_shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.CUBE)
        sub_shape.width = 500
        sub_shape.height = 700
        sub_shape.left = 0
        sub_shape.top = 0

        group.append_child(sub_shape)
        builder.insert_node(group)

        self.assertEqual(3, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)
        self.assertEqual(1, doc.get_child_nodes(aw.NodeType.GROUP_SHAPE, True).count)

        # Remove all Shape nodes from the document.
        shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)
        shapes.clear()

        # All shapes are gone, but the group shape is still in the document.
        self.assertEqual(1, doc.get_child_nodes(aw.NodeType.GROUP_SHAPE, True).count)
        self.assertEqual(0, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)

        # Remove all group shapes separately.
        group_shapes = doc.get_child_nodes(aw.NodeType.GROUP_SHAPE, True)
        group_shapes.clear()

        self.assertEqual(0, doc.get_child_nodes(aw.NodeType.GROUP_SHAPE, True).count)
        self.assertEqual(0, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)
        #ExEnd

    def test_is_inline(self):

        #ExStart
        #ExFor:ShapeBase.is_inline
        #ExSummary:Shows how to determine whether a shape is inline or floating.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Below are two wrapping types that shapes may have.
        # 1 -  Inline:
        builder.write("Hello world! ")
        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, 100, 100)
        shape.fill_color = drawing.Color.light_blue
        builder.write(" Hello again.")

        # An inline shape sits inside a paragraph among other paragraph elements, such as runs of text.
        # In Microsoft Word, we may click and drag the shape to any paragraph as if it is a character.
        # If the shape is large, it will affect vertical paragraph spacing.
        # We cannot move this shape to a place with no paragraph.
        self.assertEqual(aw.drawing.WrapType.INLINE, shape.wrap_type)
        self.assertTrue(shape.is_inline)

        # 2 -  Floating:
        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN ,200,
            aw.drawing.RelativeVerticalPosition.TOP_MARGIN ,200, 100, 100, aw.drawing.WrapType.NONE)
        shape.fill_color = drawing.Color.orange

        # A floating shape belongs to the paragraph that we insert it into,
        # which we can determine by an anchor symbol that appears when we click the shape.
        # If the shape does not have a visible anchor symbol to its left,
        # we will need to enable visible anchors via "Options" -> "Display" -> "Object Anchors".
        # In Microsoft Word, we may left click and drag this shape freely to any location.
        self.assertEqual(aw.drawing.WrapType.NONE, shape.wrap_type)
        self.assertFalse(shape.is_inline)

        doc.save(ARTIFACTS_DIR + "Shape.is_inline.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.is_inline.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.RECTANGLE, "Rectangle 100002", 100, 100, 0, 0, shape)
        self.assertEqual(drawing.Color.light_blue.to_argb(), shape.fill_color.to_argb())
        self.assertEqual(aw.drawing.WrapType.INLINE, shape.wrap_type)
        self.assertTrue(shape.is_inline)

        shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.RECTANGLE, "Rectangle 100004", 100, 100, 200, 200, shape)
        self.assertEqual(drawing.Color.orange.to_argb(), shape.fill_color.to_argb())
        self.assertEqual(aw.drawing.WrapType.NONE, shape.wrap_type)
        self.assertFalse(shape.is_inline)

    def test_bounds(self):

        #ExStart
        #ExFor:ShapeBase.bounds
        #ExFor:ShapeBase.bounds_in_points
        #ExSummary:Shows how to verify shape containing block boundaries.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_shape(aw.drawing.ShapeType.LINE, aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN, 50,
            aw.drawing.RelativeVerticalPosition.TOP_MARGIN, 50, 100, 100, aw.drawing.WrapType.NONE)
        shape.stroke_color = drawing.Color.orange

        # Even though the line itself takes up little space on the document page,
        # it occupies a rectangular containing block, the size of which we can determine using the "bounds" properties.
        self.assertEqual(drawing.RectangleF(50, 50, 100, 100), shape.bounds)
        self.assertEqual(drawing.RectangleF(50, 50, 100, 100), shape.bounds_in_points)

        # Create a group shape, and then set the size of its containing block using the "Bounds" property.
        group = aw.drawing.GroupShape(doc)
        group.bounds = drawing.RectangleF(0, 100, 250, 250)

        self.assertEqual(drawing.RectangleF(0, 100, 250, 250), group.bounds_in_points)

        # Create a rectangle, verify the size of its bounding block, and then add it to the group shape.
        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.RECTANGLE)
        shape.width = 100
        shape.height = 100
        shape.left = 700
        shape.top = 700

        self.assertEqual(drawing.RectangleF(700, 700, 100, 100), shape.bounds_in_points)

        group.append_child(shape)

        # The group shape's coordinate plane has its origin on the top left-hand side corner of its containing block,
        # and the x and y coordinates of (1000, 1000) on the bottom right-hand side corner.
        # Our group shape is 250x250pt in size, so every 4pt on the group shape's coordinate plane
        # translates to 1pt in the document body's coordinate plane.
        # Every shape that we insert will also shrink in size by a factor of 4.
        # The change in the shape's "bounds_in_points" property will reflect this.
        self.assertEqual(drawing.RectangleF(175, 275, 25, 25), shape.bounds_in_points)

        doc.first_section.body.first_paragraph.append_child(group)

        # Insert a shape and place it outside of the bounds of the group shape's containing block.
        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.RECTANGLE)
        shape.width = 100
        shape.height = 100
        shape.left = 1000
        shape.top = 1000

        group.append_child(shape)

        # The group shape's footprint in the document body has increased, but the containing block remains the same.
        self.assertEqual(drawing.RectangleF(0, 100, 250, 250), group.bounds_in_points)
        self.assertEqual(drawing.RectangleF(250, 350, 25, 25), shape.bounds_in_points)

        doc.save(ARTIFACTS_DIR + "Shape.bounds.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.bounds.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.LINE, "Line 100002", 100, 100, 50, 50, shape)
        self.assertEqual(drawing.Color.orange.to_argb(), shape.stroke_color.to_argb())
        self.assertEqual(drawing.RectangleF(50, 50, 100, 100), shape.bounds_in_points)

        group = doc.get_child(aw.NodeType.GROUP_SHAPE, 0, True).as_group_shape()

        self.assertEqual(drawing.RectangleF(0, 100, 250, 250), group.bounds)
        self.assertEqual(drawing.RectangleF(0, 100, 250, 250), group.bounds_in_points)
        self.assertEqual(drawing.Size(1000, 1000), group.coord_size)
        self.assertEqual(drawing.Point(0, 0), group.coord_origin)

        shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.RECTANGLE, "", 100, 100, 700, 700, shape)
        self.assertEqual(drawing.RectangleF(175, 275, 25, 25), shape.bounds_in_points)

        shape = doc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.RECTANGLE, "", 100, 100, 1000, 1000, shape)
        self.assertEqual(drawing.RectangleF(250, 350, 25, 25), shape.bounds_in_points)

    def test_flip_shape_orientation(self):

        #ExStart
        #ExFor:ShapeBase.flip_orientation
        #ExFor:FlipOrientation
        #ExSummary:Shows how to flip a shape on an axis.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert an image shape and leave its orientation in its default state.
        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN, 100,
            aw.drawing.RelativeVerticalPosition.TOP_MARGIN, 100, 100, 100, aw.drawing.WrapType.NONE)
        shape.image_data.set_image(IMAGE_DIR + "Logo.jpg")

        self.assertEqual(aw.drawing.FlipOrientation.NONE, shape.flip_orientation)

        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN, 250,
            aw.drawing.RelativeVerticalPosition.TOP_MARGIN, 100, 100, 100, aw.drawing.WrapType.NONE)
        shape.image_data.set_image(IMAGE_DIR + "Logo.jpg")

        # Set the "flip_orientation" property to "FlipOrientation.HORIZONTAL" to flip the second shape on the y-axis,
        # making it into a horizontal mirror image of the first shape.
        shape.flip_orientation = aw.drawing.FlipOrientation.HORIZONTAL

        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN, 100,
            aw.drawing.RelativeVerticalPosition.TOP_MARGIN, 250, 100, 100, aw.drawing.WrapType.NONE)
        shape.image_data.set_image(IMAGE_DIR + "Logo.jpg")

        # Set the "flip_orientation" property to "FlipOrientation.VERTICAL" to flip the third shape on the x-axis,
        # making it into a vertical mirror image of the first shape.
        shape.flip_orientation = aw.drawing.FlipOrientation.VERTICAL

        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN, 250,
            aw.drawing.RelativeVerticalPosition.TOP_MARGIN, 250, 100, 100, aw.drawing.WrapType.NONE)
        shape.image_data.set_image(IMAGE_DIR + "Logo.jpg")

        # Set the "flip_orientation" property to "FlipOrientation.BOTH" to flip the fourth shape on both the x and y axes,
        # making it into a horizontal and vertical mirror image of the first shape.
        shape.flip_orientation = aw.drawing.FlipOrientation.BOTH

        doc.save(ARTIFACTS_DIR + "Shape.flip_shape_orientation.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.flip_shape_orientation.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.RECTANGLE, "Rectangle 100002", 100, 100, 100, 100, shape)
        self.assertEqual(aw.drawing.FlipOrientation.NONE, shape.flip_orientation)

        shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.RECTANGLE, "Rectangle 100004", 100, 100, 100, 250, shape)
        self.assertEqual(aw.drawing.FlipOrientation.HORIZONTAL, shape.flip_orientation)

        shape = doc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.RECTANGLE, "Rectangle 100006", 100, 100, 250, 100, shape)
        self.assertEqual(aw.drawing.FlipOrientation.VERTICAL, shape.flip_orientation)

        shape = doc.get_child(aw.NodeType.SHAPE, 3, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.RECTANGLE, "Rectangle 100008", 100, 100, 250, 250, shape)
        self.assertEqual(aw.drawing.FlipOrientation.BOTH, shape.flip_orientation)

    def test_fill(self):

        #ExStart
        #ExFor:ShapeBase.fill
        #ExFor:Shape.fill_color
        #ExFor:Shape.stroke_color
        #ExFor:Fill
        #ExFor:Fill.opacity
        #ExSummary:Shows how to fill a shape with a solid color.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Write some text, and then cover it with a floating shape.
        builder.font.size = 32
        builder.writeln("Hello world!")

        shape = builder.insert_shape(aw.drawing.ShapeType.CLOUD_CALLOUT, aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN, 25,
            aw.drawing.RelativeVerticalPosition.TOP_MARGIN, 25, 250, 150, aw.drawing.WrapType.NONE)

        # Use the "stroke_color" property to set the color of the outline of the shape.
        shape.stroke_color = drawing.Color.cadet_blue

        # Use the "fill_color" property to set the color of the inside area of the shape.
        shape.fill_color = drawing.Color.light_blue

        # The "opacity" property determines how transparent the color is on a 0-1 scale,
        # with 1 being fully opaque, and 0 being invisible.
        # The shape fill by default is fully opaque, so we cannot see the text that this shape is on top of.
        self.assertEqual(1.0, shape.fill.opacity)

        # Set the shape fill color's opacity to a lower value so that we can see the text underneath it.
        shape.fill.opacity = 0.3

        doc.save(ARTIFACTS_DIR + "Shape.fill.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.fill.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.CLOUD_CALLOUT, "CloudCallout 100002", 250.0, 150.0, 25.0, 25.0, shape)
        colorWithOpacity = drawing.Color.from_argb(255 * shape.fill.opacity,
            drawing.Color.light_blue.r,
            drawing.Color.light_blue.g,
            drawing.Color.light_blue.b)
        self.assertEqual(colorWithOpacity.to_argb(), shape.fill_color.to_argb())
        self.assertEqual(drawing.Color.cadet_blue.to_argb(), shape.stroke_color.to_argb())
        self.assertAlmostEqual(0.3, shape.fill.opacity, delta=0.01)

    def test_texture_fill(self):

        #ExStart
        #ExFor:Fill.texture_alignment
        #ExFor:TextureAlignment
        #ExSummary:Shows how to fill and tiling the texture inside the shape.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, 80, 80)

        # Apply texture alignment to the shape fill.
        shape.fill.preset_textured(aw.drawing.PresetTexture.CANVAS)
        shape.fill.texture_alignment = aw.drawing.TextureAlignment.TOP_RIGHT

        # Use the compliance option to define the shape using DML if you want to get "texture_alignment"
        # property after the document saves.
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.compliance = aw.saving.OoxmlCompliance.ISO29500_2008_STRICT

        doc.save(ARTIFACTS_DIR + "Shape.texture_fill.docx", save_options)
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.texture_fill.docx")

        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.assertEqual(aw.drawing.TextureAlignment.TOP_RIGHT, shape.fill.texture_alignment)

    def test_gradient_fill(self):

        #ExStart
        #ExFor:Fill.one_color_gradient(Color, GradientStyle, GradientVariant, float)
        #ExFor:Fill.one_color_gradient(GradientStyle, GradientVariant, float)
        #ExFor:Fill.two_color_gradient(Color, Color, GradientStyle, GradientVariant)
        #ExFor:Fill.two_color_gradient(GradientStyle, GradientVariant)
        #ExFor:Fill.gradient_style
        #ExFor:Fill.gradient_variant
        #ExFor:Fill.gradient_angle
        #ExFor:GradientStyle
        #ExFor:GradientVariant
        #ExSummary:Shows how to fill a shape with a gradients.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, 80, 80)
        # Apply One-color gradient fill to the shape with "fore_color" of gradient fill.
        shape.fill.one_color_gradient(drawing.Color.red, aw.drawing.GradientStyle.HORIZONTAL, aw.drawing.GradientVariant.VARIANT2, 0.1)

        self.assertEqual(drawing.Color.red.to_argb(), shape.fill.fore_color.to_argb())
        self.assertEqual(aw.drawing.GradientStyle.HORIZONTAL, shape.fill.gradient_style)
        self.assertEqual(aw.drawing.GradientVariant.VARIANT2, shape.fill.gradient_variant)
        self.assertEqual(270, shape.fill.gradient_angle)

        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, 80, 80)
        # Apply Two-color gradient fill to the shape.
        shape.fill.two_color_gradient(aw.drawing.GradientStyle.FROM_CORNER, aw.drawing.GradientVariant.VARIANT4)
        # Change "back_color" of gradient fill.
        shape.fill.back_color = drawing.Color.yellow
        # Note that changes "gradient_angle" for "GradientStyle.FROM_CORNER/GradientStyle.FROM_CENTER"
        # gradient fill don't get any effect, it will work only for linear gradient.
        shape.fill.gradient_angle = 15

        self.assertEqual(drawing.Color.yellow.to_argb(), shape.fill.back_color.to_argb())
        self.assertEqual(aw.drawing.GradientStyle.FROM_CORNER, shape.fill.gradient_style)
        self.assertEqual(aw.drawing.GradientVariant.VARIANT4, shape.fill.gradient_variant)
        self.assertEqual(0, shape.fill.gradient_angle)

        # Use the compliance option to define the shape using DML if you want to get "gradient_style",
        # "gradient_variant" and "gradient_angle" properties after the document saves.
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.compliance = aw.saving.OoxmlCompliance.ISO29500_2008_STRICT

        doc.save(ARTIFACTS_DIR + "Shape.gradient_fill.docx", save_options)
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.gradient_fill.docx")
        first_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.assertEqual(drawing.Color.red.to_argb(), first_shape.fill.fore_color.to_argb())
        self.assertEqual(aw.drawing.GradientStyle.HORIZONTAL, first_shape.fill.gradient_style)
        self.assertEqual(aw.drawing.GradientVariant.VARIANT2, first_shape.fill.gradient_variant)
        self.assertEqual(270, first_shape.fill.gradient_angle)

        second_shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()

        self.assertEqual(drawing.Color.yellow.to_argb(), second_shape.fill.back_color.to_argb())
        self.assertEqual(aw.drawing.GradientStyle.FROM_CORNER, second_shape.fill.gradient_style)
        self.assertEqual(aw.drawing.GradientVariant.VARIANT4, second_shape.fill.gradient_variant)
        self.assertEqual(0, second_shape.fill.gradient_angle)

    def test_gradient_stops(self):

        #ExStart
        #ExFor:Fill.gradient_stops
        #ExFor:GradientStopCollection
        #ExFor:GradientStopCollection.insert(int, GradientStop)
        #ExFor:GradientStopCollection.add(GradientStop)
        #ExFor:GradientStopCollection.remove_at(int)
        #ExFor:GradientStopCollection.remove(GradientStop)
        #ExFor:GradientStopCollection.__getitem__(int)
        #ExFor:GradientStopCollection.count
        #ExFor:GradientStop.__init__(Color, float)
        #ExFor:GradientStop.__init__(Color, float, float)
        #ExFor:GradientStop.color
        #ExFor:GradientStop.position
        #ExFor:GradientStop.transparency
        #ExFor:GradientStop.remove
        #ExSummary:Shows how to add gradient stops to the gradient fill.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, 80, 80)
        shape.fill.two_color_gradient(drawing.Color.green, drawing.Color.red, aw.drawing.GradientStyle.HORIZONTAL, aw.drawing.GradientVariant.VARIANT2)

        # Get gradient stops collection.
        gradient_stops = shape.fill.gradient_stops

        # Change first gradient stop.
        gradient_stops[0].color = drawing.Color.aqua
        gradient_stops[0].position = 0.1
        gradient_stops[0].transparency = 0.25

        # Add new gradient stop to the end of collection.
        gradient_stop = aw.drawing.GradientStop(drawing.Color.brown, 0.5)
        gradient_stops.add(gradient_stop)

        # Remove gradient stop at index 1.
        gradient_stops.remove_at(1)
        # And insert new gradient stop at the same index 1.
        gradient_stops.insert(1, aw.drawing.GradientStop(drawing.Color.chocolate, 0.75, 0.3))

        # Remove last gradient stop in the collection.
        gradient_stop = gradient_stops[2]
        gradient_stops.remove(gradient_stop)

        self.assertEqual(2, gradient_stops.count)

        self.assertEqual(drawing.Color.aqua.to_argb(), gradient_stops[0].color.to_argb())
        self.assertAlmostEqual(0.1, gradient_stops[0].position, delta=0.01)
        self.assertAlmostEqual(0.25, gradient_stops[0].transparency, delta=0.01)

        self.assertEqual(drawing.Color.chocolate.to_argb(), gradient_stops[1].color.to_argb())
        self.assertAlmostEqual(0.75, gradient_stops[1].position, delta=0.01)
        self.assertAlmostEqual(0.3, gradient_stops[1].transparency, delta=0.01)

        # Use the compliance option to define the shape using DML
        # if you want to get "gradient_stops" property after the document saves.
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.compliance = aw.saving.OoxmlCompliance.ISO29500_2008_STRICT

        doc.save(ARTIFACTS_DIR + "Shape.gradient_stops.docx", save_options)
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.gradient_stops.docx")

        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        gradient_stops = shape.fill.gradient_stops

        self.assertEqual(2, gradient_stops.count)

        self.assertEqual(drawing.Color.aqua.to_argb(), gradient_stops[0].color.to_argb())
        self.assertAlmostEqual(0.1, gradient_stops[0].position, delta=0.01)
        self.assertAlmostEqual(0.25, gradient_stops[0].transparency, delta=0.01)

        self.assertEqual(drawing.Color.chocolate.to_argb(), gradient_stops[1].color.to_argb())
        self.assertAlmostEqual(0.75, gradient_stops[1].position, delta=0.01)
        self.assertAlmostEqual(0.3, gradient_stops[1].transparency, delta=0.01)

    def test_fill_pattern(self):

        #ExStart
        #ExFor:Fill.patterned(PatternType)
        #ExFor:Fill.patterned(PatternType, Color, Color)
        #ExSummary:Shows how to set pattern for a shape.
        doc = aw.Document(MY_DIR + "Shape stroke pattern border.docx")

        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        fill = shape.fill

        print("Pattern value is:", fill.pattern)

        # There are several ways specified fill to a pattern.
        # 1 -  Apply pattern to the shape fill:
        fill.patterned(aw.drawing.PatternType.DIAGONAL_BRICK)

        # 2 -  Apply pattern with foreground and background colors to the shape fill:
        fill.patterned(aw.drawing.PatternType.DIAGONAL_BRICK, drawing.Color.aqua, drawing.Color.bisque)

        doc.save(ARTIFACTS_DIR + "Shape.fill_pattern.docx")
        #ExEnd

    def test_title(self):

        #ExStart
        #ExFor:ShapeBase.title
        #ExSummary:Shows how to set the title of a shape.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create a shape, give it a title, and then add it to the document.
        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.CUBE)
        shape.width = 200
        shape.height = 200
        shape.title = "My cube"

        builder.insert_node(shape)

        # When we save a document with a shape that has a title,
        # Aspose.Words will store that title in the shape's Alt Text.
        doc.save(ARTIFACTS_DIR + "Shape.title.docx")

        doc = aw.Document(ARTIFACTS_DIR + "Shape.title.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.assertEqual("", shape.title)
        self.assertEqual("Title: My cube", shape.alternative_text)
        #ExEnd

        self.verify_shape(aw.drawing.ShapeType.CUBE, "", 200.0, 200.0, 0.0, 0.0, shape)

    def test_replace_textboxes_with_images(self):

        #ExStart
        #ExFor:WrapSide
        #ExFor:ShapeBase.wrap_side
        #ExFor:NodeCollection
        #ExFor:CompositeNode.insert_after(Node,Node)
        #ExFor:NodeCollection.to_array
        #ExSummary:Shows how to replace all textbox shapes with image shapes.
        doc = aw.Document(MY_DIR + "Textboxes in drawing canvas.docx")

        shapes = [node.as_shape() for node in doc.get_child_nodes(aw.NodeType.SHAPE, True)]

        self.assertEqual(3, len([shape for shape in shapes if shape.shape_type == aw.drawing.ShapeType.TEXT_BOX]))
        self.assertEqual(1, len([shape for shape in shapes if shape.shape_type == aw.drawing.ShapeType.IMAGE]))

        for shape in shapes:
            if shape.shape_type == aw.drawing.ShapeType.TEXT_BOX:
                replacement_shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.IMAGE)
                replacement_shape.image_data.set_image(IMAGE_DIR + "Logo.jpg")
                replacement_shape.left = shape.left
                replacement_shape.top = shape.top
                replacement_shape.width = shape.width
                replacement_shape.height = shape.height
                replacement_shape.relative_horizontal_position = shape.relative_horizontal_position
                replacement_shape.relative_vertical_position = shape.relative_vertical_position
                replacement_shape.horizontal_alignment = shape.horizontal_alignment
                replacement_shape.vertical_alignment = shape.vertical_alignment
                replacement_shape.wrap_type = shape.wrap_type
                replacement_shape.wrap_side = shape.wrap_side

                shape.parent_node.insert_after(replacement_shape, shape)
                shape.remove()

        shapes = [node.as_shape() for node in doc.get_child_nodes(aw.NodeType.SHAPE, True)]

        self.assertEqual(0, len([shape for shape in shapes if shape.shape_type == aw.drawing.ShapeType.TEXT_BOX]))
        self.assertEqual(4, len([shape for shape in shapes if shape.shape_type == aw.drawing.ShapeType.IMAGE]))

        doc.save(ARTIFACTS_DIR + "Shape.replace_textboxes_with_images.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.replace_textboxes_with_images.docx")
        out_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.assertEqual(aw.drawing.WrapSide.BOTH, out_shape.wrap_side)

    def test_create_text_box(self):

        #ExStart
        #ExFor:Shape.__init__(DocumentBase,ShapeType)
        #ExFor:Story.first_paragraph
        #ExFor:Shape.first_paragraph
        #ExFor:ShapeBase.wrap_type
        #ExSummary:Shows how to create and format a text box.
        doc = aw.Document()

        # Create a floating text box.
        text_box = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_BOX)
        text_box.wrap_type = aw.drawing.WrapType.NONE
        text_box.height = 50
        text_box.width = 200

        # Set the horizontal, and vertical alignment of the text inside the shape.
        text_box.horizontal_alignment = aw.drawing.HorizontalAlignment.CENTER
        text_box.vertical_alignment = aw.drawing.VerticalAlignment.TOP

        # Add a paragraph to the text box and add a run of text that the text box will display.
        text_box.append_child(aw.Paragraph(doc))
        para = text_box.first_paragraph
        para.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        run = aw.Run(doc)
        run.text = "Hello world!"
        para.append_child(run)

        doc.first_section.body.first_paragraph.append_child(text_box)

        doc.save(ARTIFACTS_DIR + "Shape.create_text_box.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.create_text_box.docx")
        text_box = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.TEXT_BOX, "", 200.0, 50.0, 0.0, 0.0, text_box)
        self.assertEqual(aw.drawing.WrapType.NONE, text_box.wrap_type)
        self.assertEqual(aw.drawing.HorizontalAlignment.CENTER, text_box.horizontal_alignment)
        self.assertEqual(aw.drawing.VerticalAlignment.TOP, text_box.vertical_alignment)
        self.assertEqual("Hello world!", text_box.get_text().strip())

    def test_z_order(self):

        #ExStart
        #ExFor:ShapeBase.z_order
        #ExSummary:Shows how to manipulate the order of shapes.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert three different colored rectangles that partially overlap each other.
        # When we insert a shape that overlaps another shape, Aspose.Words places the newer shape on top of the old one.
        # The light green rectangle will overlap the light blue rectangle and partially obscure it,
        # and the light blue rectangle will obscure the orange rectangle.
        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN, 100,
            aw.drawing.RelativeVerticalPosition.TOP_MARGIN, 100, 200, 200, aw.drawing.WrapType.NONE)
        shape.fill_color = drawing.Color.orange

        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN, 150,
            aw.drawing.RelativeVerticalPosition.TOP_MARGIN, 150, 200, 200, aw.drawing.WrapType.NONE)
        shape.fill_color = drawing.Color.light_blue

        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN, 200,
            aw.drawing.RelativeVerticalPosition.TOP_MARGIN, 200, 200, 200, aw.drawing.WrapType.NONE)
        shape.fill_color = drawing.Color.light_green

        shapes = [node.as_shape() for node in doc.get_child_nodes(aw.NodeType.SHAPE, True)]

        # The "z_order" property of a shape determines its stacking priority among other overlapping shapes.
        # If two overlapping shapes have different "z_order" values,
        # Microsoft Word will place the shape with a higher value over the shape with the lower value.
        # Set the "z_order" values of our shapes to place the first orange rectangle over the second light blue one
        # and the second light blue rectangle over the third light green rectangle.
        # This will reverse their original stacking order.
        shapes[0].z_order = 3
        shapes[1].z_order = 2
        shapes[2].z_order = 1

        doc.save(ARTIFACTS_DIR + "Shape.z_order.docx")
        #ExEnd

    def test_get_active_x_control_properties(self):

        #ExStart
        #ExFor:OleControl
        #ExFor:OleControl.is_forms2_ole_control
        #ExFor:OleControl.name
        #ExFor:OleFormat.ole_control
        #ExFor:Forms2OleControl
        #ExFor:Forms2OleControl.caption
        #ExFor:Forms2OleControl.value
        #ExFor:Forms2OleControl.enabled
        #ExFor:Forms2OleControl.type
        #ExFor:Forms2OleControl.child_nodes
        #ExSummary:Shows how to verify the properties of an ActiveX control.
        doc = aw.Document(MY_DIR + "ActiveX controls.docx")

        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        ole_control = shape.ole_format.ole_control

        self.assertIsNone(ole_control.name)

        if ole_control.is_forms2_ole_control:

            check_box = ole_control.as_forms2_ole_control()
            self.assertEqual("Первый", check_box.caption)
            self.assertEqual("0", check_box.value)
            self.assertEqual(True, check_box.enabled)
            self.assertEqual(aw.drawing.ole.Forms2OleControlType.CHECK_BOX, check_box.type)
            self.assertIsNone(check_box.child_nodes)

        #ExEnd

    def test_get_ole_object_raw_data(self):

        #ExStart
        #ExFor:OleFormat.get_raw_data
        #ExSummary:Shows how to access the raw data of an embedded OLE object.
        doc = aw.Document(MY_DIR + "OLE objects.docx")

        for shape in doc.get_child_nodes(aw.NodeType.SHAPE, True):
            ole_format = shape.as_shape().ole_format
            if ole_format is not None:
                if ole_format.is_link:
                    print("This is a linked object")
                else:
                    print("This is an embedded object")
                ole_raw_data = ole_format.get_raw_data()

                self.assertEqual(24576, len(ole_raw_data))

        #ExEnd

    def test_linked_chart_source_full_name(self):
        #ExStart
        #ExFor:Chart.source_full_name
        #ExSummary:Shows how to get the full name of the external xls/xlsx document if the chart is linked.
        doc = aw.Document(MY_DIR + "Shape with linked chart.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        source_fullname = shape.chart.source_full_name
        self.assertTrue(source_fullname.find("Examples\\Data\\Spreadsheet.xlsx") != -1)
        #ExEnd

    def test_ole_control(self):

        #ExStart
        #ExFor:OleFormat
        #ExFor:OleFormat.auto_update
        #ExFor:OleFormat.is_locked
        #ExFor:OleFormat.prog_id
        #ExFor:OleFormat.save(BytesIO)
        #ExFor:OleFormat.save(str)
        #ExFor:OleFormat.suggested_extension
        #ExSummary:Shows how to extract embedded OLE objects into files.
        doc = aw.Document(MY_DIR + "OLE spreadsheet.docm")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        # The OLE object in the first shape is a Microsoft Excel spreadsheet.
        ole_format = shape.ole_format

        self.assertEqual("Excel.Sheet.12", ole_format.prog_id)

        # Our object is neither auto updating nor locked from updates.
        self.assertFalse(ole_format.auto_update)
        self.assertEqual(False, ole_format.is_locked)

        # If we plan on saving the OLE object to a file in the local file system,
        # we can use the "suggested_extension" property to determine which file extension to apply to the file.
        self.assertEqual(".xlsx", ole_format.suggested_extension)

        # Below are two ways of saving an OLE object to a file in the local file system.
        # 1 -  Save it via a stream:
        with open(ARTIFACTS_DIR + "OLE spreadsheet extracted via stream" + ole_format.suggested_extension, "wb") as file:
            ole_format.save(file)

        # 2 -  Save it directly to a filename:
        ole_format.save(ARTIFACTS_DIR + "OLE spreadsheet saved directly" + ole_format.suggested_extension)
        #ExEnd

        self.assertLess(8000, os.path.getsize(ARTIFACTS_DIR + "OLE spreadsheet extracted via stream.xlsx"))
        self.assertLess(8000, os.path.getsize(ARTIFACTS_DIR + "OLE spreadsheet saved directly.xlsx"))

    def test_ole_links(self):

        #ExStart
        #ExFor:OleFormat.icon_caption
        #ExFor:OleFormat.get_ole_entry(str)
        #ExFor:OleFormat.is_link
        #ExFor:OleFormat.ole_icon
        #ExFor:OleFormat.source_full_name
        #ExFor:OleFormat.source_item
        #ExSummary:Shows how to insert linked and unlinked OLE objects.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Embed a Microsoft Visio drawing into the document as an OLE object.
        builder.insert_ole_object(IMAGE_DIR + "Microsoft Visio drawing.vsd", "Package", False, False, None)

        # Insert a link to the file in the local file system and display it as an icon.
        builder.insert_ole_object(IMAGE_DIR + "Microsoft Visio drawing.vsd", "Package", True, True, None)

        # Inserting OLE objects creates shapes that store these objects.
        shapes = [node.as_shape() for node in doc.get_child_nodes(aw.NodeType.SHAPE, True)]

        self.assertEqual(2, len(shapes))
        self.assertEqual(2, len([shape for shape in shapes if shape.shape_type == aw.drawing.ShapeType.OLE_OBJECT]))

        # If a shape contains an OLE object, it will have a valid "ole_format" property,
        # which we can use to verify some aspects of the shape.
        ole_format = shapes[0].ole_format

        self.assertEqual(False, ole_format.is_link)
        self.assertEqual(False, ole_format.ole_icon)

        ole_format = shapes[1].ole_format

        self.assertEqual(True, ole_format.is_link)
        self.assertEqual(True, ole_format.ole_icon)

        self.assertTrue(ole_format.source_full_name.endswith("Images/Microsoft Visio drawing.vsd"))
        self.assertEqual("", ole_format.source_item)

        self.assertEqual("Microsoft Visio drawing.vsd", ole_format.icon_caption)

        doc.save(ARTIFACTS_DIR + "Shape.ole_links.docx")

        # If the object contains OLE data, we can access it using a stream.
        stream = ole_format.get_ole_entry("\u0001CompObj")
        stream.seek(0)
        ole_entry_bytes = stream.read()
        self.assertEqual(76, len(ole_entry_bytes))

        #ExEnd

    def test_ole_control_collection(self):

        #ExStart
        #ExFor:OleFormat.clsid
        #ExFor:Forms2OleControlCollection
        #ExFor:Forms2OleControlCollection.count
        #ExFor:Forms2OleControlCollection.__getitem__(int)
        #ExSummary:Shows how to access an OLE control embedded in a document and its child controls.
        doc = aw.Document(MY_DIR + "OLE ActiveX controls.docm")

        # Shapes store and display OLE objects in the document's body.
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.assertEqual("6e182020-f460-11ce-9bcd-00aa00608e01", str(shape.ole_format.clsid))

        ole_control = shape.ole_format.ole_control.as_forms2_ole_control()

        # Some OLE controls may contain child controls, such as the one in this document with three options buttons.
        ole_control_collection = ole_control.child_nodes

        self.assertEqual(3, ole_control_collection.count)

        self.assertEqual("C#", ole_control_collection[0].caption)
        self.assertEqual("1", ole_control_collection[0].value)

        self.assertEqual("Visual Basic", ole_control_collection[1].caption)
        self.assertEqual("0", ole_control_collection[1].value)

        self.assertEqual("Delphi", ole_control_collection[2].caption)
        self.assertEqual("0", ole_control_collection[2].value)
        #ExEnd

    def test_suggested_file_name(self):

        #ExStart
        #ExFor:OleFormat.suggested_file_name
        #ExSummary:Shows how to get an OLE object's suggested file name.
        doc = aw.Document(MY_DIR + "OLE shape.rtf")

        ole_shape = doc.first_section.body.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        # OLE objects can provide a suggested filename and extension,
        # which we can use when saving the object's contents into a file in the local file system.
        suggested_file_name = ole_shape.ole_format.suggested_file_name

        self.assertEqual("CSV.csv", suggested_file_name)

        with open(ARTIFACTS_DIR + suggested_file_name, "wb") as file_stream:
            ole_shape.ole_format.save(file_stream)

        #ExEnd

    def test_object_did_not_have_suggested_file_name(self):

        doc = aw.Document(MY_DIR + "ActiveX controls.docx")

        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        self.assertEqual(shape.ole_format.suggested_file_name, "")

    def test_resolution_default_values(self):

        image_options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)

        self.assertEqual(96, image_options.horizontal_resolution)
        self.assertEqual(96, image_options.vertical_resolution)

    def test_render_office_math(self):

        #ExStart
        #ExFor:ImageSaveOptions.scale
        #ExFor:OfficeMath.get_math_renderer
        #ExFor:NodeRendererBase.save(str,ImageSaveOptions)
        #ExSummary:Shows how to render an Office Math object into an image file in the local file system.
        doc = aw.Document(MY_DIR + "Office math.docx")

        math = doc.get_child(aw.NodeType.OFFICE_MATH, 0, True).as_office_math()

        # Create an "ImageSaveOptions" object to pass to the node renderer's "save" method to modify
        # how it renders the OfficeMath node into an image.
        save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)

        # Set the "scale" property to 5 to render the object to five times its original size.
        save_options.scale = 5

        math.get_math_renderer().save(ARTIFACTS_DIR + "Shape.render_office_math.png", save_options)
        #ExEnd

        self.verify_image(795, 87, filename=(ARTIFACTS_DIR + "Shape.render_office_math.png"))

    def test_office_math_display_exception(self):

        doc = aw.Document(MY_DIR + "Office math.docx")

        office_math = doc.get_child(aw.NodeType.OFFICE_MATH, 0, True).as_office_math()
        office_math.display_type = aw.math.OfficeMathDisplayType.DISPLAY

        with self.assertRaises(Exception):
            office_math.justification = aw.math.OfficeMathJustification.INLINE

    def test_office_math_default_value(self):

        doc = aw.Document(MY_DIR + "Office math.docx")

        office_math = doc.get_child(aw.NodeType.OFFICE_MATH, 6, True).as_office_math()

        self.assertEqual(aw.math.OfficeMathDisplayType.INLINE, office_math.display_type)
        self.assertEqual(aw.math.OfficeMathJustification.INLINE, office_math.justification)

    def test_office_math(self):

        #ExStart
        #ExFor:OfficeMath
        #ExFor:OfficeMath.display_type
        #ExFor:OfficeMath.equation_xml_encoding
        #ExFor:OfficeMath.justification
        #ExFor:OfficeMath.node_type
        #ExFor:OfficeMath.parent_paragraph
        #ExFor:OfficeMathDisplayType
        #ExFor:OfficeMathJustification
        #ExSummary:Shows how to set office math display formatting.
        doc = aw.Document(MY_DIR + "Office math.docx")

        office_math = doc.get_child(aw.NodeType.OFFICE_MATH, 0, True).as_office_math()

        # OfficeMath nodes that are children of other OfficeMath nodes are always inline.
        # The node we are working with is the base node to change its location and display type.
        self.assertEqual(aw.math.MathObjectType.O_MATH_PARA, office_math.math_object_type)
        self.assertEqual(aw.NodeType.OFFICE_MATH, office_math.node_type)
        self.assertEqual(office_math.parent_node, office_math.parent_paragraph)

        # OOXML and WML formats use the "equation_xml_encoding" property.
        self.assertIsNone(office_math.equation_xml_encoding)

        # Change the location and display type of the OfficeMath node.
        office_math.display_type = aw.math.OfficeMathDisplayType.DISPLAY
        office_math.justification = aw.math.OfficeMathJustification.LEFT

        doc.save(ARTIFACTS_DIR + "Shape.office_math.docx")
        #ExEnd

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "Shape.office_math.docx", GOLDS_DIR + "Shape.OfficeMath Gold.docx"))

    def test_cannot_be_set_display_with_inline_justification(self):

        doc = aw.Document(MY_DIR + "Office math.docx")

        office_math = doc.get_child(aw.NodeType.OFFICE_MATH, 0, True).as_office_math()
        office_math.display_type = aw.math.OfficeMathDisplayType.DISPLAY

        with self.assertRaises(Exception):
            office_math.justification = aw.math.OfficeMathJustification.INLINE

    def test_cannot_be_set_inline_display_with_justification(self):

        doc = aw.Document(MY_DIR + "Office math.docx")

        office_math = doc.get_child(aw.NodeType.OFFICE_MATH, 0, True).as_office_math()
        office_math.display_type = aw.math.OfficeMathDisplayType.INLINE

        with self.assertRaises(Exception):
            office_math.justification = aw.math.OfficeMathJustification.CENTER

    def test_office_math_display_nested_objects(self):

        doc = aw.Document(MY_DIR + "Office math.docx")

        office_math = doc.get_child(aw.NodeType.OFFICE_MATH, 0, True).as_office_math()

        self.assertEqual(aw.math.OfficeMathDisplayType.DISPLAY, office_math.display_type)
        self.assertEqual(aw.math.OfficeMathJustification.CENTER, office_math.justification)

    def def_work_with_math_object_type(self):

        parameters = [
            (0, aw.math.MathObjectType.O_MATH_PARA),
            (1, aw.math.MathObjectType.O_MATH),
            (2, aw.math.MathObjectType.SUPERSCRIPT),
            (3, aw.math.MathObjectType.ARGUMENT),
            (4, aw.math.MathObjectType.SUPERSCRIPT_PART)]

        for index, object_type in parameters:
            with self.subTest(index=index, object_type=object_type):
                doc = aw.Document(MY_DIR + "Office math.docx")

                office_math = doc.get_child(aw.NodeType.OFFICE_MATH, index, True).as_office_math()
                self.assertEqual(object_type, office_math.math_object_type)

    def test_aspect_ratio(self):

        for lock_aspect_ratio in (True, False):
            with self.subTest(lock_aspect_ratio=lock_aspect_ratio):
                #ExStart
                #ExFor:ShapeBase.aspect_ratio_locked
                #ExSummary:Shows how to lock/unlock a shape's aspect ratio.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Insert a shape. If we open this document in Microsoft Word, we can left click the shape to reveal
                # eight sizing handles around its perimeter, which we can click and drag to change its size.
                shape = builder.insert_image(IMAGE_DIR + "Logo.jpg")

                # Set the "aspect_ratio_locked" property to "True" to preserve the shape's aspect ratio
                # when using any of the four diagonal sizing handles, which change both the image's height and width.
                # Using any orthogonal sizing handles that either change the height or width will still change the aspect ratio.
                # Set the "aspect_ratio_locked" property to "False" to allow us to
                # freely change the image's aspect ratio with all sizing handles.
                shape.aspect_ratio_locked = lock_aspect_ratio

                doc.save(ARTIFACTS_DIR + "Shape.aspect_ratio.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "Shape.aspect_ratio.docx")
                shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

                self.assertEqual(lock_aspect_ratio, shape.aspect_ratio_locked)

    def test_markup_language_by_default(self):

        #ExStart
        #ExFor:ShapeBase.markup_language
        #ExFor:ShapeBase.size_in_points
        #ExSummary:Shows how to verify a shape's size and markup language.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_image(IMAGE_DIR + "Transparent background logo.png")

        self.assertEqual(aw.drawing.ShapeMarkupLanguage.DML, shape.markup_language)
        self.assertEqual(drawing.SizeF(300, 300), shape.size_in_points)
        #ExEnd

    def test_markup_language_for_different_ms_word_versions(self):
        parameters = [
            (aw.settings.MsWordVersion.WORD2000, aw.drawing.ShapeMarkupLanguage.VML),
            (aw.settings.MsWordVersion.WORD2002, aw.drawing.ShapeMarkupLanguage.VML),
            (aw.settings.MsWordVersion.WORD2003, aw.drawing.ShapeMarkupLanguage.VML),
            (aw.settings.MsWordVersion.WORD2007, aw.drawing.ShapeMarkupLanguage.VML),
            (aw.settings.MsWordVersion.WORD2010, aw.drawing.ShapeMarkupLanguage.DML),
            (aw.settings.MsWordVersion.WORD2013, aw.drawing.ShapeMarkupLanguage.DML),
            (aw.settings.MsWordVersion.WORD2016, aw.drawing.ShapeMarkupLanguage.DML)]
        for ms_word_version, shape_markup_language in parameters:
            with self.subTest(ms_word_version=ms_word_version, shape_markup_language=shape_markup_language):
                doc = aw.Document()
                doc.compatibility_options.optimize_for(ms_word_version)

                builder = aw.DocumentBuilder(doc)
                builder.insert_image(IMAGE_DIR + "Transparent background logo.png")

                for shape in doc.get_child_nodes(aw.NodeType.SHAPE, True):
                    shape = shape.as_shape()
                    self.assertEqual(shape_markup_language, shape.markup_language)

    def test_stroke(self):

        #ExStart
        #ExFor:Stroke
        #ExFor:Stroke.on
        #ExFor:Stroke.weight
        #ExFor:Stroke.join_style
        #ExFor:Stroke.line_style
        #ExFor:ShapeLineStyle
        #ExSummary:Shows how change stroke properties.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN, 100,
            aw.drawing.RelativeVerticalPosition.TOP_MARGIN, 100, 200, 200, aw.drawing.WrapType.NONE)

        # Basic shapes, such as the rectangle, have two visible parts.
        # 1 -  The fill, which applies to the area within the outline of the shape:
        shape.fill.fore_color = drawing.Color.white

        # 2 -  The stroke, which marks the outline of the shape:
        # Modify various properties of this shape's stroke.
        stroke = shape.stroke
        stroke.on = True
        stroke.weight = 5
        stroke.color = drawing.Color.red
        stroke.dash_style = aw.drawing.DashStyle.SHORT_DASH_DOT_DOT
        stroke.join_style = aw.drawing.JoinStyle.MITER
        stroke.end_cap = aw.drawing.EndCap.SQUARE
        stroke.line_style = aw.drawing.ShapeLineStyle.TRIPLE

        doc.save(ARTIFACTS_DIR + "Shape.stroke.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.stroke.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        stroke = shape.stroke

        self.assertEqual(True, stroke.on)
        self.assertEqual(5, stroke.weight)
        self.assertEqual(drawing.Color.red.to_argb(), stroke.color.to_argb())
        self.assertEqual(aw.drawing.DashStyle.SHORT_DASH_DOT_DOT, stroke.dash_style)
        self.assertEqual(aw.drawing.JoinStyle.MITER, stroke.join_style)
        self.assertEqual(aw.drawing.EndCap.SQUARE, stroke.end_cap)
        self.assertEqual(aw.drawing.ShapeLineStyle.TRIPLE, stroke.line_style)

    def test_insert_ole_object_as_html_file(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_ole_object("http://www.aspose.com", "htmlfile", True, False, None)

        doc.save(ARTIFACTS_DIR + "Shape.insert_ole_object_as_html_file.docx")

    #def test_insert_ole_package(self):

    #    #ExStart
    #    #ExFor:OlePackage
    #    #ExFor:OleFormat.ole_package
    #    #ExFor:OlePackage.file_name
    #    #ExFor:OlePackage.display_name
    #    #ExSummary:Shows how insert an OLE object into a document.
    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    # OLE objects allow us to open other files in the local file system using another installed application
    #    # in our operating system by double-clicking on the shape that contains the OLE object in the document body.
    #    # In this case, our external file will be a ZIP archive.
    #    with open(DATABASE_DIR + "cat001.zip", 'rb') as file:
    #        zip_file_bytes = file.read()

    #    with io.BytesIO(zip_file_bytes) as stream:
    #        shape = builder.insert_ole_object(stream, "Package", True, None)

    #        shape.ole_format.ole_package.file_name = "Package file name.zip"
    #        shape.ole_format.ole_package.display_name = "Package display name.zip"

    #    doc.save(ARTIFACTS_DIR + "Shape.InsertOlePackage.docx")
    #    #ExEnd

    #    doc = aw.Document(ARTIFACTS_DIR + "Shape.InsertOlePackage.docx")
    #    get_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

    #    self.assertEqual("Package file name.zip", get_shape.ole_format.ole_package.file_name)
    #    self.assertEqual("Package display name.zip", get_shape.ole_format.ole_package.display_name)

    def test_get_access_to_ole_package(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        ole_object = builder.insert_ole_object(MY_DIR + "Spreadsheet.xlsx", False, False, None)
        ole_object_as_ole_package = builder.insert_ole_object(MY_DIR + "Spreadsheet.xlsx", "Excel.Sheet", False, False, None)

        self.assertEqual(None, ole_object.ole_format.ole_package)
        self.assertIsInstance(ole_object_as_ole_package.ole_format.ole_package, aw.drawing.OlePackage)

    def test_resize(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, 200, 300)
        shape.height = 300
        shape.width = 500
        shape.rotation = 30

        doc.save(ARTIFACTS_DIR + "Shape.resize.docx")

    def test_calendar(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_table()
        builder.row_format.height = 100
        builder.row_format.height_rule = aw.HeightRule.EXACTLY

        for i in range(31):
            if i != 0 and i % 7 == 0:
                builder.end_row()
            builder.insert_cell()
            builder.write("Cell contents")

        builder.end_table()

        runs = doc.get_child_nodes(aw.NodeType.RUN, True)
        num = 1

        for run in runs:
            run = run.as_run()
            watermark = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_PLAIN_TEXT)
            watermark.relative_horizontal_position = aw.drawing.RelativeHorizontalPosition.PAGE
            watermark.relative_vertical_position = aw.drawing.RelativeVerticalPosition.PAGE
            watermark.width = 30
            watermark.height = 30
            watermark.horizontal_alignment = aw.drawing.HorizontalAlignment.CENTER
            watermark.vertical_alignment = aw.drawing.VerticalAlignment.CENTER
            watermark.rotation = -40

            watermark.fill.fore_color = drawing.Color.gainsboro
            watermark.stroke_color = drawing.Color.gainsboro

            watermark.text_path.text = str(num)
            watermark.text_path.font_family = "Arial"

            watermark.name = "Watermark_" + str(num)
            num += 1

            watermark.behind_text = True

            builder.move_to(run)
            builder.insert_node(watermark)

        doc.save(ARTIFACTS_DIR + "Shape.calendar.docx")

        doc = aw.Document(ARTIFACTS_DIR + "Shape.calendar.docx")
        shapes = [node.as_shape() for node in doc.get_child_nodes(aw.NodeType.SHAPE, True)]

        self.assertEqual(31, len(shapes))

        for shape in shapes:
            self.verify_shape(aw.drawing.ShapeType.TEXT_PLAIN_TEXT, "Watermark_" + str(shapes.index(shape) + 1),
                30.0, 30.0, 0.0, 0.0, shape)

    def test_is_layout_in_cell(self):

        for is_layout_in_cell in (False, True):
            with self.subTest(is_layout_in_cell=is_layout_in_cell):
                #ExStart
                #ExFor:ShapeBase.is_layout_in_cell
                #ExSummary:Shows how to determine how to display a shape in a table cell.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                table = builder.start_table()
                builder.insert_cell()
                builder.insert_cell()
                builder.end_table()

                table_style = doc.styles.add(aw.StyleType.TABLE, "MyTableStyle1").as_table_style()
                table_style.bottom_padding = 20
                table_style.left_padding = 10
                table_style.right_padding = 10
                table_style.top_padding = 20
                table_style.borders.color = drawing.Color.black
                table_style.borders.line_style = aw.LineStyle.SINGLE

                table.style = table_style

                builder.move_to(table.first_row.first_cell.first_paragraph)

                shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, aw.drawing.RelativeHorizontalPosition.LEFT_MARGIN, 50,
                    aw.drawing.RelativeVerticalPosition.TOP_MARGIN, 100, 100, 100, aw.drawing.WrapType.NONE)

                # Set the "is_layout_in_cell" property to "True" to display the shape as an inline element inside the cell's paragraph.
                # The coordinate origin that will determine the shape's location will be the top left corner of the shape's cell.
                # If we re-size the cell, the shape will move to maintain the same position starting from the cell's top left.
                # Set the "is_layout_in_cell" property to "False" to display the shape as an independent floating shape.
                # The coordinate origin that will determine the shape's location will be the top left corner of the page,
                # and the shape will not respond to any re-sizing of its cell.
                shape.is_layout_in_cell = is_layout_in_cell

                # We can only apply the "is_layout_in_cell" property to floating shapes.
                shape.wrap_type = aw.drawing.WrapType.NONE

                doc.save(ARTIFACTS_DIR + "Shape.is_layout_in_cell.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "Shape.is_layout_in_cell.docx")
                table = doc.first_section.body.tables[0]
                shape = table.first_row.first_cell.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

                self.assertEqual(is_layout_in_cell, shape.is_layout_in_cell)

    def test_shape_insertion(self):

        #ExStart
        #ExFor:DocumentBuilder.insert_shape(ShapeType,RelativeHorizontalPosition,float,RelativeVerticalPosition,float,float,float,WrapType)
        #ExFor:DocumentBuilder.insert_shape(ShapeType,float,float)
        #ExFor:OoxmlCompliance
        #ExFor:OoxmlSaveOptions.compliance
        #ExSummary:Shows how to insert DML shapes into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Below are two wrapping types that shapes may have.
        # 1 -  Floating:
        builder.insert_shape(aw.drawing.ShapeType.TOP_CORNERS_ROUNDED, aw.drawing.RelativeHorizontalPosition.PAGE, 100,
                aw.drawing.RelativeVerticalPosition.PAGE, 100, 50, 50, aw.drawing.WrapType.NONE)

        # 2 -  Inline:
        builder.insert_shape(aw.drawing.ShapeType.DIAGONAL_CORNERS_ROUNDED, 50, 50)

        # If you need to create "non-primitive" shapes, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
        # TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, or DiagonalCornersRounded,
        # then save the document with "Strict" or "Transitional" compliance, which allows saving shape as DML.
        save_options = aw.saving.OoxmlSaveOptions(aw.SaveFormat.DOCX)
        save_options.compliance = aw.saving.OoxmlCompliance.ISO29500_2008_TRANSITIONAL

        doc.save(ARTIFACTS_DIR + "Shape.shape_insertion.docx", save_options)
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.shape_insertion.docx")
        shapes = [node.as_shape() for node in doc.get_child_nodes(aw.NodeType.SHAPE, True)]

        self.verify_shape(aw.drawing.ShapeType.TOP_CORNERS_ROUNDED, "TopCornersRounded 100002", 50.0, 50.0, 100.0, 100.0, shapes[0])
        self.verify_shape(aw.drawing.ShapeType.DIAGONAL_CORNERS_ROUNDED, "DiagonalCornersRounded 100004", 50.0, 50.0, 0.0, 0.0, shapes[1])

    ##ExStart
    ##ExFor:Shape.accept(DocumentVisitor)
    ##ExFor:Shape.chart
    ##ExFor:Shape.extrusion_enabled
    ##ExFor:Shape.filled
    ##ExFor:Shape.has_chart
    ##ExFor:Shape.ole_format
    ##ExFor:Shape.shadow_enabled
    ##ExFor:Shape.story_type
    ##ExFor:Shape.stroke_color
    ##ExFor:Shape.stroked
    ##ExFor:Shape.stroke_weight
    ##ExSummary:Shows how to iterate over all the shapes in a document.
    #def test_visit_shapes(self):

    #    doc = aw.Document(MY_DIR + "Revision shape.docx")
    #    self.assertEqual(2, doc.get_child_nodes(aw.NodeType.SHAPE, True).count) #ExSKip

    #    visitor = ExShape.ShapeAppearancePrinter()
    #    doc.accept(visitor)

    #    print(visitor.get_text())

    #class ShapeAppearancePrinter(aw.DocumentVisitor):
    #    """Logs appearance-related information about visited shapes."""

    #    def __init__(self):
    #        self.shapes_visited = 0
    #        self.text_indent_level = 0
    #        self.string_builder = io.StringIO()

    #    def append_line(text: str):
    #        """Appends a line to the StringBuilder with one prepended tab character for each indent level."""

    #        for i in range(self.text_indent_level):
    #            self.string_builder.write('\t')

    #        self.string_builder.write(text)
    #        self.string_builder.write('\n')

    #    def get_text(self):
    #        """Return all the text that the StringIO has accumulated."""

    #        return f"Shapes visited: {self.shapes_visited}\n{self.string_builder.getvalue()}"

    #    def visit_shape_start(self, shape: aw.drawing.Shape) -> aw.VisitorAction:
    #        """Called when this visitor visits the start of a Shape node."""

    #        self.append_line(f"Shape found: {shape.shape_type}")

    #        self.text_indent_level += 1

    #        if shape.has_chart:
    #            self.append_line(f"Has chart: {shape.chart.title.text}")

    #        self.append_line(f"Extrusion enabled: {shape.extrusion_enabled}")
    #        self.append_line(f"Shadow enabled: {shape.shadow_enabled}")
    #        self.append_line(f"StoryType: {shape.story_type}")

    #        if shape.stroked:
    #            self.assertEqual(shape.stroke.color, shape.stroke_color)
    #            self.append_line(f"Stroke colors: {shape.stroke.color}, {shape.stroke.color2}")
    #            self.append_line(f"Stroke weight: {shape.stroke_weight}")

    #        if shape.filled:
    #            self.append_line(f"Filled: {shape.fill_color}")

    #        if shape.ole_format is not None:
    #            self.append_line(f"Ole found of type: {shape.ole_format.prog_id}")

    #        if shape.signature_line is not None:
    #            self.append_line(f"Found signature line for: {shape.signature_line.signer}, {shape.signature_line.signer_title}")

    #        return aw.VisitorAction.CONTINUE

    #    def visit_shape_end(shape: aw.drawing.Shape) -> aw.VisitorAction:
    #        """Called when this visitor visits the end of a Shape node."""

    #        self.text_indent_level -= 1
    #        self.shapes_visited += 1
    #        aw.append_line(f"End of {shape.shape_type}")

    #        return aw.VisitorAction.CONTINUE

    #    def visit_group_shape_start(group_shape: aw.drawing.GroupShape) -> aw.VisitorAction:
    #        """Called when this visitor visits the start of a GroupShape node."""

    #        self.append_line(f"Shape group found: {group_shape.shape_type}")
    #        self.text_indent_level += 1

    #        return aw.VisitorAction.CONTINUE

    #    def visit_group_shape_end(groupShape: aw.drawing.GroupShape) -> aw.VisitorAction:
    #        """Called when this visitor visits the end of a GroupShape node."""

    #        self.text_indent_level -= 1
    #        self.append_line(f"End of {group_shape.shape_type}")

    #        return aw.VisitorAction.CONTINUE

    ##ExEnd

    def test_signature_line(self):

        #ExStart
        #ExFor:Shape.signature_line
        #ExFor:ShapeBase.is_signature_line
        #ExFor:SignatureLine
        #ExFor:SignatureLine.allow_comments
        #ExFor:SignatureLine.default_instructions
        #ExFor:SignatureLine.email
        #ExFor:SignatureLine.instructions
        #ExFor:SignatureLine.show_date
        #ExFor:SignatureLine.signer
        #ExFor:SignatureLine.signer_title
        #ExSummary:Shows how to create a line for a signature and insert it into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        options = aw.SignatureLineOptions()
        options.allow_comments = True
        options.default_instructions = True
        options.email = "john.doe@management.com"
        options.instructions = "Please sign here"
        options.show_date = True
        options.signer = "John Doe"
        options.signer_title = "Senior Manager"

        # Insert a shape that will contain a signature line, whose appearance we will
        # customize using the "SignatureLineOptions" object we have created above.
        # If we insert a shape whose coordinates originate at the bottom right hand corner of the page,
        # we will need to supply negative x and y coordinates to bring the shape into view.
        shape = builder.insert_signature_line(options, aw.drawing.RelativeHorizontalPosition.RIGHT_MARGIN, -170.0,
            aw.drawing.RelativeVerticalPosition.BOTTOM_MARGIN, -60.0, aw.drawing.WrapType.NONE)

        self.assertTrue(shape.is_signature_line)

        # Verify the properties of our signature line via its Shape object.
        signature_line = shape.signature_line

        self.assertEqual("john.doe@management.com", signature_line.email)
        self.assertEqual("John Doe", signature_line.signer)
        self.assertEqual("Senior Manager", signature_line.signer_title)
        self.assertEqual("Please sign here", signature_line.instructions)
        self.assertTrue(signature_line.show_date)
        self.assertTrue(signature_line.allow_comments)
        self.assertTrue(signature_line.default_instructions)

        doc.save(ARTIFACTS_DIR + "Shape.signature_line.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.signature_line.docx")
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.IMAGE, "", 192.75, 96.75, -60.0, -170.0, shape)
        self.assertTrue(shape.is_signature_line)

        signature_line = shape.signature_line

        self.assertEqual("john.doe@management.com", signature_line.email)
        self.assertEqual("John Doe", signature_line.signer)
        self.assertEqual("Senior Manager", signature_line.signer_title)
        self.assertEqual("Please sign here", signature_line.instructions)
        self.assertTrue(signature_line.show_date)
        self.assertTrue(signature_line.allow_comments)
        self.assertTrue(signature_line.default_instructions)
        self.assertFalse(signature_line.is_signed)
        self.assertFalse(signature_line.is_valid)

    def test_text_box_layout_flow(self):

        layouts = [
            aw.drawing.LayoutFlow.VERTICAL,
            aw.drawing.LayoutFlow.HORIZONTAL,
            aw.drawing.LayoutFlow.HORIZONTAL_IDEOGRAPHIC,
            aw.drawing.LayoutFlow.BOTTOM_TO_TOP,
            aw.drawing.LayoutFlow.TOP_TO_BOTTOM,
            aw.drawing.LayoutFlow.TOP_TO_BOTTOM_IDEOGRAPHIC]

        for layout_flow in layouts:
            with self.subTest(layout_flow=layout_flow):
                #ExStart
                #ExFor:Shape.text_box
                #ExFor:Shape.last_paragraph
                #ExFor:TextBox
                #ExFor:TextBox.layout_flow
                #ExSummary:Shows how to set the orientation of text inside a text box.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                text_box_shape = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 150, 100)
                text_box = text_box_shape.text_box

                # Move the document builder to inside the TextBox and add text.
                builder.move_to(text_box_shape.last_paragraph)
                builder.writeln("Hello world!")
                builder.write("Hello again!")

                # Set the "layout_flow" property to set an orientation for the text contents of this text box.
                text_box.layout_flow = layout_flow

                doc.save(ARTIFACTS_DIR + "Shape.text_box_layout_flow.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "Shape.text_box_layout_flow.docx")
                text_box_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

                self.verify_shape(aw.drawing.ShapeType.TEXT_BOX, "TextBox 100002", 150.0, 100.0, 0.0, 0.0, text_box_shape)

                if layout_flow in [aw.drawing.LayoutFlow.BOTTOM_TO_TOP,
                                   aw.drawing.LayoutFlow.HORIZONTAL,
                                   aw.drawing.LayoutFlow.TOP_TO_BOTTOM_IDEOGRAPHIC,
                                   aw.drawing.LayoutFlow.VERTICAL]:
                    expected_layout_flow = layout_flow
                elif layout_flow == aw.drawing.LayoutFlow.TOP_TO_BOTTOM:
                    expected_layout_flow = aw.drawing.LayoutFlow.VERTICAL
                else:
                    expected_layout_flow = aw.drawing.LayoutFlow.HORIZONTAL

                self.verify_text_box(expected_layout_flow, False, aw.drawing.TextBoxWrapMode.SQUARE, 3.6, 3.6, 7.2, 7.2, text_box_shape.text_box)
                self.assertEqual("Hello world!\rHello again!", text_box_shape.get_text().strip())

    def test_text_box_fit_shape_to_text(self):

        #ExStart
        #ExFor:TextBox
        #ExFor:TextBox.fit_shape_to_text
        #ExSummary:Shows how to get a text box to resize itself to fit its contents tightly.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        text_box_shape = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 150, 100)
        text_box = text_box_shape.text_box

        # Apply these values to both these members to get the parent shape to fit
        # tightly around the text contents, ignoring the dimensions we have set.
        text_box.fit_shape_to_text = True
        text_box.text_box_wrap_mode = aw.drawing.TextBoxWrapMode.NONE

        builder.move_to(text_box_shape.last_paragraph)
        builder.write("Text fit tightly inside textbox.")

        doc.save(ARTIFACTS_DIR + "Shape.text_box_fit_shape_to_text.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.text_box_fit_shape_to_text.docx")
        text_box_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.TEXT_BOX, "TextBox 100002", 150.0, 100.0, 0.0, 0.0, text_box_shape)
        self.verify_text_box(aw.drawing.LayoutFlow.HORIZONTAL, True, aw.drawing.TextBoxWrapMode.NONE, 3.6, 3.6, 7.2, 7.2, text_box_shape.text_box)
        self.assertEqual("Text fit tightly inside textbox.", text_box_shape.get_text().strip())

    def test_text_box_margins(self):

        #ExStart
        #ExFor:TextBox
        #ExFor:TextBox.internal_margin_bottom
        #ExFor:TextBox.internal_margin_left
        #ExFor:TextBox.internal_margin_right
        #ExFor:TextBox.internal_margin_top
        #ExSummary:Shows how to set internal margins for a text box.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert another textbox with specific margins.
        text_box_shape = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 100, 100)
        text_box = text_box_shape.text_box
        text_box.internal_margin_top = 15
        text_box.internal_margin_bottom = 15
        text_box.internal_margin_left = 15
        text_box.internal_margin_right = 15

        builder.move_to(text_box_shape.last_paragraph)
        builder.write("Text placed according to textbox margins.")

        doc.save(ARTIFACTS_DIR + "Shape.text_box_margins.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.text_box_margins.docx")
        text_box_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.verify_shape(aw.drawing.ShapeType.TEXT_BOX, "TextBox 100002", 100.0, 100.0, 0.0, 0.0, text_box_shape)
        self.verify_text_box(aw.drawing.LayoutFlow.HORIZONTAL, False, aw.drawing.TextBoxWrapMode.SQUARE, 15.0, 15.0, 15.0, 15.0, text_box_shape.text_box)
        self.assertEqual("Text placed according to textbox margins.", text_box_shape.get_text().strip())

    def test_text_box_сontents_wrap_mode(self):

        for text_box_wrap_mode in (aw.drawing.TextBoxWrapMode.NONE, aw.drawing.TextBoxWrapMode.SQUARE):
            with self.subTest(text_box_wrap_mode=text_box_wrap_mode):
                #ExStart
                #ExFor:TextBox.text_box_wrap_mode
                #ExFor:TextBoxWrapMode
                #ExSummary:Shows how to set a wrapping mode for the contents of a text box.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                text_box_shape = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 300, 300)
                text_box = text_box_shape.text_box

                # Set the "text_box_wrap_mode" property to "TextBoxWrapMode.NONE" to increase the text box's width
                # to accommodate text, should it be large enough.
                # Set the "text_box_wrap_mode" property to "TextBoxWrapMode.SQUARE" to
                # wrap all text inside the text box, preserving its dimensions.
                text_box.text_box_wrap_mode = text_box_wrap_mode

                builder.move_to(text_box_shape.last_paragraph)
                builder.font.size = 32
                builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")

                doc.save(ARTIFACTS_DIR + "Shape.text_box_сontents_wrap_mode.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "Shape.text_box_сontents_wrap_mode.docx")
                text_box_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

                self.verify_shape(aw.drawing.ShapeType.TEXT_BOX, "TextBox 100002", 300.0, 300.0, 0.0, 0.0, text_box_shape)
                self.verify_text_box(aw.drawing.LayoutFlow.HORIZONTAL, False, text_box_wrap_mode, 3.6, 3.6, 7.2, 7.2, text_box_shape.text_box)
                self.assertEqual("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.", text_box_shape.get_text().strip())

    def test_text_box_shape_type(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Set compatibility options to correctly using of VerticalAnchor property.
        doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2016)

        text_box_shape = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 100, 100)
        # Not all formats are compatible with this one.
        # For most of the incompatible formats, AW generated warnings on save, so use doc.WarningCallback to check it.
        text_box_shape.text_box.vertical_anchor = aw.drawing.TextBoxAnchor.BOTTOM

        builder.move_to(text_box_shape.last_paragraph)
        builder.write("Text placed bottom")

        doc.save(ARTIFACTS_DIR + "Shape.text_box_shape_type.docx")

    def test_create_link_between_text_boxes(self):

        #ExStart
        #ExFor:TextBox.is_valid_link_target(TextBox)
        #ExFor:TextBox.next
        #ExFor:TextBox.previous
        #ExFor:TextBox.break_forward_link
        #ExSummary:Shows how to link text boxes.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        text_box_shape1 = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 100, 100)
        text_box1 = text_box_shape1.text_box
        builder.writeln()

        text_box_shape2 = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 100, 100)
        text_box2 = text_box_shape2.text_box
        builder.writeln()

        text_box_shape3 = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 100, 100)
        text_box3 = text_box_shape3.text_box
        builder.writeln()

        text_box_shape4 = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 100, 100)
        text_box4 = text_box_shape4.text_box

        # Create links between some of the text boxes.
        if text_box1.is_valid_link_target(text_box2):
            text_box1.next = text_box2

        if text_box2.is_valid_link_target(text_box3):
            text_box2.next = text_box3

        # Only an empty text box may have a link.
        self.assertTrue(text_box3.is_valid_link_target(text_box4))

        builder.move_to(text_box_shape4.last_paragraph)
        builder.write("Hello world!")

        self.assertFalse(text_box3.is_valid_link_target(text_box4))

        if text_box1.next is not None and text_box1.previous is None:
            print("This TextBox is the head of the sequence")

        if text_box2.next is not None and text_box2.previous is None:
            print("This TextBox is the middle of the sequence")

        if text_box3.next is None and text_box3.previous is not None:
            print("This TextBox is the tail of the sequence")

            # Break the forward link between text_box2 and text_box3, and then verify that they are no longer linked.
            text_box3.previous.break_forward_link()

            self.assertIsNone(text_box2.next)
            self.assertIsNone(text_box3.previous)

        doc.save(ARTIFACTS_DIR + "Shape.create_link_between_text_boxes.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Shape.create_link_between_text_boxes.docx")
        shapes = [node.as_shape() for node in doc.get_child_nodes(aw.NodeType.SHAPE, True)]

        self.verify_shape(aw.drawing.ShapeType.TEXT_BOX, "TextBox 100002", 100.0, 100.0, 0.0, 0.0, shapes[0])
        self.verify_text_box(aw.drawing.LayoutFlow.HORIZONTAL, False, aw.drawing.TextBoxWrapMode.SQUARE, 3.6, 3.6, 7.2, 7.2, shapes[0].text_box)
        self.assertEqual("", shapes[0].get_text().strip())

        self.verify_shape(aw.drawing.ShapeType.TEXT_BOX, "TextBox 100004", 100.0, 100.0, 0.0, 0.0, shapes[1])
        self.verify_text_box(aw.drawing.LayoutFlow.HORIZONTAL, False, aw.drawing.TextBoxWrapMode.SQUARE, 3.6, 3.6, 7.2, 7.2, shapes[1].text_box)
        self.assertEqual("", shapes[1].get_text().strip())

        self.verify_shape(aw.drawing.ShapeType.RECTANGLE, "TextBox 100006", 100.0, 100.0, 0.0, 0.0, shapes[2])
        self.verify_text_box(aw.drawing.LayoutFlow.HORIZONTAL, False, aw.drawing.TextBoxWrapMode.SQUARE, 3.6, 3.6, 7.2, 7.2, shapes[2].text_box)
        self.assertEqual("", shapes[2].get_text().strip())

        self.verify_shape(aw.drawing.ShapeType.TEXT_BOX, "TextBox 100008", 100.0, 100.0, 0.0, 0.0, shapes[3])
        self.verify_text_box(aw.drawing.LayoutFlow.HORIZONTAL, False, aw.drawing.TextBoxWrapMode.SQUARE, 3.6, 3.6, 7.2, 7.2, shapes[3].text_box)
        self.assertEqual("Hello world!", shapes[3].get_text().strip())

    def test_vertical_anchor(self):

        for vertical_anchor in (aw.drawing.TextBoxAnchor.TOP, aw.drawing.TextBoxAnchor.MIDDLE, aw.drawing.TextBoxAnchor.BOTTOM):
            with self.subTest(vertical_anchor=vertical_anchor):
                #ExStart
                #ExFor:CompatibilityOptions
                #ExFor:CompatibilityOptions.optimize_for(MsWordVersion)
                #ExFor:TextBoxAnchor
                #ExFor:TextBox.vertical_anchor
                #ExSummary:Shows how to vertically align the text contents of a text box.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                shape = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 200, 200)

                # Set the "vertical_anchor" property to "TextBoxAnchor.TOP" to
                # align the text in this text box with the top side of the shape.
                # Set the "vertical_anchor" property to "TextBoxAnchor.MIDDLE" to
                # align the text in this text box to the center of the shape.
                # Set the "vertical_anchor" property to "TextBoxAnchor.BOTTOM" to
                # align the text in this text box to the bottom of the shape.
                shape.text_box.vertical_anchor = vertical_anchor

                builder.move_to(shape.first_paragraph)
                builder.write("Hello world!")

                # The vertical aligning of text inside text boxes is available from Microsoft Word 2007 onwards.
                doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2007)
                doc.save(ARTIFACTS_DIR + "Shape.vertical_anchor.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "Shape.vertical_anchor.docx")
                shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

                self.verify_shape(aw.drawing.ShapeType.TEXT_BOX, "TextBox 100002", 200.0, 200.0, 0.0, 0.0, shape)
                self.verify_text_box(aw.drawing.LayoutFlow.HORIZONTAL, False, aw.drawing.TextBoxWrapMode.SQUARE, 3.6, 3.6, 7.2, 7.2, shape.text_box)
                self.assertEqual(vertical_anchor, shape.text_box.vertical_anchor)
                self.assertEqual("Hello world!", shape.get_text().strip())

    #ExStart
    #ExFor:Shape.text_path
    #ExFor:ShapeBase.is_word_art
    #ExFor:TextPath
    #ExFor:TextPath.bold
    #ExFor:TextPath.fit_path
    #ExFor:TextPath.fit_shape
    #ExFor:TextPath.font_family
    #ExFor:TextPath.italic
    #ExFor:TextPath.kerning
    #ExFor:TextPath.on
    #ExFor:TextPath.reverse_rows
    #ExFor:TextPath.rotate_letters
    #ExFor:TextPath.same_letter_heights
    #ExFor:TextPath.shadow
    #ExFor:TextPath.small_caps
    #ExFor:TextPath.spacing
    #ExFor:TextPath.strike_through
    #ExFor:TextPath.text
    #ExFor:TextPath.text_path_alignment
    #ExFor:TextPath.trim
    #ExFor:TextPath.underline
    #ExFor:TextPath.x_scale
    #ExFor:TextPathAlignment
    #ExSummary:Shows how to work with WordArt.
    def test_insert_text_paths(self):

        doc = aw.Document()

        # Insert a WordArt object to display text in a shape that we can re-size and move by using the mouse in Microsoft Word.
        # Provide a "ShapeType" as an argument to set a shape for the WordArt.
        shape = self.append_word_art(doc, "Hello World! This text is bold, and italic.",
            "Arial", 480, 24, drawing.Color.white, drawing.Color.black, aw.drawing.ShapeType.TEXT_PLAIN_TEXT)

        # Apply the "bold" and "italic" formatting settings to the text using the respective properties.
        shape.text_path.bold = True
        shape.text_path.italic = True

        # Below are various other text formatting-related properties.
        self.assertFalse(shape.text_path.underline)
        self.assertFalse(shape.text_path.shadow)
        self.assertFalse(shape.text_path.strike_through)
        self.assertFalse(shape.text_path.reverse_rows)
        self.assertFalse(shape.text_path.x_scale)
        self.assertFalse(shape.text_path.trim)
        self.assertFalse(shape.text_path.small_caps)

        self.assertEqual(36.0, shape.text_path.size)
        self.assertEqual("Hello World! This text is bold, and italic.", shape.text_path.text)
        self.assertEqual(aw.drawing.ShapeType.TEXT_PLAIN_TEXT, shape.shape_type)

        # Use the "on" property to show/hide the text.
        shape = self.append_word_art(doc, "On set to \"True\"", "Calibri", 150, 24, drawing.Color.yellow, drawing.Color.red, aw.drawing.ShapeType.TEXT_PLAIN_TEXT)
        shape.text_path.on = True

        shape = self.append_word_art(doc, "On set to \"False\"", "Calibri", 150, 24, drawing.Color.yellow, drawing.Color.purple, aw.drawing.ShapeType.TEXT_PLAIN_TEXT)
        shape.text_path.on = False

        # Use the "kerning" property to enable/disable kerning spacing between certain characters.
        shape = self.append_word_art(doc, "Kerning: VAV", "Times New Roman", 90, 24, drawing.Color.orange, drawing.Color.red, aw.drawing.ShapeType.TEXT_PLAIN_TEXT)
        shape.text_path.kerning = True

        shape = self.append_word_art(doc, "No kerning: VAV", "Times New Roman", 100, 24, drawing.Color.orange, drawing.Color.red, aw.drawing.ShapeType.TEXT_PLAIN_TEXT)
        shape.text_path.kerning = False

        # Use the "spacing" property to set the custom spacing between characters on a scale from 0.0 (none) to 1.0 (default).
        shape = self.append_word_art(doc, "Spacing set to 0.1", "Calibri", 120, 24, drawing.Color.blue_violet, drawing.Color.blue, aw.drawing.ShapeType.TEXT_CASCADE_DOWN)
        shape.text_path.spacing = 0.1

        # Set the "rotate_letters" property to "True" to rotate each character 90 degrees counterclockwise.
        shape = self.append_word_art(doc, "RotateLetters", "Calibri", 200, 36, drawing.Color.green_yellow, drawing.Color.green, aw.drawing.ShapeType.TEXT_WAVE)
        shape.text_path.rotate_letters = True

        # Set the "same_letter_heights" property to "True" to get the x-height of each character to equal the cap height.
        shape = self.append_word_art(doc, "Same character height for lower and UPPER case", "Calibri", 300, 24, drawing.Color.deep_sky_blue, drawing.Color.dodger_blue, aw.drawing.ShapeType.TEXT_SLANT_UP)
        shape.text_path.same_letter_heights = True

        # By default, the text's size will always scale to fit the containing shape's size, overriding the text size setting.
        shape = self.append_word_art(doc, "FitShape on", "Calibri", 160, 24, drawing.Color.light_blue, drawing.Color.blue, aw.drawing.ShapeType.TEXT_PLAIN_TEXT)
        self.assertTrue(shape.text_path.fit_shape)
        shape.text_path.size = 24.0

        # If we set the "fit_shape: property to "False", the text will keep the size
        # which the "size" property specifies regardless of the size of the shape.
        # Use the "text_path_alignment" property also to align the text to a side of the shape.
        shape = self.append_word_art(doc, "FitShape off", "Calibri", 160, 24, drawing.Color.light_blue, drawing.Color.blue, aw.drawing.ShapeType.TEXT_PLAIN_TEXT)
        shape.text_path.fit_shape = False
        shape.text_path.size = 24.0
        shape.text_path.text_path_alignment = aw.drawing.TextPathAlignment.RIGHT

        doc.save(ARTIFACTS_DIR + "Shape.insert_text_paths.docx")
        self._test_insert_text_paths(ARTIFACTS_DIR + "Shape.insert_text_paths.docx") #ExSkip

    def append_word_art(self, doc: aw.Document, text: str, text_font_family: str, shape_width: float, shape_height: float, word_art_fill: drawing.Color, line: drawing.Color, word_art_shape_type: aw.drawing.ShapeType) -> aw.drawing.Shape:
        """Insert a new paragraph with a WordArt shape inside it."""

        # Create an inline Shape, which will serve as a container for our WordArt.
        # The shape can only be a valid WordArt shape if we assign a WordArt-designated ShapeType to it.
        # These types will have "WordArt object" in the description,
        # and their enumerator constant names will all start with "text".
        shape = aw.drawing.Shape(doc, word_art_shape_type)
        shape.wrap_type = aw.drawing.WrapType.INLINE
        shape.width = shape_width
        shape.height = shape_height
        shape.fill_color = word_art_fill
        shape.stroke_color = line

        shape.text_path.text = text
        shape.text_path.font_family = text_font_family

        para = doc.first_section.body.append_child(aw.Paragraph(doc)).as_paragraph()
        para.append_child(shape)
        return shape

    #ExEnd

    def _test_insert_text_paths(self, filename: str):

        doc = aw.Document(filename)
        shapes = [node.as_shape() for node in doc.get_child_nodes(aw.NodeType.SHAPE, True)]

        self.verify_shape(aw.drawing.ShapeType.TEXT_PLAIN_TEXT, "", 480, 24, 0.0, 0.0, shapes[0])
        self.assertTrue(shapes[0].text_path.bold)
        self.assertTrue(shapes[0].text_path.italic)

        self.verify_shape(aw.drawing.ShapeType.TEXT_PLAIN_TEXT, "", 150, 24, 0.0, 0.0, shapes[1])
        self.assertTrue(shapes[1].text_path.on)

        self.verify_shape(aw.drawing.ShapeType.TEXT_PLAIN_TEXT, "", 150, 24, 0.0, 0.0, shapes[2])
        self.assertFalse(shapes[2].text_path.on)

        self.verify_shape(aw.drawing.ShapeType.TEXT_PLAIN_TEXT, "", 90, 24, 0.0, 0.0, shapes[3])
        self.assertTrue(shapes[3].text_path.kerning)

        self.verify_shape(aw.drawing.ShapeType.TEXT_PLAIN_TEXT, "", 100, 24, 0.0, 0.0, shapes[4])
        self.assertFalse(shapes[4].text_path.kerning)

        self.verify_shape(aw.drawing.ShapeType.TEXT_CASCADE_DOWN, "", 120, 24, 0.0, 0.0, shapes[5])
        self.assertAlmostEqual(0.1, shapes[5].text_path.spacing, delta=0.01)

        self.verify_shape(aw.drawing.ShapeType.TEXT_WAVE, "", 200, 36, 0.0, 0.0, shapes[6])
        self.assertTrue(shapes[6].text_path.rotate_letters)

        self.verify_shape(aw.drawing.ShapeType.TEXT_SLANT_UP, "", 300, 24, 0.0, 0.0, shapes[7])
        self.assertTrue(shapes[7].text_path.same_letter_heights)

        self.verify_shape(aw.drawing.ShapeType.TEXT_PLAIN_TEXT, "", 160, 24, 0.0, 0.0, shapes[8])
        self.assertTrue(shapes[8].text_path.fit_shape)
        self.assertEqual(24.0, shapes[8].text_path.size)

        self.verify_shape(aw.drawing.ShapeType.TEXT_PLAIN_TEXT, "", 160, 24, 0.0, 0.0, shapes[9])
        self.assertFalse(shapes[9].text_path.fit_shape)
        self.assertEqual(24.0, shapes[9].text_path.size)
        self.assertEqual(aw.drawing.TextPathAlignment.RIGHT, shapes[9].text_path.text_path_alignment)

    def test_shape_revision(self):

        #ExStart
        #ExFor:ShapeBase.is_delete_revision
        #ExFor:ShapeBase.is_insert_revision
        #ExSummary:Shows how to work with revision shapes.
        doc = aw.Document()

        self.assertFalse(doc.track_revisions)

        # Insert an inline shape without tracking revisions, which will make this shape not a revision of any kind.
        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.CUBE)
        shape.wrap_type = aw.drawing.WrapType.INLINE
        shape.width = 100.0
        shape.height = 100.0
        doc.first_section.body.first_paragraph.append_child(shape)

        # Start tracking revisions and then insert another shape, which will be a revision.
        doc.start_track_revisions("John Doe")

        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.SUN)
        shape.wrap_type = aw.drawing.WrapType.INLINE
        shape.width = 100.0
        shape.height = 100.0
        doc.first_section.body.first_paragraph.append_child(shape)

        shapes = [node.as_shape() for node in doc.get_child_nodes(aw.NodeType.SHAPE, True)]

        self.assertEqual(2, len(shapes))

        shapes[0].remove()

        # Since we removed that shape while we were tracking changes,
        # the shape persists in the document and counts as a delete revision.
        # Accepting this revision will remove the shape permanently, and rejecting it will keep it in the document.
        self.assertEqual(aw.drawing.ShapeType.CUBE, shapes[0].shape_type)
        self.assertTrue(shapes[0].is_delete_revision)

        # And we inserted another shape while tracking changes, so that shape will count as an insert revision.
        # Accepting this revision will assimilate this shape into the document as a non-revision,
        # and rejecting the revision will remove this shape permanently.
        self.assertEqual(aw.drawing.ShapeType.SUN, shapes[1].shape_type)
        self.assertTrue(shapes[1].is_insert_revision)
        #ExEnd

    def test_move_revisions(self):

        #ExStart
        #ExFor:ShapeBase.is_move_from_revision
        #ExFor:ShapeBase.is_move_to_revision
        #ExSummary:Shows how to identify move revision shapes.
        # A move revision is when we move an element in the document body by cut-and-pasting it in Microsoft Word while
        # tracking changes. If we involve an inline shape in such a text movement, that shape will also be a revision.
        # Copying-and-pasting or moving floating shapes do not create move revisions.
        doc = aw.Document(MY_DIR + "Revision shape.docx")

        # Move revisions consist of pairs of "Move from", and "Move to" revisions. We moved in this document in one shape,
        # but until we accept or reject the move revision, there will be two instances of that shape.
        shapes = [node.as_shape() for node in doc.get_child_nodes(aw.NodeType.SHAPE, True)]

        self.assertEqual(2, len(shapes))

        # This is the "Move to" revision, which is the shape at its arrival destination.
        # If we accept the revision, this "Move to" revision shape will disappear,
        # and the "Move from" revision shape will remain.
        self.assertFalse(shapes[0].is_move_from_revision)
        self.assertTrue(shapes[0].is_move_to_revision)

        # This is the "Move from" revision, which is the shape at its original location.
        # If we accept the revision, this "Move from" revision shape will disappear,
        # and the "Move to" revision shape will remain.
        self.assertTrue(shapes[1].is_move_from_revision)
        self.assertFalse(shapes[1].is_move_to_revision)
        #ExEnd

    def test_adjust_with_effects(self):

        #ExStart
        #ExFor:ShapeBase.adjust_with_effects(RectangleF)
        #ExFor:ShapeBase.bounds_with_effects
        #ExSummary:Shows how to check how a shape's bounds are affected by shape effects.
        doc = aw.Document(MY_DIR + "Shape shadow effect.docx")

        shapes = [node.as_shape() for node in doc.get_child_nodes(aw.NodeType.SHAPE, True)]

        self.assertEqual(2, len(shapes))

        # The two shapes are identical in terms of dimensions and shape type.
        self.assertEqual(shapes[0].width, shapes[1].width)
        self.assertEqual(shapes[0].height, shapes[1].height)
        self.assertEqual(shapes[0].shape_type, shapes[1].shape_type)

        # The first shape has no effects, and the second one has a shadow and thick outline.
        # These effects make the size of the second shape's silhouette bigger than that of the first.
        # Even though the rectangle's size shows up when we click on these shapes in Microsoft Word,
        # the visible outer bounds of the second shape are affected by the shadow and outline and thus are bigger.
        # We can use the "adjust_with_effects" method to see the true size of the shape.
        self.assertEqual(0.0, shapes[0].stroke_weight)
        self.assertEqual(20.0, shapes[1].stroke_weight)
        self.assertFalse(shapes[0].shadow_enabled)
        self.assertTrue(shapes[1].shadow_enabled)

        shape = shapes[0]

        # Create a drawing.RectangleF object, representing a rectangle,
        # which we could potentially use as the coordinates and bounds for a shape.
        rectangle_f = drawing.RectangleF(200, 200, 1000, 1000)

        # Run this method to get the size of the rectangle adjusted for all our shape effects.
        rectangle_f_out = shape.adjust_with_effects(rectangle_f)

        # Since the shape has no border-changing effects, its boundary dimensions are unaffected.
        self.assertEqual(200, rectangle_f_out.x)
        self.assertEqual(200, rectangle_f_out.y)
        self.assertEqual(1000, rectangle_f_out.width)
        self.assertEqual(1000, rectangle_f_out.height)

        # Verify the final extent of the first shape, in points.
        self.assertEqual(0, shape.bounds_with_effects.x)
        self.assertEqual(0, shape.bounds_with_effects.y)
        self.assertEqual(147, shape.bounds_with_effects.width)
        self.assertEqual(147, shape.bounds_with_effects.height)

        shape = shapes[1]
        rectangle_f = drawing.RectangleF(200, 200, 1000, 1000)
        rectangle_f_out = shape.adjust_with_effects(rectangle_f)

        # The shape effects have moved the apparent top left corner of the shape slightly.
        self.assertEqual(171.5, rectangle_f_out.x)
        self.assertEqual(167, rectangle_f_out.y)

        # The effects have also affected the visible dimensions of the shape.
        self.assertEqual(1045, rectangle_f_out.width)
        self.assertEqual(1132, rectangle_f_out.height)

        # The effects have also affected the visible bounds of the shape.
        self.assertEqual(-28.5, shape.bounds_with_effects.x)
        self.assertEqual(-33, shape.bounds_with_effects.y)
        self.assertEqual(192, shape.bounds_with_effects.width)
        self.assertEqual(279, shape.bounds_with_effects.height)
        #ExEnd

    def test_render_all_shapes(self):

        #ExStart
        #ExFor:ShapeBase.get_shape_renderer
        #ExFor:NodeRendererBase.save(BytesIO,ImageSaveOptions)
        #ExSummary:Shows how to use a shape renderer to export shapes to files in the local file system.
        doc = aw.Document(MY_DIR + "Various shapes.docx")
        shapes = [node.as_shape() for node in doc.get_child_nodes(aw.NodeType.SHAPE, True)]

        self.assertEqual(7, len(shapes))

        # There are 7 shapes in the document, including one group shape with 2 child shapes.
        # We will render every shape to an image file in the local file system
        # while ignoring the group shapes since they have no appearance.
        # This will produce 6 image files.
        for shape in doc.get_child_nodes(aw.NodeType.SHAPE, True):
            shape = shape.as_shape()
            renderer = shape.get_shape_renderer()
            options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
            renderer.save(ARTIFACTS_DIR + "Shape.render_all_shapes." + shape.name + ".png", options)

        #ExEnd

    def test_document_has_smart_art_object(self):

        #ExStart
        #ExFor:Shape.has_smart_art
        #ExSummary:Shows how to count the number of shapes in a document with SmartArt objects.
        doc = aw.Document(MY_DIR + "SmartArt.docx")

        number_of_smart_art_shapes = len([node for node in doc.get_child_nodes(aw.NodeType.SHAPE, True) if node.as_shape().has_smart_art])

        self.assertEqual(2, number_of_smart_art_shapes)
        #ExEnd

    def test_office_math_renderer(self):

        #ExStart
        #ExFor:NodeRendererBase
        #ExFor:NodeRendererBase.bounds_in_points
        #ExFor:NodeRendererBase.get_bounds_in_pixels(float,float)
        #ExFor:NodeRendererBase.get_bounds_in_pixels(float,float,float)
        #ExFor:NodeRendererBase.get_opaque_bounds_in_pixels(float,float)
        #ExFor:NodeRendererBase.get_opaque_bounds_in_pixels(float,float,float)
        #ExFor:NodeRendererBase.get_size_in_pixels(float,float)
        #ExFor:NodeRendererBase.get_size_in_pixels(float,float,float)
        #ExFor:NodeRendererBase.opaque_bounds_in_points
        #ExFor:NodeRendererBase.size_in_points
        #ExFor:OfficeMathRenderer
        #ExFor:OfficeMathRenderer.__init__(OfficeMath)
        #ExSummary:Shows how to measure and scale shapes.
        doc = aw.Document(MY_DIR + "Office math.docx")

        office_math = doc.get_child(aw.NodeType.OFFICE_MATH, 0, True).as_office_math()
        renderer = aw.rendering.OfficeMathRenderer(office_math)

        # Verify the size of the image that the OfficeMath object will create when we render it.
        self.assertAlmostEqual(119.0, renderer.size_in_points.width, delta=0.2)
        self.assertAlmostEqual(13.0, renderer.size_in_points.height, delta=0.1)

        self.assertAlmostEqual(119.0, renderer.bounds_in_points.width, delta=0.2)
        self.assertAlmostEqual(13.0, renderer.bounds_in_points.height, delta=0.1)

        # Shapes with transparent parts may contain different values in the "opaque_bounds_in_points" properties.
        self.assertAlmostEqual(119.0, renderer.opaque_bounds_in_points.width, delta=0.2)
        self.assertAlmostEqual(14.2, renderer.opaque_bounds_in_points.height, delta=0.1)

        # Get the shape size in pixels, with linear scaling to a specific DPI.
        bounds = renderer.get_bounds_in_pixels(1.0, 96.0)

        self.assertEqual(159, bounds.width)
        self.assertEqual(18, bounds.height)

        # Get the shape size in pixels, but with a different DPI for the horizontal and vertical dimensions.
        bounds = renderer.get_bounds_in_pixels(1.0, 96.0, 150.0)
        self.assertEqual(159, bounds.width)
        self.assertEqual(28, bounds.height)

        # The opaque bounds may vary here also.
        bounds = renderer.get_opaque_bounds_in_pixels(1.0, 96.0)

        self.assertEqual(159, bounds.width)
        self.assertEqual(18, bounds.height)

        bounds = renderer.get_opaque_bounds_in_pixels(1.0, 96.0, 150.0)

        self.assertEqual(159, bounds.width)
        self.assertEqual(30, bounds.height)
        #ExEnd

    def test_shape_types(self):

        #ExStart
        #ExFor:ShapeType
        #ExSummary:Shows how Aspose.Words identify shapes.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_shape(aw.drawing.ShapeType.HEPTAGON, aw.drawing.RelativeHorizontalPosition.PAGE, 0,
            aw.drawing.RelativeVerticalPosition.PAGE, 0, 0, 0, aw.drawing.WrapType.NONE)

        builder.insert_shape(aw.drawing.ShapeType.CLOUD, aw.drawing.RelativeHorizontalPosition.RIGHT_MARGIN, 0,
            aw.drawing.RelativeVerticalPosition.PAGE, 0, 0, 0, aw.drawing.WrapType.NONE)

        builder.insert_shape(aw.drawing.ShapeType.MATH_PLUS, aw.drawing.RelativeHorizontalPosition.RIGHT_MARGIN, 0,
            aw.drawing.RelativeVerticalPosition.PAGE, 0, 0, 0, aw.drawing.WrapType.NONE)

        # To correct identify shape types you need to work with shapes as DML.
        save_options = aw.saving.OoxmlSaveOptions(aw.SaveFormat.DOCX)

        # "Strict" or "Transitional" compliance allows to save shape as DML.
        save_options.compliance = aw.saving.OoxmlCompliance.ISO29500_2008_TRANSITIONAL

        doc.save(ARTIFACTS_DIR + "Shape.shape_types.docx", save_options)
        doc = aw.Document(ARTIFACTS_DIR + "Shape.shape_types.docx")

        shapes = [node.as_shape() for node in doc.get_child_nodes(aw.NodeType.SHAPE, True)]

        for shape in shapes:
            print(shape.shape_type)

        #ExEnd

    def test_is_decorative(self):

        #ExStart
        #ExFor:ShapeBase.is_decorative
        #ExSummary:Shows how to set that the shape is decorative.
        doc = aw.Document(MY_DIR + "Decorative shapes.docx")

        shape = doc.get_child_nodes(aw.NodeType.SHAPE, True)[0].as_shape()
        self.assertTrue(shape.is_decorative)

        # If "alternative_text" is not empty, the shape cannot be decorative.
        # That's why our value has changed to 'False'.
        shape.alternative_text = "Alternative text."
        self.assertFalse(shape.is_decorative)

        builder = aw.DocumentBuilder(doc)

        builder.move_to_document_end()
        # Create a new shape as decorative.
        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, 100, 100)
        shape.is_decorative = True

        doc.save(ARTIFACTS_DIR + "Shape.is_decorative.docx")
        #ExEnd

    def test_fill_image(self):

        #ExStart
        #ExFor:Fill.set_image(str)
        #ExSummary:Shows how to set shape fill type as image.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # There are several ways of setting image.
        shape = builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, 80, 80)
        # 1 -  Using a local system filename:
        shape.fill.set_image(IMAGE_DIR + "Logo.jpg")
        doc.save(ARTIFACTS_DIR + "Shape.fill_image.file_name.docx")

        # 2 -  Load a file into a byte array:
        with open(IMAGE_DIR + "Logo.jpg", "rb") as stream:
            shape.fill.set_image(stream.read())

        doc.save(ARTIFACTS_DIR + "Shape.fill_image.byte_array.docx")

        # 3 -  From a stream:
        with open(IMAGE_DIR + "Logo.jpg", "rb") as stream:
            shape.fill.set_image(stream)

        doc.save(ARTIFACTS_DIR + "Shape.fill_image.stream.docx")
        #ExEnd

    def test_shadow_format(self):
        #ExStart
        #ExFor:ShadowFormat.Visible
        #ExFor:ShadowFormat.Clear()
        #ExFor:ShadowType
        #ExSummary:Shows how to work with a shadow formatting for the shape.
        doc = aw.Document(MY_DIR + "Shape stroke pattern border.docx")
        shape = doc.get_child_nodes(aw.NodeType.SHAPE, True)[0].as_shape()
        if shape.shadow_format.visible and shape.shadow_format.type == awd.ShadowType.SHADOW2:
            shape.shadow_format.type == awd.ShadowType.SHADOW7
        if shape.shadow_format.type == awd.ShadowType.SHADOW_MIXED:
            shape.shadow_format.clear()
        #ExEnd
        pass
