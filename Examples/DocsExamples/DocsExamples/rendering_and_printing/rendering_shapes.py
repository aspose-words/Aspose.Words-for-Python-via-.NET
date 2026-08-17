import io

import aspose.words as aw
import aspose.pydrawing as drawing

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class RenderingShapes(DocsExamplesBase):

    def test_render_shape_as_emf(self):

        doc = aw.Document(MY_DIR + "Rendering.docx")
        # Retrieve the target shape from the document.
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        #ExStart:RenderShapeAsEmf
        #GistId:d95a7c191b62bdce78605ee22d39b9ab
        render = shape.get_shape_renderer()
        image_options = aw.saving.ImageSaveOptions(aw.SaveFormat.EMF)
        image_options.scale = 1.5

        render.save(ARTIFACTS_DIR + "RenderShape.render_shape_as_emf.emf", image_options)
        #ExEnd:RenderShapeAsEmf

    def test_render_shape_as_jpeg(self):

        doc = aw.Document(MY_DIR + "Rendering.docx")

        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        #ExStart:RenderShapeAsJpeg
        #GistId:d95a7c191b62bdce78605ee22d39b9ab
        render = aw.rendering.ShapeRenderer(shape)
        image_options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)
        # Output the image in gray scale
        image_options.image_color_mode = aw.saving.ImageColorMode.GRAYSCALE
        # Reduce the brightness a bit (default is 0.5).
        image_options.image_brightness = 0.45

        with io.BytesIO() as stream:
            render.save(stream, image_options)

            with open(ARTIFACTS_DIR + "RenderShape.render_shape_as_jpeg.jpg", "wb") as output:
                output.write(stream.getbuffer())
        #ExEnd:RenderShapeAsJpeg

    def test_find_shape_sizes(self):

        doc = aw.Document(MY_DIR + "Rendering.docx")

        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        #ExStart:FindShapeSizes
        #GistId:d95a7c191b62bdce78605ee22d39b9ab
        shape_rendered_size = shape.get_shape_renderer().get_size_in_pixels(1.0, 96.0)

        print(f"Shape rendered size: {shape_rendered_size.width} x {shape_rendered_size.height} px")
        #ExEnd:FindShapeSizes

    def test_render_cell_to_image(self):

        doc = aw.Document(MY_DIR + "Rendering.docx")

        #ExStart:RenderCellToImage
        cell = doc.get_child(aw.NodeType.CELL, 2, True).as_cell()
        tmp = RenderingShapes.convert_to_image(doc, cell)
        tmp.save(ARTIFACTS_DIR + "RenderShape.render_cell_to_image.png")
        #ExEnd:RenderCellToImage

    def test_render_row_to_image(self):

        doc = aw.Document(MY_DIR + "Rendering.docx")

        #ExStart:RenderRowToImage
        row = doc.get_child(aw.NodeType.ROW, 0, True).as_row()
        tmp = RenderingShapes.convert_to_image(doc, row)
        tmp.save(ARTIFACTS_DIR + "RenderShape.render_row_to_image.png")
        #ExEnd:RenderRowToImage

    def test_render_paragraph_to_image(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        #ExStart:RenderParagraphToImage
        text_box_shape = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 150, 100)

        builder.move_to(text_box_shape.last_paragraph)
        builder.write("Vertical text")

        options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
        options.paper_color = drawing.Color.light_pink

        tmp = RenderingShapes.convert_to_image(doc, text_box_shape.last_paragraph)
        tmp.save(ARTIFACTS_DIR + "RenderShape.render_paragraph_to_image.png")
        #ExEnd:RenderParagraphToImage

    def test_render_shape_image(self):

        doc = aw.Document(MY_DIR + "Rendering.docx")

        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        #ExStart:RenderShapeImage
        #GistId:7fc867ac8ef1b729b6f70580fbc5b3f9
        shape.get_shape_renderer().save(
            ARTIFACTS_DIR + "RenderShape.render_shape_image.jpg",
            aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG))
        #ExEnd:RenderShapeImage

    @staticmethod
    def convert_to_image(doc: aw.Document, node: aw.CompositeNode) -> aw.Document:
        """Renders any node in a document into an image."""

        tmp = RenderingShapes.create_temporary_document(doc, node)
        RenderingShapes.append_node_content(tmp, node)
        RenderingShapes.adjust_document_layout(tmp)
        return tmp

    @staticmethod
    def create_temporary_document(doc: aw.Document, node: aw.CompositeNode) -> aw.Document:
        """Creates a temporary document for further rendering."""

        tmp = doc.clone(False).as_document()
        tmp.sections.add(tmp.import_node(node.get_ancestor(aw.NodeType.SECTION), False,
                                         aw.ImportFormatMode.USE_DESTINATION_STYLES))
        tmp.first_section.append_child(aw.Body(tmp))
        tmp.first_section.page_setup.top_margin = 0
        tmp.first_section.page_setup.bottom_margin = 0

        return tmp

    @staticmethod
    def append_node_content(tmp: aw.Document, node: aw.CompositeNode):
        """Adds a node to a temporary document."""

        if node.node_type == aw.NodeType.HEADER_FOOTER:
            for hf_node in node.get_child_nodes(aw.NodeType.ANY, False):
                tmp.first_section.body.append_child(
                    tmp.import_node(hf_node, True, aw.ImportFormatMode.USE_DESTINATION_STYLES))
        else:
            RenderingShapes.append_non_header_footer_content(tmp, node)

    @staticmethod
    def append_non_header_footer_content(tmp: aw.Document, node: aw.CompositeNode):

        # The Python API returns ancestors as a generic CompositeNode, so the story/shape
        # boundary is detected by node type rather than by the runtime class.
        story_types = (aw.NodeType.BODY, aw.NodeType.HEADER_FOOTER, aw.NodeType.COMMENT,
                       aw.NodeType.FOOTNOTE, aw.NodeType.SHAPE, aw.NodeType.GROUP_SHAPE)

        parent_node = node.parent_node
        while parent_node is not None and parent_node.node_type not in story_types:
            parent = parent_node.clone(False).as_composite_node()
            parent.append_child(node.clone(True))
            node = parent

            parent_node = parent_node.parent_node

        tmp.first_section.body.append_child(
            tmp.import_node(node, True, aw.ImportFormatMode.USE_DESTINATION_STYLES))

    @staticmethod
    def adjust_document_layout(tmp: aw.Document):
        """Adjusts the layout of the document to fit the content area."""

        enumerator = aw.layout.LayoutEnumerator(tmp)
        rect = RenderingShapes.calculate_visible_rect(enumerator, drawing.RectangleF.EMPTY)

        tmp.first_section.page_setup.page_height = rect.height
        tmp.update_page_layout()

    @staticmethod
    def calculate_visible_rect(enumerator: aw.layout.LayoutEnumerator,
                               rect: drawing.RectangleF) -> drawing.RectangleF:
        """Calculates the visible area of the content."""

        result = rect
        while True:
            if enumerator.move_first_child():
                if enumerator.type in (aw.layout.LayoutEntityType.LINE, aw.layout.LayoutEntityType.SPAN):
                    result = enumerator.rectangle if result.is_empty \
                        else drawing.RectangleF.union(result, enumerator.rectangle)
                result = RenderingShapes.calculate_visible_rect(enumerator, result)
                enumerator.move_parent()

            if not enumerator.move_next():
                break

        return result
