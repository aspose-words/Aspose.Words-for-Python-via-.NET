from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR, IMAGES_DIR

import aspose.words as aw
import aspose.pydrawing as drawing

class WorkWithWatermark(DocsExamplesBase):

    def test_add_text_watermark_with_specific_options(self):

        #ExStart:AddTextWatermarkWithSpecificOptions
        doc = aw.Document(MY_DIR + "Document.docx")

        options = aw.TextWatermarkOptions()

        options.font_family = "Arial"
        options.font_size = 36
        options.color = drawing.Color.black
        options.layout = aw.WatermarkLayout.HORIZONTAL
        options.is_semitrasparent = False

        doc.watermark.set_text("Test", options)

        doc.save(ARTIFACTS_DIR + "WorkWithWatermark.add_text_watermark_with_specific_options.docx")
        #ExEnd:AddTextWatermarkWithSpecificOptions

    def test_add_image_watermark_with_specific_options(self):

        #ExStart:AddImageWatermarkWithSpecificOptions
        doc = aw.Document(MY_DIR + "Document.docx")

        options = aw.ImageWatermarkOptions()

        options.scale = 5
        options.is_washout = False

        doc.watermark.set_image(IMAGES_DIR + "Transparent background logo.png", options)

        doc.save(ARTIFACTS_DIR + "WorkWithWatermark.add_image_watermark.docx")
        #ExEnd:AddImageWatermarkWithSpecificOptions

    def test_remove_watermark_from_document(self):

        #ExStart:RemoveWatermarkFromDocument
        doc = aw.Document()

        # Add a plain text watermark.
        doc.watermark.set_text("Aspose Watermark")

        # If we wish to edit the text formatting using it as a watermark,
        # we can do so by passing a TextWatermarkOptions object when creating the watermark.
        text_watermark_options = aw.TextWatermarkOptions()
        text_watermark_options.font_family = "Arial"
        text_watermark_options.font_size = 36
        text_watermark_options.color = drawing.Color.black
        text_watermark_options.layout = aw.WatermarkLayout.DIAGONAL
        text_watermark_options.is_semitrasparent = False

        doc.watermark.set_text("Aspose Watermark", text_watermark_options)

        doc.save(ARTIFACTS_DIR + "Document.text_watermark.docx")

        # We can remove a watermark from a document like this.
        if doc.watermark.type == aw.WatermarkType.TEXT:
            doc.watermark.remove()

        doc.save(ARTIFACTS_DIR + "WorkWithWatermark.remove_watermark_from_document.docx")
        #ExEnd:RemoveWatermarkFromDocument

    #ExStart:AddWatermark
    def test_add_and_remove_watermark(self):

        doc = aw.Document(MY_DIR + "Document.docx")

        self.insert_watermark_text(doc, "CONFIDENTIAL")
        doc.save(ARTIFACTS_DIR + "TestFile.watermark.docx")

        self.remove_watermark_text(doc)
        doc.save(ARTIFACTS_DIR + "WorkWithWatermark.remove_watermark.docx")

    def insert_watermark_text(self, doc: aw.Document, watermark_text: str):
        """Inserts a watermark into a document.
        
        :param doc: The input document.
        :param watermark_text: Text of the watermark.
        """

        # Create a watermark shape, this will be a WordArt shape.
        watermark = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_PLAIN_TEXT)
        watermark.name = "Watermark"

        watermark.text_path.text = watermark_text
        watermark.text_path.font_family = "Arial"
        watermark.width = 500
        watermark.height = 100

        # Text will be directed from the bottom-left to the top-right corner.
        watermark.rotation = -40

        # Remove the following two lines if you need a solid black text.
        watermark.fill_color = drawing.Color.gray
        watermark.stroke_color = drawing.Color.gray

        # Place the watermark in the page center.
        watermark.relative_horizontal_position = aw.drawing.RelativeHorizontalPosition.PAGE
        watermark.relative_vertical_position = aw.drawing.RelativeVerticalPosition.PAGE
        watermark.wrap_type = aw.drawing.WrapType.NONE
        watermark.vertical_alignment = aw.drawing.VerticalAlignment.CENTER
        watermark.horizontal_alignment = aw.drawing.HorizontalAlignment.CENTER

        # Create a new paragraph and append the watermark to this paragraph.
        watermark_para = aw.Paragraph(doc)
        watermark_para.append_child(watermark)

        # Insert the watermark into all headers of each document section.
        for sect in doc.sections:
            sect = sect.as_section()
            # There could be up to three different headers in each section.
            # Since we want the watermark to appear on all pages, insert it into all headers.
            self.insert_watermark_into_header(watermark_para, sect, aw.HeaderFooterType.HEADER_PRIMARY)
            self.insert_watermark_into_header(watermark_para, sect, aw.HeaderFooterType.HEADER_FIRST)
            self.insert_watermark_into_header(watermark_para, sect, aw.HeaderFooterType.HEADER_EVEN)

    def insert_watermark_into_header(self, watermark_para: aw.Paragraph, sect: aw.Section, header_type: aw.HeaderFooterType):

        header = sect.headers_footers.get_by_header_footer_type(header_type)
        if header is None:
            # There is no header of the specified type in the current section, so we need to create it.
            header = aw.HeaderFooter(sect.document, header_type)
            sect.headers_footers.add(header)

        # Insert a clone of the watermark into the header.
        header.append_child(watermark_para.clone(True))

    #ExEnd:AddWatermark

    #ExStart:RemoveWatermark
    def remove_watermark_text(self, doc: aw.Document):

        for header_footer in doc.get_child_nodes(aw.NodeType.HEADER_FOOTER, True):
            header_footer = header_footer.as_header_footer()

            for shape in header_footer.get_child_nodes(aw.NodeType.SHAPE, True):
                shape = shape.as_shape()
                if "WaterMark" in shape.name:
                    shape.remove()
    #ExEnd:RemoveWatermark
