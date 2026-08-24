import aspose.words as aw
import aspose.pydrawing as drawing
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class DocumentFormatting(DocsExamplesBase):

    def test_space_between_asian_and_latin_text(self):

        #ExStart:SpaceBetweenAsianAndLatinText
        #GistId:0eb0780ab42a1b0032793a6eb510da35
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        paragraph_format = builder.paragraph_format
        paragraph_format.add_space_between_far_east_and_alpha = True
        paragraph_format.add_space_between_far_east_and_digit = True

        builder.writeln("Automatically adjust space between Asian and Latin text")
        builder.writeln("Automatically adjust space between Asian text and numbers")

        doc.save(ARTIFACTS_DIR + "DocumentFormatting.space_between_asian_and_latin_text.docx")
        #ExEnd:SpaceBetweenAsianAndLatinText

    def test_asian_typography_line_break_group(self):

        #ExStart:AsianTypographyLineBreakGroup
        #GistId:0eb0780ab42a1b0032793a6eb510da35
        doc = aw.Document(MY_DIR + "Asian typography.docx")

        paragraph_format = doc.first_section.body.paragraphs[0].paragraph_format
        paragraph_format.far_east_line_break_control = False
        paragraph_format.word_wrap = True
        paragraph_format.hanging_punctuation = False

        doc.save(ARTIFACTS_DIR + "DocumentFormatting.asian_typography_line_break_group.docx")
        #ExEnd:AsianTypographyLineBreakGroup

    def test_paragraph_formatting(self):

        #ExStart:ParagraphFormatting
        #GistId:3782e77b237fd3303b01a130ae46f958
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        paragraph_format = builder.paragraph_format
        paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        paragraph_format.left_indent = 50
        paragraph_format.right_indent = 50
        paragraph_format.space_after = 25

        builder.writeln(
            "I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.")
        builder.writeln(
            "I'm another nice formatted paragraph. I'm intended to demonstrate how the space after paragraph looks like.")

        doc.save(ARTIFACTS_DIR + "DocumentFormatting.paragraph_formatting.docx")
        #ExEnd:ParagraphFormatting

    def test_multilevel_list_formatting(self):

        #ExStart:MultilevelListFormatting
        #GistId:7aba3b36b61737610167905e1bd5f350
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.list_format.apply_number_default()
        builder.writeln("Item 1")
        builder.writeln("Item 2")

        builder.list_format.list_indent()
        builder.writeln("Item 2.1")
        builder.writeln("Item 2.2")

        builder.list_format.list_indent()
        builder.writeln("Item 2.2.1")
        builder.writeln("Item 2.2.2")

        builder.list_format.list_outdent()
        builder.writeln("Item 2.3")

        builder.list_format.list_outdent()
        builder.writeln("Item 3")

        builder.list_format.remove_numbers()

        doc.save(ARTIFACTS_DIR + "DocumentFormatting.multilevel_list_formatting.docx")
        #ExEnd:MultilevelListFormatting

    def test_apply_paragraph_style(self):

        #ExStart:ApplyParagraphStyle
        #GistId:3782e77b237fd3303b01a130ae46f958
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.TITLE
        builder.write("Hello")

        doc.save(ARTIFACTS_DIR + "DocumentFormatting.apply_paragraph_style.docx")
        #ExEnd:ApplyParagraphStyle

    def test_apply_borders_and_shading_to_paragraph(self):

        #ExStart:ApplyBordersAndShadingToParagraph
        #GistId:3782e77b237fd3303b01a130ae46f958
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        borders = builder.paragraph_format.borders
        borders.distance_from_text = 20
        borders.get_by_border_type(aw.BorderType.LEFT).line_style = aw.LineStyle.DOUBLE
        borders.get_by_border_type(aw.BorderType.RIGHT).line_style = aw.LineStyle.DOUBLE
        borders.get_by_border_type(aw.BorderType.TOP).line_style = aw.LineStyle.DOUBLE
        borders.get_by_border_type(aw.BorderType.BOTTOM).line_style = aw.LineStyle.DOUBLE

        shading = builder.paragraph_format.shading
        shading.texture = aw.TextureIndex.TEXTURE_DIAGONAL_CROSS
        shading.background_pattern_color = drawing.Color.light_coral
        shading.foreground_pattern_color = drawing.Color.light_salmon

        builder.write("I'm a formatted paragraph with double border and nice shading.")

        doc.save(ARTIFACTS_DIR + "DocumentFormatting.apply_borders_and_shading_to_paragraph.doc")
        #ExEnd:ApplyBordersAndShadingToParagraph

    def test_change_asian_paragraph_spacing_and_indents(self):

        #ExStart:ChangeAsianParagraphSpacingAndIndents
        doc = aw.Document(MY_DIR + "Asian typography.docx")

        paragraph_format = doc.first_section.body.first_paragraph.paragraph_format
        paragraph_format.character_unit_left_indent = 10       # ParagraphFormat.left_indent will be updated
        paragraph_format.character_unit_right_indent = 10      # ParagraphFormat.right_indent will be updated
        paragraph_format.character_unit_first_line_indent = 20 # ParagraphFormat.first_line_indent will be updated
        paragraph_format.line_unit_before = 5                  # ParagraphFormat.space_before will be updated
        paragraph_format.line_unit_after = 10                  # ParagraphFormat.space_after will be updated

        doc.save(ARTIFACTS_DIR + "DocumentFormatting.change_asian_paragraph_spacing_and_indents.doc")
        #ExEnd:ChangeAsianParagraphSpacingAndIndents

    def test_snap_to_grid(self):

        #ExStart:SnapToGrid
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Optimize the layout when typing in Asian characters.
        par = doc.first_section.body.first_paragraph
        par.paragraph_format.snap_to_grid = True

        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod " +
                        "tempor incididunt ut labore et dolore magna aliqua.")

        par.runs[0].font.snap_to_grid = True

        doc.save(ARTIFACTS_DIR + "Paragraph.snap_to_grid.docx")
        #ExEnd:SnapToGrid

    def test_get_paragraph_style_separator(self):

        #ExStart:GetParagraphStyleSeparator
        #GistId:3782e77b237fd3303b01a130ae46f958
        doc = aw.Document(MY_DIR + "Document.docx")

        for paragraph in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True):
            paragraph = paragraph.as_paragraph()
            if paragraph.break_is_style_separator:
                print("Separator Found!")
        #ExEnd:GetParagraphStyleSeparator

    #ExStart:GetParagraphLines
    #GistId:3782e77b237fd3303b01a130ae46f958
    def test_get_paragraph_lines(self):

        doc = aw.Document(MY_DIR + "Properties.docx")

        collector = aw.layout.LayoutCollector(doc)
        enumerator = aw.layout.LayoutEnumerator(doc)
        for paragraph in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True):
            self.process_paragraph(paragraph.as_paragraph(), collector, enumerator)

    @staticmethod
    def get_position(enumerator):
        """Returns the identity of the current layout entity.

        LayoutEnumerator has no readable "current" property in Python, so a position
        is identified by the page it is on and the bounds it occupies."""
        rectangle = enumerator.rectangle

        return (enumerator.page_index, rectangle.x, rectangle.y, rectangle.width, rectangle.height)

    @staticmethod
    def get_stop_position(paragraph, collector, enumerator):
        previous_node = paragraph.previous_sibling
        if previous_node is None:
            return None

        if previous_node.node_type == aw.NodeType.PARAGRAPH:
            enumerator.set_current(collector, previous_node.as_paragraph())  # Para break.
            enumerator.move_parent()  # Last line.

            return DocumentFormatting.get_position(enumerator)

        if previous_node.node_type == aw.NodeType.TABLE:
            table = previous_node.as_table()
            enumerator.set_current(collector, table.last_row.last_cell.last_paragraph)  # Cell break.
            enumerator.move_parent()  # Cell.
            enumerator.move_parent()  # Row.

            return DocumentFormatting.get_position(enumerator)

        raise RuntimeError("Unsupported node type encountered.")

    @staticmethod
    def count_lines(enumerator, stop_position):
        """We move from line to line in a paragraph.
        When paragraph spans multiple pages the we will follow across them."""
        count = 1
        while DocumentFormatting.get_position(enumerator) != stop_position:
            if not enumerator.move_previous_logical():
                break
            count += 1

        return count

    @staticmethod
    def get_truncated_text(text):
        MAX_CHARS = 16

        return f"{text[:MAX_CHARS]}..." if len(text) > MAX_CHARS else text

    @staticmethod
    def process_paragraph(paragraph, collector, enumerator):
        try:
            enumerator.set_current(collector, paragraph)  # Para break.
        except RuntimeError:
            return  # There is no layout entity for this paragraph.

        stop_position = DocumentFormatting.get_stop_position(paragraph, collector, enumerator)

        enumerator.set_current(collector, paragraph)
        enumerator.move_parent()

        line_count = DocumentFormatting.count_lines(enumerator, stop_position)

        paragraph_text = DocumentFormatting.get_truncated_text(paragraph.get_text())
        print(f"Paragraph '{paragraph_text}' has {line_count} line(-s).")
    #ExEnd:GetParagraphLines
