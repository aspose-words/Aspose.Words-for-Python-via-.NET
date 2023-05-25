import aspose.words as aw
import aspose.pydrawing as drawing
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithTableStylesAndFormatting(DocsExamplesBase):

    def test_distance_between_table_surrounding_text(self):
        #ExStart:DistanceBetweenTableSurroundingText
        #GistId:8df1ad0825619cab7c80b571c6e6ba99
        doc = aw.Document(MY_DIR + "Tables.docx")

        print("\nGet distance between table left, right, bottom, top and the surrounding text.")
        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        print(table.distance_top)
        print(table.distance_bottom)
        print(table.distance_right)
        print(table.distance_left)
        #ExEnd:DistanceBetweenTableSurroundingText

    def test_apply_outline_border(self):

        #ExStart:ApplyOutlineBorder
        #GistId:770bf20bd617f3cb80031a74cc6c9b73
        #ExStart:InlineTablePosition
        #GistId:8df1ad0825619cab7c80b571c6e6ba99
        doc = aw.Document(MY_DIR + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()
        # Align the table to the center of the page.
        table.alignment = aw.tables.TableAlignment.CENTER
        #ExEnd: InlineTablePosition
        # Clear any existing borders from the table.
        table.clear_borders()

        # Set a green border around the table but not inside.
        table.set_border(aw.BorderType.LEFT, aw.LineStyle.SINGLE, 1.5, drawing.Color.green, True)
        table.set_border(aw.BorderType.RIGHT, aw.LineStyle.SINGLE, 1.5, drawing.Color.green, True)
        table.set_border(aw.BorderType.TOP, aw.LineStyle.SINGLE, 1.5, drawing.Color.green, True)
        table.set_border(aw.BorderType.BOTTOM, aw.LineStyle.SINGLE, 1.5, drawing.Color.green, True)

        # Fill the cells with a light green solid color.
        table.set_shading(aw.TextureIndex.TEXTURE_SOLID, drawing.Color.light_green, drawing.Color.empty())

        doc.save(ARTIFACTS_DIR + "WorkingWithTableStylesAndFormatting.apply_outline_border.docx")
        #ExEnd:ApplyOutlineBorder

    def test_build_table_with_borders(self):

        #ExStart:BuildTableWithBorders
        #GistId:770bf20bd617f3cb80031a74cc6c9b73
        doc = aw.Document(MY_DIR + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        # Clear any existing borders from the table.
        table.clear_borders()

        # Set a green border around and inside the table.
        table.set_borders(aw.LineStyle.SINGLE, 1.5, drawing.Color.green)

        doc.save(ARTIFACTS_DIR + "WorkingWithTableStylesAndFormatting.build_table_with_borders.docx")
        #ExEnd:BuildTableWithBorders

    def test_modify_row_formatting(self):

        #ExStart:ModifyRowFormatting
        #GistId:770bf20bd617f3cb80031a74cc6c9b73
        doc = aw.Document(MY_DIR + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        # Retrieve the first row in the table.
        first_row = table.first_row
        first_row.row_format.borders.line_style = aw.LineStyle.NONE
        first_row.row_format.height_rule = aw.HeightRule.AUTO
        first_row.row_format.allow_break_across_pages = True
        #ExEnd:ModifyRowFormatting

    def test_apply_row_formatting(self):

        #ExStart:ApplyRowFormatting
        #GistId:770bf20bd617f3cb80031a74cc6c9b73
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()

        row_format = builder.row_format
        row_format.height = 100
        row_format.height_rule = aw.HeightRule.EXACTLY

        # These formatting properties are set on the table and are applied to all rows in the table.
        table.left_padding = 30
        table.right_padding = 30
        table.top_padding = 30
        table.bottom_padding = 30

        builder.writeln("I'm a wonderful formatted row.")

        builder.end_row()
        builder.end_table()

        doc.save(ARTIFACTS_DIR + "WorkingWithTableStylesAndFormatting.apply_row_formatting.docx")
        #ExEnd:ApplyRowFormatting

    def test_cell_padding(self):

        #ExStart:CellPadding
        #GistId:770bf20bd617f3cb80031a74cc6c9b73
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_table()
        builder.insert_cell()

        # Sets the amount of space (in points) to add to the left/top/right/bottom of the cell's contents.
        builder.cell_format.set_paddings(30, 50, 30, 50)
        builder.writeln("I'm a wonderful formatted cell.")

        builder.end_row()
        builder.end_table()

        doc.save(ARTIFACTS_DIR + "WorkingWithTableStylesAndFormatting.cell_padding.docx")
        #ExEnd:CellPadding

    def test_modify_cell_formatting(self):
        """Shows how to modify formatting of a table cell."""

        #ExStart:ModifyCellFormatting
        #GistId:770bf20bd617f3cb80031a74cc6c9b73
        doc = aw.Document(MY_DIR + "Tables.docx")
        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        first_cell = table.first_row.first_cell
        first_cell.cell_format.width = 30
        first_cell.cell_format.orientation = aw.TextOrientation.DOWNWARD
        first_cell.cell_format.shading.foreground_pattern_color = drawing.Color.light_green
        #ExEnd:ModifyCellFormatting

    def test_format_table_and_cell_with_different_borders(self):

        #ExStart:FormatTableAndCellWithDifferentBorders
        #GistId:770bf20bd617f3cb80031a74cc6c9b73
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()

        # Set the borders for the entire table.
        table.set_borders(aw.LineStyle.SINGLE, 2.0, drawing.Color.black)

        # Set the cell shading for this cell.
        builder.cell_format.shading.background_pattern_color = drawing.Color.red
        builder.writeln("Cell #1")

        builder.insert_cell()

        # Specify a different cell shading for the second cell.
        builder.cell_format.shading.background_pattern_color = drawing.Color.green
        builder.writeln("Cell #2")

        builder.end_row()

        # Clear the cell formatting from previous operations.
        builder.cell_format.clear_formatting()

        builder.insert_cell()

        # Create larger borders for the first cell of this row. This will be different
        # compared to the borders set for the table.
        builder.cell_format.borders.left.line_width = 4.0
        builder.cell_format.borders.right.line_width = 4.0
        builder.cell_format.borders.top.line_width = 4.0
        builder.cell_format.borders.bottom.line_width = 4.0
        builder.writeln("Cell #3")

        builder.insert_cell()
        builder.cell_format.clear_formatting()
        builder.writeln("Cell #4")

        doc.save(ARTIFACTS_DIR + "WorkingWithTableStylesAndFormatting.format_table_and_cell_with_different_borders.docx")
        #ExEnd:FormatTableAndCellWithDifferentBorders

    def test_table_title_and_description(self):

        #ExStart:TableTitleAndDescription
        #GistId:458eb4fd5bd1de8b06fab4d1ef1acdc6
        doc = aw.Document(MY_DIR + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()
        table.title = "Test title"
        table.description = "Test description"

        options = aw.saving.OoxmlSaveOptions()
        options.compliance = aw.saving.OoxmlCompliance.ISO29500_2008_STRICT

        doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2016)

        doc.save(ARTIFACTS_DIR + "WorkingWithTableStylesAndFormatting.table_title_and_description.docx", options)
        #ExEnd:TableTitleAndDescription

    def test_allow_cell_spacing(self):

        #ExStart:AllowCellSpacing
        #GistId:770bf20bd617f3cb80031a74cc6c9b73
        doc = aw.Document(MY_DIR + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()
        table.allow_cell_spacing = True
        table.cell_spacing = 2

        doc.save(ARTIFACTS_DIR + "WorkingWithTableStylesAndFormatting.allow_cell_spacing.docx")
        #ExEnd:AllowCellSpacing

    def test_build_table_with_style(self):

        #ExStart:BuildTableWithStyle
        #GistId:93b92a7e6f2f4bbfd9177dd7fcecbd8c
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()

        # We must insert at least one row first before setting any table formatting.
        builder.insert_cell()

        # Set the table style used based on the unique style identifier.
        table.style_identifier = aw.StyleIdentifier.MEDIUM_SHADING1_ACCENT1

        # Apply which features should be formatted by the style.
        table.style_options = aw.tables.TableStyleOptions.FIRST_COLUMN | aw.tables.TableStyleOptions.ROW_BANDS | aw.tables.TableStyleOptions.FIRST_ROW
        table.auto_fit(aw.tables.AutoFitBehavior.AUTO_FIT_TO_CONTENTS)

        builder.writeln("Item")
        builder.cell_format.right_padding = 40
        builder.insert_cell()
        builder.writeln("Quantity (kg)")
        builder.end_row()

        builder.insert_cell()
        builder.writeln("Apples")
        builder.insert_cell()
        builder.writeln("20")
        builder.end_row()

        builder.insert_cell()
        builder.writeln("Bananas")
        builder.insert_cell()
        builder.writeln("40")
        builder.end_row()

        builder.insert_cell()
        builder.writeln("Carrots")
        builder.insert_cell()
        builder.writeln("50")
        builder.end_row()

        doc.save(ARTIFACTS_DIR + "WorkingWithTableStylesAndFormatting.build_table_with_style.docx")
        #ExEnd:BuildTableWithStyle

    def test_expand_formatting_on_cells_and_row_from_style(self):

        #ExStart:ExpandFormattingOnCellsAndRowFromStyle
        #GistId:93b92a7e6f2f4bbfd9177dd7fcecbd8c
        doc = aw.Document(MY_DIR + "Tables.docx")

        # Get the first cell of the first table in the document.
        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()
        first_cell = table.first_row.first_cell

        # First print the color of the cell shading.
        # This should be empty as the current shading is stored in the table style.
        cell_shading_before = first_cell.cell_format.shading.background_pattern_color
        print("Cell shading before style expansion:", cell_shading_before)

        doc.expand_table_styles_to_direct_formatting()

        # Now print the cell shading after expanding table styles.
        # A blue background pattern color should have been applied from the table style.
        cell_shading_after = first_cell.cell_format.shading.background_pattern_color
        print("Cell shading after style expansion:", cell_shading_after)
        #ExEnd:ExpandFormattingOnCellsAndRowFromStyle

    def test_create_table_style(self):

        #ExStart:CreateTableStyle
        #GistId:93b92a7e6f2f4bbfd9177dd7fcecbd8c
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Name")
        builder.insert_cell()
        builder.write("Value")
        builder.end_row()
        builder.insert_cell()
        builder.insert_cell()
        builder.end_table()

        table_style = doc.styles.add(aw.StyleType.TABLE, "MyTableStyle1").as_table_style()
        table_style.borders.line_style = aw.LineStyle.DOUBLE
        table_style.borders.line_width = 1
        table_style.left_padding = 18
        table_style.right_padding = 18
        table_style.top_padding = 12
        table_style.bottom_padding = 12

        table.style = table_style

        doc.save(ARTIFACTS_DIR + "WorkingWithTableStylesAndFormatting.create_table_style.docx")
        #ExEnd:CreateTableStyle

    def test_define_conditional_formatting(self):

        #ExStart:DefineConditionalFormatting
        #GistId:93b92a7e6f2f4bbfd9177dd7fcecbd8c
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Name")
        builder.insert_cell()
        builder.write("Value")
        builder.end_row()
        builder.insert_cell()
        builder.insert_cell()
        builder.end_table()

        table_style = doc.styles.add(aw.StyleType.TABLE, "MyTableStyle1").as_table_style()
        table_style.conditional_styles.first_row.shading.background_pattern_color = drawing.Color.green_yellow
        table_style.conditional_styles.first_row.shading.texture = aw.TextureIndex.TEXTURE_NONE

        table.style = table_style

        doc.save(ARTIFACTS_DIR + "WorkingWithTableStylesAndFormatting.define_conditional_formatting.docx")
        #ExEnd:DefineConditionalFormatting

    def test_set_table_cell_formatting(self):

        #ExStart:DocumentBuilderSetTableCellFormatting
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_table()
        builder.insert_cell()

        cell_format = builder.cell_format
        cell_format.width = 250
        cell_format.left_padding = 30
        cell_format.right_padding = 30
        cell_format.top_padding = 30
        cell_format.bottom_padding = 30

        builder.writeln("I'm a wonderful formatted cell.")

        builder.end_row()
        builder.end_table()

        doc.save(ARTIFACTS_DIR + "WorkingWithTableStylesAndFormatting.document_builder_set_table_cell_formatting.docx")
        #ExEnd:DocumentBuilderSetTableCellFormatting

    def test_set_table_row_formatting(self):

        #ExStart:DocumentBuilderSetTableRowFormatting
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()

        row_format = builder.row_format
        row_format.height = 100
        row_format.height_rule = aw.HeightRule.EXACTLY

        # These formatting properties are set on the table and are applied to all rows in the table.
        table.left_padding = 30
        table.right_padding = 30
        table.top_padding = 30
        table.bottom_padding = 30

        builder.writeln("I'm a wonderful formatted row.")

        builder.end_row()
        builder.end_table()

        doc.save(ARTIFACTS_DIR + "WorkingWithTableStylesAndFormatting.document_builder_set_table_row_formatting.docx")
        #ExEnd:DocumentBuilderSetTableRowFormatting
