# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR
from document_helper import DocumentHelper

class ExTable(ApiExampleBase):

    def test_create_table(self):

        #ExStart
        #ExFor:Table
        #ExFor:Row
        #ExFor:Cell
        #ExFor:Table.__init__(DocumentBase)
        #ExSummary:Shows how to create a table.
        doc = aw.Document()
        table = aw.tables.Table(doc)
        doc.first_section.body.append_child(table)

        # Tables contain rows, which contain cells, which may have paragraphs
        # with typical elements such as runs, shapes, and even other tables.
        # Calling the "ensure_minimum" method on a table will ensure that
        # the table has at least one row, cell, and paragraph.
        first_row = aw.tables.Row(doc)
        table.append_child(first_row)

        first_cell = aw.tables.Cell(doc)
        first_row.append_child(first_cell)

        paragraph = aw.Paragraph(doc)
        first_cell.append_child(paragraph)

        # Add text to the first call in the first row of the table.
        run = aw.Run(doc, "Hello world!")
        paragraph.append_child(run)

        doc.save(ARTIFACTS_DIR + "Table.create_table.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Table.create_table.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(1, table.rows.count)
        self.assertEqual(1, table.first_row.cells.count)
        self.assertEqual("Hello world!\a\a", table.get_text().strip())

    def test_padding(self):

        #ExStart
        #ExFor:Table.left_padding
        #ExFor:Table.right_padding
        #ExFor:Table.top_padding
        #ExFor:Table.bottom_padding
        #ExSummary:Shows how to configure content padding in a table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Row 1, cell 1.")
        builder.insert_cell()
        builder.write("Row 1, cell 2.")
        builder.end_table()

        # For every cell in the table, set the distance between its contents and each of its borders.
        # This table will maintain the minimum padding distance by wrapping text.
        table.left_padding = 30
        table.right_padding = 60
        table.top_padding = 10
        table.bottom_padding = 90
        table.preferred_width = aw.tables.PreferredWidth.from_points(250)

        doc.save(ARTIFACTS_DIR + "Table.padding.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Table.padding.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(30.0, table.left_padding)
        self.assertEqual(60.0, table.right_padding)
        self.assertEqual(10.0, table.top_padding)
        self.assertEqual(90.0, table.bottom_padding)

    def test_row_cell_format(self):

        #ExStart
        #ExFor:Row.row_format
        #ExFor:RowFormat
        #ExFor:Cell.cell_format
        #ExFor:CellFormat
        #ExFor:CellFormat.shading
        #ExSummary:Shows how to modify the format of rows and cells in a table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("City")
        builder.insert_cell()
        builder.write("Country")
        builder.end_row()
        builder.insert_cell()
        builder.write("London")
        builder.insert_cell()
        builder.write("U.K.")
        builder.end_table()

        # Use the first row's "row_format" property to modify the formatting
        # of the contents of all cells in this row.
        row_format = table.first_row.row_format
        row_format.height = 25
        row_format.borders.bottom.color = drawing.Color.red

        # Use the "cell_format" property of the first cell in the last row to modify the formatting of that cell's contents.
        cell_format = table.last_row.first_cell.cell_format
        cell_format.width = 100
        cell_format.shading.background_pattern_color = drawing.Color.orange

        doc.save(ARTIFACTS_DIR + "Table.row_cell_format.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Table.row_cell_format.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual("City\aCountry\a\aLondon\aU.K.\a\a", table.get_text().strip())

        row_format = table.first_row.row_format

        self.assertEqual(25.0, row_format.height)
        self.assertEqual(drawing.Color.red.to_argb(), row_format.borders.bottom.color.to_argb())

        cell_format = table.last_row.first_cell.cell_format

        self.assertEqual(110.8, cell_format.width)
        self.assertEqual(drawing.Color.orange.to_argb(), cell_format.shading.background_pattern_color.to_argb())

    def test_display_content_of_tables(self):

        #ExStart
        #ExFor:Cell
        #ExFor:CellCollection
        #ExFor:CellCollection.__getitem__(int)
        #ExFor:CellCollection.to_array
        #ExFor:Row
        #ExFor:Row.cells
        #ExFor:RowCollection
        #ExFor:RowCollection.__getitem__(int)
        #ExFor:RowCollection.to_array
        #ExFor:Table
        #ExFor:Table.rows
        #ExFor:TableCollection.__getitem__(int)
        #ExFor:TableCollection.to_array
        #ExSummary:Shows how to iterate through all tables in the document and print the contents of each cell.
        doc = aw.Document(MY_DIR + "Tables.docx")
        tables = doc.first_section.body.tables

        self.assertEqual(2, len(tables.to_array()))

        for i in range(tables.count):
            print(f"Start of Table {i}")

            rows = tables[i].rows

            # We can use the "to_array" method on a row collection to clone it into an array.
            self.assertSequenceEqual(list(rows), rows.to_array())
            #Assert.are_not_same(rows, rows.to_array())

            for j in range(rows.count):
                print(f"\tStart of Row {j}")

                cells = rows[j].cells

                # We can use the "to_array" method on a cell collection to clone it into an array.
                self.assertSequenceEqual(list(cells), cells.to_array())
                #Assert.are_not_same(cells, cells.to_array())

                for k in range(cells.count):
                    cell_text = cells[k].to_string(aw.SaveFormat.TEXT).strip()
                    print(f"\t\tContents of Cell:{k} = \"{cell_text}\"")

                print(f"\tEnd of Row {j}")

            print(f"End of Table {i}\n")
        #ExEnd

    #ExStart
    #ExFor:Node.get_ancestor(NodeType)
    #ExFor:Node.get_ancestor(type)
    #ExFor:Table.node_type
    #ExFor:Cell.tables
    #ExFor:TableCollection
    #ExFor:NodeCollection.count
    #ExSummary:Shows how to find out if a tables are nested.
    def test_calculate_depth_of_nested_tables(self):

        doc = aw.Document(MY_DIR + "Nested tables.docx")
        tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
        self.assertEqual(5, tables.count) #ExSkip

        for i in range(tables.count):
            table = tables[i].as_table()

            # Find out if any cells in the table have other tables as children.
            count = self.get_child_table_count(table)
            print(f"Table #{i} has {count} tables directly within its cells")

            # Find out if the table is nested inside another table, and, if so, at what depth.
            table_depth = self.get_nested_depth_of_table(table)

            if table_depth > 0:
                print(f"Table #{i} is nested inside another table at depth of {table_depth}")
            else:
                print("Table #{i} is a non nested table (is not a child of another table)")

    @staticmethod
    def get_nested_depth_of_table(table: aw.tables.Table) -> int:
        """Calculates what level a table is nested inside other tables.

        :return: An integer indicating the nesting depth of the table (number of parent table nodes).
        """

        depth = 0
        parent = table.get_ancestor(table.node_type)

        while parent is not None:
            depth += 1
            parent = parent.get_ancestor(table.node_type)

        return depth

    @staticmethod
    def get_child_table_count(table: aw.tables.Table) -> int:
        """Determines if a table contains any immediate child table within its cells.

        Do not recursively traverse through those tables to check for further tables.

        :return: Returns true if at least one child cell contains a table.
                 Returns false if no cells in the table contain a table.
        """

        child_table_count = 0

        for row in table.rows:
            row = row.as_row()

            for cell in row.cells:
                cell = cell.as_cell()

                child_tables = cell.tables

                if child_tables.count > 0:
                    child_table_count += 1

        return child_table_count

    #ExEnd

    def test_ensure_table_minimum(self):

        #ExStart
        #ExFor:Table.ensure_minimum
        #ExSummary:Shows how to ensure that a table node contains the nodes we need to add content.
        doc = aw.Document()
        table = aw.tables.Table(doc)
        doc.first_section.body.append_child(table)

        # Tables contain rows, which contain cells, which may contain paragraphs
        # with typical elements such as runs, shapes, and even other tables.
        # Our new table has none of these nodes, and we cannot add contents to it until it does.
        self.assertEqual(0, table.get_child_nodes(aw.NodeType.ANY, True).count)

        # Calling the "ensure_minimum" method on a table will ensure that
        # the table has at least one row and one cell with an empty paragraph.
        table.ensure_minimum()
        table.first_row.first_cell.first_paragraph.append_child(aw.Run(doc, "Hello world!"))
        #ExEnd

        self.assertEqual(4, table.get_child_nodes(aw.NodeType.ANY, True).count)

    def test_ensure_row_minimum(self):

        #ExStart
        #ExFor:Row.ensure_minimum
        #ExSummary:Shows how to ensure a row node contains the nodes we need to begin adding content to it.
        doc = aw.Document()
        table = aw.tables.Table(doc)
        doc.first_section.body.append_child(table)
        row = aw.tables.Row(doc)
        table.append_child(row)

        # Rows contain cells, containing paragraphs with typical elements such as runs, shapes, and even other tables.
        # Our new row has none of these nodes, and we cannot add contents to it until it does.
        self.assertEqual(0, row.get_child_nodes(aw.NodeType.ANY, True).count)

        # Calling the "ensure_minimum" method on a table will ensure that
        # the table has at least one cell with an empty paragraph.
        row.ensure_minimum()
        row.first_cell.first_paragraph.append_child(aw.Run(doc, "Hello world!"))
        #ExEnd

        self.assertEqual(3, row.get_child_nodes(aw.NodeType.ANY, True).count)

    def test_ensure_cell_minimum(self):

        #ExStart
        #ExFor:Cell.ensure_minimum
        #ExSummary:Shows how to ensure a cell node contains the nodes we need to begin adding content to it.
        doc = aw.Document()
        table = aw.tables.Table(doc)
        doc.first_section.body.append_child(table)
        row = aw.tables.Row(doc)
        table.append_child(row)
        cell = aw.tables.Cell(doc)
        row.append_child(cell)

        # Cells may contain paragraphs with typical elements such as runs, shapes, and even other tables.
        # Our new cell does not have any paragraphs, and we cannot add contents such as run and shape nodes to it until it does.
        self.assertEqual(0, cell.get_child_nodes(aw.NodeType.ANY, True).count)

        # Calling the "ensure_minimum" method on a cell will ensure that
        # the cell has at least one empty paragraph, which we can then add contents to.
        cell.ensure_minimum()
        cell.first_paragraph.append_child(aw.Run(doc, "Hello world!"))
        #ExEnd

        self.assertEqual(2, cell.get_child_nodes(aw.NodeType.ANY, True).count)

    def test_set_outline_borders(self):

        #ExStart
        #ExFor:Table.alignment
        #ExFor:TableAlignment
        #ExFor:Table.clear_borders
        #ExFor:Table.clear_shading
        #ExFor:Table.set_border
        #ExFor:TextureIndex
        #ExFor:Table.set_shading
        #ExSummary:Shows how to apply an outline border to a table.
        doc = aw.Document(MY_DIR + "Tables.docx")
        table = doc.first_section.body.tables[0]

        # Align the table to the center of the page.
        table.alignment = aw.tables.TableAlignment.CENTER

        # Clear any existing borders and shading from the table.
        table.clear_borders()
        table.clear_shading()

        # Add green borders to the outline of the table.
        table.set_border(aw.BorderType.LEFT, aw.LineStyle.SINGLE, 1.5, drawing.Color.green, True)
        table.set_border(aw.BorderType.RIGHT, aw.LineStyle.SINGLE, 1.5, drawing.Color.green, True)
        table.set_border(aw.BorderType.TOP, aw.LineStyle.SINGLE, 1.5, drawing.Color.green, True)
        table.set_border(aw.BorderType.BOTTOM, aw.LineStyle.SINGLE, 1.5, drawing.Color.green, True)

        # Fill the cells with a light green solid color.
        table.set_shading(aw.TextureIndex.TEXTURE_SOLID, drawing.Color.light_green, drawing.Color.empty())

        doc.save(ARTIFACTS_DIR + "Table.set_outline_borders.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Table.set_outline_borders.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(aw.tables.TableAlignment.CENTER, table.alignment)

        borders = table.first_row.row_format.borders

        self.assertEqual(drawing.Color.green.to_argb(), borders.top.color.to_argb())
        self.assertEqual(drawing.Color.green.to_argb(), borders.left.color.to_argb())
        self.assertEqual(drawing.Color.green.to_argb(), borders.right.color.to_argb())
        self.assertEqual(drawing.Color.green.to_argb(), borders.bottom.color.to_argb())
        self.assertNotEqual(drawing.Color.green.to_argb(), borders.horizontal.color.to_argb())
        self.assertNotEqual(drawing.Color.green.to_argb(), borders.vertical.color.to_argb())
        self.assertEqual(drawing.Color.light_green.to_argb(), table.first_row.first_cell.cell_format.shading.foreground_pattern_color.to_argb())

    def test_set_borders(self):

        #ExStart
        #ExFor:Table.set_borders
        #ExSummary:Shows how to format of all of a table's borders at once.
        doc = aw.Document(MY_DIR + "Tables.docx")
        table = doc.first_section.body.tables[0]

        # Clear all existing borders from the table.
        table.clear_borders()

        # Set a single green line to serve as every outer and inner border of this table.
        table.set_borders(aw.LineStyle.SINGLE, 1.5, drawing.Color.green)

        doc.save(ARTIFACTS_DIR + "Table.set_borders.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Table.set_borders.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(drawing.Color.green.to_argb(), table.first_row.row_format.borders.top.color.to_argb())
        self.assertEqual(drawing.Color.green.to_argb(), table.first_row.row_format.borders.left.color.to_argb())
        self.assertEqual(drawing.Color.green.to_argb(), table.first_row.row_format.borders.right.color.to_argb())
        self.assertEqual(drawing.Color.green.to_argb(), table.first_row.row_format.borders.bottom.color.to_argb())
        self.assertEqual(drawing.Color.green.to_argb(), table.first_row.row_format.borders.horizontal.color.to_argb())
        self.assertEqual(drawing.Color.green.to_argb(), table.first_row.row_format.borders.vertical.color.to_argb())

    def test_row_format(self):

        #ExStart
        #ExFor:RowFormat
        #ExFor:Row.row_format
        #ExSummary:Shows how to modify formatting of a table row.
        doc = aw.Document(MY_DIR + "Tables.docx")
        table = doc.first_section.body.tables[0]

        # Use the first row's "row_format" property to set formatting that modifies that entire row's appearance.
        first_row = table.first_row
        first_row.row_format.borders.line_style = aw.LineStyle.NONE
        first_row.row_format.height_rule = aw.HeightRule.AUTO
        first_row.row_format.allow_break_across_pages = True

        doc.save(ARTIFACTS_DIR + "Table.row_format.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Table.row_format.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(aw.LineStyle.NONE, table.first_row.row_format.borders.line_style)
        self.assertEqual(aw.HeightRule.AUTO, table.first_row.row_format.height_rule)
        self.assertTrue(table.first_row.row_format.allow_break_across_pages)

    def test_cell_format(self):

        #ExStart
        #ExFor:CellFormat
        #ExFor:Cell.cell_format
        #ExSummary:Shows how to modify formatting of a table cell.
        doc = aw.Document(MY_DIR + "Tables.docx")
        table = doc.first_section.body.tables[0]
        first_cell = table.first_row.first_cell

        # Use a cell's "cell_format" property to set formatting that modifies the appearance of that cell.
        first_cell.cell_format.width = 30
        first_cell.cell_format.orientation = aw.TextOrientation.DOWNWARD
        first_cell.cell_format.shading.foreground_pattern_color = drawing.Color.light_green

        doc.save(ARTIFACTS_DIR + "Table.cell_format.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Table.cell_format.docx")

        table = doc.first_section.body.tables[0]
        self.assertEqual(30, table.first_row.first_cell.cell_format.width)
        self.assertEqual(aw.TextOrientation.DOWNWARD, table.first_row.first_cell.cell_format.orientation)
        self.assertEqual(drawing.Color.light_green.to_argb(), table.first_row.first_cell.cell_format.shading.foreground_pattern_color.to_argb())

    def test_get_distance(self):

        #ExStart
        #ExFor:Table.distance_bottom
        #ExFor:Table.distance_left
        #ExFor:Table.distance_right
        #ExFor:Table.distance_top
        #ExSummary:Shows the minimum distance operations between table boundaries and text.
        doc = aw.Document(MY_DIR + "Table wrapped by text.docx")

        table = doc.first_section.body.tables[0]

        self.assertEqual(25.9, table.distance_top)
        self.assertEqual(25.9, table.distance_bottom)
        self.assertEqual(17.3, table.distance_left)
        self.assertEqual(17.3, table.distance_right)
        #ExEnd

    def test_borders(self):

        #ExStart
        #ExFor:Table.clear_borders
        #ExSummary:Shows how to remove all borders from a table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Hello world!")
        builder.end_table()

        # Modify the color and thickness of the top border.
        top_border = table.first_row.row_format.borders.top
        table.set_border(aw.BorderType.TOP, aw.LineStyle.DOUBLE, 1.5, drawing.Color.red, True)

        self.assertEqual(1.5, top_border.line_width)
        self.assertEqual(drawing.Color.red.to_argb(), top_border.color.to_argb())
        self.assertEqual(aw.LineStyle.DOUBLE, top_border.line_style)

        # Clear the borders of all cells in the table, and then save the document.
        table.clear_borders()
        self.assertNotEqual(drawing.Color.empty().to_argb(), top_border.color.to_argb()) #ExSkip
        doc.save(ARTIFACTS_DIR + "Table.borders.docx")

        # Verify the values of the table's properties after re-opening the document.
        doc = aw.Document(ARTIFACTS_DIR + "Table.borders.docx")
        table = doc.first_section.body.tables[0]
        top_border = table.first_row.row_format.borders.top

        self.assertEqual(0.0, top_border.line_width)
        self.assertEqual(drawing.Color.empty().to_argb(), top_border.color.to_argb())
        self.assertEqual(aw.LineStyle.NONE, top_border.line_style)
        #ExEnd

    def test_replace_cell_text(self):

        #ExStart
        #ExFor:Range.replace(str,str,FindReplaceOptions)
        #ExSummary:Shows how to replace all instances of String of text in a table and cell.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Carrots")
        builder.insert_cell()
        builder.write("50")
        builder.end_row()
        builder.insert_cell()
        builder.write("Potatoes")
        builder.insert_cell()
        builder.write("50")
        builder.end_table()

        options = aw.replacing.FindReplaceOptions()
        options.match_case = True
        options.find_whole_words_only = True

        # Perform a find-and-replace operation on an entire table.
        table.range.replace("Carrots", "Eggs", options)

        # Perform a find-and-replace operation on the last cell of the last row of the table.
        table.last_row.last_cell.range.replace("50", "20", options)

        self.assertEqual("Eggs\a50\a\a" + "Potatoes\a20\a\a", table.get_text().strip())
        #ExEnd

    def test_remove_paragraph_text_and_mark(self):

        for is_smart_paragraph_break_replacement in (True, False):
            with self.subTest(is_smart_paragraph_break_replacement=is_smart_paragraph_break_replacement):
                #ExStart
                #ExFor:FindReplaceOptions.smart_paragraph_break_replacement
                #ExSummary:Shows how to remove paragraph from a table cell with a nested table.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Create table with paragraph and inner table in first cell.
                builder.start_table()
                builder.insert_cell()
                builder.write("TEXT1")
                builder.start_table()
                builder.insert_cell()
                builder.end_table()
                builder.end_table()
                builder.writeln()

                options = aw.replacing.FindReplaceOptions()
                # When the following option is set to 'True', Aspose.Words will remove paragraph's text
                # completely with its paragraph mark. Otherwise, Aspose.Words will mimic Word and remove
                # only paragraph's text and leaves the paragraph mark intact (when a table follows the text).
                options.smart_paragraph_break_replacement = is_smart_paragraph_break_replacement
                doc.range.replace("TEXT1&p", "", options)

                doc.save(ARTIFACTS_DIR + "Table.remove_paragraph_text_and_mark.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "Table.remove_paragraph_text_and_mark.docx")

                self.assertEqual(1 if is_smart_paragraph_break_replacement else 2,
                    doc.first_section.body.tables[0].rows[0].cells[0].paragraphs.count)

    def test_print_table_range(self):

        doc = aw.Document(MY_DIR + "Tables.docx")

        table = doc.first_section.body.tables[0]

        # The range text will include control characters such as "\a" for a cell.
        # You can call ToString on the desired node to retrieve the plain text content.

        # Print the plain text range of the table to the screen.
        print("Contents of the table: ")
        print(table.range.text)

        # Print the contents of the second row to the screen.
        print("\nContents of the row: ")
        print(table.rows[1].range.text)

        # Print the contents of the last cell in the table to the screen.
        print("\nContents of the cell: ")
        print(table.last_row.last_cell.range.text)

        self.assertEqual("\aColumn 1\aColumn 2\aColumn 3\aColumn 4\a\a", table.rows[1].range.text)
        self.assertEqual("Cell 12 contents\a", table.last_row.last_cell.range.text)

    def test_clone_table(self):

        doc = aw.Document(MY_DIR + "Tables.docx")

        table = doc.first_section.body.tables[0]

        table_clone = table.clone(True).as_table()

        # Insert the cloned table into the document after the original.
        table.parent_node.insert_after(table_clone, table)

        # Insert an empty paragraph between the two tables.
        table.parent_node.insert_after(aw.Paragraph(doc), table)

        doc.save(ARTIFACTS_DIR + "Table.clone_table.doc")

        self.assertEqual(3, doc.get_child_nodes(aw.NodeType.TABLE, True).count)
        self.assertEqual(table.range.text, table_clone.range.text)

        for cell in table_clone.get_child_nodes(aw.NodeType.CELL, True):
            cell = cell.as_cell()
            cell.remove_all_children()

        self.assertEqual("", table_clone.to_string(aw.SaveFormat.TEXT).strip())

    def test_allow_break_across_pages(self):

        for allow_break_across_pages in (False, True):
            with self.subTest(allow_break_across_pages=allow_break_across_pages):
                #ExStart
                #ExFor:RowFormat.allow_break_across_pages
                #ExSummary:Shows how to disable rows breaking across pages for every row in a table.
                doc = aw.Document(MY_DIR + "Table spanning two pages.docx")
                table = doc.first_section.body.tables[0]

                # Set the "allow_break_across_pages" property to "False" to keep the row
                # in one piece if a table spans two pages, which break up along that row.
                # If the row is too big to fit in one page, Microsoft Word will push it down to the next page.
                # Set the "allow_break_across_pages" property to "True" to allow the row to break up across two pages.
                for row in table:
                    row = row.as_row()
                    row.row_format.allow_break_across_pages = allow_break_across_pages

                doc.save(ARTIFACTS_DIR + "Table.allow_break_across_pages.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "Table.allow_break_across_pages.docx")
                table = doc.first_section.body.tables[0]

                self.assertEqual(3, len([row for row in table if row.as_row().row_format.allow_break_across_pages == allow_break_across_pages]))

    def test_allow_auto_fit_on_table(self):

        for allow_auto_fit in (False, True):
            with self.subTest(allow_auto_fit=allow_auto_fit):
                #ExStart
                #ExFor:Table.allow_auto_fit
                #ExSummary:Shows how to enable/disable automatic table cell resizing.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                table = builder.start_table()
                builder.insert_cell()
                builder.cell_format.preferred_width = aw.tables.PreferredWidth.from_points(100)
                builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                              "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")

                builder.insert_cell()
                builder.cell_format.preferred_width = aw.tables.PreferredWidth.AUTO
                builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                              "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")
                builder.end_row()
                builder.end_table()

                # Set the "allow_auto_fit" property to "False" to get the table to maintain the dimensions
                # of all its rows and cells, and truncate contents if they get too large to fit.
                # Set the "allow_auto_fit" property to "True" to allow the table to change its cells' width and height
                # to accommodate their contents.
                table.allow_auto_fit = allow_auto_fit

                doc.save(ARTIFACTS_DIR + "Table.allow_auto_fit_on_table.html")
                #ExEnd

                with open(ARTIFACTS_DIR + "Table.allow_auto_fit_on_table.html", 'rb') as file:
                    text = file.read().decode('utf-8')
                    if allow_auto_fit:
                        self.assertIn(
                            "<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-right:0.5pt single\">",
                            text)
                        self.assertIn(
                            "<td style=\"border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-left:0.5pt single\">",
                            text)
                    else:
                        self.assertIn(
                            "<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-right:0.5pt single\">",
                            text)
                        self.assertIn(
                            "<td style=\"width:7.2pt; border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-left:0.5pt single\">",
                            text)

    def test_keep_table_together(self):

        #ExStart
        #ExFor:ParagraphFormat.keep_with_next
        #ExFor:Row.is_last_row
        #ExFor:Paragraph.is_end_of_cell
        #ExFor:Paragraph.is_in_cell
        #ExFor:Cell.parent_row
        #ExFor:Cell.paragraphs
        #ExSummary:Shows how to set a table to stay together on the same page.
        doc = aw.Document(MY_DIR + "Table spanning two pages.docx")
        table = doc.first_section.body.tables[0]

        # Enabling keep_with_next for every paragraph in the table except for the
        # last ones in the last row will prevent the table from splitting across multiple pages.
        for cell in table.get_child_nodes(aw.NodeType.CELL, True):
            cell = cell.as_cell()
            for para in cell.paragraphs:
                para = para.as_paragraph()

                self.assertTrue(para.is_in_cell)

                if not cell.parent_row.is_last_row and para.is_end_of_cell:
                    para.paragraph_format.keep_with_next = True

        doc.save(ARTIFACTS_DIR + "Table.keep_table_together.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Table.keep_table_together.docx")
        table = doc.first_section.body.tables[0]

        for para in table.get_child_nodes(aw.NodeType.PARAGRAPH, True):
            para = para.as_paragraph()
            if para.is_end_of_cell and para.parent_node.as_cell().parent_row.is_last_row:
                self.assertFalse(para.paragraph_format.keep_with_next)
            else:
                self.assertTrue(para.paragraph_format.keep_with_next)

    def test_get_index_of_table_elements(self):

        #ExStart
        #ExFor:NodeCollection.index_of(Node)
        #ExSummary:Shows how to get the index of a node in a collection.
        doc = aw.Document(MY_DIR + "Tables.docx")

        table = doc.first_section.body.tables[0]
        all_tables = doc.get_child_nodes(aw.NodeType.TABLE, True)

        self.assertEqual(0, all_tables.index_of(table))

        row = table.rows[2]

        self.assertEqual(2, table.index_of(row))

        cell = row.last_cell

        self.assertEqual(4, row.index_of(cell))
        #ExEnd

    def test_get_preferred_width_type_and_value(self):

        #ExStart
        #ExFor:PreferredWidthType
        #ExFor:PreferredWidth.type
        #ExFor:PreferredWidth.value
        #ExSummary:Shows how to verify the preferred width type and value of a table cell.
        doc = aw.Document(MY_DIR + "Tables.docx")

        table = doc.first_section.body.tables[0]
        first_cell = table.first_row.first_cell

        self.assertEqual(aw.tables.PreferredWidthType.PERCENT, first_cell.cell_format.preferred_width.type)
        self.assertEqual(11.16, first_cell.cell_format.preferred_width.value)
        #ExEnd

    def test_allow_cell_spacing(self):

        for allow_cell_spacing in (False, True):
            with self.subTest(allow_cell_spacing=allow_cell_spacing):
                #ExStart
                #ExFor:Table.allow_cell_spacing
                #ExFor:Table.cell_spacing
                #ExSummary:Shows how to enable spacing between individual cells in a table.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                table = builder.start_table()
                builder.insert_cell()
                builder.write("Animal")
                builder.insert_cell()
                builder.write("Class")
                builder.end_row()
                builder.insert_cell()
                builder.write("Dog")
                builder.insert_cell()
                builder.write("Mammal")
                builder.end_table()

                table.cell_spacing = 3

                # Set the "allow_cell_spacing" property to "True" to enable spacing between cells
                # with a magnitude equal to the value of the "cell_spacing" property, in points.
                # Set the "allow_cell_spacing" property to "False" to disable cell spacing
                # and ignore the value of the "cell_spacing" property.
                table.allow_cell_spacing = allow_cell_spacing

                doc.save(ARTIFACTS_DIR + "Table.allow_cell_spacing.html")

                # Adjusting the "cell_spacing" property will automatically enable cell spacing.
                table.cell_spacing = 5

                self.assertTrue(table.allow_cell_spacing)
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "Table.allow_cell_spacing.html")
                table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

                self.assertEqual(allow_cell_spacing, table.allow_cell_spacing)

                if allow_cell_spacing:
                    self.assertEqual(3.0, table.cell_spacing)
                else:
                    self.assertEqual(0.0, table.cell_spacing)

                with open(ARTIFACTS_DIR + "Table.allow_cell_spacing.html", 'rb') as file:
                    text = file.read().decode('utf-8')
                    if allow_cell_spacing:
                        self.assertIn(
                            "<td style=\"border-style:solid; border-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border:0.5pt single\">",
                            text)
                    else:
                        self.assertIn(
                            "<td style=\"border-right-style:solid; border-right-width:0.75pt; border-bottom-style:solid; border-bottom-width:0.75pt; " +
                            "padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-bottom:0.5pt single; -aw-border-right:0.5pt single\">",
                            text)

    #ExStart
    #ExFor:Table
    #ExFor:Row
    #ExFor:Cell
    #ExFor:Table.__init__(DocumentBase)
    #ExFor:Table.title
    #ExFor:Table.description
    #ExFor:Row.__init__(DocumentBase)
    #ExFor:Cell.__init__(DocumentBase)
    #ExFor:Cell.first_paragraph
    #ExSummary:Shows how to build a nested table without using a document builder.
    def test_create_nested_table(self):

        doc = aw.Document()

        # Create the outer table with three rows and four columns, and then add it to the document.
        outer_table = ExTable.create_table(doc, 3, 4, "Outer Table")
        doc.first_section.body.append_child(outer_table)

        # Create another table with two rows and two columns and then insert it into the first table's first cell.
        inner_table = ExTable.create_table(doc, 2, 2, "Inner Table")
        outer_table.first_row.first_cell.append_child(inner_table)

        doc.save(ARTIFACTS_DIR + "Table.create_nested_table.docx")
        self.create_and_test_nested_table(aw.Document(ARTIFACTS_DIR + "Table.create_nested_table.docx")) #ExSkip

    @staticmethod
    def create_table(doc: aw.Document, row_count: int, cell_count: int, cell_text: str) -> aw.tables.Table:
        """Creates a new table in the document with the given dimensions and text in each cell."""
        table = aw.tables.Table(doc)

        for row_id in range(1, row_count + 1):
            row = aw.tables.Row(doc)
            table.append_child(row)

            for cell_id in range(1, cell_count + 1):
                cell = aw.tables.Cell(doc)
                cell.append_child(aw.Paragraph(doc))
                cell.first_paragraph.append_child(aw.Run(doc, cell_text))

                row.append_child(cell)

        # You can use the "title" and "description" properties to add a title and description respectively to your table.
        # The table must have at least one row before we can use these properties.
        # These properties are meaningful for ISO / IEC 29500 compliant .docx documents (see the OoxmlCompliance class).
        # If we save the document to pre-ISO/IEC 29500 formats, Microsoft Word ignores these properties.
        table.title = "Aspose table title"
        table.description = "Aspose table description"

        return table

    #ExEnd

    def create_and_test_nested_table(self, doc: aw.Document):

        outer_table = doc.first_section.body.tables[0]
        inner_table = doc.get_child(aw.NodeType.TABLE, 1, True).as_table()

        self.assertEqual(2, doc.get_child_nodes(aw.NodeType.TABLE, True).count)
        self.assertEqual(1, outer_table.first_row.first_cell.tables.count)
        self.assertEqual(16, outer_table.get_child_nodes(aw.NodeType.CELL, True).count)
        self.assertEqual(4, inner_table.get_child_nodes(aw.NodeType.CELL, True).count)
        self.assertEqual("Aspose table title", inner_table.title)
        self.assertEqual("Aspose table description", inner_table.description)

    #ExStart
    #ExFor:CellFormat.horizontal_merge
    #ExFor:CellFormat.vertical_merge
    #ExFor:CellMerge
    #ExSummary:Prints the horizontal and vertical merge type of a cell.
    def test_check_cells_merged(self):

        doc = aw.Document(MY_DIR + "Table with merged cells.docx")
        table = doc.first_section.body.tables[0]

        for row in table.rows:
            row = row.as_row()
            for cell in row.cells:
                cell = cell.as_cell()
                print(self.print_cell_merge_type(cell))
        self.assertEqual("The cell at R1, C1 is vertically merged", self.print_cell_merge_type(table.first_row.first_cell)) #ExSkip

    @staticmethod
    def print_cell_merge_type(cell: aw.tables.Cell) -> str:

        is_horizontally_merged = cell.cell_format.horizontal_merge != aw.tables.CellMerge.NONE
        is_vertically_merged = cell.cell_format.vertical_merge != aw.tables.CellMerge.NONE
        cell_location = f"R{cell.parent_row.parent_table.index_of(cell.parent_row) + 1}, C{cell.parent_row.index_of(cell) + 1}"

        if is_horizontally_merged and is_vertically_merged:
            return f"The cell at {cell_location} is both horizontally and vertically merged"
        if is_horizontally_merged:
            return f"The cell at {cell_location} is horizontally merged."
        if is_vertically_merged:
            return f"The cell at {cell_location} is vertically merged"
        return f"The cell at {cell_location} is not merged"

    #ExEnd

    def test_merge_cell_range(self):

        doc = aw.Document(MY_DIR + "Tables.docx")

        table = doc.first_section.body.tables[0]

        # We want to merge the range of cells found in between these two cells.
        cell_start_range = table.rows[2].cells[2]
        cell_end_range = table.rows[3].cells[3]

        # Merge all the cells between the two specified cells into one.
        self.merge_cells(cell_start_range, cell_end_range)

        doc.save(ARTIFACTS_DIR + "Table.merge_cell_range.doc")

        merged_cells_count = 0
        for node in table.get_child_nodes(aw.NodeType.CELL, True):
            cell = node.as_cell()
            if (cell.cell_format.horizontal_merge != aw.tables.CellMerge.NONE or
                cell.cell_format.vertical_merge != aw.tables.CellMerge.NONE):
                merged_cells_count += 1

        self.assertEqual(4, merged_cells_count)
        self.assertTrue(table.rows[2].cells[2].cell_format.horizontal_merge == aw.tables.CellMerge.FIRST)
        self.assertTrue(table.rows[2].cells[2].cell_format.vertical_merge == aw.tables.CellMerge.FIRST)
        self.assertTrue(table.rows[3].cells[3].cell_format.horizontal_merge == aw.tables.CellMerge.PREVIOUS)
        self.assertTrue(table.rows[3].cells[3].cell_format.vertical_merge == aw.tables.CellMerge.PREVIOUS)

    @staticmethod
    def merge_cells(start_cell: aw.tables.Cell, end_cell: aw.tables.Cell):
        """Merges the range of cells found between the two specified cells both horizontally and vertically.
        Can span over multiple rows."""

        parent_table = start_cell.parent_row.parent_table

        # Find the row and cell indices for the start and end cells.
        start_cell_pos = drawing.Point(
            start_cell.parent_row.index_of(start_cell),
            parent_table.index_of(start_cell.parent_row))
        end_cell_pos = drawing.Point(
            end_cell.parent_row.index_of(end_cell),
            parent_table.index_of(end_cell.parent_row))

        # Create a range of cells to be merged based on these indices.
        # Inverse each index if the end cell is before the start cell.
        merge_range = drawing.Rectangle(
            min(start_cell_pos.x, end_cell_pos.x),
            min(start_cell_pos.y, end_cell_pos.y),
            abs(end_cell_pos.x - start_cell_pos.x) + 1,
            abs(end_cell_pos.y - start_cell_pos.y) + 1)

        for row in parent_table.rows:
            row = row.as_row()

            for cell in row.cells:
                cell = cell.as_cell()

                current_pos = drawing.Point(row.index_of(cell), parent_table.index_of(row))

                # Check if the current cell is inside our merge range, then merge it.
                if merge_range.contains(current_pos):
                    cell.cell_format.horizontal_merge = aw.tables.CellMerge.FIRST if current_pos.x == merge_range.x else aw.tables.CellMerge.PREVIOUS
                    cell.cell_format.vertical_merge = aw.tables.CellMerge.FIRST if current_pos.y == merge_range.y else aw.tables.CellMerge.PREVIOUS

    def test_combine_tables(self):

        #ExStart
        #ExFor:Cell.cell_format
        #ExFor:CellFormat.borders
        #ExFor:Table.rows
        #ExFor:Table.first_row
        #ExFor:CellFormat.clear_formatting
        #ExFor:CompositeNode.has_child_nodes
        #ExSummary:Shows how to combine the rows from two tables into one.
        doc = aw.Document(MY_DIR + "Tables.docx")

        # Below are two ways of getting a table from a document.
        # 1 -  From the "tables" collection of a Body node:
        first_table = doc.first_section.body.tables[0]

        # 2 -  Using the "get_child" method:
        second_table = doc.get_child(aw.NodeType.TABLE, 1, True).as_table()

        # Append all rows from the current table to the next.
        while second_table.has_child_nodes:
            first_table.rows.add(second_table.first_row)

        # Remove the empty table container.
        second_table.remove()

        doc.save(ARTIFACTS_DIR + "Table.combine_tables.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Table.combine_tables.docx")

        self.assertEqual(1, doc.get_child_nodes(aw.NodeType.TABLE, True).count)
        self.assertEqual(9, doc.first_section.body.tables[0].rows.count)
        self.assertEqual(42, doc.first_section.body.tables[0].get_child_nodes(aw.NodeType.CELL, True).count)

    def test_split_table(self):

        doc = aw.Document(MY_DIR + "Tables.docx")

        first_table = doc.first_section.body.tables[0]

        # We will split the table at the third row (inclusive).
        row = first_table.rows[2]

        # Create a new container for the split table.
        table = first_table.clone(False).as_table()

        # Insert the container after the original.
        first_table.parent_node.insert_after(table, first_table)

        # Add a buffer paragraph to ensure the tables stay apart.
        first_table.parent_node.insert_after(aw.Paragraph(doc), first_table)

        current_row = None
        while current_row != row:
            current_row = first_table.last_row
            table.prepend_child(current_row)

        doc = DocumentHelper.save_open(doc)

        self.assertEqual(row, table.first_row)
        self.assertEqual(2, first_table.rows.count)
        self.assertEqual(3, table.rows.count)
        self.assertEqual(3, doc.get_child_nodes(aw.NodeType.TABLE, True).count)

    def test_wrap_text(self):

        #ExStart
        #ExFor:Table.text_wrapping
        #ExFor:TextWrapping
        #ExSummary:Shows how to work with table text wrapping.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Cell 1")
        builder.insert_cell()
        builder.write("Cell 2")
        builder.end_table()
        table.preferred_width = aw.tables.PreferredWidth.from_points(300)

        builder.font.size = 16
        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")

        # Set the "text_wrapping" property to "TextWrapping.AROUND" to get the table to wrap text around it,
        # and push it down into the paragraph below by setting the position.
        table.text_wrapping = aw.tables.TextWrapping.AROUND
        table.absolute_horizontal_distance = 100
        table.absolute_vertical_distance = 20

        doc.save(ARTIFACTS_DIR + "Table.wrap_text.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Table.wrap_text.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(aw.tables.TextWrapping.AROUND, table.text_wrapping)
        self.assertEqual(100.0, table.absolute_horizontal_distance)
        self.assertEqual(20.0, table.absolute_vertical_distance)

    def test_get_floating_table_properties(self):

        #ExStart
        #ExFor:Table.horizontal_anchor
        #ExFor:Table.vertical_anchor
        #ExFor:Table.allow_overlap
        #ExFor:ShapeBase.allow_overlap
        #ExSummary:Shows how to work with floating tables properties.
        doc = aw.Document(MY_DIR + "Table wrapped by text.docx")

        table = doc.first_section.body.tables[0]

        if table.text_wrapping == aw.tables.TextWrapping.AROUND:
            self.assertEqual(aw.drawing.RelativeHorizontalPosition.MARGIN, table.horizontal_anchor)
            self.assertEqual(aw.drawing.RelativeVerticalPosition.PARAGRAPH, table.vertical_anchor)
            self.assertEqual(False, table.allow_overlap)

            # Only MARGIN, PAGE, COLUMN available in RelativeHorizontalPosition for horizontal_anchor setter.
            # The Exception will be thrown for any other values.
            table.horizontal_anchor = aw.drawing.RelativeHorizontalPosition.COLUMN

            # Only MARGIN, PAGE, PARAGRAPH available in RelativeVerticalPosition for vertical_anchor setter.
            # The Exception will be thrown for any other values.
            table.vertical_anchor = aw.drawing.RelativeVerticalPosition.PAGE

        #ExEnd

    def test_change_floating_table_properties(self):

        #ExStart
        #ExFor:Table.relative_horizontal_alignment
        #ExFor:Table.relative_vertical_alignment
        #ExFor:Table.absolute_horizontal_distance
        #ExFor:Table.absolute_vertical_distance
        #ExSummary:Shows how set the location of floating tables.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Table 1, cell 1")
        builder.end_table()
        table.preferred_width = aw.tables.PreferredWidth.from_points(300)

        # Set the table's location to a place on the page, such as, in this case, the bottom right corner.
        table.relative_vertical_alignment = aw.drawing.VerticalAlignment.BOTTOM
        table.relative_horizontal_alignment = aw.drawing.HorizontalAlignment.RIGHT

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Table 2, cell 1")
        builder.end_table()
        table.preferred_width = aw.tables.PreferredWidth.from_points(300)

        # We can also set a horizontal and vertical offset in points from the paragraph's location where we inserted the table.
        table.absolute_vertical_distance = 50
        table.absolute_horizontal_distance = 100

        doc.save(ARTIFACTS_DIR + "Table.change_floating_table_properties.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Table.change_floating_table_properties.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(aw.drawing.VerticalAlignment.BOTTOM, table.relative_vertical_alignment)
        self.assertEqual(aw.drawing.HorizontalAlignment.RIGHT, table.relative_horizontal_alignment)

        table = doc.get_child(aw.NodeType.TABLE, 1, True).as_table()

        self.assertEqual(50.0, table.absolute_vertical_distance)
        self.assertEqual(100.0, table.absolute_horizontal_distance)

    def test_table_style_creation(self):

        #ExStart
        #ExFor:Table.bidi
        #ExFor:Table.cell_spacing
        #ExFor:Table.style
        #ExFor:Table.style_name
        #ExFor:TableStyle
        #ExFor:TableStyle.allow_break_across_pages
        #ExFor:TableStyle.bidi
        #ExFor:TableStyle.cell_spacing
        #ExFor:TableStyle.bottom_padding
        #ExFor:TableStyle.left_padding
        #ExFor:TableStyle.right_padding
        #ExFor:TableStyle.top_padding
        #ExFor:TableStyle.shading
        #ExFor:TableStyle.borders
        #ExFor:TableStyle.vertical_alignment
        #ExSummary:Shows how to create custom style settings for the table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Name")
        builder.insert_cell()
        builder.write("مرحبًا")
        builder.end_row()
        builder.insert_cell()
        builder.insert_cell()
        builder.end_table()

        table_style = doc.styles.add(aw.StyleType.TABLE, "MyTableStyle1").as_table_style()
        table_style.allow_break_across_pages = True
        table_style.bidi = True
        table_style.cell_spacing = 5
        table_style.bottom_padding = 20
        table_style.left_padding = 5
        table_style.right_padding = 10
        table_style.top_padding = 20
        table_style.shading.background_pattern_color = drawing.Color.antique_white
        table_style.borders.color = drawing.Color.blue
        table_style.borders.line_style = aw.LineStyle.DOT_DASH
        table_style.vertical_alignment = aw.tables.CellVerticalAlignment.CENTER

        table.style = table_style

        # Setting the style properties of a table may affect the properties of the table itself.
        self.assertTrue(table.bidi)
        self.assertEqual(5.0, table.cell_spacing)
        self.assertEqual("MyTableStyle1", table.style_name)

        doc.save(ARTIFACTS_DIR + "Table.table_style_creation.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Table.table_style_creation.docx")
        table = doc.first_section.body.tables[0]

        self.assertTrue(table.bidi)
        self.assertEqual(5.0, table.cell_spacing)
        self.assertEqual("MyTableStyle1", table.style_name)
        self.assertEqual(20.0, table_style.bottom_padding)
        self.assertEqual(5.0, table_style.left_padding)
        self.assertEqual(10.0, table_style.right_padding)
        self.assertEqual(20.0, table_style.top_padding)
        self.assertEqual(6, len([b for b in table.first_row.row_format.borders if b.color.to_argb() == drawing.Color.blue.to_argb()]))
        self.assertEqual(aw.tables.CellVerticalAlignment.CENTER, table_style.vertical_alignment)

        table_style = doc.styles.get_by_name("MyTableStyle1").as_table_style()

        self.assertTrue(table_style.allow_break_across_pages)
        self.assertTrue(table_style.bidi)
        self.assertEqual(5.0, table_style.cell_spacing)
        self.assertEqual(20.0, table_style.bottom_padding)
        self.assertEqual(5.0, table_style.left_padding)
        self.assertEqual(10.0, table_style.right_padding)
        self.assertEqual(20.0, table_style.top_padding)
        self.assertEqual(drawing.Color.antique_white.to_argb(), table_style.shading.background_pattern_color.to_argb())
        self.assertEqual(drawing.Color.blue.to_argb(), table_style.borders.color.to_argb())
        self.assertEqual(aw.LineStyle.DOT_DASH, table_style.borders.line_style)
        self.assertEqual(aw.tables.CellVerticalAlignment.CENTER, table_style.vertical_alignment)

    def test_set_table_alignment(self):

        #ExStart
        #ExFor:TableStyle.alignment
        #ExFor:TableStyle.left_indent
        #ExSummary:Shows how to set the position of a table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Below are two ways of aligning a table horizontally.
        # 1 -  Use the "alignment" property to align it to a location on the page, such as the center:
        table_style = doc.styles.add(aw.StyleType.TABLE, "MyTableStyle1").as_table_style()
        table_style.alignment = aw.tables.TableAlignment.CENTER
        table_style.borders.color = drawing.Color.blue
        table_style.borders.line_style = aw.LineStyle.SINGLE

        # Insert a table and apply the style we created to it.
        table = builder.start_table()
        builder.insert_cell()
        builder.write("Aligned to the center of the page")
        builder.end_table()
        table.preferred_width = aw.tables.PreferredWidth.from_points(300)

        table.style = table_style

        # 2 -  Use the "left_indent" to specify an indent from the left margin of the page:
        table_style = doc.styles.add(aw.StyleType.TABLE, "MyTableStyle2").as_table_style()
        table_style.left_indent = 55
        table_style.borders.color = drawing.Color.green
        table_style.borders.line_style = aw.LineStyle.SINGLE

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Aligned according to left indent")
        builder.end_table()
        table.preferred_width = aw.tables.PreferredWidth.from_points(300)

        table.style = table_style

        doc.save(ARTIFACTS_DIR + "Table.set_table_alignment.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Table.set_table_alignment.docx")

        table_style = doc.styles.get_by_name("MyTableStyle1").as_table_style()

        self.assertEqual(aw.tables.TableAlignment.CENTER, table_style.alignment)
        self.assertEqual(table_style, doc.first_section.body.tables[0].style)

        table_style = doc.styles.get_by_name("MyTableStyle2").as_table_style()

        self.assertEqual(55.0, table_style.left_indent)
        self.assertEqual(table_style, doc.get_child(aw.NodeType.TABLE, 1, True).as_table().style)

    def test_conditional_styles(self):

        #ExStart
        #ExFor:ConditionalStyle
        #ExFor:ConditionalStyle.shading
        #ExFor:ConditionalStyle.borders
        #ExFor:ConditionalStyle.paragraph_format
        #ExFor:ConditionalStyle.bottom_padding
        #ExFor:ConditionalStyle.left_padding
        #ExFor:ConditionalStyle.right_padding
        #ExFor:ConditionalStyle.top_padding
        #ExFor:ConditionalStyle.font
        #ExFor:ConditionalStyle.type
        #ExFor:ConditionalStyleCollection.__iter__
        #ExFor:ConditionalStyleCollection.first_row
        #ExFor:ConditionalStyleCollection.last_row
        #ExFor:ConditionalStyleCollection.last_column
        #ExFor:ConditionalStyleCollection.count
        #ExFor:ConditionalStyleCollection
        #ExFor:ConditionalStyleCollection.bottom_left_cell
        #ExFor:ConditionalStyleCollection.bottom_right_cell
        #ExFor:ConditionalStyleCollection.even_column_banding
        #ExFor:ConditionalStyleCollection.even_row_banding
        #ExFor:ConditionalStyleCollection.first_column
        #ExFor:ConditionalStyleCollection.__getitem__(ConditionalStyleType)
        #ExFor:ConditionalStyleCollection.__getitem__(int)
        #ExFor:ConditionalStyleCollection.odd_column_banding
        #ExFor:ConditionalStyleCollection.odd_row_banding
        #ExFor:ConditionalStyleCollection.top_left_cell
        #ExFor:ConditionalStyleCollection.top_right_cell
        #ExFor:ConditionalStyleType
        #ExFor:TableStyle.conditional_styles
        #ExSummary:Shows how to work with certain area styles of a table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Cell 1")
        builder.insert_cell()
        builder.write("Cell 2")
        builder.end_row()
        builder.insert_cell()
        builder.write("Cell 3")
        builder.insert_cell()
        builder.write("Cell 4")
        builder.end_table()

        # Create a custom table style.
        table_style = doc.styles.add(aw.StyleType.TABLE, "MyTableStyle1").as_table_style()

        # Conditional styles are formatting changes that affect only some of the table's cells
        # based on a predicate, such as the cells being in the last row.
        # Below are three ways of accessing a table style's conditional styles from the "conditional_styles" collection.
        # 1 -  By style type:
        table_style.conditional_styles[aw.ConditionalStyleType.FIRST_ROW].shading.background_pattern_color = drawing.Color.alice_blue

        # 2 -  By index:
        table_style.conditional_styles[0].borders.color = drawing.Color.black
        table_style.conditional_styles[0].borders.line_style = aw.LineStyle.DOT_DASH
        self.assertEqual(aw.ConditionalStyleType.FIRST_ROW, table_style.conditional_styles[0].type)

        # 3 -  As a property:
        table_style.conditional_styles.first_row.paragraph_format.alignment = aw.ParagraphAlignment.CENTER

        # Apply padding and text formatting to conditional styles.
        table_style.conditional_styles.last_row.bottom_padding = 10
        table_style.conditional_styles.last_row.left_padding = 10
        table_style.conditional_styles.last_row.right_padding = 10
        table_style.conditional_styles.last_row.top_padding = 10
        table_style.conditional_styles.last_column.font.bold = True

        # List all possible style conditions.
        for conditional_style in table_style.conditional_styles:
            if conditional_style is not None:
                print(conditional_style.type)

        # Apply the custom style, which contains all conditional styles, to the table.
        table.style = table_style

        # Our style applies some conditional styles by default.
        self.assertEqual(aw.tables.TableStyleOptions.FIRST_ROW | aw.tables.TableStyleOptions.FIRST_COLUMN | aw.tables.TableStyleOptions.ROW_BANDS,
            table.style_options)

        # We will need to enable all other styles ourselves via the "style_options" property.
        table.style_options = table.style_options | aw.tables.TableStyleOptions.LAST_ROW | aw.tables.TableStyleOptions.LAST_COLUMN

        doc.save(ARTIFACTS_DIR + "Table.conditional_styles.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Table.conditional_styles.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(aw.tables.TableStyleOptions.DEFAULT | aw.tables.TableStyleOptions.LAST_ROW | aw.tables.TableStyleOptions.LAST_COLUMN, table.style_options)
        conditional_styles = doc.styles.get_by_name("MyTableStyle1").as_table_style().conditional_styles

        self.assertEqual(aw.ConditionalStyleType.FIRST_ROW, conditional_styles[0].type)
        self.assertEqual(drawing.Color.alice_blue.to_argb(), conditional_styles[0].shading.background_pattern_color.to_argb())
        self.assertEqual(drawing.Color.black.to_argb(), conditional_styles[0].borders.color.to_argb())
        self.assertEqual(aw.LineStyle.DOT_DASH, conditional_styles[0].borders.line_style)
        self.assertEqual(aw.ParagraphAlignment.CENTER, conditional_styles[0].paragraph_format.alignment)

        self.assertEqual(aw.ConditionalStyleType.LAST_ROW, conditional_styles[2].type)
        self.assertEqual(10.0, conditional_styles[2].bottom_padding)
        self.assertEqual(10.0, conditional_styles[2].left_padding)
        self.assertEqual(10.0, conditional_styles[2].right_padding)
        self.assertEqual(10.0, conditional_styles[2].top_padding)

        self.assertEqual(aw.ConditionalStyleType.LAST_COLUMN, conditional_styles[3].type)
        self.assertTrue(conditional_styles[3].font.bold)

    def test_clear_table_style_formatting(self):

        #ExStart
        #ExFor:ConditionalStyle.clear_formatting
        #ExFor:ConditionalStyleCollection.clear_formatting
        #ExSummary:Shows how to reset conditional table styles.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("First row")
        builder.end_row()
        builder.insert_cell()
        builder.write("Last row")
        builder.end_table()

        table_style = doc.styles.add(aw.StyleType.TABLE, "MyTableStyle1").as_table_style()
        table.style = table_style

        # Set the table style to color the borders of the first row of the table in red.
        table_style.conditional_styles.first_row.borders.color = drawing.Color.red

        # Set the table style to color the borders of the last row of the table in blue.
        table_style.conditional_styles.last_row.borders.color = drawing.Color.blue

        # Below are two ways of using the "clear_formatting" method to clear the conditional styles.
        # 1 -  Clear the conditional styles for a specific part of a table:
        table_style.conditional_styles[0].clear_formatting()

        self.assertEqual(drawing.Color.empty(), table_style.conditional_styles.first_row.borders.color)

        # 2 -  Clear the conditional styles for the entire table:
        table_style.conditional_styles.clear_formatting()

        self.assertTrue(all(s.borders.color == drawing.Color.empty()
                            for s in table_style.conditional_styles))
        #ExEnd

    def test_alternating_row_styles(self):

        #ExStart
        #ExFor:TableStyle.column_stripe
        #ExFor:TableStyle.row_stripe
        #ExSummary:Shows how to create conditional table styles that alternate between rows.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # We can configure a conditional style of a table to apply a different color to the row/column,
        # based on whether the row/column is even or odd, creating an alternating color pattern.
        # We can also apply a number n to the row/column banding,
        # meaning that the color alternates after every n rows/columns instead of one.
        # Create a table where single columns and rows will band the columns will banded in threes.
        table = builder.start_table()
        for i in range(15):
            for j in range(4):
                builder.insert_cell()
                builder.writeln(f"{'Even' if j % 2 == 0 else 'Odd'} column.")
                builder.write(f"Row banding {'start' if i % 3 == 0 else 'continuation'}.")
            builder.end_row()
        builder.end_table()

        # Apply a line style to all the borders of the table.
        table_style = doc.styles.add(aw.StyleType.TABLE, "MyTableStyle1").as_table_style()
        table_style.borders.color = drawing.Color.black
        table_style.borders.line_style = aw.LineStyle.DOUBLE

        # Set the two colors, which will alternate over every 3 rows.
        table_style.row_stripe = 3
        table_style.conditional_styles[aw.ConditionalStyleType.ODD_ROW_BANDING].shading.background_pattern_color = drawing.Color.light_blue
        table_style.conditional_styles[aw.ConditionalStyleType.EVEN_ROW_BANDING].shading.background_pattern_color = drawing.Color.light_cyan

        # Set a color to apply to every even column, which will override any custom row coloring.
        table_style.column_stripe = 1
        table_style.conditional_styles[aw.ConditionalStyleType.EVEN_COLUMN_BANDING].shading.background_pattern_color = drawing.Color.light_salmon

        table.style = table_style

        # The "style_options" property enables row banding by default.
        self.assertEqual(aw.tables.TableStyleOptions.FIRST_ROW | aw.tables.TableStyleOptions.FIRST_COLUMN | aw.tables.TableStyleOptions.ROW_BANDS,
            table.style_options)

        # Use the "style_options" property also to enable column banding.
        table.style_options = table.style_options | aw.tables.TableStyleOptions.COLUMN_BANDS

        doc.save(ARTIFACTS_DIR + "Table.alternating_row_styles.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Table.alternating_row_styles.docx")
        table = doc.first_section.body.tables[0]
        table_style = doc.styles.get_by_name("MyTableStyle1").as_table_style()

        self.assertEqual(table_style, table.style)
        self.assertEqual(table.style_options | aw.tables.TableStyleOptions.COLUMN_BANDS, table.style_options)

        self.assertEqual(drawing.Color.black.to_argb(), table_style.borders.color.to_argb())
        self.assertEqual(aw.LineStyle.DOUBLE, table_style.borders.line_style)
        self.assertEqual(3, table_style.row_stripe)
        self.assertEqual(drawing.Color.light_blue.to_argb(), table_style.conditional_styles[aw.ConditionalStyleType.ODD_ROW_BANDING].shading.background_pattern_color.to_argb())
        self.assertEqual(drawing.Color.light_cyan.to_argb(), table_style.conditional_styles[aw.ConditionalStyleType.EVEN_ROW_BANDING].shading.background_pattern_color.to_argb())
        self.assertEqual(1, table_style.column_stripe)
        self.assertEqual(drawing.Color.light_salmon.to_argb(), table_style.conditional_styles[aw.ConditionalStyleType.EVEN_COLUMN_BANDING].shading.background_pattern_color.to_argb())

    def test_convert_to_horizontally_merged_cells(self):

        #ExStart
        #ExFor:Table.convert_to_horizontally_merged_cells
        #ExSummary:Shows how to convert cells horizontally merged by width to cells merged by CellFormat.horizontal_merge.
        doc = aw.Document(MY_DIR + "Table with merged cells.docx")

        # Microsoft Word does not write merge flags anymore, defining merged cells by width instead.
        # Aspose.Words by default define only 5 cells in a row, and none of them have the horizontal merge flag,
        # even though there were 7 cells in the row before the horizontal merging took place.
        table = doc.first_section.body.tables[0]
        row = table.rows[0]

        self.assertEqual(5, row.cells.count)
        self.assertTrue(all(c.as_cell().cell_format.horizontal_merge == aw.tables.CellMerge.NONE
                            for c in row.cells))

        # Use the "convert_to_horizontally_merged_cells" method to convert cells horizontally merged
        # by its width to the cell horizontally merged by flags.
        # Now, we have 7 cells, and some of them have horizontal merge values.
        table.convert_to_horizontally_merged_cells()
        row = table.rows[0]

        self.assertEqual(7, row.cells.count)

        self.assertEqual(aw.tables.CellMerge.NONE, row.cells[0].cell_format.horizontal_merge)
        self.assertEqual(aw.tables.CellMerge.FIRST, row.cells[1].cell_format.horizontal_merge)
        self.assertEqual(aw.tables.CellMerge.PREVIOUS, row.cells[2].cell_format.horizontal_merge)
        self.assertEqual(aw.tables.CellMerge.NONE, row.cells[3].cell_format.horizontal_merge)
        self.assertEqual(aw.tables.CellMerge.FIRST, row.cells[4].cell_format.horizontal_merge)
        self.assertEqual(aw.tables.CellMerge.PREVIOUS, row.cells[5].cell_format.horizontal_merge)
        self.assertEqual(aw.tables.CellMerge.NONE, row.cells[6].cell_format.horizontal_merge)
        #ExEnd
