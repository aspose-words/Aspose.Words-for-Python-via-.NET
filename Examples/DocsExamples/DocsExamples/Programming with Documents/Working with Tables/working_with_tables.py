import unittest
import os
import sys
import math

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw
import aspose.pydrawing as drawing

class WorkingWithTables(docs_base.DocsExamplesBase):

    def test_remove_column(self):

        #ExStart:RemoveColumn
        doc = aw.Document(docs_base.my_dir + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 1, True).as_table()

        column = self.Column(table, 2)
        column.remove()
        #ExEnd:RemoveColumn


    def test_insert_blank_column(self):

        #ExStart:InsertBlankColumn
        doc = aw.Document(docs_base.my_dir + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        #ExStart:GetPlainText
        column = self.Column(table, 0)
        # Print the plain text of the column to the screen.
        print(column.to_txt())
        #ExEnd:GetPlainText

        # Create a new column to the left of this column.
        # This is the same as using the "Insert Column Before" command in Microsoft Word.
        new_column = column.insert_column_before()

        for cell in new_column.get_column_cells():
            cell.first_paragraph.append_child(aw.Run(doc, f"Column Text {new_column.index_of(cell)}"))
        #ExEnd:InsertBlankColumn


    #ExStart:ColumnClass
    # <summary>
    # Represents a facade object for a column of a table in a Microsoft Word document.
    # </summary>
    class Column:

        def __init__(self, table: aw.tables.Table, column_index: int):

            if table is None:
                raise ValueError("table")
            self.table = table
            self.column_index = column_index

        # <summary>
        # Returns the index of the given cell in the column.
        # </summary>
        def index_of(self, cell: aw.tables.Cell):
            return self.get_column_cells().index(cell)


        # <summary>
        # Inserts a brand new column before this column into the table.
        # </summary>
        def insert_column_before(self):

            column_cells = self.get_column_cells()

            if len(column_cells) == 0:
                raise ValueError("Column must not be empty")

            # Create a clone of this column.
            for cell in column_cells:
                cell.parent_row.insert_before(cell.clone(False), cell)

            # This is the new column.
            column = self.__class__(column_cells[0].parent_row.parent_table, self.column_index)

            # We want to make sure that the cells are all valid to work with (have at least one paragraph).
            for cell in column.get_column_cells():
                cell.ensure_minimum()

            # Increase the index which this column represents since there is now one extra column in front.
            self.column_index += 1

            return column


        # <summary>
        # Removes the column from the table.
        # </summary>
        def remove(self):

            for cell in self.get_column_cells():
                cell.remove()


        # <summary>
        # Returns the text of the column.
        # </summary>
        def to_txt(self):

            result = ""

            for cell in self.get_column_cells():
                result += cell.to_string(aw.SaveFormat.TEXT)

            return result


        # <summary>
        # Provides an up-to-date collection of cells which make up the column represented by this facade.
        # </summary>
        def get_column_cells(self):

            column_cells = []

            for row in self.table.rows:

                cell = row.as_row().cells[self.column_index]
                if cell is not None:
                    column_cells.append(cell)

            return column_cells


    #ExEnd:ColumnClass

    def test_auto_fit_table_to_contents(self):

        #ExStart:AutoFitTableToContents
        doc = aw.Document(docs_base.my_dir + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()
        table.auto_fit(aw.tables.AutoFitBehavior.AUTO_FIT_TO_CONTENTS)

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.auto_fit_table_to_contents.docx")
        #ExEnd:AutoFitTableToContents


    def test_auto_fit_table_to_fixed_column_widths(self):

        #ExStart:AutoFitTableToFixedColumnWidths
        doc = aw.Document(docs_base.my_dir + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()
        # Disable autofitting on this table.
        table.auto_fit(aw.tables.AutoFitBehavior.FIXED_COLUMN_WIDTHS)

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.auto_fit_table_to_fixed_column_widths.docx")
        #ExEnd:AutoFitTableToFixedColumnWidths


    def test_auto_fit_table_to_page_width(self):

        #ExStart:AutoFitTableToPageWidth
        doc = aw.Document(docs_base.my_dir + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()
        # Autofit the first table to the page width.
        table.auto_fit(aw.tables.AutoFitBehavior.AUTO_FIT_TO_WINDOW)

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.auto_fit_table_to_window.docx")
        #ExEnd:AutoFitTableToPageWidth


    #def test_build_table_from_data_table(self):

    #    #ExStart:BuildTableFromDataTable
    #    doc = aw.Document()
    #    # We can position where we want the table to be inserted and specify any extra formatting to the table.
    #    builder = aw.DocumentBuilder(doc)

    #    # We want to rotate the page landscape as we expect a wide table.
    #    doc.first_section.page_setup.orientation = Orientation.landscape

    #    DataSet ds = new DataSet()
    #    ds.read_xml(docs_base.my_dir + "List of people.xml")
    #    # Retrieve the data from our data source, which is stored as a DataTable.
    #    DataTable dataTable = ds.tables[0]

    #    # Build a table in the document from the data contained in the DataTable.
    #    Table table = ImportTableFromDataTable(builder, dataTable, True)

    #    # We can apply a table style as a very quick way to apply formatting to the entire table.
    #    table.style_identifier = StyleIdentifier.medium_list_2_accent_1
    #    table.style_options = TableStyleOptions.first_row | TableStyleOptions.row_bands | TableStyleOptions.last_column

    #    # For our table, we want to remove the heading for the image column.
    #    table.first_row.last_cell.remove_all_children()

    #    doc.save(docs_base.artifacts_dir + "WorkingWithTables.build_table_from_data_table.docx")
    #    #ExEnd:BuildTableFromDataTable


    ##ExStart:ImportTableFromDataTable
    ## <summary>
    ## Imports the content from the specified DataTable into a new Aspose.words Table object.
    ## The table is inserted at the document builder's current position and using the current builder's formatting if any is defined.
    ## </summary>
    #public Table ImportTableFromDataTable(DocumentBuilder builder, DataTable dataTable,
    #    bool importColumnHeadings)

    #    Table table = builder.start_table()

    #    # Check if the columns' names from the data source are to be included in a header row.
    #    if (importColumnHeadings)

    #        # Store the original values of these properties before changing them.
    #        bool boldValue = builder.font.bold
    #        ParagraphAlignment paragraphAlignmentValue = builder.paragraph_format.alignment

    #        # Format the heading row with the appropriate properties.
    #        builder.font.bold = True
    #        builder.paragraph_format.alignment = ParagraphAlignment.center

    #        # Create a new row and insert the name of each column into the first row of the table.
    #        foreach (DataColumn column in dataTable.columns)

    #            builder.insert_cell()
    #            builder.writeln(column.column_name)


    #        builder.end_row()

    #        # Restore the original formatting.
    #        builder.font.bold = boldValue
    #        builder.paragraph_format.alignment = paragraphAlignmentValue


    #    foreach (DataRow dataRow in dataTable.rows)

    #        foreach (object item in dataRow.item_array)

    #            # Insert a new cell for each object.
    #            builder.insert_cell()

    #            switch (item.get_type().name)

    #                case "DateTime":
    #                    # Define a custom format for dates and times.
    #                    DateTime dateTime = (DateTime) item
    #                    builder.write(dateTime.to_string("MMMM d, yyyy"))
    #                    break
    #                default:
    #                    # By default any other item will be inserted as text.
    #                    builder.write(item.to_string())
    #                    break


    #        # After we insert all the data from the current record, we can end the table row.
    #        builder.end_row()


    #    # We have finished inserting all the data from the DataTable, we can end the table.
    #    builder.end_table()

    #    return table

    ##ExEnd:ImportTableFromDataTable

    def test_clone_complete_table(self):

        #ExStart:CloneCompleteTable
        doc = aw.Document(docs_base.my_dir + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        # Clone the table and insert it into the document after the original.
        table_clone = table.clone(True).as_table()
        table.parent_node.insert_after(table_clone, table)

        # Insert an empty paragraph between the two tables,
        # or else they will be combined into one upon saving this has to do with document validation.
        table.parent_node.insert_after(aw.Paragraph(doc), table)

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.clone_complete_table.docx")
        #ExEnd:CloneCompleteTable


    def test_clone_last_row(self):

        #ExStart:CloneLastRow
        doc = aw.Document(docs_base.my_dir + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        cloned_row = table.last_row.clone(True).as_row()
        # Remove all content from the cloned row's cells. This makes the row ready for new content to be inserted into.
        for cell in cloned_row.cells:
            cell = cell.as_cell()
            cell.remove_all_children()

        table.append_child(cloned_row)

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.clone_last_row.docx")
        #ExEnd:CloneLastRow


    def test_finding_index(self):

        doc = aw.Document(docs_base.my_dir + "Tables.docx")

        #ExStart:RetrieveTableIndex
        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        all_tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
        table_index = all_tables.index_of(table)
        #ExEnd:RetrieveTableIndex
        print(f"\nTable index is {table_index}")

        #ExStart:RetrieveRowIndex
        row_index = table.index_of(table.last_row)
        #ExEnd:RetrieveRowIndex
        print(f"\nRow index is {row_index}")

        row = table.last_row
        #ExStart:RetrieveCellIndex
        cell_index = row.index_of(row.cells[4])
        #ExEnd:RetrieveCellIndex
        print(f"\nCell index is {cell_index}")


    def test_insert_table_directly(self):

        #ExStart:InsertTableDirectly
        doc = aw.Document()

        # We start by creating the table object. Note that we must pass the document object
        # to the constructor of each node. This is because every node we create must belong
        # to some document.
        table = aw.tables.Table(doc)
        doc.first_section.body.append_child(table)

        # Here we could call EnsureMinimum to create the rows and cells for us. This method is used
        # to ensure that the specified node is valid. In this case, a valid table should have at least one Row and one cell.

        # Instead, we will handle creating the row and table ourselves.
        # This would be the best way to do this if we were creating a table inside an algorithm.
        row = aw.tables.Row(doc)
        row.row_format.allow_break_across_pages = True
        table.append_child(row)

        # We can now apply any auto fit settings.
        table.auto_fit(aw.tables.AutoFitBehavior.FIXED_COLUMN_WIDTHS)

        cell = aw.tables.Cell(doc)
        cell.cell_format.shading.background_pattern_color = drawing.Color.light_blue
        cell.cell_format.width = 80
        cell.append_child(aw.Paragraph(doc))
        cell.first_paragraph.append_child(aw.Run(doc, "Row 1, Cell 1 Text"))

        row.append_child(cell)

        # We would then repeat the process for the other cells and rows in the table.
        # We can also speed things up by cloning existing cells and rows.
        row.append_child(cell.clone(False))
        row.last_cell.append_child(aw.Paragraph(doc))
        row.last_cell.first_paragraph.append_child(aw.Run(doc, "Row 1, Cell 2 Text"))

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.insert_table_directly.docx")
        #ExEnd:InsertTableDirectly


    def test_insert_table_from_html(self):

        #ExStart:InsertTableFromHtml
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Note that AutoFitSettings does not apply to tables inserted from HTML.
        builder.insert_html("<table>" +
                            "<tr>" +
                            "<td>Row 1, Cell 1</td>" +
                            "<td>Row 1, Cell 2</td>" +
                            "</tr>" +
                            "<tr>" +
                            "<td>Row 2, Cell 2</td>" +
                            "<td>Row 2, Cell 2</td>" +
                            "</tr>" +
                            "</table>")

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.insert_table_from_html.docx")
        #ExEnd:InsertTableFromHtml


    def test_create_simple_table(self):

        #ExStart:CreateSimpleTable
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Start building the table.
        builder.start_table()
        builder.insert_cell()
        builder.write("Row 1, Cell 1 Content.")

        # Build the second cell.
        builder.insert_cell()
        builder.write("Row 1, Cell 2 Content.")

        # Call the following method to end the row and start a new row.
        builder.end_row()

        # Build the first cell of the second row.
        builder.insert_cell()
        builder.write("Row 2, Cell 1 Content")

        # Build the second cell.
        builder.insert_cell()
        builder.write("Row 2, Cell 2 Content.")
        builder.end_row()

        # Signal that we have finished building the table.
        builder.end_table()

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.create_simple_table.docx")
        #ExEnd:CreateSimpleTable


    def test_formatted_table(self):

        #ExStart:FormattedTable
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()

        # Table wide formatting must be applied after at least one row is present in the table.
        table.left_indent = 20.0

        # Set height and define the height rule for the header row.
        builder.row_format.height = 40.0
        builder.row_format.height_rule = aw.HeightRule.AT_LEAST

        builder.cell_format.shading.background_pattern_color = drawing.Color.from_argb(198, 217, 241)
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        builder.font.size = 16
        builder.font.name = "Arial"
        builder.font.bold = True

        builder.cell_format.width = 100.0
        builder.write("Header Row,\n Cell 1")

        # We don't need to specify this cell's width because it's inherited from the previous cell.
        builder.insert_cell()
        builder.write("Header Row,\n Cell 2")

        builder.insert_cell()
        builder.cell_format.width = 200.0
        builder.write("Header Row,\n Cell 3")
        builder.end_row()

        builder.cell_format.shading.background_pattern_color = drawing.Color.white
        builder.cell_format.width = 100.0
        builder.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.CENTER

        # Reset height and define a different height rule for table body.
        builder.row_format.height = 30.0
        builder.row_format.height_rule = aw.HeightRule.AUTO
        builder.insert_cell()

        # Reset font formatting.
        builder.font.size = 12
        builder.font.bold = False

        builder.write("Row 1, Cell 1 Content")
        builder.insert_cell()
        builder.write("Row 1, Cell 2 Content")

        builder.insert_cell()
        builder.cell_format.width = 200.0
        builder.write("Row 1, Cell 3 Content")
        builder.end_row()

        builder.insert_cell()
        builder.cell_format.width = 100.0
        builder.write("Row 2, Cell 1 Content")

        builder.insert_cell()
        builder.write("Row 2, Cell 2 Content")

        builder.insert_cell()
        builder.cell_format.width = 200.0
        builder.write("Row 2, Cell 3 Content.")
        builder.end_row()
        builder.end_table()

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.formatted_table.docx")
        #ExEnd:FormattedTable


    def test_nested_table(self):

        #ExStart:NestedTable
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        cell = builder.insert_cell()
        builder.writeln("Outer Table Cell 1")

        builder.insert_cell()
        builder.writeln("Outer Table Cell 2")

        # This call is important to create a nested table within the first table.
        # Without this call, the cells inserted below will be appended to the outer table.
        builder.end_table()

        # Move to the first cell of the outer table.
        builder.move_to(cell.first_paragraph)

        # Build the inner table.
        builder.insert_cell()
        builder.writeln("Inner Table Cell 1")
        builder.insert_cell()
        builder.writeln("Inner Table Cell 2")
        builder.end_table()

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.nested_table.docx")
        #ExEnd:NestedTable


    def test_combine_rows(self):

        #ExStart:CombineRows
        doc = aw.Document(docs_base.my_dir + "Tables.docx")

        # The rows from the second table will be appended to the end of the first table.
        first_table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()
        second_table = doc.get_child(aw.NodeType.TABLE, 1, True).as_table()

        # Append all rows from the current table to the next tables
        # with different cell count and widths can be joined into one table.
        while second_table.has_child_nodes:
            first_table.rows.add(second_table.first_row)

        second_table.remove()

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.combine_rows.docx")
        #ExEnd:CombineRows


    def test_split_table(self):

        #ExStart:SplitTable
        doc = aw.Document(docs_base.my_dir + "Tables.docx")

        first_table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        # We will split the table at the third row (inclusive).
        row = first_table.rows[2]

        # Create a new container for the split table.
        table = first_table.clone(False).as_table()

        # Insert the container after the original.
        first_table.parent_node.insert_after(table, first_table)

        # Add a buffer paragraph to ensure the tables stay apart.
        first_table.parent_node.insert_after(aw.Paragraph(doc), first_table)


        while True:
            current_row = first_table.last_row
            table.prepend_child(current_row)
            if current_row == row:
                break


        doc.save(docs_base.artifacts_dir + "WorkingWithTables.split_table.docx")
        #ExEnd:SplitTable


    def test_row_format_disable_break_across_pages(self):

        #ExStart:RowFormatDisableBreakAcrossPages
        doc = aw.Document(docs_base.my_dir + "Table spanning two pages.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        # Disable breaking across pages for all rows in the table.
        for row in table.rows:
            row.as_row().row_format.allow_break_across_pages = False

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.row_format_disable_break_across_pages.docx")
        #ExEnd:RowFormatDisableBreakAcrossPages


    def test_keep_table_together(self):

        #ExStart:KeepTableTogether
        doc = aw.Document(docs_base.my_dir + "Table spanning two pages.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        # We need to enable KeepWithNext for every paragraph in the table to keep it from breaking across a page,
        # except for the last paragraphs in the last row of the table.
        for cell in table.get_child_nodes(aw.NodeType.CELL, True):
            cell = cell.as_cell()
            cell.ensure_minimum()

            for para in cell.paragraphs:
                para = para.as_paragraph()
                if not (cell.parent_row.is_last_row and para.is_end_of_cell):
                    para.paragraph_format.keep_with_next = True


        doc.save(docs_base.artifacts_dir + "WorkingWithTables.keep_table_together.docx")
        #ExEnd:KeepTableTogether


    def test_check_cells_merged(self):

        #ExStart:CheckCellsMerged
        doc = aw.Document(docs_base.my_dir + "Table with merged cells.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        for row in table.rows:
            for cell in row.as_row().cells:
                print(self.print_cell_merge_type(cell.as_cell()))


        #ExEnd:CheckCellsMerged


    #ExStart:PrintCellMergeType
    @staticmethod
    def print_cell_merge_type(cell: aw.tables.Cell):

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

    #ExEnd:PrintCellMergeType

    def test_vertical_merge(self):

        #ExStart:VerticalMerge
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_cell()
        builder.cell_format.vertical_merge = aw.tables.CellMerge.FIRST
        builder.write("Text in merged cells.")

        builder.insert_cell()
        builder.cell_format.vertical_merge = aw.tables.CellMerge.NONE
        builder.write("Text in one cell")
        builder.end_row()

        builder.insert_cell()
        # This cell is vertically merged to the cell above and should be empty.
        builder.cell_format.vertical_merge = aw.tables.CellMerge.PREVIOUS

        builder.insert_cell()
        builder.cell_format.vertical_merge = aw.tables.CellMerge.NONE
        builder.write("Text in another cell")
        builder.end_row()
        builder.end_table()

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.vertical_merge.docx")
        #ExEnd:VerticalMerge


    def test_horizontal_merge(self):

        #ExStart:HorizontalMerge
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_cell()
        builder.cell_format.horizontal_merge = aw.tables.CellMerge.FIRST
        builder.write("Text in merged cells.")

        builder.insert_cell()
        # This cell is merged to the previous and should be empty.
        builder.cell_format.horizontal_merge = aw.tables.CellMerge.PREVIOUS
        builder.end_row()

        builder.insert_cell()
        builder.cell_format.horizontal_merge = aw.tables.CellMerge.NONE
        builder.write("Text in one cell.")

        builder.insert_cell()
        builder.write("Text in another cell.")
        builder.end_row()
        builder.end_table()

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.horizontal_merge.docx")
        #ExEnd:HorizontalMerge


    def test_merge_cell_range(self):

        #ExStart:MergeCellRange
        doc = aw.Document(docs_base.my_dir + "Table with merged cells.docx")

        table = doc.first_section.body.tables[0]

        # We want to merge the range of cells found inbetween these two cells.
        cell_start_range = table.rows[0].cells[0]
        cell_end_range = table.rows[1].cells[1]

        # Merge all the cells between the two specified cells into one.
        self.merge_cells(cell_start_range, cell_end_range)

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.merge_cell_range.docx")
        #ExEnd:MergeCellRange


    def test_convert_to_horizontally_merged_cells(self):

        #ExStart:ConvertToHorizontallyMergedCells
        doc = aw.Document(docs_base.my_dir + "Table with merged cells.docx")

        table = doc.first_section.body.tables[0]
        # Now merged cells have appropriate merge flags.
        table.convert_to_horizontally_merged_cells()
        #ExEnd:ConvertToHorizontallyMergedCells


    #ExStart:MergeCells
    @staticmethod
    def merge_cells(start_cell: aw.tables.Cell, end_cell: aw.tables.Cell):

        parent_table = start_cell.parent_row.parent_table

        # Find the row and cell indices for the start and end cell.
        start_cell_pos = drawing.Point(start_cell.parent_row.index_of(start_cell), parent_table.index_of(start_cell.parent_row))
        end_cell_pos = drawing.Point(end_cell.parent_row.index_of(end_cell), parent_table.index_of(end_cell.parent_row))

        # Create a range of cells to be merged based on these indices.
        # Inverse each index if the end cell is before the start cell.
        merge_range = drawing.Rectangle(min(start_cell_pos.x, end_cell_pos.x),
            min(start_cell_pos.y, end_cell_pos.y),
            abs(end_cell_pos.x - start_cell_pos.x) + 1, abs(end_cell_pos.y - start_cell_pos.y) + 1)

        for row in parent_table.rows:
            row = row.as_row()
            for cell in row.cells:

                cell = cell.as_cell()
                current_pos = drawing.Point(row.index_of(cell), parent_table.index_of(row))

                # Check if the current cell is inside our merge range, then merge it.
                if merge_range.contains(current_pos):

                    cell.cell_format.horizontal_merge = aw.tables.CellMerge.FIRST if (current_pos.x == merge_range.x) else aw.tables.CellMerge.PREVIOUS

                    cell.cell_format.vertical_merge = aw.tables.CellMerge.FIRST if (current_pos.y == merge_range.y) else aw.tables.CellMerge.PREVIOUS


    #ExEnd:MergeCells


    def test_repeat_rows_on_subsequent_pages(self):

        #ExStart:RepeatRowsOnSubsequentPages
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_table()
        builder.row_format.heading_format = True
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        builder.cell_format.width = 100
        builder.insert_cell()
        builder.writeln("Heading row 1")
        builder.end_row()
        builder.insert_cell()
        builder.writeln("Heading row 2")
        builder.end_row()

        builder.cell_format.width = 50
        builder.paragraph_format.clear_formatting()

        for i in range(0, 50):

            builder.insert_cell()
            builder.row_format.heading_format = False
            builder.write("Column 1 Text")
            builder.insert_cell()
            builder.write("Column 2 Text")
            builder.end_row()


        doc.save(docs_base.artifacts_dir + "WorkingWithTables.repeat_rows_on_subsequent_pages.docx")
        #ExEnd:RepeatRowsOnSubsequentPages


    def test_auto_fit_to_page_width(self):

        #ExStart:AutoFitToPageWidth
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a table with a width that takes up half the page width.
        table = builder.start_table()

        builder.insert_cell()
        table.preferred_width = aw.tables.PreferredWidth.from_percent(50)
        builder.writeln("Cell #1")

        builder.insert_cell()
        builder.writeln("Cell #2")

        builder.insert_cell()
        builder.writeln("Cell #3")

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.auto_fit_to_page_width.docx")
        #ExEnd:AutoFitToPageWidth


    def test_preferred_width_settings(self):

        #ExStart:PreferredWidthSettings
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a table row made up of three cells which have different preferred widths.
        builder.start_table()

        # Insert an absolute sized cell.
        builder.insert_cell()
        builder.cell_format.preferred_width = aw.tables.PreferredWidth.from_points(40)
        builder.cell_format.shading.background_pattern_color = drawing.Color.light_yellow
        builder.writeln("Cell at 40 points width")

        # Insert a relative (percent) sized cell.
        builder.insert_cell()
        builder.cell_format.preferred_width = aw.tables.PreferredWidth.from_percent(20)
        builder.cell_format.shading.background_pattern_color = drawing.Color.light_blue
        builder.writeln("Cell at 20% width")

        # Insert a auto sized cell.
        builder.insert_cell()
        builder.cell_format.preferred_width = aw.tables.PreferredWidth.AUTO
        builder.cell_format.shading.background_pattern_color = drawing.Color.light_green
        builder.writeln(
            "Cell automatically sized. The size of this cell is calculated from the table preferred width.")
        builder.writeln("In this case the cell will fill up the rest of the available space.")

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.preferred_width_settings.docx")
        #ExEnd:PreferredWidthSettings


    def test_retrieve_preferred_width_type(self):

        #ExStart:RetrievePreferredWidthType
        doc = aw.Document(docs_base.my_dir + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()
        #ExStart:AllowAutoFit
        table.allow_auto_fit = True
        #ExEnd:AllowAutoFit

        first_cell = table.first_row.first_cell
        type = first_cell.cell_format.preferred_width.type
        value = first_cell.cell_format.preferred_width.value
        #ExEnd:RetrievePreferredWidthType


    def test_get_table_position(self):

        #ExStart:GetTablePosition
        doc = aw.Document(docs_base.my_dir + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        if table.text_wrapping == aw.tables.TextWrapping.AROUND:

            print(table.relative_horizontal_alignment)
            print(table.relative_vertical_alignment)

        else:

            print(table.alignment)

        #ExEnd:GetTablePosition


    def test_get_floating_table_position(self):

        #ExStart:GetFloatingTablePosition
        doc = aw.Document(docs_base.my_dir + "Table wrapped by text.docx")

        for  table in doc.first_section.body.tables:
            table = table.as_table()
            # If the table is floating type, then print its positioning properties.
            if table.text_wrapping == aw.tables.TextWrapping.AROUND:

                print(table.horizontal_anchor)
                print(table.vertical_anchor)
                print(table.absolute_horizontal_distance)
                print(table.absolute_vertical_distance)
                print(table.allow_overlap)
                print(table.absolute_horizontal_distance)
                print(table.relative_vertical_alignment)
                print("..............................")


        #ExEnd:GetFloatingTablePosition


    def test_floating_table_position(self):

        #ExStart:FloatingTablePosition
        doc = aw.Document(docs_base.my_dir + "Table wrapped by text.docx")

        table = doc.first_section.body.tables[0]
        table.absolute_horizontal_distance = 10
        table.relative_vertical_alignment = aw.drawing.VerticalAlignment.CENTER

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.floating_table_position.docx")
        #ExEnd:FloatingTablePosition


    def test_set_relative_horizontal_or_vertical_position(self):

        #ExStart:SetRelativeHorizontalOrVerticalPosition
        doc = aw.Document(docs_base.my_dir + "Table wrapped by text.docx")

        table = doc.first_section.body.tables[0]
        table.horizontal_anchor = aw.drawing.RelativeHorizontalPosition.COLUMN
        table.vertical_anchor = aw.drawing.RelativeVerticalPosition.PAGE

        doc.save(docs_base.artifacts_dir + "WorkingWithTables.set_floating_table_position.docx")
        #ExEnd:SetRelativeHorizontalOrVerticalPosition


if __name__ == '__main__':
    unittest.main()
