# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

from typing import List

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExTableColumn(ApiExampleBase):

    class Column:
        """Represents a facade object for a column of a table in a Microsoft Word document."""

        def __init__(self, table: aw.tables.Table, column_index: int):
            if table is None:
                raise ValueError("table is <None>")

            self.table = table
            self.column_index = column_index

        @staticmethod
        def from_index(table: aw.tables.Table, column_index: int):
            """Returns a new column facade from the table and supplied zero-based index."""

            return ExTableColumn.Column(table, column_index)

        @property
        def cells(self) -> List[aw.tables.Cell]:
            """Returns the cells which make up the column."""

            return self.get_column_cells()

        def index_of(self, cell: aw.tables.Cell) -> int:
            """Returns the index of the given cell in the column."""

            return self.get_column_cells().index(cell)

        def insert_column_before(self) -> "Column":
            """Inserts a new column before this column into the table."""

            column_cells = self.cells

            if len(column_cells) == 0:
                raise ValueError("Column must not be empty")

            # Create a clone of this column
            for cell in column_cells:
                cell.parent_row.insert_before(cell.clone(False), cell)

            new_column = ExTableColumn.Column(column_cells[0].parent_row.parent_table, self.column_index)

            # We want to make sure that the cells are all valid to work with (have at least one paragraph).
            for cell in new_column.cells:
                cell.ensure_minimum()

            # Increment the index of this column represents since there is a new column before it.
            self.column_index += 1

            return new_column

        def remove(self):
            """Removes the column from the table."""

            for cell in self.cells:
                cell.remove()

        def to_txt(self) -> str:
            """Returns the text of the column."""

            txt = ""

            for cell in self.cells:
                txt += cell.to_string(aw.SaveFormat.TEXT)

            return txt

        def get_column_cells(self) -> List[aw.tables.Cell]:
            """Provides an up-to-date collection of cells which make up the column represented by this facade."""

            column_cells = []

            for row in self.table.rows:
                row = row.as_row()

                cell = row.cells[self.column_index]
                if cell is not None:
                    column_cells.append(cell)

            return column_cells

    def test_remove_column_from_table(self):

        doc = aw.Document(MY_DIR + "Tables.docx")
        table = doc.get_child(aw.NodeType.TABLE, 1, True).as_table()

        column = ExTableColumn.Column.from_index(table, 2)
        column.remove()

        doc.save(ARTIFACTS_DIR + "TableColumn.remove_column.doc")

        self.assertEqual(16, table.get_child_nodes(aw.NodeType.CELL, True).count)
        self.assertEqual("Cell 7 contents", table.rows[2].cells[2].to_string(aw.SaveFormat.TEXT).strip())
        self.assertEqual("Cell 11 contents", table.last_row.cells[2].to_string(aw.SaveFormat.TEXT).strip())

    def test_insert(self):

        doc = aw.Document(MY_DIR + "Tables.docx")
        table = doc.get_child(aw.NodeType.TABLE, 1, True).as_table()

        column = ExTableColumn.Column.from_index(table, 1)

        # Create a new column to the left of this column.
        # This is the same as using the "Insert Column Before" command in Microsoft Word.
        new_column = column.insert_column_before()

        # Add some text to each cell in the column.
        for cell in new_column.cells:
            cell.first_paragraph.append_child(aw.Run(doc, "Column Text " + str(new_column.index_of(cell))))

        doc.save(ARTIFACTS_DIR + "TableColumn.insert.doc")

        self.assertEqual(24, table.get_child_nodes(aw.NodeType.CELL, True).count)
        self.assertEqual("Column Text 0", table.first_row.cells[1].to_string(aw.SaveFormat.TEXT).strip())
        self.assertEqual("Column Text 3", table.last_row.cells[1].to_string(aw.SaveFormat.TEXT).strip())

    def test_table_column_to_txt(self):

        doc = aw.Document(MY_DIR + "Tables.docx")
        table = doc.get_child(aw.NodeType.TABLE, 1, True).as_table()

        column = ExTableColumn.Column.from_index(table, 0)
        print(column.to_txt())

        self.assertEqual("\rRow 1\rRow 2\rRow 3\r", column.to_txt())
