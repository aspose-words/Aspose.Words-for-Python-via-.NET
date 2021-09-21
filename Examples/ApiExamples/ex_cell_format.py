import unittest

import api_example_base as aeb
from document_helper import DocumentHelper

import aspose.words as aw

class ExCellFormat(aeb.ApiExampleBase):
    
    def test_vertical_merge(self) :
        
        #ExStart
        #ExFor:DocumentBuilder.end_row
        #ExFor:CellMerge
        #ExFor:CellFormat.vertical_merge
        #ExSummary:Shows how to merge table cells vertically.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a cell into the first column of the first row.
        # This cell will be the first in a range of vertically merged cells.
        builder.insert_cell()
        builder.cell_format.vertical_merge = aw.tables.CellMerge.FIRST
        builder.write("Text in merged cells.")

        # Insert a cell into the second column of the first row, then end the row.
        # Also, configure the builder to disable vertical merging in created cells.
        builder.insert_cell()
        builder.cell_format.vertical_merge = aw.tables.CellMerge.NONE
        builder.write("Text in unmerged cell.")
        builder.end_row()

        # Insert a cell into the first column of the second row. 
        # Instead of adding text contents, we will merge this cell with the first cell that we added directly above.
        builder.insert_cell()
        builder.cell_format.vertical_merge = aw.tables.CellMerge.PREVIOUS

        # Insert another independent cell in the second column of the second row.
        builder.insert_cell()
        builder.cell_format.vertical_merge = aw.tables.CellMerge.NONE
        builder.write("Text in unmerged cell.")
        builder.end_row()
        builder.end_table()

        doc.save(aeb.artifacts_dir + "CellFormat.vertical_merge.docx")
        #ExEnd

        doc = aw.Document(aeb.artifacts_dir + "CellFormat.vertical_merge.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(aw.tables.CellMerge.FIRST, table.rows[0].cells[0].cell_format.vertical_merge)
        self.assertEqual(aw.tables.CellMerge.PREVIOUS, table.rows[1].cells[0].cell_format.vertical_merge)
        self.assertEqual("Text in merged cells.", table.rows[0].cells[0].get_text().strip("\a"))
        self.assertNotEqual(table.rows[0].cells[0].get_text(), table.rows[1].cells[0].get_text())
        

    def test_horizontal_merge(self) :
        
        #ExStart
        #ExFor:CellMerge
        #ExFor:CellFormat.horizontal_merge
        #ExSummary:Shows how to merge table cells horizontally.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a cell into the first column of the first row.
        # This cell will be the first in a range of horizontally merged cells.
        builder.insert_cell()
        builder.cell_format.horizontal_merge = aw.tables.CellMerge.FIRST
        builder.write("Text in merged cells.")

        # Insert a cell into the second column of the first row. Instead of adding text contents,
        # we will merge this cell with the first cell that we added directly to the left.
        builder.insert_cell()
        builder.cell_format.horizontal_merge = aw.tables.CellMerge.PREVIOUS
        builder.end_row()

        # Insert two more unmerged cells to the second row.
        builder.cell_format.horizontal_merge = aw.tables.CellMerge.NONE
        builder.insert_cell()
        builder.write("Text in unmerged cell.")
        builder.insert_cell()
        builder.write("Text in unmerged cell.")
        builder.end_row()
        builder.end_table()

        doc.save(aeb.artifacts_dir + "CellFormat.horizontal_merge.docx")
        #ExEnd

        doc = aw.Document(aeb.artifacts_dir + "CellFormat.horizontal_merge.docx")
        table = doc.first_section.body.tables[0]

        self.assertEqual(1, table.rows[0].cells.count)
        self.assertEqual(aw.tables.CellMerge.NONE, table.rows[0].cells[0].cell_format.horizontal_merge)
        self.assertEqual("Text in merged cells.", table.rows[0].cells[0].get_text().strip('\a'))
        

    def test_padding(self) :
        
        #ExStart
        #ExFor:CellFormat.set_paddings
        #ExSummary:Shows how to pad the contents of a cell with whitespace.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Set a padding distance (in points) between the border and the text contents
        # of each table cell we create with the document builder. 
        builder.cell_format.set_paddings(5, 10, 40, 50)

        # Create a table with one cell whose contents will have whitespace padding.
        builder.start_table()
        builder.insert_cell()
        builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                        "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.")

        doc.save(aeb.artifacts_dir + "CellFormat.padding.docx")
        #ExEnd

        doc = aw.Document(aeb.artifacts_dir + "CellFormat.padding.docx")

        table = doc.first_section.body.tables[0]
        cell = table.rows[0].cells[0]

        self.assertEqual(5, cell.cell_format.left_padding)
        self.assertEqual(10, cell.cell_format.top_padding)
        self.assertEqual(40, cell.cell_format.right_padding)
        self.assertEqual(50, cell.cell_format.bottom_padding)
        
    
if __name__ == '__main__':
    unittest.main()    