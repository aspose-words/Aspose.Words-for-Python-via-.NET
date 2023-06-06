# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import io
import unittest

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR

@unittest.skip("type 'aspose.words.DocumentVisitor' is not an acceptable base type ")
class ExDocumentVisitor(ApiExampleBase):

    #ExStart
    #ExFor:Document.accept(DocumentVisitor)
    #ExFor:Body.accept(DocumentVisitor)
    #ExFor:SubDocument.accept(DocumentVisitor)
    #ExFor:DocumentVisitor
    #ExFor:DocumentVisitor.visit_run(Run)
    #ExFor:DocumentVisitor.visit_document_end(Document)
    #ExFor:DocumentVisitor.visit_document_start(Document)
    #ExFor:DocumentVisitor.visit_section_end(Section)
    #ExFor:DocumentVisitor.visit_section_start(Section)
    #ExFor:DocumentVisitor.visit_body_start(Body)
    #ExFor:DocumentVisitor.visit_body_end(Body)
    #ExFor:DocumentVisitor.visit_paragraph_start(Paragraph)
    #ExFor:DocumentVisitor.visit_paragraph_end(Paragraph)
    #ExFor:DocumentVisitor.visit_sub_document(SubDocument)
    #ExSummary:Shows how to use a document visitor to print a document's node structure.
    def test_doc_structure_to_text(self):

        doc = aw.Document(MY_DIR + "DocumentVisitor-compatible features.docx")
        visitor = ExDocumentVisitor.DocStructurePrinter()

        # When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        # and then traverses all the node's children in a depth-first manner.
        # The visitor can read and modify each visited node.
        doc.accept(visitor)

        print(visitor.get_text())
        self._test_doc_structure_to_text(visitor) #ExSkip


    class DocStructurePrinter(aw.DocumentVisitor):
        """Traverses a node's tree of child nodes.
        Creates a map of this tree in the form of a string."""

        def __init__(self):

            aw.DocumentVisitor.__init__(self)

            self.accepting_node_child_tree = io.StringIO()
            self.doc_traversal_depth = 0

        def get_text(self):

            return self.accepting_node_child_tree.getvalue()

        def visit_document_start(self, doc: aw.Document) -> aw.VisitorAction:
            """Called when a Document node is encountered."""

            child_node_count = doc.get_child_nodes(aw.NodeType.ANY, True).count

            self._indent_and_append_line("[Document start] Child nodes: " + child_node_count)
            self.doc_traversal_depth += 1

            # Allow the visitor to continue visiting other nodes.
            return aw.VisitorAction.CONTINUE

        def visit_document_end(self, doc: aw.Document) -> aw.VisitorAction:
            """Called after all the child nodes of a Document node have been visited."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[Document end]")

            return aw.VisitorAction.CONTINUE

        def visit_section_start(self, section: aw.Section) -> aw.VisitorAction:
            """Called when a Section node is encountered in the document."""

            # Get the index of our section within the document.
            doc_sections = section.document.get_child_nodes(aw.NodeType.SECTION, False)
            section_index = doc_sections.index_of(section)

            self._indent_and_append_line("[Section start] Section index: " + section_index)
            self.doc_traversal_depth += 1

            return aw.VisitorAction.CONTINUE

        def visit_section_end(self, section: aw.Section) -> aw.VisitorAction:
            """Called after all the child nodes of a Section node have been visited."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[Section end]")

            return aw.VisitorAction.CONTINUE

        def visit_body_start(self, body: aw.Body) -> aw.VisitorAction:
            """Called when a Body node is encountered in the document."""

            paragraph_count = body.paragraphs.count
            self._indent_and_append_line("[Body start] Paragraphs: " + paragraph_count)
            self.doc_traversal_depth += 1

            return aw.VisitorAction.CONTINUE

        def visit_body_end(self, body: aw.Body) -> aw.VisitorAction:
            """Called after all the child nodes of a Body node have been visited."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[Body end]")

            return aw.VisitorAction.CONTINUE

        def visit_paragraph_start(self, paragraph: aw.Paragraph) -> aw.VisitorAction:
            """Called when a Paragraph node is encountered in the document."""

            self._indent_and_append_line("[Paragraph start]")
            self.doc_traversal_depth += 1

            return aw.VisitorAction.CONTINUE

        def visit_paragraph_end(self, paragraph: aw.Paragraph) -> aw.VisitorAction:
            """Called after all the child nodes of a Paragraph node have been visited."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[Paragraph end]")

            return aw.VisitorAction.CONTINUE

        def visit_run(self, run: aw.Run) -> aw.VisitorAction:
            """Called when a Run node is encountered in the document."""

            self._indent_and_append_line("[Run] \"" + run.get_text() + "\"")

            return aw.VisitorAction.CONTINUE

        def visit_sub_document(self, sub_document: aw.SubDocument) -> aw.VisitorAction:
            """Called when a SubDocument node is encountered in the document."""

            self._indent_and_append_line("[SubDocument]")

            return aw.VisitorAction.CONTINUE

        def _indent_and_append_line(self, text: str):
            """Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree."""

            for i in range(self.doc_traversal_depth):
                self.accepting_node_child_tree.write("|  ")

            self.accepting_node_child_tree.write(text + "\n")

    #ExEnd

    def _test_doc_structure_to_text(self, visitor: ExDocumentVisitor.DocStructurePrinter):

        visitor_text = visitor.get_text()

        self.assertIn("[Document start]", visitor_text)
        self.assertIn("[Document end]", visitor_text)
        self.assertIn("[Section start]", visitor_text)
        self.assertIn("[Section end]", visitor_text)
        self.assertIn("[Body start]", visitor_text)
        self.assertIn("[Body end]", visitor_text)
        self.assertIn("[Paragraph start]", visitor_text)
        self.assertIn("[Paragraph end]", visitor_text)
        self.assertIn("[Run]", visitor_text)
        self.assertIn("[SubDocument]", visitor_text)

    #ExStart
    #ExFor:Cell.accept(DocumentVisitor)
    #ExFor:Cell.is_first_cell
    #ExFor:Cell.is_last_cell
    #ExFor:DocumentVisitor.visit_table_end(Table)
    #ExFor:DocumentVisitor.visit_table_start(Table)
    #ExFor:DocumentVisitor.visit_row_end(Row)
    #ExFor:DocumentVisitor.visit_row_start(Row)
    #ExFor:DocumentVisitor.visit_cell_start(Cell)
    #ExFor:DocumentVisitor.visit_cell_end(Cell)
    #ExFor:Row.accept(DocumentVisitor)
    #ExFor:Row.first_cell
    #ExFor:Row.get_text
    #ExFor:Row.is_first_row
    #ExFor:Row.last_cell
    #ExFor:Row.parent_table
    #ExSummary:Shows how to print the node structure of every table in a document.
    def test_table_to_text(self):

        doc = aw.Document(MY_DIR + "DocumentVisitor-compatible features.docx")
        visitor = ExDocumentVisitor.TableStructurePrinter()

        # When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        # and then traverses all the node's children in a depth-first manner.
        # The visitor can read and modify each visited node.
        doc.accept(visitor)

        print(visitor.get_text())
        self._test_table_to_text(visitor) #ExSkip

    class TableStructurePrinter(aw.DocumentVisitor):
        """Traverses a node's non-binary tree of child nodes.
        Creates a map in the form of a string of all encountered Table nodes and their children."""

        def __init__(self):

            aw.DocumentVisitor.__init__(self)

            self.visited_tables = io.StringIO()
            self.visitor_is_inside_table = False
            self.doc_traversal_depth = 0

        def get_text(self):

            return self.visited_tables.getvalue()

        def visit_run(self, run: aw.Run) -> aw.VisitorAction:
            """Called when a Run node is encountered in the document.
            Runs that are not within tables are not recorded."""

            if self.visitor_is_inside_table:
                self._indent_and_append_line("[Run] \"" + run.get_text() + "\"")

            return aw.VisitorAction.CONTINUE

        def visit_table_start(self, table: aw.Table) -> aw.VisitorAction:
            """Called when a Table is encountered in the document."""

            rows = 0
            columns = 0

            if table.rows.count > 0:
                rows = table.rows.count
                columns = table.first_row.count

            self._indent_and_append_line("[Table start] Size: " + rows + "x" + columns)
            self.doc_traversal_depth += 1
            self.visitor_is_inside_table = True

            return aw.VisitorAction.CONTINUE

        def visit_table_end(self, table: aw.Table) -> aw.VisitorAction:
            """Called after all the child nodes of a Table node have been visited."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[Table end]")
            self.visitor_is_inside_table = False

            return aw.VisitorAction.CONTINUE

        def visit_row_start(self, row: aw.Row) -> aw.VisitorAction:
            """Called when a Row node is encountered in the document."""

            row_contents = row.get_text().rstrip("\u0007 ").replace("\u0007", ", ")
            row_width = row.index_of(row.last_cell) + 1
            row_index = row.parent_table.index_of(row)
            if row.is_first_row and row.is_last_row:
                row_status_in_table = "only"
            elif row.is_first_row:
                row_status_in_table = "first"
            elif row.is_last_row:
                row_status_in_table = "last"
            else:
                row_status_in_table = ""

            if row_status_in_table != "":
                row_status_in_table = f", the {row_status_in_table} row in this table,"

            row_index += 1
            self._indent_and_append_line(f"[Row start] Row #{row_index}{row_status_in_table} width {row_width}, \"{row_contents}\"")
            self.doc_traversal_depth += 1

            return aw.VisitorAction.CONTINUE

        def visit_row_end(self, row: aw.Row) -> aw.VisitorAction:
            """Called after all the child nodes of a Row node have been visited."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[Row end]")

            return aw.VisitorAction.CONTINUE

        def visit_cell_start(self, cell: aw.Cell) -> aw.VisitorAction:
            """Called when a Cell node is encountered in the document."""

            row = cell.parent_row
            table = row.parent_table
            if cell.is_first_cell and cell.is_last_cell:
                cell_status_in_row = "only"
            elif cell.is_first_cell:
                cell_status_in_row = "first"
            elif cell.is_last_cell:
                cell_status_in_row = "last"
            else:
                cell_status_in_row = ""

            if cell_status_in_row != "":
                cell_status_in_row = f", the {cell_status_in_row} cell in this row"

            self._indent_and_append_line(f"[Cell start] Row {table.index_of(row) + 1}, Col {row.index_of(cell) + 1}{cell_status_in_row}")
            self.doc_traversal_depth += 1

            return aw.VisitorAction.CONTINUE

        def visit_cell_end(self, cell: aw.Cell) -> aw.VisitorAction:
            """Called after all the child nodes of a Cell node have been visited."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[Cell end]")
            return aw.VisitorAction.CONTINUE

        def _indent_and_append_line(self, text: str):
            """Append a line to the output, and indent it depending on how deep the visitor is
            into the current table's tree of child nodes."""

            for i in range(self.doc_traversal_depth):
                self.visited_tables.write("|  ")

            self.visited_tables.write(text + "\n")

    #ExEnd

    def _test_table_to_text(self, visitor: ExDocumentVisitor.TableStructurePrinter):

        visitor_text = visitor.get_text()

        self.assertIn("[Table start]", visitor_text)
        self.assertIn("[Table end]", visitor_text)
        self.assertIn("[Row start]", visitor_text)
        self.assertIn("[Row end]", visitor_text)
        self.assertIn("[Cell start]", visitor_text)
        self.assertIn("[Cell end]", visitor_text)
        self.assertIn("[Run]", visitor_text)

    #ExStart
    #ExFor:DocumentVisitor.visit_comment_start(Comment)
    #ExFor:DocumentVisitor.visit_comment_end(Comment)
    #ExFor:DocumentVisitor.visit_comment_range_end(CommentRangeEnd)
    #ExFor:DocumentVisitor.visit_comment_range_start(CommentRangeStart)
    #ExSummary:Shows how to print the node structure of every comment and comment range in a document.
    def test_comments_to_text(self):

        doc = aw.Document(MY_DIR + "DocumentVisitor-compatible features.docx")
        visitor = ExDocumentVisitor.CommentStructurePrinter()

        # When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        # and then traverses all the node's children in a depth-first manner.
        # The visitor can read and modify each visited node.
        doc.accept(visitor)

        print(visitor.get_text())
        self._test_comments_to_text(visitor) #ExSkip

    class CommentStructurePrinter(aw.DocumentVisitor):
        """Traverses a node's non-binary tree of child nodes.
        Creates a map in the form of a string of all encountered Comment/CommentRange nodes and their children."""

        def __init__(self):

            aw.DocumentVisitor.__init__(self)

            self.builder = io.StringIO()
            self.visitor_is_inside_comment = False
            self.doc_traversal_depth = 0

        def get_text(self):

            return self.builder.getvalue()

        def visit_run(self, run: aw.Run) -> aw.VisitorAction:
            """Called when a Run node is encountered in the document.
            A Run is only recorded if it is a child of a Comment or CommentRange node."""

            if self.visitor_is_inside_comment:
                self._indent_and_append_line("[Run] \"" + run.get_text() + "\"")

            return aw.VisitorAction.CONTINUE

        def visit_comment_range_start(self, comment_range_start: aw.CommentRangeStart) -> aw.VisitorAction:
            """Called when a CommentRangeStart node is encountered in the document."""

            self._indent_and_append_line("[Comment range start] ID: " + comment_range_start.id)
            self.doc_traversal_depth += 1
            self.visitor_is_inside_comment = True

            return aw.VisitorAction.CONTINUE

        def visit_comment_range_end(self, comment_range_end: aw.CommentRangeEnd) -> aw.VisitorAction:
            """Called when a CommentRangeEnd node is encountered in the document."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[Comment range end]")
            self.visitor_is_inside_comment = False

            return aw.VisitorAction.CONTINUE

        def visit_comment_start(self, comment: aw.Comment) -> aw.VisitorAction:
            """Called when a Comment node is encountered in the document."""

            self._indent_and_append_line(
                f"[Comment start] For comment range ID {comment.Id}, By {comment.Author} on {comment.DateTime}")
            self.doc_traversal_depth += 1
            self.visitor_is_inside_comment = True

            return aw.VisitorAction.CONTINUE

        def visit_comment_end(self, comment: aw.Comment) -> aw.VisitorAction:
            """Called after all the child nodes of a Comment node have been visited."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[Comment end]")
            self.visitor_is_inside_comment = False

            return aw.VisitorAction.CONTINUE

        def _indent_and_append_line(self, text: str):
            """Append a line to the output, and indent it depending on how deep the visitor is
            into a comment/comment range's tree of child nodes."""

            for i in range(self.doc_traversal_depth):
                self.builder.write("|  ")

            self.builder.write(text + "\n")

    #ExEnd

    def _test_comments_to_text(self, visitor: ExDocumentVisitor.CommentStructurePrinter):

        visitor_text = visitor.get_text()

        self.assertIn("[Comment range start]", visitor_text)
        self.assertIn("[Comment range end]", visitor_text)
        self.assertIn("[Comment start]", visitor_text)
        self.assertIn("[Comment end]", visitor_text)
        self.assertIn("[Run]", visitor_text)

    #ExStart
    #ExFor:DocumentVisitor.visit_field_start
    #ExFor:DocumentVisitor.visit_field_end
    #ExFor:DocumentVisitor.visit_field_separator
    #ExSummary:Shows how to print the node structure of every field in a document.
    def test_field_to_text(self):

        doc = aw.Document(MY_DIR + "DocumentVisitor-compatible features.docx")
        visitor = ExDocumentVisitor.FieldStructurePrinter()

        # When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        # and then traverses all the node's children in a depth-first manner.
        # The visitor can read and modify each visited node.
        doc.accept(visitor)

        print(visitor.get_text())
        self._test_field_to_text(visitor) #ExSkip

    class FieldStructurePrinter(aw.DocumentVisitor):
        """Traverses a node's non-binary tree of child nodes.
        Creates a map in the form of a string of all encountered Field nodes and their children."""

        def __init__(self):

            aw.DocumentVisitor.__init__(self)

            self.builder = io.StringIO()
            self.visitor_is_inside_field = False
            self.doc_traversal_depth = 0

        def get_text(self):

            return self.builder.getvalue()

        def visit_run(self, run: aw.Run) -> aw.VisitorAction:
            """Called when a Run node is encountered in the document."""

            if self.visitor_is_inside_field:
                self._indent_and_append_line("[Run] \"" + run.get_text() + "\"")

            return aw.VisitorAction.CONTINUE

        def visit_field_start(self, field_start: aw.fields.FieldStart) -> aw.VisitorAction:
            """Called when a FieldStart node is encountered in the document."""

            self._indent_and_append_line("[Field start] FieldType: " + field_start.field_type)
            self.doc_traversal_depth += 1
            self.visitor_is_inside_field = True

            return aw.VisitorAction.CONTINUE

        def visit_field_end(self, field_end: aw.fields.FieldEnd) -> aw.VisitorAction:
            """Called when a FieldEnd node is encountered in the document."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[Field end]")
            self.visitor_is_inside_field = False

            return aw.VisitorAction.CONTINUE

        def visit_field_separator(self, field_separator: aw.fields.FieldSeparator) -> aw.VisitorAction:
            """Called when a FieldSeparator node is encountered in the document."""

            self._indent_and_append_line("[FieldSeparator]")

            return aw.VisitorAction.CONTINUE

        def _indent_and_append_line(self, text: str):
            """Append a line to the output, and indent it depending on how deep the visitor is
            into the field's tree of child nodes."""

            for i in range(self.doc_traversal_depth):
                self.builder.write("|  ")

            self.builder.write(text + "\n")

    #ExEnd

    def _test_field_to_text(self, visitor: aw.fields.FieldStructurePrinter):

        visitor_text = visitor.get_text()

        self.assertIn("[Field start]", visitor_text)
        self.assertIn("[Field end]", visitor_text)
        self.assertIn("[FieldSeparator]", visitor_text)
        self.assertIn("[Run]", visitor_text)

    #ExStart
    #ExFor:DocumentVisitor.visit_header_footer_start(HeaderFooter)
    #ExFor:DocumentVisitor.visit_header_footer_end(HeaderFooter)
    #ExFor:HeaderFooter.accept(DocumentVisitor)
    #ExFor:HeaderFooterCollection.to_array
    #ExFor:Run.accept(DocumentVisitor)
    #ExFor:Run.get_text
    #ExSummary:Shows how to print the node structure of every header and footer in a document.
    def test_header_footer_to_text(self):

        doc = aw.Document(MY_DIR + "DocumentVisitor-compatible features.docx")
        visitor = ExDocumentVisitor.HeaderFooterStructurePrinter()

        # When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        # and then traverses all the node's children in a depth-first manner.
        # The visitor can read and modify each visited node.
        doc.accept(visitor)

        print(visitor.get_text())

        # An alternative way of accessing a document's header/footers section-by-section is by accessing the collection.
        header_footers = doc.first_section.headers_footers.to_array()
        self.assertEqual(3, len(header_footers))
        self._test_header_footer_to_text(visitor) #ExSkip

    class HeaderFooterStructurePrinter(aw.DocumentVisitor):
        """Traverses a node's non-binary tree of child nodes.
        Creates a map in the form of a string of all encountered HeaderFooter nodes and their children."""

        def __init__(self):

            aw.DocumentVisitor.__init__(self)

            self.builder = io.StringIO()
            self.visitor_is_inside_header_footer = False
            self.doc_traversal_depth = 0

        def get_text(self):

            return self.builder.getvalue()

        def visit_run(self, run: aw.Run) -> aw.VisitorAction:
            """Called when a Run node is encountered in the document."""

            if self.visitor_is_inside_header_footer:
                self._indent_and_append_line("[Run] \"" + run.get_text() + "\"")

            return aw.VisitorAction.CONTINUE

        def visit_header_footer_start(self, header_footer: aw.HeaderFooter) -> aw.VisitorAction:
            """Called when a HeaderFooter node is encountered in the document."""

            self._indent_and_append_line("[HeaderFooter start] HeaderFooterType: " + header_footer.header_footer_type)
            self.doc_traversal_depth += 1
            self.visitor_is_inside_header_footer = True

            return aw.VisitorAction.CONTINUE

        def visit_header_footer_end(self, header_footer: aw.HeaderFooter) -> aw.VisitorAction:
            """Called after all the child nodes of a HeaderFooter node have been visited."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[HeaderFooter end]")
            self.visitor_is_inside_header_footer = False

            return aw.VisitorAction.CONTINUE

        def _indent_and_append_line(self, text: str):
            """Append a line to the output, and indent it depending on how deep the visitor is into the document tree."""

            for i in range(self.doc_traversal_depth):
                self.builder.write("|  ")

            self.builder.write(text + "\n")

    #ExEnd

    def _test_header_footer_to_text(self, visitor: aw.HeaderFooterStructurePrinter):

        visitor_text = visitor.get_text()

        self.assertIn("[HeaderFooter start] HeaderFooterType: HeaderPrimary", visitor_text)
        self.assertIn("[HeaderFooter end]", visitor_text)
        self.assertIn("[HeaderFooter start] HeaderFooterType: HeaderFirst", visitor_text)
        self.assertIn("[HeaderFooter start] HeaderFooterType: HeaderEven", visitor_text)
        self.assertIn("[HeaderFooter start] HeaderFooterType: FooterPrimary", visitor_text)
        self.assertIn("[HeaderFooter start] HeaderFooterType: FooterFirst", visitor_text)
        self.assertIn("[HeaderFooter start] HeaderFooterType: FooterEven", visitor_text)
        self.assertIn("[Run]", visitor_text)

    #ExStart
    #ExFor:DocumentVisitor.visit_editable_range_end(EditableRangeEnd)
    #ExFor:DocumentVisitor.visit_editable_range_start(EditableRangeStart)
    #ExSummary:Shows how to print the node structure of every editable range in a document.
    def test_editable_range_to_text(self):

        doc = aw.Document(MY_DIR + "DocumentVisitor-compatible features.docx")
        visitor = ExDocumentVisitor.EditableRangeStructurePrinter()

        # When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        # and then traverses all the node's children in a depth-first manner.
        # The visitor can read and modify each visited node.
        doc.accept(visitor)

        print(visitor.get_text())
        self._test_editable_range_to_text(visitor) #ExSkip

    class EditableRangeStructurePrinter(aw.DocumentVisitor):
        """Traverses a node's non-binary tree of child nodes.
        Creates a map in the form of a string of all encountered EditableRange nodes and their children."""

        def __init__(self):

            aw.DocumentVisitor.__init__(self)

            self.builder = io.StringIO()
            self.visitor_is_inside_editable_range = False
            self.doc_traversal_depth = 0

        def get_text(self) -> str:
            """Gets the plain text of the document that was accumulated by the visitor."""

            return self.builder.getvalue()

        def visit_run(self, run: aw.Run) -> aw.VisitorAction:
            """Called when a Run node is encountered in the document."""

            # We want to print the contents of runs, but only if they are inside shapes, as they would be in the case of text boxes
            if self.visitor_is_inside_editable_range:
                self._indent_and_append_line("[Run] \"" + run.get_text() + "\"")

            return aw.VisitorAction.CONTINUE

        def visit_editable_range_start(self, editable_range_start: aw.EditableRangeStart) -> aw.VisitorAction:
            """Called when an EditableRange node is encountered in the document."""

            self._indent_and_append_line("[EditableRange start] ID: " + editable_range_start.id + " Owner: " +
                                editable_range_start.editable_range.single_user)
            self.doc_traversal_depth += 1
            self.visitor_is_inside_editable_range = True

            return aw.VisitorAction.CONTINUE

        def visit_editable_range_end(self, editable_range_end: aw.EditableRangeEnd) -> aw.VisitorAction:
            """Called when the visiting of a EditableRange node is ended."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[EditableRange end]")
            self.visitor_is_inside_editable_range = False

            return aw.VisitorAction.CONTINUE

        def _indent_and_append_line(self, text: str):
            """Append a line to the output and indent it depending on how deep the visitor is into the document tree."""

            for i in range(self.doc_traversal_depth):
                self.builder.write("|  ")

            self.builder.write(text + "\n")

    #ExEnd

    def _test_editable_range_to_text(self, visitor: ExDocumentVisitor.EditableRangeStructurePrinter):

        visitor_text = visitor.get_text()

        self.assertIn("[EditableRange start]", visitor_text)
        self.assertIn("[EditableRange end]", visitor_text)
        self.assertIn("[Run]", visitor_text)

    #ExStart
    #ExFor:DocumentVisitor.visit_footnote_end(Footnote)
    #ExFor:DocumentVisitor.visit_footnote_start(Footnote)
    #ExFor:Footnote.accept(DocumentVisitor)
    #ExSummary:Shows how to print the node structure of every footnote in a document.
    def test_footnote_to_text(self):

        doc = aw.Document(MY_DIR + "DocumentVisitor-compatible features.docx")
        visitor = ExDocumentVisitor.FootnoteStructurePrinter()

        # When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        # and then traverses all the node's children in a depth-first manner.
        # The visitor can read and modify each visited node.
        doc.accept(visitor)

        print(visitor.get_text())
        self._test_footnote_to_text(visitor) #ExSkip

    class FootnoteStructurePrinter(aw.DocumentVisitor):
        """Traverses a node's non-binary tree of child nodes.
        Creates a map in the form of a string of all encountered Footnote nodes and their children."""

        def __init__(self):

            aw.DocumentVisitor.__init__(self)

            self.builder = io.StringIO()
            self.visitor_is_inside_footnote = False
            self.doc_traversal_depth = 0

        def get_text(self) -> str:
            """Gets the plain text of the document that was accumulated by the visitor."""

            return self.builder.getvalue()

        def visit_footnote_start(self, footnote: aw.Footnote) -> aw.VisitorAction:
            """Called when a Footnote node is encountered in the document."""

            self._indent_and_append_line("[Footnote start] Type: " + footnote.footnote_type)
            self.doc_traversal_depth += 1
            self.visitor_is_inside_footnote = True

            return aw.VisitorAction.CONTINUE

        def visit_footnote_end(self, footnote: aw.Footnote) -> aw.VisitorAction:
            """Called after all the child nodes of a Footnote node have been visited."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[Footnote end]")
            self.visitor_is_inside_footnote = False

            return aw.VisitorAction.CONTINUE

        def visit_run(self, run: aw.Run) -> aw.VisitorAction:
            """Called when a Run node is encountered in the document."""

            if self.visitor_is_inside_footnote:
                self._indent_and_append_line("[Run] \"" + run.get_text() + "\"")

            return aw.VisitorAction.CONTINUE

        def _indent_and_append_line(self, text: str):
            """Append a line to the output and indent it depending on how deep the visitor is into the document tree."""

            for i in range(self.doc_traversal_depth):
                self.builder.write("|  ")

            self.builder.write(text + "\n")

    #ExEnd

    def _test_footnote_to_text(self, visitor: aw.FootnoteStructurePrinter):

        visitor_text = visitor.get_text()

        self.assertIn("[Footnote start] Type: Footnote", visitor_text)
        self.assertIn("[Footnote end]", visitor_text)
        self.assertIn("[Run]", visitor_text)

    #ExStart
    #ExFor:DocumentVisitor.visit_office_math_end(OfficeMath)
    #ExFor:DocumentVisitor.visit_office_math_start(OfficeMath)
    #ExFor:MathObjectType
    #ExFor:OfficeMath.accept(DocumentVisitor)
    #ExFor:OfficeMath.math_object_type
    #ExSummary:Shows how to print the node structure of every office math node in a document.
    def test_office_math_to_text(self):

        doc = aw.Document(MY_DIR + "DocumentVisitor-compatible features.docx")
        visitor = ExDocumentVisitor.OfficeMathStructurePrinter()

        # When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        # and then traverses all the node's children in a depth-first manner.
        # The visitor can read and modify each visited node.
        doc.accept(visitor)

        print(visitor.get_text())
        self._test_office_math_to_text(visitor) #ExSkip

    class OfficeMathStructurePrinter(aw.DocumentVisitor):
        """Traverses a node's non-binary tree of child nodes.
        Creates a map in the form of a string of all encountered OfficeMath nodes and their children."""

        def __init__(self):

            aw.DocumentVisitor.__init__(self)

            self.builder = io.StringIO()
            self.visitor_is_inside_office_math = False
            self.doc_traversal_depth = 0

        def get_text(self) -> str:
            """Gets the plain text of the document that was accumulated by the visitor."""

            return self.builder.getvalue()

        def visit_run(self, run: aw.Run) -> aw.VisitorAction:
            """Called when a Run node is encountered in the document."""

            if self.visitor_is_inside_office_math:
                self._indent_and_append_line("[Run] \"" + run.get_text() + "\"")

            return aw.VisitorAction.CONTINUE

        def visit_office_math_start(self, office_math: aw.OfficeMath) -> aw.VisitorAction:
            """Called when an OfficeMath node is encountered in the document."""

            self._indent_and_append_line("[OfficeMath start] Math object type: " + office_math.math_object_type)
            self.doc_traversal_depth += 1
            self.visitor_is_inside_office_math = True

            return aw.VisitorAction.CONTINUE

        def visit_office_math_end(self, office_math: aw.OfficeMath) -> aw.VisitorAction:
            """Called after all the child nodes of an OfficeMath node have been visited."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[OfficeMath end]")
            self.visitor_is_inside_office_math = False

            return aw.VisitorAction.CONTINUE

        def _indent_and_append_line(self, text: str):
            """Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree."""

            for i in range(self.doc_traversal_depth):
                self.builder.write("|  ")

            self.builder.write(text + "\n")

    #ExEnd

    def _test_office_math_to_text(self, visitor: aw.OfficeMathStructurePrinter):

        visitor_text = visitor.get_text()

        self.assertIn("[OfficeMath start] Math object type: OMathPara", visitor_text)
        self.assertIn("[OfficeMath start] Math object type: OMath", visitor_text)
        self.assertIn("[OfficeMath start] Math object type: Argument", visitor_text)
        self.assertIn("[OfficeMath start] Math object type: Supercript", visitor_text)
        self.assertIn("[OfficeMath start] Math object type: SuperscriptPart", visitor_text)
        self.assertIn("[OfficeMath start] Math object type: Fraction", visitor_text)
        self.assertIn("[OfficeMath start] Math object type: Numerator", visitor_text)
        self.assertIn("[OfficeMath start] Math object type: Denominator", visitor_text)
        self.assertIn("[OfficeMath end]", visitor_text)
        self.assertIn("[Run]", visitor_text)

    #ExStart
    #ExFor:DocumentVisitor.visit_smart_tag_end(SmartTag)
    #ExFor:DocumentVisitor.visit_smart_tag_start(SmartTag)
    #ExSummary:Shows how to print the node structure of every smart tag in a document.
    def test_smart_tag_to_text(self):

        doc = aw.Document(MY_DIR + "Smart tags.doc")
        visitor = ExDocumentVisitor.SmartTagStructurePrinter()

        # When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        # and then traverses all the node's children in a depth-first manner.
        # The visitor can read and modify each visited node.
        doc.accept(visitor)

        print(visitor.get_text())
        self._test_smart_tag_to_text(visitor) #ExSkip

    class SmartTagStructurePrinter(aw.DocumentVisitor):
        """Traverses a node's non-binary tree of child nodes.
        Creates a map in the form of a string of all encountered SmartTag nodes and their children."""

        def __init__(self):

            aw.DocumentVisitor.__init__(self)

            self.builder = io.StringIO()
            self.visitor_is_inside_smart_tag = False
            self.doc_traversal_depth = 0

        def get_text(self) -> str:
            """Gets the plain text of the document that was accumulated by the visitor."""

            return self.builder.getvalue()

        def visit_run(self, run: aw.Run) -> aw.VisitorAction:
            """Called when a Run node is encountered in the document."""

            if self.visitor_is_inside_smart_tag:
                self._indent_and_append_line("[Run] \"" + run.get_text() + "\"")

            return aw.VisitorAction.CONTINUE

        def visit_smart_tag_start(self, smart_tag: aw.SmartTag) -> aw.VisitorAction:
            """Called when a SmartTag node is encountered in the document."""

            self._indent_and_append_line("[SmartTag start] Name: " + smart_tag.element)
            self.doc_traversal_depth += 1
            self.visitor_is_inside_smart_tag = True

            return aw.VisitorAction.CONTINUE

        def visit_smart_tag_end(self, smart_tag: aw.SmartTag) -> aw.VisitorAction:
            """Called after all the child nodes of a SmartTag node have been visited."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[SmartTag end]")
            self.visitor_is_inside_smart_tag = False

            return aw.VisitorAction.CONTINUE

        def _indent_and_append_line(self, text: str):
            """Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree."""

            for i in range(self.doc_traversal_depth):
                self.builder.write("|  ")

            self.builder.write(text + "\n")

    #ExEnd

    def _test_smart_tag_to_text(self, visitor: aw.SmartTagStructurePrinter):

        visitor_text = visitor.get_text()

        self.assertIn("[SmartTag start] Name: address", visitor_text)
        self.assertIn("[SmartTag start] Name: Street", visitor_text)
        self.assertIn("[SmartTag start] Name: PersonName", visitor_text)
        self.assertIn("[SmartTag start] Name: title", visitor_text)
        self.assertIn("[SmartTag start] Name: GivenName", visitor_text)
        self.assertIn("[SmartTag start] Name: Sn", visitor_text)
        self.assertIn("[SmartTag start] Name: stockticker", visitor_text)
        self.assertIn("[SmartTag start] Name: date", visitor_text)
        self.assertIn("[SmartTag end]", visitor_text)
        self.assertIn("[Run]", visitor_text)

    #ExStart
    #ExFor:StructuredDocumentTag.accept(DocumentVisitor)
    #ExFor:DocumentVisitor.visit_structured_document_tag_end(StructuredDocumentTag)
    #ExFor:DocumentVisitor.visit_structured_document_tag_start(StructuredDocumentTag)
    #ExSummary:Shows how to print the node structure of every structured document tag in a document.
    def test_structured_document_tag_to_text(self):

        doc = aw.Document(MY_DIR + "DocumentVisitor-compatible features.docx")
        visitor = ExDocumentVisitor.StructuredDocumentTagNodePrinter()

        # When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        # and then traverses all the node's children in a depth-first manner.
        # The visitor can read and modify each visited node.
        doc.accept(visitor)

        print(visitor.get_text())
        self._test_structured_document_tag_to_text(visitor) #ExSkip

    class StructuredDocumentTagNodePrinter(aw.DocumentVisitor):
        """Traverses a node's non-binary tree of child nodes.
        Creates a map in the form of a string of all encountered StructuredDocumentTag nodes and their children."""

        def __init__(self):

            aw.DocumentVisitor.__init__(self)

            self.builder = io.StringIO()
            self.visitor_is_inside_structured_document_tag = False
            self.doc_traversal_depth = 0

        def get_text(self) -> str:
            """Gets the plain text of the document that was accumulated by the visitor."""

            return self.builder.getvalue()

        def visit_run(self, run: aw.Run) -> aw.VisitorAction:
            """Called when a Run node is encountered in the document."""

            if self.visitor_is_inside_structured_document_tag:
                self._indent_and_append_line("[Run] \"" + run.get_text() + "\"")

            return aw.VisitorAction.CONTINUE

        def visit_structured_document_tag_start(self, sdt: aw.StructuredDocumentTag) -> aw.VisitorAction:
            """Called when a StructuredDocumentTag node is encountered in the document."""

            self._indent_and_append_line("[StructuredDocumentTag start] Title: " + sdt.title)
            self.doc_traversal_depth += 1

            return aw.VisitorAction.CONTINUE

        def visit_structured_document_tag_end(self, sdt: aw.StructuredDocumentTag) -> aw.VisitorAction:
            """Called after all the child nodes of a StructuredDocumentTag node have been visited."""

            self.doc_traversal_depth -= 1
            self._indent_and_append_line("[StructuredDocumentTag end]")

            return aw.VisitorAction.CONTINUE

        def _indent_and_append_line(self, text: str):
            """Append a line to the output and indent it depending on how deep the visitor is into the document tree."""

            for i in range(self.doc_traversal_depth):
                self.builder.write("|  ")

            self.builder.write(text + "\n")

    #ExEnd

    def _test_structured_document_tag_to_text(self, visitor: aw.StructuredDocumentTagNodePrinter):

        visitor_text = visitor.get_text()

        self.assertIn("[StructuredDocumentTag start]", visitor_text)
        self.assertIn("[StructuredDocumentTag end]", visitor_text)
