# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw
from api_example_base import ApiExampleBase, TEMP_DIR

class DocumentHelper(ApiExampleBase):

    def create_document_without_dummy_text() -> aw.Document:
        """Create simple document without run in the paragraph"""
        doc = aw.Document()

        # Remove the previous changes of the document
        doc.remove_all_children()

        # Set the document author
        doc.built_in_document_properties.author = "Test Author"

        # Create paragraph without run
        builder = aw.DocumentBuilder(doc)
        builder.writeln()

        return doc


    @staticmethod
    def find_text_in_file(path: str, expression: str):
        with open(path, "rt", encoding="utf-8") as file:
            if expression in file.read():
                print("true")

    @staticmethod
    def create_simple_document(template_text: str) -> aw.Document:
        """Create new document template for reporting engine"""
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write(template_text)

        return doc

    @staticmethod
    def create_document_fill_with_dummy_text():
        """Create new document with text."""
        doc = aw.Document()

        # Remove the previous changes of the document
        doc.remove_all_children()

        # Set the document author
        doc.built_in_document_properties.author = "Test Author"

        builder = aw.DocumentBuilder(doc)

        builder.write("Page ")
        builder.insert_field("PAGE", "")
        builder.write(" of ")
        builder.insert_field("NUMPAGES", "")

        # Insert new table with two rows and two cells
        DocumentHelper.insert_table(builder)

        builder.writeln("Hello World!")

        #  Continued on page 2 of the document content
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # Insert TOC entries
        DocumentHelper.insert_toc(builder)

        return doc

    @staticmethod
    def create_template_document_with_draw_objects(template_text: str, shape_type: aw.drawing.ShapeType) -> aw.Document:
        """Create new document with textbox shape and some query."""
        doc = aw.Document()

        # Create textbox shape.
        shape = aw.drawing.Shape(doc, shape_type)
        shape.width = 431.5
        shape.height = 346.35

        paragraph = aw.Paragraph(doc)
        paragraph.append_child(aw.Run(doc, template_text))

        # Insert paragraph into the textbox.
        shape.append_child(paragraph)

        # Insert textbox into the document.
        doc.first_section.body.first_paragraph.append_child(shape)

        return doc

    @staticmethod
    def compare_docs(file_path_doc_1: str, file_path_doc_2: str):
        """Compare word documents.

            file_path_doc_1 -- First document path

            file_path_doc_2 -- Second document path

            Returns result of compare document
        """
        doc1 = aw.Document(file_path_doc_1)
        doc2 = aw.Document(file_path_doc_2)

        return doc1.get_text() == doc2.get_text()

    def insert_new_run(doc: aw.Document, text: str, para_index: int) -> aw.Run:
        """Insert run into the current document.
        
        :param doc: Current document.
        :param text: Custom text.
        :param para_index: Paragraph index."""
        para = DocumentHelper.get_paragraph(doc, para_index)

        run = aw.Run(doc)
        run.text = text

        para.append_child(run)

        return run

    @staticmethod
    def insert_builder_text(builder: aw.DocumentBuilder, text_strings):
        """Insert text into the current document.
        
        :param builder: Current document builder.
        :param text_strings: Custom text."""
        for text_string in text_strings :
            builder.writeln(text_string)

    @staticmethod
    def get_paragraph_text(doc: aw.Document, para_index: int) -> str:
        """Get paragraph text of the current document.
        
        :param doc: Current document.
        :param para_index: Paragraph number from collection."""
        return doc.first_section.body.paragraphs[para_index].get_text()

    @staticmethod
    def insert_table(builder: aw.DocumentBuilder):
        """Insert new table in the document.
        
        :param builder: Current document builder."""
        # Start creating a new table
        table = builder.start_table()

        # Insert Row 1 Cell 1
        builder.insert_cell()
        builder.write("Date")

        # Set width to fit the table contents
        table.auto_fit(aw.tables.AutoFitBehavior.AUTO_FIT_TO_CONTENTS)

        # Insert Row 1 Cell 2
        builder.insert_cell()
        builder.write(" ")

        builder.end_row()

        # Insert Row 2 Cell 1
        builder.insert_cell()
        builder.write("Author")

        # Insert Row 2 Cell 2
        builder.insert_cell()
        builder.write(" ")

        builder.end_row()

        builder.end_table()

        return table

    @staticmethod
    def insert_toc(builder: aw.DocumentBuilder):
        """Insert TOC entries in the document.
        
        :param builder: The builder."""
        # Creating TOC entries
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1

        builder.writeln("Heading 1")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING2

        builder.writeln("Heading 1.1")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING4

        builder.writeln("Heading 1.1.1.1")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING5

        builder.writeln("Heading 1.1.1.1.1")

        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING9

        builder.writeln("Heading 1.1.1.1.1.1.1.1.1")

    def get_section_text(doc: aw.Document, sec_index: int) -> str:
        """Get section text of the current document.
        
        :param doc: Current document.
        :param sec_index: Section number from collection."""
        return doc.sections[sec_index].get_text()

    @staticmethod
    def get_paragraph(doc: aw.Document, para_index: int) -> aw.Paragraph:
        """Get paragraph of the current document.
        
        :param doc: Current document.
        :param para_index: Paragraph number from collection."""
        return doc.first_section.body.paragraphs[para_index]

    # <summary>
    # Save the document to a file, immediately re-open it and return the newly opened version
    # </summary>
    # <remarks>
    # Used for testing how document features are preserved after saving/loading
    # </remarks>
    # <param name="doc">The document we wish to re-open</param>
    @staticmethod
    def save_open(doc):
        tmp_file_name = TEMP_DIR + "tmp.docx"
        doc.save(tmp_file_name, aw.SaveFormat.DOCX)
        return aw.Document(tmp_file_name)

    # Rude workaround to get style by name.
    @staticmethod
    def get_style_by_name(doc, style_name):
        for s in doc.styles:
            if s.name == style_name:
                return s
        return None

    # Rude workaround to get style by name.
    @staticmethod
    def get_vba_module_by_name(doc, module_name):
        for m in doc.vba_project.modules:
            if m.name == module_name:
                return m
        return None
