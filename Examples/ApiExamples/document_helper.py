import api_example_base as aeb
import aspose.words as aw


class DocumentHelper(aeb.ApiExampleBase):

    # <summary>
    # Create simple document without run in the paragraph
    # </summary>
    #    def CreateDocumentWithoutDummyText() -> aw.Document :
    #        doc = aw.Document()

    # Remove the previous changes of the document
    #        doc.remove_all_children()

    # Set the document author
    #        doc.built_in_document_properties.author = "Test Author"

    # Create paragraph without run
    #        builder = aw.DocumentBuilder(doc)
    #        builder.writeln()

    #        return doc


    #    def FindTextInFile(path : str, expression : str) :
    #        with open(str) as f:
    #            if expression in f.read():
    #                print("true")

    # <summary>
    # Create new document template for reporting engine
    # </summary>
    #    def CreateSimpleDocument(string templateText) -> Document :
    #        doc = aw.Document()
    #        builder = aw.DocumentBuilder(doc)

    #        builder.write(templateText)

    #        return doc

    # <summary>
    # Create new document with text
    # </summary>
    def create_document_fill_with_dummy_text(self):
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
        self.insert_table(builder)

        builder.writeln("Hello World!")

        #  Continued on page 2 of the document content
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # Insert TOC entries
        self.insert_toc(builder)

        return doc

    # <summary>
    # Create new document with textbox shape and some query
    # </summary>
    #    def CreateTemplateDocumentWithDrawObjects(templateText : str, shapeType : awd.ShapeType) -> aw.Document :
    #        doc = aw.Document()

    #  Create textbox shape.
    #        shape = awd.Shape(doc, shapeType)
    #        shape.width = 431.5
    #        shape.height = 346.35

    #        paragraph = aw.Paragraph(doc)
    #        paragraph.append_child(aw.Run(doc, templateText))

    #  Insert paragraph into the textbox.
    #        shape.append_child(paragraph)

    #  Insert textbox into the document.
    #        doc.first_section.body.first_paragraph.append_child(shape)

    #        return doc

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

    # <summary>
    # Insert run into the current document
    # </summary>
    # <param name="doc">Current document</param>
    # <param name="text">Custom text</param>
    # <param name="paraIndex">Paragraph index</param>
    def insert_new_run(self, doc, text, para_index):
        para = self.get_paragraph(doc, para_index)

        run = aw.Run(doc)
        run.text = text

        para.append_child(run)

        return run

    # <summary>
    # Insert text into the current document
    # </summary>
    # <param name="builder">Current document builder</param>
    # <param name="textStrings">Custom text</param>
    #    def InsertBuilderText(builder, textStrings) :
    #        for  textString in textStrings :
    #            builder.writeln(textString)

    # <summary>
    # Get paragraph text of the current document
    # </summary>
    # <param name="doc">Current document</param>
    # <param name="paraIndex">Paragraph number from collection</param>
    @staticmethod
    def get_paragraph_text(doc, para_index):
        return doc.first_section.body.paragraphs[para_index].get_text()

    # <summary>
    # Insert new table in the document
    # </summary>
    # <param name="builder">Current document builder</param>
    @staticmethod
    def insert_table(builder):
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

    # <summary>
    # Insert TOC entries in the document
    # </summary>
    # <param name="builder">
    # The builder.
    # </param>
    @staticmethod
    def insert_toc(builder):
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

    # <summary>
    # Get section text of the current document
    # </summary>
    # <param name="doc">Current document</param>
    # <param name="secIndex">Section number from collection</param>
    @staticmethod
    def get_section_text(doc, sec_index):
        return doc.sections[sec_index].get_text()

    # <summary>
    # Get paragraph of the current document
    # </summary>
    # <param name="doc">Current document</param>
    # <param name="paraIndex">Paragraph number from collection</param>
    @staticmethod
    def get_paragraph(doc, para_index):
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
        tmp_file_name = aeb.temp_dir + "tmp.docx"
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
