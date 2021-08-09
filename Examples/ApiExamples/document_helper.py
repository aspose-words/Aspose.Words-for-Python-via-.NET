import ApiExampleBase as aeb
import aspose.words as aw
import aspose.words.drawing as awd

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


    # <summary>
    # Create new document with text
    # </summary>
#    def  CreateDocumentFillWithDummyText() -> Document :
#        doc = aw.Document()

        # Remove the previous changes of the document
#        doc.remove_all_children()

        # Set the document author
#        doc.built_in_document_properties.author = "Test Author"

#        builder = aw.DocumentBuilder(doc)

#        builder.write("Page ")
#        builder.InsertField("PAGE", "")
#        builder.write(" of ")
#        builder.InsertField("NUMPAGES", "")

        # Insert new table with two rows and two cells
#        InsertTable(builder)

#        builder.writeln("Hello World!")

        #  Continued on page 2 of the document content
#        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # Insert TOC entries
#        InsertToc(builder)

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


    # <summary>
    # Compare word documents
    # </summary>
    # <param name="filePathDoc1">First document path</param>
    # <param name="filePathDoc2">Second document path</param>
    # <returns>Result of compare document</returns>
    def compare_docs(filePathDoc1, filePathDoc2) :
        doc1 = aw.Document(filePathDoc1)
        doc2 = aw.Document(filePathDoc2)

        return (doc1.get_text() == doc2.get_text())


    # <summary>
    # Insert run into the current document
    # </summary>
    # <param name="doc">Current document</param>
    # <param name="text">Custom text</param>
    # <param name="paraIndex">Paragraph index</param>
#    def InsertNewRun(doc : aw.Document, text : str, paraIndex : int) -> aw.Run :
#        para = GetParagraph(doc, paraIndex)

#        run = aw.Run(doc)
#        run.text = text

#        para.append_child(run)

#        return run


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
#    def GetParagraphText(doc, paraIndex) -> str :
#        return doc.first_section.body.paragraphs[paraIndex].get_text()


    # <summary>
    # Insert new table in the document
    # </summary>
    # <param name="builder">Current document builder</param>
#    def InsertTable(builder) -> aw.Table :
        # Start creating a new table
#        table = builder.start_table()

        # Insert Row 1 Cell 1
#        builder.insert_cell()
#        builder.write("Date")

        # Set width to fit the table contents
#        table.auto_fit(aw.AutoFitBehavior.AUTO_FIT_TO_CONTENTS)

        # Insert Row 1 Cell 2
#        builder.insert_cell()
#        builder.write(" ")

#        builder.end_row()

        # Insert Row 2 Cell 1
#        builder.insert_cell()
#        builder.write("Author")

        # Insert Row 2 Cell 2
#        builder.insert_cell()
#        builder.write(" ")

#        builder.end_row()

#        builder.end_table()

#        return table


    # <summary>
    # Insert TOC entries in the document
    # </summary>
    # <param name="builder">
    # The builder.
    # </param>
#    def InsertToc(DocumentBuilder builder)
        #  Creating TOC entries
#        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1

#        builder.writeln("Heading 1")

#        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING2

#        builder.writeln("Heading 1.1")

#        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING4

#        builder.writeln("Heading 1.1.1.1")

#        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING5

#        builder.writeln("Heading 1.1.1.1.1")

#        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING9

#        builder.writeln("Heading 1.1.1.1.1.1.1.1.1")


    # <summary>
    # Get section text of the current document
    # </summary>
    # <param name="doc">Current document</param>
    # <param name="secIndex">Section number from collection</param>
#    def GetSectionText(doc, secIndex) -> str :
#        return doc.sections[secIndex].get_text()


    # <summary>
    # Get paragraph of the current document
    # </summary>
    # <param name="doc">Current document</param>
    # <param name="paraIndex">Paragraph number from collection</param>
#    def GetParagraph(doc, paraIndex) -> aw.Paragraph :
#        return doc.first_section.body.paragraphs[paraIndex]

    # <summary>
    # Save the document to a file, immediately re-open it and return the newly opened version
    # </summary>
    # <remarks>
    # Used for testing how document features are preserved after saving/loading
    # </remarks>
    # <param name="doc">The document we wish to re-open</param>
    def save_open(doc):
        tmpFileName = aeb.TempDir + "tmp.docx"
        doc.save(tmpFileName, aw.SaveFormat.DOCX)
        return aw.Document(tmpFileName)

    # Rude workaround to get style by name.
    def get_style_by_name(doc, styleName) :
        for s in doc.styles :
            if s.name == styleName : 
                return s
        return None

    # Rude workaround to get style by name.
    def get_vba_module_by_name(doc, moduleName) :
        for m in doc.vba_project.modules :
            if m.name == moduleName : 
                return m
        return None     
